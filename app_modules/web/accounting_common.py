from .runs_formats import *
from ..core.foundation import _DEC_ZERO, _DEC_TWO

def normalize_family_code(value):
    text_value = str(value or '').strip()
    if text_value.endswith('.0'):
        text_value = text_value[:-2]
    return text_value


@lru_cache(maxsize=1)
def load_family_catalog_names():
    names = {}
    if not os.path.exists(FAMILY_CATALOG_PATH):
        return names
    try:
        df = pd.read_excel(FAMILY_CATALOG_PATH, dtype=str)
        cols = normalize_columns(df.columns)
        code_col = cols.get('C_FAM') or cols.get('CODIGO') or cols.get('COD_FAM') or cols.get('FAMILIA')
        name_col = cols.get('NOM_FAM') or cols.get('NOMBRE') or cols.get('DESCRIPCION') or cols.get('NOM')
        if not code_col or not name_col:
            return names
        for _, row in df.iterrows():
            code = normalize_family_code(row.get(code_col))
            name = str(row.get(name_col) or '').strip()
            if code and name and code.lower() != 'nan':
                names[code] = name
    except Exception:
        return names
    return names


def get_accounting_template_path():
    for path in ACCOUNTING_TEMPLATE_CANDIDATES:
        if os.path.exists(path):
            return path
    return ACCOUNTING_TEMPLATE_CANDIDATES[0]


def normalize_month_year(month_raw, year_raw):
    now = now_local_dt()
    try:
        month = int(str(month_raw or '').strip())
    except Exception:
        month = now.month
    try:
        year = int(str(year_raw or '').strip())
    except Exception:
        year = now.year
    month = min(12, max(1, month))
    year = min(2100, max(2000, year))
    return month, year


def sanitize_filename(text):
    raw = unicodedata.normalize('NFD', str(text or ''))
    raw = ''.join(ch for ch in raw if unicodedata.category(ch) != 'Mn')
    allowed = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ .'
    clean = ''.join(ch if ch in allowed else '_' for ch in raw).strip()
    return clean or 'reporte'


def accounting_base_folder(year, month):
    return os.path.join(REPORTS_DIR, 'accounting_monthly', 'bases', str(year), f'{month:02d}')


def read_accounting_base_dataframe(file_storage):
    file_name = str(getattr(file_storage, 'filename', '') or '').lower()
    if file_name.endswith('.csv'):
        return pd.read_csv(file_storage)
    return pd.read_excel(file_storage)


def extract_accounting_base_rows(df):
    cols = normalize_columns(df.columns)
    if 'C_ACT' not in cols:
        raise ValueError('El archivo debe contener la columna C_ACT')
    ordered_cols = list(df.columns)
    rows = []

    for _, row in df.iterrows():
        c_act_val = get_cell(row, cols, 'C_ACT')
        c_act_raw = str(c_act_val).strip() if c_act_val is not None else ''
        c_act = normalize_override_asset_code(c_act_raw)
        if not c_act:
            continue
        raw_payload = {}
        for col_name in ordered_cols:
            key = str(col_name).strip().upper()
            raw_payload[key] = None if pd.isna(row[col_name]) else serialize_raw_value_for_json(row[col_name])
        fam_code = normalize_family_code(get_cell(row, cols, 'C_FAM'))
        fam_name = str(get_cell(row, cols, 'NOM_FAM') or '').strip()
        rows.append({
            'c_act': c_act,
            'c_fam': fam_code,
            'nom_fam': fam_name,
            'costo': try_float(get_cell(row, cols, 'COSTO')),
            'saldo': try_float(get_cell(row, cols, 'SALDO')),
            'raw_row_json': json.dumps(raw_payload, ensure_ascii=False, default=str),
        })
    return rows


def serialize_raw_value_for_json(v):
    if isinstance(v, pd.Timestamp):
        return v.isoformat()
    if hasattr(v, 'item'):
        try:
            return v.item()
        except Exception:
            return v
    return v


def get_latest_accounting_base(month, year):
    return AccountingMonthlyBase.query.filter_by(
        period_month=month,
        period_year=year,
        status='active',
    ).order_by(AccountingMonthlyBase.id.desc()).first()


def build_overrides_summary_for_rows(asset_rows, selected_overrides):
    assets_by_code = {}
    for _, c_act, c_fam, nom_fam, costo, saldo in asset_rows:
        key = normalize_override_asset_code(c_act)
        assets_by_code.setdefault(key, []).append({
            'c_act': key,
            'c_fam': normalize_family_code(c_fam),
            'nom_fam': str(nom_fam or '').strip(),
            'costo': to_decimal_amount(costo, default=Decimal('0')).quantize(_DEC_TWO, rounding=ROUND_HALF_UP),
        })
    rows = []
    total_delta = _DEC_ZERO
    for code in sorted(selected_overrides.keys()):
        override_cost = selected_overrides.get(code, _DEC_ZERO).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
        matches = assets_by_code.get(code, [])
        base_cost = matches[0]['costo'] if matches else None
        delta = (override_cost - base_cost) if base_cost is not None else None
        if delta is not None:
            delta = delta.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
            total_delta += delta
        rows.append({
            'c_act': code,
            'present': bool(matches),
            'hits': len(matches),
            'c_fam': (matches[0]['c_fam'] if matches else ''),
            'nom_fam': (matches[0]['nom_fam'] if matches else ''),
            'base_cost': (float(base_cost) if base_cost is not None else None),
            'override_cost': float(override_cost),
            'delta': (float(delta) if delta is not None else None),
        })
    summary = {
        'rows': rows,
        'totals': {
            'configured': len(selected_overrides),
            'present': sum(1 for r in rows if r['present']),
            'missing': sum(1 for r in rows if not r['present']),
            'duplicates': sum(1 for r in rows if r['hits'] > 1),
            'delta_total': float(total_delta.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)),
        },
    }
    return summary


def persist_accounting_report_file(content, file_name, period_label, month, year, report_title, period_id=None, accounting_base_id=None, overrides_summary=None):
    reports_folder = os.path.join(REPORTS_DIR, 'accounting_monthly', str(year), f'{month:02d}')
    os.makedirs(reports_folder, exist_ok=True)

    stamped_name = f"{os.path.splitext(file_name)[0]}_{now_local_dt().strftime('%Y%m%d%H%M%S')}.xlsx"
    file_path = os.path.join(reports_folder, sanitize_filename(stamped_name))
    with open(file_path, 'wb') as f:
        f.write(content)

    report_row = GeneratedReport(
        report_type='accounting_monthly',
        title=report_title or 'Informe de conciliacion activos fijos - contabilidad',
        period_id=period_id,
        period_label=period_label,
        accounting_base_id=accounting_base_id,
        overrides_summary_json=(json.dumps(overrides_summary, ensure_ascii=False) if overrides_summary else None),
        file_name=os.path.basename(file_path),
        file_path=file_path,
        generated_at=now_iso(),
    )
    db.session.add(report_row)
    db.session.commit()


def persist_generated_report_file(content, report_type, title, period_label, file_name, folder_group, year=None, month=None, period_id=None):
    yy = str(year or now_local_dt().year)
    mm = f"{int(month):02d}" if month else '00'
    reports_folder = os.path.join(REPORTS_DIR, folder_group, yy, mm)
    os.makedirs(reports_folder, exist_ok=True)

    stamped_name = f"{os.path.splitext(file_name)[0]}_{now_local_dt().strftime('%Y%m%d%H%M%S')}{os.path.splitext(file_name)[1] or ''}"
    stamped_name = sanitize_filename(stamped_name)
    file_path = os.path.join(reports_folder, stamped_name)
    with open(file_path, 'wb') as f:
        f.write(content)

    row = GeneratedReport(
        report_type=report_type,
        title=title,
        period_id=period_id,
        period_label=period_label,
        file_name=stamped_name,
        file_path=file_path,
        generated_at=now_iso(),
    )
    db.session.add(row)
    db.session.commit()


def asset_raw_payload(asset):
    if asset.raw_row_json:
        try:
            payload = json.loads(asset.raw_row_json)
            if isinstance(payload, dict):
                return payload
        except Exception:
            pass
    return {
        'C_ACT': asset.c_act,
        'NOM': asset.nom or '',
        'C_FAM': asset.c_fam or '',
        'NOM_FAM': asset.nom_fam or '',
        'MODELO': asset.modelo or '',
        'REF': asset.ref or '',
        'SERIE': asset.serie or '',
        'NOM_MARCA': asset.nom_marca or '',
        'C_TIAC': asset.c_tiac or '',
        'DESC_TIAC': asset.desc_tiac or '',
        'DES_SUBTIAC': asset.desc_subtiac or '',
        'DEPRECIA': asset.deprecia or '',
        'VIDA_UTIL': asset.vida_util or '',
        'DES_UBI': asset.des_ubi or '',
        'NOM_CCOS': asset.nom_ccos or '',
        'NOM_RESP': asset.nom_resp or '',
        'EST': asset.est or '',
        'COSTO': to_number(asset.costo),
        'SALDO': to_number(asset.saldo),
        'FECHA_COMPRA': asset.fecha_compra or '',
        'CODIGO_INTELIGENTE': asset.codigo_inteligente or '',
        'SUBTIPO_CODIGO': asset.subtipo_codigo or '',
        'COLOR': asset.color or '',
        'NIT_PROVEEDOR': asset.nit_proveedor or '',
        'DESCRIPCION_PROVEEDOR': asset.desc_proveedor or '',
        'FORMA_ADQUISICION': asset.forma_adquisicion or '',
        'EN_GARANTIA': asset.en_garantia or '',
        'ENTIDAD_GARANTIA': asset.entidad_garantia or '',
        'GARANTIA_DESDE': asset.garantia_desde or '',
        'GARANTIA_HASTA': asset.garantia_hasta or '',
        'AGENCIA': asset.agencia or '',
        'CENTRO_COSTO_CODIGO': asset.centro_costo_code or '',
    }


def _pick_first_value(payload, keys):
    normalized_payload = {normalize_lookup_key(k): v for k, v in payload.items()}
    for key in keys:
        val = payload.get(key)
        if val is None:
            val = normalized_payload.get(normalize_lookup_key(key))
        if val is None:
            continue
        txt = str(val).strip()
        if txt:
            return txt
    return ''


def normalize_search_text(value):
    txt = str(value or '').strip().lower()
    if not txt:
        return ''
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')
    txt = re.sub(r'[^a-z0-9]+', ' ', txt)
    return re.sub(r'\s+', ' ', txt).strip()


def tokenize_search_text(value, min_len=ASSET_ASSIST_OCR_MIN_TOKEN_SIZE):
    txt = normalize_search_text(value)
    if not txt:
        return []
    tokens = []
    seen = set()
    for raw in txt.split():
        token = raw.strip()
        if len(token) < min_len:
            continue
        if token.isdigit() and len(token) < 4:
            continue
        if token in seen:
            continue
        seen.add(token)
        tokens.append(token)
    return tokens


def parse_detected_labels(raw_json):
    if not raw_json:
        return []
    try:
        payload = json.loads(raw_json)
    except Exception:
        return []
    if not isinstance(payload, list):
        return []
    out = []
    for item in payload[:15]:
        if isinstance(item, str):
            label = item
            score = None
        elif isinstance(item, dict):
            label = item.get('label')
            score = item.get('score')
        else:
            continue
        label_txt = normalize_search_text(label)
        if not label_txt:
            continue
        out.append({'label': label_txt, 'score': to_number(score)})
    return out


def get_assist_rule(rule_key):
    for rule in ASSET_ASSIST_CATEGORY_RULES:
        if rule.get('key') == rule_key:
            return rule
    return None


def category_keyword_matches(text_blob, rule):
    blob = normalize_search_text(text_blob)
    include_hits = []
    exclude_hits = []
    for kw in rule.get('keywords', []):
        token = normalize_search_text(kw)
        if token and token in blob:
            include_hits.append(token)
    for kw in rule.get('exclude_keywords', []):
        token = normalize_search_text(kw)
        if token and token in blob:
            exclude_hits.append(token)
    return include_hits, exclude_hits


def classify_asset_type_signal(ocr_tokens, ocr_text, detected_labels):
    signal_text = ' '.join(ocr_tokens or [])
    if ocr_text:
        signal_text = (signal_text + ' ' + normalize_search_text(ocr_text)).strip()
    label_tokens = [normalize_search_text(d.get('label')) for d in (detected_labels or []) if d.get('label')]
    ranked = []
    for rule in ASSET_ASSIST_CATEGORY_RULES:
        points = 0.0
        reasons = []
        for kw in rule.get('keywords', []):
            kw_txt = normalize_search_text(kw)
            if kw_txt and kw_txt in signal_text:
                points += 28.0
                reasons.append(f"texto:{kw_txt}")
        for kw in rule.get('exclude_keywords', []):
            kw_txt = normalize_search_text(kw)
            if kw_txt and kw_txt in signal_text:
                points -= 35.0
                reasons.append(f"exclusion:{kw_txt}")
        for lbl in rule.get('model_labels', []):
            lbl_txt = normalize_search_text(lbl)
            for item in detected_labels or []:
                det_label = normalize_search_text(item.get('label'))
                if not det_label:
                    continue
                if lbl_txt == det_label or lbl_txt in det_label or det_label in lbl_txt:
                    det_score = float(item.get('score') or 0.0)
                    points += 16.0 + (det_score * 20.0)
                    reasons.append(f"vision:{det_label}")
                    break
        ranked.append({
            'key': rule.get('key') or '',
            'label': rule.get('label') or '',
            'points': round(points, 4),
            'reasons': reasons[:10],
        })

    ranked.sort(key=lambda x: x['points'], reverse=True)
    best = ranked[0] if ranked else None

    if not best or best['points'] <= 0:
        return {
            'key': '',
            'label': 'Sin clasificar',
            'confidence': 0,
            'reasons': [],
            'detected_labels': [x for x in label_tokens if x],
            'ranked': ranked[:3],
            'candidate_keys': [],
        }

    confidence = min(99, int(round(best['points'])))
    candidate_keys = [best['key']] if best.get('key') else []
    if len(ranked) > 1 and ranked[1].get('points', 0) > 0:
        delta = best['points'] - ranked[1]['points']
        # Si dos tipos son cercanos, habilita ambos para evitar perder candidatos del tipo real.
        if delta <= 12 and ranked[1].get('key'):
            candidate_keys.append(ranked[1]['key'])

    return {
        'key': best['key'],
        'label': best['label'],
        'confidence': confidence,
        'reasons': best['reasons'][:8],
        'detected_labels': [x for x in label_tokens if x],
        'ranked': ranked[:3],
        'candidate_keys': candidate_keys,
    }


def extract_text_signals_from_image(file_bytes):
    result = {
        'ocr_text': '',
        'ocr_tokens': [],
        'serial_candidates': [],
        'ocr_engine': 'none',
        'ocr_available': False,
    }
    if not file_bytes:
        return result

    try:
        import pytesseract
        from PIL import Image as PILImage
        from PIL import ImageOps, ImageEnhance
    except ImportError:
        return result

    try:
        img = PILImage.open(BytesIO(file_bytes))
        img.load()
    except Exception:
        return result

    variants = []
    try:
        base = img.convert('RGB')
        variants.append(base)
        gray = ImageOps.grayscale(base)
        variants.append(gray)
        sharp = ImageEnhance.Sharpness(gray).enhance(1.8)
        variants.append(sharp)
        high = ImageEnhance.Contrast(gray).enhance(2.0)
        variants.append(high)
    except Exception:
        variants = [img]

    best_txt = ''
    for variant in variants[:4]:
        try:
            txt = pytesseract.image_to_string(variant, lang='eng+spa')
        except Exception:
            try:
                txt = pytesseract.image_to_string(variant)
            except Exception:
                txt = ''
        if len(txt or '') > len(best_txt):
            best_txt = txt or ''

    clean_text = best_txt.strip()
    if not clean_text:
        return result

    serials = re.findall(r'[A-Z0-9\-]{6,}', str(clean_text).upper())
    serial_candidates = []
    seen = set()
    for token in serials:
        t = token.strip().upper()
        if not t or t in seen:
            continue
        seen.add(t)
        serial_candidates.append(t)

    result['ocr_text'] = clean_text[:4000]
    result['ocr_tokens'] = tokenize_search_text(clean_text)
    result['serial_candidates'] = serial_candidates[:20]
    result['ocr_engine'] = 'pytesseract'
    result['ocr_available'] = True
    return result


def build_asset_assist_text_blob(asset):
    values = [
        asset.c_act, asset.nom, asset.nom_fam, asset.desc_tiac, asset.desc_subtiac,
        asset.modelo, asset.ref, asset.serie, asset.nom_marca, asset.nom_ccos, asset.des_ubi,
    ]
    return normalize_search_text(' '.join(str(v or '') for v in values))


def _similarity_ratio(left, right):
    a = normalize_search_text(left)
    b = normalize_search_text(right)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def score_asset_assisted_candidate(asset, category_info, ocr_tokens, serial_candidates, service_hint, location_hint):
    score = 0.0
    reasons = []
    blob = build_asset_assist_text_blob(asset)
    category_key = category_info.get('key') if category_info else ''
    category_conf = to_number(category_info.get('confidence') if category_info else 0)
    candidate_keys = category_info.get('candidate_keys') if category_info else []
    if not isinstance(candidate_keys, list):
        candidate_keys = []

    category_alignment = 0.0
    if category_key:
        for key in candidate_keys or [category_key]:
            rule = get_assist_rule(key)
            if not rule:
                continue
            include_hits, exclude_hits = category_keyword_matches(blob, rule)
            align_points = (len(include_hits) * 24.0) - (len(exclude_hits) * 30.0)
            if align_points > category_alignment:
                category_alignment = align_points
            if include_hits:
                score += min(60.0, 30.0 + len(include_hits) * 8.0)
                reasons.append(f"tipo:{rule.get('key')}:{include_hits[0]}")
            if exclude_hits:
                score -= min(80.0, 24.0 + len(exclude_hits) * 14.0)
                reasons.append(f"descarta:{rule.get('key')}:{exclude_hits[0]}")

    for token in (ocr_tokens or [])[:40]:
        if token in blob:
            score += 4.0
            reasons.append(f"ocr:{token}")

    serial_fields = [asset.serie, asset.ref, asset.modelo, asset.c_act]
    serial_blob = normalize_search_text(' '.join(str(v or '') for v in serial_fields))
    for serial in serial_candidates or []:
        serial_txt = normalize_search_text(serial)
        if not serial_txt:
            continue
        if serial_txt in serial_blob:
            score += 55.0
            reasons.append(f"serial:{serial_txt}")

    svc = str(service_hint or '').strip()
    if svc:
        svc_norm = normalize_search_text(svc)
        if svc_norm and svc_norm == normalize_search_text(asset.nom_ccos):
            score += 16.0
            reasons.append('servicio:exacto')
        elif svc_norm and svc_norm in normalize_search_text(asset.nom_ccos):
            score += 9.0
            reasons.append('servicio:parcial')

    loc = str(location_hint or '').strip()
    if loc:
        loc_norm = normalize_search_text(loc)
        if loc_norm:
            ratio = _similarity_ratio(loc_norm, asset.des_ubi or '')
            if ratio >= 0.85:
                score += 12.0
                reasons.append('ubicacion:alta')
            elif ratio >= 0.65:
                score += 7.0
                reasons.append('ubicacion:media')

    # En detecciones con confianza alta, forzar coherencia de tipo para reducir falsos positivos.
    if category_key and category_conf >= 55 and category_alignment < 1:
        score -= 70.0
        reasons.append('penalizacion:tipo_incompatible')

    # Prioriza activos de mayor impacto economico para acelerar decision operativa.
    value_weight = min(max(to_number(asset.costo), 0.0), 30000000.0) / 30000000.0
    score += value_weight * 4.0
    return score, reasons[:10]


def build_asset_assisted_candidates(max_candidates, category_info, ocr_tokens, serial_candidates, service_hint, location_hint):
    query = Asset.query.filter(Asset.estado_inventario == 'No encontrado')
    rows = query.limit(5000).all()

    scored = []
    for asset in rows:
        score, reasons = score_asset_assisted_candidate(
            asset=asset,
            category_info=category_info,
            ocr_tokens=ocr_tokens,
            serial_candidates=serial_candidates,
            service_hint=service_hint,
            location_hint=location_hint,
        )
        if score <= -45:
            continue
        scored.append((score, reasons, asset))

    scored.sort(key=lambda x: (x[0], to_number(x[2].costo), x[2].id), reverse=True)
    limit = max(1, min(max_candidates, 40))
    top = scored[:limit]

    # Si el tipo ya fue detectado y faltan candidatos, completa con activos del mismo tipo
    # aunque tengan evidencia OCR baja, para no ocultar opciones reales.
    if len(top) < limit and category_info and category_info.get('key'):
        need = limit - len(top)
        candidate_keys = category_info.get('candidate_keys') or [category_info.get('key')]
        supplemental = []
        for asset in rows:
            code = str(asset.c_act or '').strip()
            if not code or any(code == str(t[2].c_act or '').strip() for t in top):
                continue
            blob = build_asset_assist_text_blob(asset)
            for key in candidate_keys:
                rule = get_assist_rule(key)
                if not rule:
                    continue
                include_hits, exclude_hits = category_keyword_matches(blob, rule)
                if include_hits and not exclude_hits:
                    supplemental.append((0.5, [f"relleno_tipo:{key}"], asset))
                    break
            if len(supplemental) >= need:
                break
        top.extend(supplemental[:need])
        top.sort(key=lambda x: (x[0], to_number(x[2].costo), x[2].id), reverse=True)

    out = []
    for score, reasons, asset in top:
        out.append({
            'score': round(score, 2),
            'match_reasons': reasons,
            'asset': {
                'codigo': asset.c_act or '',
                'descripcion': asset.nom or '',
                'familia': asset.nom_fam or '',
                'marca': asset.nom_marca or '',
                'modelo': asset.modelo or '',
                'serial_ref': asset.serie or asset.ref or '',
                'servicio': asset.nom_ccos or '',
                'ubicacion': asset.des_ubi or '',
                'responsable': asset.nom_resp or '',
                'estado_inventario': asset.estado_inventario or '',
                'costo': round(to_number(asset.costo), 2),
            }
        })
    return out, len(rows)


def _token_match_ratio(tokens, text_blob):
    if not tokens:
        return 1.0, []
    blob = normalize_search_text(text_blob)
    matched = [tok for tok in tokens if tok in blob]
    return (len(matched) / max(1, len(tokens))), matched


def build_quick_lookup_candidates(rule_key, service_hint, location_hint, query_text, technical_text, subtype_text, limit):
    query = Asset.query.filter(Asset.estado_inventario == 'No encontrado')
    rows = query.limit(8000).all()
    rule = get_assist_rule(rule_key) if rule_key else None
    query_tokens = tokenize_search_text(query_text)
    technical_tokens = tokenize_search_text(technical_text)
    subtype_tokens = tokenize_search_text(subtype_text)

    filtered = []
    svc_filter = normalize_search_text(service_hint)
    loc_filter = normalize_search_text(location_hint)

    for asset in rows:
        blob = build_asset_assist_text_blob(asset)
        reasons = []
        include_hits = []
        exclude_hits = []
        group_full = str(asset.tipo_activo_cache or '').strip() or classify_asset_group(asset)
        major_type = normalize_asset_major_type(group_full)

        if rule:
            selected_major = normalize_asset_major_type(rule.get('label'))
            if major_type != selected_major:
                continue
            reasons.append(f"tipo:{selected_major}")
            include_hits, exclude_hits = category_keyword_matches(blob, rule)
            if include_hits:
                reasons.append(f"tipo_ref:{include_hits[0]}")

        if svc_filter:
            asset_svc = normalize_search_text(asset.nom_ccos)
            if svc_filter != asset_svc and svc_filter not in asset_svc:
                continue
            reasons.append('servicio')

        if loc_filter:
            asset_loc = normalize_search_text(asset.des_ubi)
            if loc_filter != asset_loc and loc_filter not in asset_loc:
                continue
            reasons.append('ubicacion')

        subtype_blob = normalize_search_text(' '.join([
            str(asset.desc_subtiac or ''),
            str(asset.nom_fam or ''),
            str(asset.nom or ''),
        ]))
        if subtype_tokens:
            subtype_ratio, subtype_matches = _token_match_ratio(subtype_tokens, subtype_blob)
            if subtype_ratio < 0.34:
                continue
            reasons.append(f"subtipo:{'/'.join(subtype_matches[:3])}")

        # Para descripcion y tecnico: usa cobertura parcial para no dejar pocos activos.
        if query_tokens:
            desc_ratio, desc_matches = _token_match_ratio(query_tokens, blob)
            if desc_ratio < 0.34:
                continue
            reasons.append(f"descripcion:{'/'.join(desc_matches[:3])}")

        technical_blob = normalize_search_text(' '.join([
            str(asset.nom_marca or ''),
            str(asset.modelo or ''),
            str(asset.serie or ''),
            str(asset.ref or ''),
            str(asset.c_act or ''),
            str(asset.nom or ''),
            str(asset.desc_subtiac or ''),
        ]))
        if technical_tokens:
            tech_ratio, tech_matches = _token_match_ratio(technical_tokens, technical_blob)
            if tech_ratio < 0.34:
                continue
            reasons.append(f"tecnico:{'/'.join(tech_matches[:3])}")

        # Cuando no hay filtros de texto y solo se define tipo, incluye todo el tipo.
        if not reasons:
            continue

        # Orden simple: mas filtros cumplidos primero, luego costo y codigo.
        score = float(len(reasons))
        filtered.append((score, reasons, asset))

    filtered.sort(key=lambda x: (x[0], to_number(x[2].costo), str(x[2].c_act or '')), reverse=True)
    requested = max(10, min(parse_int(limit, default=30) or 30, 80))
    top = filtered[:requested]

    out = []
    for score, reasons, asset in top[:requested]:
        out.append({
            'score': round(score, 2),
            'type_match': bool(rule_key),
            'match_reasons': reasons[:8],
            'asset': {
                'codigo': asset.c_act or '',
                'descripcion': asset.nom or '',
                'familia': asset.nom_fam or '',
                'tipo': asset.desc_tiac or '',
                'subtipo': asset.desc_subtiac or '',
                'marca': asset.nom_marca or '',
                'modelo': asset.modelo or '',
                'serial_ref': asset.serie or asset.ref or '',
                'servicio': asset.nom_ccos or '',
                'ubicacion': asset.des_ubi or '',
                'responsable': asset.nom_resp or '',
                'estado_inventario': asset.estado_inventario or '',
                'costo': round(to_number(asset.costo), 2),
            }
        })
    return out, len(rows)


