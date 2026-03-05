from ..core.foundation import *

@app.route('/')
def index():
    ensure_db()
    return render_template('home.html')


@app.route('/inventario')
def inventario_page():
    ensure_db()
    return render_template('inventario.html')


@app.route('/jornadas')
def jornadas_page():
    ensure_db()
    return render_template('jornadas.html')


@app.route('/formatos')
def formatos_page():
    ensure_db()
    return render_template('formatos.html')


@app.route('/bajas')
def bajas_page():
    ensure_db()
    return render_template('bajas.html')


@app.route('/dashboard')
def dashboard_page():
    ensure_db()
    return render_template('dashboard.html')


@app.route('/informes')
def informes_page():
    ensure_db()
    return render_template('informes.html')


@app.route('/cronograma')
def cronograma_page():
    ensure_db()
    return render_template('cronograma.html')


@app.route('/novedades')
def novedades_page():
    ensure_db()
    return render_template('novedades.html')


@app.route('/hoja_vida')
def hoja_vida_page():
    ensure_db()
    categories = [
        {'key': str(r.get('key') or ''), 'label': str(r.get('label') or '')}
        for r in ASSET_ASSIST_CATEGORY_RULES
        if str(r.get('key') or '').strip()
    ]
    return render_template('hoja_vida.html', assist_categories=categories)


@app.route('/documentos')
def documentos_page():
    ensure_db()
    return render_template('documentos.html')


@app.route('/logo')
def logo_file():
    logo_path = os.path.join(BASE_DIR, 'logo.png')
    if not os.path.exists(logo_path):
        return jsonify({'error': 'Logo no encontrado'}), 404
    return send_file(logo_path, mimetype='image/png')


@app.route('/import', methods=['POST'])
def import_file():
    ensure_db()
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400

    filename = f.filename.lower()
    try:
        if filename.endswith('.csv'):
            df = pd.read_csv(f)
        else:
            df = pd.read_excel(f)
    except Exception as e:
        return jsonify({'error': f'Error leyendo archivo: {e}'}), 400

    cols = normalize_columns(df.columns)
    if 'C_ACT' not in cols:
        return jsonify({'error': 'El archivo debe contener la columna C_ACT'}), 400

    imported = 0
    updated = 0
    ordered_cols = list(df.columns)

    def serialize_raw_value(v):
        if pd.isna(v):
            return None
        if isinstance(v, pd.Timestamp):
            return v.isoformat()
        if hasattr(v, 'item'):
            try:
                return v.item()
            except Exception:
                pass
        return v

    for _, row in df.iterrows():
        c_act_val = get_cell(row, cols, 'C_ACT')
        c_act = str(c_act_val).strip() if c_act_val is not None else ''
        if not c_act:
            continue
        raw_payload = {}
        for col_name in ordered_cols:
            key = str(col_name).strip().upper()
            raw_payload[key] = serialize_raw_value(row[col_name])

        asset = Asset.query.filter_by(c_act=c_act).first()
        data = {
            'nom': get_cell(row, cols, 'NOM'),
            'modelo': get_cell(row, cols, 'MODELO'),
            'ref': get_cell(row, cols, 'REF'),
            'serie': get_cell(row, cols, 'SERIE'),
            'nom_marca': get_cell(row, cols, 'NOM_MARCA'),
            'c_fam': get_cell(row, cols, 'C_FAM'),
            'nom_fam': get_cell(row, cols, 'NOM_FAM'),
            'c_tiac': get_cell(row, cols, 'C_TIAC'),
            'desc_tiac': get_cell(row, cols, 'DESC_TIAC'),
            'desc_subtiac': get_cell_first(row, cols, [
                'DES_SUBTIAC',
                'DESC_SUBTIAC',
                'DES_SUB_TIA',
                'DESC_SUB_TIA',
                'SUBTIAC',
                'DESC_SUBTIPO_ACTIVO',
                'DES_SUBTIPO_ACTIVO',
            ]),
            'deprecia': get_cell_first(row, cols, [
                'DEPRECIA',
                'DEP',
                'SE_DEPRECIA',
            ]),
            'vida_util': get_cell_first(row, cols, [
                'VIDA_UTIL',
                'VIDA UTIL',
                'V_UTIL',
            ]),
            'des_ubi': get_cell(row, cols, 'DES_UBI'),
            'nom_ccos': get_cell(row, cols, 'NOM_CCOS'),
            'nom_resp': get_cell(row, cols, 'NOM_RESP'),
            'est': get_cell(row, cols, 'EST'),
            'costo': try_float(get_cell(row, cols, 'COSTO')),
            'saldo': try_float(get_cell(row, cols, 'SALDO')),
            'fecha_compra': get_cell(row, cols, 'FECHA_COMPRA'),
            'codigo_inteligente': get_cell_first(row, cols, [
                'CODINTELIGENTE', 'CODIGO_INTELIGENTE', 'COD_INTELIGENTE', 'CODIGO INTELIGENTE',
            ]),
            'subtipo_codigo': get_cell_first(row, cols, [
                'SUBTIPO', 'SUBTIPO_ACTIVO', 'COD_SUBTIPO', 'COD_SUBTIPO_ACTIVO',
            ]),
            'color': get_cell_first(row, cols, [
                'COLOR', 'COLORES',
            ]),
            'nit_proveedor': get_cell_first(row, cols, [
                'NIT_PROVEEDOR', 'NIT PROVEEDOR', 'NIT',
            ]),
            'desc_proveedor': get_cell_first(row, cols, [
                'DESCRIPCION_PROVEEDOR', 'DESCRIPCION DEL PROVEEDOR', 'PROVEEDOR',
            ]),
            'forma_adquisicion': get_cell_first(row, cols, [
                'FORMA_ADQUISICION', 'FORMA DE ADQUISICION', 'ADQUISICION',
            ]),
            'en_garantia': get_cell_first(row, cols, [
                'EN_GARANTIA', 'GARANTIA',
            ]),
            'entidad_garantia': get_cell_first(row, cols, [
                'ENTIDAD', 'ENTIDAD_GARANTIA',
            ]),
            'garantia_desde': get_cell_first(row, cols, [
                'GARANTIA_DESDE', 'DESDE',
            ]),
            'garantia_hasta': get_cell_first(row, cols, [
                'GARANTIA_HASTA', 'HASTA',
            ]),
            'agencia': get_cell_first(row, cols, [
                'AGENCIA',
            ]),
            'centro_costo_code': get_cell_first(row, cols, [
                'C_CCOS', 'COD_CENTRO_COSTO', 'CENTRO_COSTO',
            ]),
            'raw_row_json': json.dumps(raw_payload, ensure_ascii=False, default=str),
        }

        if asset:
            for k, v in data.items():
                if v is not None and v != 'nan':
                    setattr(asset, k, v)
            refresh_asset_type_cache(asset)
            updated += 1
        else:
            asset = Asset(c_act=c_act, **{k: v for k, v in data.items() if v is not None})
            refresh_asset_type_cache(asset)
            db.session.add(asset)
            imported += 1
    # Persist metadata so UI can show the currently imported base at all times.
    set_system_meta('last_import_file_name', str(f.filename or '').strip())
    set_system_meta('last_import_at', now_iso())
    set_system_meta('last_import_imported', str(imported))
    set_system_meta('last_import_updated', str(updated))

    db.session.commit()
    bump_assets_revision()
    invalidate_accounting_report_cache()
    return jsonify({'imported': imported, 'updated': updated})


@app.route('/import/status')
def import_status():
    ensure_db()

    file_name = (get_system_meta('last_import_file_name', '') or '').strip()
    imported_at = (get_system_meta('last_import_at', '') or '').strip()
    imported_raw = get_system_meta('last_import_imported', '0')
    updated_raw = get_system_meta('last_import_updated', '0')
    try:
        imported = int(str(imported_raw).strip())
    except Exception:
        imported = 0
    try:
        updated = int(str(updated_raw).strip())
    except Exception:
        updated = 0

    has_assets = db.session.query(Asset.id).first() is not None
    has_import = bool(file_name or imported_at or has_assets)

    return jsonify({
        'has_import': has_import,
        'file_name': file_name,
        'imported_at': imported_at,
        'imported_at_local': format_dt_local(imported_at) if imported_at else '',
        'imported': imported,
        'updated': updated,
    })


@app.route('/export_pdf')
def export_pdf():
    ensure_db()
    service = request.args.get('service')
    q = Asset.query
    if service:
        q = q.filter(Asset.nom_ccos == service)
    assets = [a.to_dict() for a in q.all()]
    if not assets:
        return jsonify({'error': 'No assets for given filter'}), 400

    out = BytesIO()
    c = canvas.Canvas(out, pagesize=letter)
    width, height = letter
    x_margin = 40
    y = height - 40
    c.setFont('Helvetica-Bold', 14)
    c.drawString(x_margin, y, f'A22 - Inventario - {service or "Todos"}')
    y -= 24
    c.setFont('Helvetica', 10)

    headers = [
        ('C_ACT', 'COD ACTIVO'),
        ('NOM', 'DESCRIPCION ACTIVO'),
        ('MODELO', 'MODELO'),
        ('SERIE', 'SERIAL'),
        ('DES_UBI', 'UBICACION'),
        ('NOM_RESP', 'RESPONSABLE'),
        ('estado_inventario', 'ESTADO INVENTARIO'),
    ]
    col_widths = [90, 140, 80, 80, 110, 110, 80]

    # draw header
    x = x_margin
    for i, (_, label) in enumerate(headers):
        c.drawString(x + 2, y, label)
        x += col_widths[i]
    y -= 14
    c.line(x_margin, y + 8, width - x_margin, y + 8)

    for a in assets:
        x = x_margin
        if y < 80:
            c.showPage()
            y = height - 40
        for i, (key, _) in enumerate(headers):
            text = str(a.get(key, '') or '')
            c.drawString(x + 2, y, text[:int(col_widths[i] / 6)])
            x += col_widths[i]
        y -= 14

    c.showPage()
    c.save()
    out.seek(0)
    base = f"a22_inventario_{service or 'almacen'}"
    filename = f"{clean_filename(base)}.pdf"
    return send_file(out, download_name=filename, as_attachment=True, mimetype='application/pdf')


def try_float(x):
    try:
        d = to_decimal_amount(x, default=None)
        return float(d) if d is not None else None
    except Exception:
        return None


def get_run_or_404(run_id):
    run = InventoryRun.query.get(run_id)
    if not run:
        return None, (jsonify({'error': 'Jornada no encontrada'}), 404)
    return run, None


def normalize_scan_code(raw):
    text = str(raw or '').strip().replace('\r', '').replace('\n', '')
    if not text:
        return ''
    compact = text.replace(' ', '')
    if compact and all(ch.isdigit() or ch == '.' for ch in compact):
        try:
            num = float(compact)
            if num.is_integer():
                return str(int(num))
        except Exception:
            pass
    return text


def scan_code_equals(left, right):
    a = normalize_scan_code(left)
    b = normalize_scan_code(right)
    if not a or not b:
        return False
    if a.casefold() == b.casefold():
        return True
    if a.isdigit() and b.isdigit():
        return (a.lstrip('0') or '0') == (b.lstrip('0') or '0')
    return False


def get_asset_by_code(code):
    if code is None:
        return None, None
    scan_code = normalize_scan_code(code)
    if not scan_code:
        return None, None

    for candidate in [scan_code, str(code).strip()]:
        if not candidate:
            continue
        asset = Asset.query.filter_by(c_act=candidate).first()
        if asset:
            return asset, 'C_ACT'

    if scan_code.isdigit():
        int_code = str(int(scan_code))
        asset = Asset.query.filter(Asset.c_act.in_([f'{int_code}.0', f'{int_code}.00'])).first()
        if asset:
            return asset, 'C_ACT'
        variants = Asset.query.filter(Asset.c_act.like(f'{int_code}.%')).limit(20).all()
        for row in variants:
            if scan_code_equals(row.c_act, scan_code):
                return row, 'C_ACT'

    keys = [
        'CODINTELIGENTE',
        'COD_BARRAS',
        'CODBARRAS',
        'CODIGO_BARRAS',
        'CODIGO DE BARRAS',
        'BARCODE',
        'BARRAS',
    ]
    candidates = Asset.query.filter(
        Asset.raw_row_json.isnot(None),
        Asset.raw_row_json.contains(scan_code)
    ).limit(500).all()
    for row in candidates:
        payload = asset_raw_payload(row)
        for key in keys:
            if scan_code_equals(payload.get(key), scan_code):
                return row, key
    return None, None


def get_asset_by_c_act_strict(code):
    if code is None:
        return None
    raw_code = str(code or '').strip()
    scan_code = normalize_scan_code(raw_code)
    if not scan_code:
        return None

    # Intento directo (tal cual y normalizado)
    for candidate in [raw_code, scan_code]:
        if not candidate:
            continue
        asset = Asset.query.filter_by(c_act=candidate).first()
        if asset:
            return asset

    # Intento robusto para codigos numericos: 7015, 7015.0, 07015, etc.
    if scan_code.isdigit():
        int_code = str(int(scan_code))
        candidates = Asset.query.filter(
            (Asset.c_act == int_code) |
            (Asset.c_act == f'{int_code}.0') |
            (Asset.c_act == f'{int_code}.00') |
            (Asset.c_act.like(f'{int_code}.%')) |
            (Asset.c_act.like(f'0%{int_code}'))
        ).limit(300).all()
        for row in candidates:
            if scan_code_equals(row.c_act, scan_code):
                return row
    return None


def get_or_create_default_period():
    period = InventoryPeriod.query.filter_by(name='LEGADO').first()
    if period:
        return period
    period = InventoryPeriod(
        name='LEGADO',
        period_type='historico',
        status='closed',
        created_at=now_iso(),
        notes='Periodo historico creado automaticamente para jornadas antiguas.',
    )
    db.session.add(period)
    db.session.commit()
    return period


def classify_area(service_name):
    service = (service_name or '').upper()
    assistential_keywords = [
        'URGEN', 'UCI', 'HOSP', 'QUIR', 'CIRUG', 'LAB', 'IMAGEN', 'CONSULT',
        'ODONTO', 'NEON', 'PEDIAT', 'FARM', 'SANGRE', 'RAYOS'
    ]
    administrative_keywords = [
        'ADMIN', 'GEREN', 'TALENTO', 'CONTAB', 'FINAN', 'SISTEM', 'ARCHIV',
        'ALMACE', 'JURID', 'FACTUR', 'CARTERA', 'COMPR', 'MANTEN'
    ]
    if any(k in service for k in assistential_keywords):
        return 'Asistencial'
    if any(k in service for k in administrative_keywords):
        return 'Administrativa'
    logistic_keywords = [
        'ALMACEN', 'LOGIST', 'MANTEN', 'SERVICIOS GENERALES', 'ACTIVOS FIJOS'
    ]
    if any(k in service for k in logistic_keywords):
        return 'Logistico'
    return 'Sin clasificar'


def a22_type_order_value(asset):
    group = classify_asset_group(asset).upper().strip()
    order_map = {
        'MUEBLE Y ENSER': 1,
        'BIOMEDICO': 2,
        'INDUSTRIAL': 3,
        'TECNOLOGICO': 4,
        'CONTROL': 5,
    }
    return order_map.get(group, 99)


def sort_assets_for_a22(assets):
    return sorted(
        list(assets or []),
        key=lambda a: (
            a22_type_order_value(a),
            str(classify_asset_group(a) or ''),
            str(a.c_act or ''),
        )
    )


def summarize_status(records):
    total = len(records)
    found = sum(1 for r in records if r.get('status') == 'Encontrado')
    not_found = sum(1 for r in records if r.get('status') == 'No encontrado')
    pending = max(total - found - not_found, 0)
    found_pct = round((found / total) * 100, 2) if total else 0
    not_found_pct = round((not_found / total) * 100, 2) if total else 0
    return {
        'total': total,
        'found': found,
        'not_found': not_found,
        'pending': pending,
        'found_pct': found_pct,
        'not_found_pct': not_found_pct,
    }


def to_number(value):
    try:
        if value is None:
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def to_decimal_amount(value, default=Decimal('0')):
    if value is None:
        return default
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        try:
            return Decimal(str(value))
        except Exception:
            return default
    txt = str(value).strip()
    if not txt:
        return default
    txt = txt.replace('$', '').replace(' ', '')
    if ',' in txt and '.' in txt:
        if txt.rfind(',') > txt.rfind('.'):
            txt = txt.replace('.', '').replace(',', '.')
        else:
            txt = txt.replace(',', '')
    elif ',' in txt:
        txt = txt.replace('.', '').replace(',', '.')
    try:
        return Decimal(txt)
    except Exception:
        return default


def normalize_override_asset_code(value):
    raw = str(value or '').strip()
    if not raw:
        return ''
    if raw.isdigit():
        return raw
    compact = raw.replace(',', '')
    try:
        dec = Decimal(compact)
        if dec == dec.to_integral_value():
            return str(dec.quantize(Decimal('1')))
    except Exception:
        pass
    return raw


def get_accounting_cost_overrides(month=None, year=None):
    return dict(ACCOUNTING_COST_OVERRIDES)


def accounting_overrides_signature(overrides):
    parts = [f'{k}:{overrides[k]}' for k in sorted(overrides.keys())]
    return '|'.join(parts)


def asset_book_value(asset):
    saldo = to_number(asset.saldo)
    costo = to_number(asset.costo)
    if saldo > 0:
        return saldo
    if costo > 0:
        return costo
    return 0.0


def money_text(value):
    return f"${value:,.0f}"


def classify_critical_asset(asset):
    text = ' '.join([
        str(asset.nom or ''),
        str(asset.desc_tiac or ''),
        str(asset.c_tiac or ''),
        str(asset.modelo or ''),
        str(asset.ref or ''),
    ]).upper()

    keyword_weights = {
        'VENTIL': 10,
        'VENTILADOR': 10,
        'MONITOR': 9,
        'SIGNOS': 9,
        'DESFIBR': 10,
        'INFUSION': 8,
        'ANESTES': 9,
        'RAYOS': 8,
        'RX': 8,
        'ECOGRAF': 8,
        'CAMILLA': 7,
        'UCI': 8,
        'RESPIR': 9,
        'TOMOGRAF': 10,
    }
    matched = [(k, w) for k, w in keyword_weights.items() if k in text]
    value = asset_book_value(asset)

    value_score = 0
    if value >= 80_000_000:
        value_score = 10
    elif value >= 40_000_000:
        value_score = 8
    elif value >= 20_000_000:
        value_score = 6
    elif value >= 10_000_000:
        value_score = 4
    elif value >= 5_000_000:
        value_score = 2

    key_score = max([w for _, w in matched], default=0)
    score = key_score + value_score
    is_critical = score >= 8

    reasons = []
    if matched:
        reasons.append('Tipo critico')
    if value >= 10_000_000:
        reasons.append('Alto valor')
    if not reasons and is_critical:
        reasons.append('Prioridad tecnica')

    return {
        'is_critical': is_critical,
        'score': score,
        'reasons': ', '.join(reasons) if reasons else 'Sin marca critica',
    }


def build_management_insights(payload):
    insights = []
    k = payload.get('kpis', {})
    f = payload.get('financial', {})
    by_service = payload.get('by_service', [])
    by_type = payload.get('by_type', [])
    by_service_value = payload.get('top_not_found_by_service_value', [])
    critical = payload.get('critical_not_found', [])

    if by_service:
        top_service = max(by_service, key=lambda x: x.get('not_found', 0))
        insights.append(
            f"Servicio con mas no encontrados: {top_service.get('name', 'N/D')} "
            f"({top_service.get('not_found', 0)} activos)."
        )
    if by_type:
        top_type = max(by_type, key=lambda x: x.get('not_found', 0))
        insights.append(
            f"Tipo con mayor faltante: {top_type.get('name', 'N/D')} "
            f"({top_type.get('not_found', 0)} no encontrados)."
        )
    if by_service_value:
        s = by_service_value[0]
        insights.append(
            f"Mayor impacto economico por servicio: {s.get('name', 'N/D')} "
            f"({money_text(to_number(s.get('not_found_value', 0)))})."
        )
    insights.append(
        f"Valor no encontrado total: {money_text(to_number(f.get('not_found_value', 0)))} "
        f"({f.get('not_found_value_pct', 0)}% del valor inventariado)."
    )
    if critical:
        insights.append(
            f"Activos criticos no encontrados: {len(critical)} "
            f"por {money_text(to_number(f.get('critical_not_found_value', 0)))}."
        )
    if k.get('found_pct', 0) < 95:
        insights.append("Cumplimiento inferior al 95%; se recomienda plan de choque por servicio.")
    return insights


def build_executive_narrative(payload):
    k = payload.get('kpis', {})
    f = payload.get('financial', {})
    c = payload.get('coverage', {})
    meta = payload.get('meta', {})
    period_name = (meta.get('period') or {}).get('name') or 'Periodo seleccionado'
    run_name = (meta.get('run') or {}).get('name') or 'todas las jornadas'
    service_filter = meta.get('service_filter') or 'todos los servicios del alcance'

    objetivo_general = (
        f"Evaluar el estado de los activos fijos inventariados en {period_name}, "
        f"considerando el alcance operativo definido para {run_name} y {service_filter}, "
        "con el fin de soportar decisiones de control, custodia y mejora continua."
    )
    objetivos_especificos = [
        "Cuantificar activos encontrados, no encontrados y pendientes dentro del alcance evaluado.",
        "Estimar el impacto economico de los no encontrados y priorizar riesgos criticos.",
        "Contrastar cobertura del inventario frente a la base total institucional para contexto gerencial.",
        "Identificar servicios y tipos de activo con mayor brecha para acciones correctivas.",
    ]

    total = k.get('total', 0)
    found = k.get('found', 0)
    not_found = k.get('not_found', 0)
    pending = k.get('pending', 0)
    found_pct = k.get('found_pct', 0)
    not_found_pct = k.get('not_found_pct', 0)
    not_found_value = money_text(to_number(f.get('not_found_value', 0)))
    total_value = money_text(to_number(f.get('total_value', 0)))
    scope_assets = c.get('scope_assets', total)
    base_assets = c.get('base_total_assets', total)
    scope_assets_pct = c.get('scope_assets_pct', 0)

    resumen = (
        f"En el corte analizado se evaluaron {total} activos dentro del alcance del periodo. "
        f"Se registraron {found} encontrados ({found_pct}%), {not_found} no encontrados "
        f"({not_found_pct}%) y {pending} pendientes. En terminos economicos, el valor no "
        f"encontrado asciende a {not_found_value} sobre un valor total evaluado de {total_value}. "
        f"La cobertura del alcance corresponde a {scope_assets} activos sobre una base de "
        f"{base_assets} ({scope_assets_pct}%)."
    )

    interpretacion = []
    if found_pct >= 98:
        interpretacion.append("El nivel de cumplimiento de encontrados es sobresaliente para el corte evaluado.")
    elif found_pct >= 95:
        interpretacion.append("El cumplimiento de encontrados es aceptable, con oportunidades puntuales de mejora.")
    else:
        interpretacion.append("El cumplimiento de encontrados es bajo para el estandar institucional esperado.")

    if not_found_pct >= 5:
        interpretacion.append("El porcentaje de no encontrados requiere intervencion prioritaria por riesgo operativo.")
    else:
        interpretacion.append("El porcentaje de no encontrados se mantiene en una franja controlable.")

    if c.get('base_not_in_scope_assets', 0) > 0:
        interpretacion.append(
            f"Existe una brecha de cobertura de {c.get('base_not_in_scope_assets', 0)} activos frente a la base total."
        )
    else:
        interpretacion.append("La cobertura del periodo frente a la base institucional es completa.")

    return {
        'objetivo_general': objetivo_general,
        'objetivos_especificos': objetivos_especificos,
        'resumen': resumen,
        'interpretacion': interpretacion,
    }


def build_executive_action_plan(payload):
    k = payload.get('kpis', {})
    f = payload.get('financial', {})
    by_service = payload.get('by_service', [])
    by_type = payload.get('by_type', [])

    found_pct = float(k.get('found_pct', 0) or 0)
    not_found_pct = float(k.get('not_found_pct', 0) or 0)
    not_found_value = to_number(f.get('not_found_value', 0))

    risk_level = 'BAJO'
    risk_reason = 'Cumplimiento estable y brecha controlada.'
    if found_pct < 95 or not_found_pct >= 5 or not_found_value >= 50_000_000:
        risk_level = 'ALTO'
        risk_reason = 'Riesgo operativo y economico alto por brecha de inventario.'
    elif found_pct < 98 or not_found_pct >= 2 or not_found_value >= 20_000_000:
        risk_level = 'MEDIO'
        risk_reason = 'Riesgo moderado; requiere seguimiento dirigido.'

    top_service = max(by_service, key=lambda x: x.get('not_found', 0)) if by_service else None
    top_type = max(by_type, key=lambda x: x.get('not_found', 0)) if by_type else None

    actions = [
        {
            'priority': '1 - Inmediata',
            'action': 'Plan de choque de localizacion y saneamiento de no encontrados.',
            'focus': (top_service.get('name') if top_service else 'Servicios con mayor brecha'),
            'owner': 'Lideres de servicio + Activos fijos',
            'term': '15 dias',
        },
        {
            'priority': '2 - Corto plazo',
            'action': 'Auditoria dirigida a activos de mayor valor no encontrados.',
            'focus': (top_type.get('name') if top_type else 'Tipos de activo criticos'),
            'owner': 'Control interno + Biomédica/Ingenieria',
            'term': '30 dias',
        },
        {
            'priority': '3 - Sostenimiento',
            'action': 'Estandarizar cierres por periodo y trazabilidad por jornada.',
            'focus': 'Gobierno del dato e indicadores',
            'owner': 'Activos fijos + Sistemas',
            'term': 'Trimestral',
        },
    ]

    return {
        'risk_level': risk_level,
        'risk_reason': risk_reason,
        'actions': actions,
    }


def build_executive_conclusion(payload):
    k = payload.get('kpis', {})
    f = payload.get('financial', {})
    c = payload.get('coverage', {})
    by_service = payload.get('by_service', [])
    by_type = payload.get('by_type', [])
    meta = payload.get('meta', {})

    period_name = (meta.get('period') or {}).get('name') or 'el periodo evaluado'
    run_name = (meta.get('run') or {}).get('name') or 'las jornadas del periodo'

    top_service = max(by_service, key=lambda x: x.get('not_found', 0)).get('name') if by_service else 'N/D'
    top_type = max(by_type, key=lambda x: x.get('not_found', 0)).get('name') if by_type else 'N/D'

    found_pct = k.get('found_pct', 0)
    not_found_pct = k.get('not_found_pct', 0)
    scope_assets = c.get('scope_assets', k.get('total', 0))
    base_assets = c.get('base_total_assets', k.get('total', 0))
    not_found_value = money_text(to_number(f.get('not_found_value', 0)))

    technical_line = (
        f"Para {period_name}, considerando {run_name}, el sistema evidencia una efectividad de localizacion "
        f"de {found_pct}% sobre {scope_assets} activos evaluados, con una brecha de no localizacion de "
        f"{not_found_pct}% y un impacto economico asociado de {not_found_value}. "
        f"El mayor foco de atencion se concentra en el servicio '{top_service}' y en el tipo '{top_type}'. "
        f"La cobertura operativa alcanzada frente a la base institucional es de {scope_assets}/{base_assets} activos."
    )

    governance_line = (
        "En consecuencia, este informe consolida una base tecnica confiable para la toma de decisiones "
        "estrategicas y la priorizacion de intervenciones por riesgo operativo, economico y asistencial."
    )

    commitment_line = (
        "El equipo de Almacen y Activos Fijos ratifica su compromiso integral con los objetivos institucionales, "
        "fortaleciendo la custodia, trazabilidad y sostenibilidad de los bienes muebles e inmuebles del hospital."
    )

    return f"{technical_line} {governance_line} {commitment_line}"


def classify_asset_group(asset):
    def clean(value):
        text = str(value or '').upper().strip()
        text = unicodedata.normalize('NFD', text)
        return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

    def classify_non_control_type():
        des_subtiac_local = clean(asset.desc_subtiac)
        c_tiac_local = clean(asset.c_tiac)
        text_local = ' '.join([
            des_subtiac_local,
            clean(asset.desc_tiac),
            clean(asset.nom),
            clean(asset.modelo),
            clean(asset.ref),
            clean(asset.estado_inventario),
        ])

        if des_subtiac_local:
            if any(k in des_subtiac_local for k in ['BIOMED', 'MEDIC', 'HOSPITAL', 'CLINIC']):
                return 'BIOMEDICO'
            if any(k in des_subtiac_local for k in ['MUEBLE', 'ENSER']):
                return 'MUEBLE Y ENSER'
            if any(k in des_subtiac_local for k in ['INDUSTR']):
                return 'INDUSTRIAL'
            if any(k in des_subtiac_local for k in ['TECNOLOG']):
                return 'TECNOLOGICO'

        if c_tiac_local == '2':
            return 'MUEBLE Y ENSER'

        biomed_keywords = [
            'BIOMED', 'VENTIL', 'MONITOR', 'DESFIB', 'INFUS', 'BOMBA DE INFUS',
            'RESPIR', 'ANESTES', 'ECOGRAF', 'TOMOGRAF', 'RAYOS X', 'RAYOS',
            'ELECTRO', 'ELECTROCARD', 'ELECTROBIST', 'ELECTROESTIM', 'ELECTROTERAP',
            'ECG', 'EKG', 'SIGNOS VITALES', 'INCUBAD', 'SUCCION', 'ASPIRADOR QUIRURG',
            'NEONATAL', 'CARDIO', 'DIALIS', 'ULTRASON', 'UCI',
        ]
        if any(k in text_local for k in biomed_keywords):
            return 'BIOMEDICO'

        industrial_keywords = [
            'INDUSTR', 'PLANTA', 'COMPRESOR', 'TABLERO', 'CALDERA', 'MOTOR',
            'GENERADOR', 'BOMBA HIDRAUL', 'TRANSFORMADOR', 'SUBESTACION', 'UPS INDUSTRIAL',
            'ENFRIADOR', 'CHILLER', 'TORRE DE ENFRIAMIENTO',
        ]
        if c_tiac_local == '3' or any(k in text_local for k in industrial_keywords):
            return 'INDUSTRIAL'

        furniture_keywords = [
            'MUEBLE', 'ENSER', 'SILLA', 'ESCRITORIO', 'ARCHIVADOR', 'CAMILLA', 'MESA',
            'GABINETE', 'ESTANTE', 'VITRINA', 'LOCKER',
        ]
        if any(k in text_local for k in furniture_keywords):
            return 'MUEBLE Y ENSER'

        return 'TECNOLOGICO'

    des_subtiac = clean(asset.desc_subtiac)
    deprecia = asset.deprecia
    vida_util = asset.vida_util
    text = ' '.join([
        des_subtiac,
        clean(asset.desc_tiac),
        clean(asset.nom),
        clean(asset.modelo),
        clean(asset.ref),
        clean(asset.estado_inventario),
    ])
    base_type = classify_non_control_type()

    # Regla prioritaria para activos de control:
    # si no deprecia o vida util es 0, va a CONTROL separado por subtipo.
    if is_non_depreciable(deprecia) or is_zero_useful_life(vida_util):
        return f'CONTROL - {base_type}'

    # Regla principal: clasificar desde DES_SUBTIAC.
    if des_subtiac:
        if any(k in des_subtiac for k in ['CONTROL']):
            return f'CONTROL - {base_type}'

    control_keywords = [
        'ACTIVO DE CONTROL', 'CONTROL', 'KIT CONTROL', 'EQUIPO DE CONTROL',
    ]
    if any(k in text for k in control_keywords):
        return f'CONTROL - {base_type}'

    return base_type


def refresh_asset_type_cache(asset):
    asset.tipo_activo_cache = classify_asset_group(asset)
    return asset.tipo_activo_cache


def date_only(value):
    return format_dt_local(value, '%Y-%m-%d')


def normalize_disposal_type_key(type_value):
    txt = str(type_value or '').strip().upper()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')
    if 'CONTROL' in txt:
        return 'CONTROL'
    if 'BIOMED' in txt:
        return 'BIOMEDICO'
    if 'MUEBLE' in txt:
        return 'MUEBLE Y ENSER'
    if 'INDUSTR' in txt:
        return 'INDUSTRIAL'
    if 'TECNOLOG' in txt:
        return 'TECNOLOGICO'
    return ''


def normalize_manual_disposal_type(type_value):
    txt = str(type_value or '').strip().upper()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')
    txt = re.sub(r'\s+', ' ', txt).strip()
    mapping = {
        'BIOMEDICO': 'BIOMEDICO',
        'MUEBLE Y ENSER': 'MUEBLE Y ENSER',
        'INDUSTRIAL': 'INDUSTRIAL',
        'TECNOLOGICO': 'TECNOLOGICO',
        'CONTROL - BIOMEDICO': 'CONTROL - BIOMEDICO',
        'CONTROL - MUEBLE Y ENSER': 'CONTROL - MUEBLE Y ENSER',
        'CONTROL - INDUSTRIAL': 'CONTROL - INDUSTRIAL',
        'CONTROL - TECNOLOGICO': 'CONTROL - TECNOLOGICO',
    }
    return mapping.get(txt, '')


def normalize_asset_major_type(group_value):
    txt = str(group_value or '').strip().upper()
    txt = unicodedata.normalize('NFD', txt)
    txt = ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')
    if 'BIOMED' in txt:
        return 'BIOMEDICO'
    if 'MUEBLE' in txt or 'ENSER' in txt:
        return 'MUEBLE Y ENSER'
    if 'INDUSTR' in txt:
        return 'INDUSTRIAL'
    if 'TECNOLOG' in txt:
        return 'TECNOLOGICO'
    return 'TECNOLOGICO'


def query_disposals(service=None, status=None, period_id=None):
    q = db.session.query(AssetDisposal, Asset).join(Asset, Asset.id == AssetDisposal.asset_id)
    if service:
        q = q.filter(Asset.nom_ccos == service)
    if status:
        q = q.filter(AssetDisposal.status == status)
    if period_id:
        q = q.filter(AssetDisposal.period_id == period_id)
    rows = q.order_by(AssetDisposal.id.desc()).limit(5000).all()
    items = []
    for d, a in rows:
        tipo_manual = str(a.tipo_activo_cache or '').strip()
        tipo = tipo_manual or classify_asset_group(a)
        item = {
            'id': d.id,
            'code': a.c_act or '',
            'name': a.nom or '',
            'service': a.nom_ccos or '',
            'type': tipo,
            'cost': to_number(a.costo),
            'saldo': to_number(a.saldo),
            'date': date_only(a.fecha_compra),
            'reason': d.reason or '',
            'status': d.status or '',
            'period_id': d.period_id,
        }
        items.append(item)
    return items


def summarize_disposals(rows):
    return {
        'count': len(rows),
        'total_cost': round(sum(r.get('cost', 0) for r in rows), 2),
        'total_saldo': round(sum(r.get('saldo', 0) for r in rows), 2),
    }


def write_disposal_sheet(ws, title, rows, saldo_header='SALDO POR DEPRECIAR', note_text=None):
    headers = [
        'COD ACTIVO FIJO',
        'DESCRIPCION',
        'COSTO INICIAL',
        'SALDO POR DEPRECIAR',
        'FECHA ADQUISICION',
        'MOTIVO DE BAJA',
    ]
    ws.title = title[:31]
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws['A1'].font = Font(bold=True, size=13, color='0B4F6C')
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center')

    summary = summarize_disposals(rows)
    saldo_resume_label = 'Total saldo por depreciar'
    if 'NO DEPRECIABLE' in str(saldo_header or '').upper() or 'CONTABLE' in str(saldo_header or '').upper():
        saldo_resume_label = 'Total saldo contable'
    ws.append([
        f"Total activos: {summary['count']}  |  Total costo inicial: {money_text(summary['total_cost'])}  |  {saldo_resume_label}: {money_text(summary['total_saldo'])}"
    ])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws['A2'].font = Font(bold=True, color='1E293B')
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center')

    header_row = 3
    if note_text:
        ws.append([note_text])
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
        ws['A3'].font = Font(bold=True, color='9A5F00')
        ws['A3'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        header_row = 4

    ws.append(headers)
    header_fill = PatternFill(fill_type='solid', start_color='EAF4FA', end_color='EAF4FA')
    header_font = Font(bold=True, color='0B4F6C')
    thin = Side(style='thin', color='D6E3EC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=header_row, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border

    for row in rows:
        ws.append([
            row.get('code', ''),
            row.get('name', ''),
            row.get('cost', 0),
            row.get('saldo', 0),
            row.get('date', ''),
            row.get('reason', ''),
        ])

    start_data_row = header_row + 1
    last_row = ws.max_row
    for r in range(start_data_row, last_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row=r, column=3).number_format = '"$"#,##0'
        ws.cell(row=r, column=4).number_format = '"$"#,##0'

    ws.append([
        'TOTALES',
        '',
        summary['total_cost'],
        summary['total_saldo'],
        '',
        '',
    ])
    total_row = ws.max_row
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.font = Font(bold=True, color='0B4F6C')
        cell.fill = PatternFill(fill_type='solid', start_color='F3F9FD', end_color='F3F9FD')
        cell.border = border
    ws.cell(row=total_row, column=3).number_format = '"$"#,##0'
    ws.cell(row=total_row, column=4).number_format = '"$"#,##0'

    widths = [16, 40, 20, 20, 18, 34]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f'A{start_data_row}'


def get_hospital_logo_path():
    for candidate in A22_LOGO_CANDIDATES:
        if os.path.exists(candidate):
            return candidate
    return None


def get_codificacion_path():
    for candidate in CODIFICACION_CANDIDATES:
        if os.path.exists(candidate):
            return candidate
    return None


def append_pdf_header_with_logo(story, title_text, meta_text, include_logo=True):
    logo_path = get_hospital_logo_path()
    title_style = ParagraphStyle(
        'RptTitle',
        fontName='Helvetica-Bold',
        fontSize=17,
        textColor=colors.HexColor('#0B4F6C'),
        leading=20,
    )
    meta_style = ParagraphStyle(
        'RptMeta',
        fontName='Helvetica',
        fontSize=9,
        textColor=colors.HexColor('#5A6B7B'),
        leading=12,
    )
    if include_logo and logo_path:
        logo = RLImage(logo_path, width=18 * mm, height=18 * mm)
        # Usa una tabla para fijar el logo en la esquina superior izquierda.
        head = Table([[logo, Paragraph(f"<b>{title_text}</b><br/>{meta_text}", meta_style)]], colWidths=[22 * mm, 160 * mm])
        head.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        story.append(head)
    else:
        story.append(Paragraph(title_text, title_style))
        story.append(Paragraph(meta_text, meta_style))
    story.append(Spacer(1, 6))


def make_pdf_page_header(
    logo_path,
    right_image_path=None,
    right_width_mm=44,
    right_height_mm=17,
    right_top_mm=17,
):
    logo_reader = ImageReader(logo_path) if logo_path else None
    right_reader = ImageReader(right_image_path) if right_image_path else None

    def _draw_header(c, doc):
        page_w, page_h = doc.pagesize
        c.saveState()
        if logo_reader:
            x = doc.leftMargin
            y = page_h - 16 * mm
            c.drawImage(logo_reader, x, y, width=14 * mm, height=14 * mm, preserveAspectRatio=True, mask='auto')
        if right_reader:
            # Imagen de codificacion en esquina superior derecha.
            right_w = right_width_mm * mm
            right_h = right_height_mm * mm
            x_right = page_w - doc.rightMargin - right_w
            y_right = page_h - right_top_mm * mm
            c.drawImage(right_reader, x_right, y_right, width=right_w, height=right_h, preserveAspectRatio=True, mask='auto')
        c.restoreState()

    return _draw_header


def add_logo_to_excel_sheet(ws, logo_path=None):
    if not logo_path:
        return
    try:
        logo = XLImage(logo_path)
        max_col = max(ws.max_column, 1)
        anchor_col = get_column_letter(max_col)
        logo.width = 90
        logo.height = 36
        # Evita anclar imagen sobre celdas combinadas del encabezado (fila 1),
        # lo cual puede provocar advertencias de reparacion en Excel.
        ws.add_image(logo, f'{anchor_col}2')
        current_h = ws.row_dimensions[2].height or 15
        ws.row_dimensions[2].height = max(current_h, 32)
    except Exception:
        # El reporte no debe fallar por un problema de imagen.
        return


def pdf_cell(text, styles, bold=False, align='LEFT', size=7):
    base = styles['Normal']
    style = ParagraphStyle(
        f'Cell_{align}_{size}_{1 if bold else 0}',
        parent=base,
        fontName='Helvetica-Bold' if bold else 'Helvetica',
        fontSize=size,
        leading=size + 1.5,
        alignment={'LEFT': 0, 'CENTER': 1, 'RIGHT': 2}.get(align, 0),
        wordWrap='CJK',
    )
    return Paragraph(escape(str(text or '')), style)


def reference_serial(asset):
    ref = (asset.ref or '').strip() if asset.ref else ''
    serial = (asset.serie or '').strip() if asset.serie else ''
    if ref and serial:
        return f'{ref} / {serial}'
    return ref or serial


def get_a22_scope(service=None, run_id=None, period_id=None):
    run = None
    q = Asset.query
    period = None
    if period_id:
        period = InventoryPeriod.query.get(period_id)
        if not period:
            return None, None, (jsonify({'error': 'Periodo no encontrado'}), 404)
    if run_id:
        run = InventoryRun.query.get(run_id)
        if not run:
            return None, None, (jsonify({'error': 'Jornada no encontrada'}), 404)
        if not period and run.period_id:
            period = InventoryPeriod.query.get(run.period_id)
        if period and run.period_id != period.id:
            return None, None, (jsonify({'error': 'La jornada no pertenece al periodo seleccionado'}), 400)
        q = apply_run_scope_filter(q, run)
    if service:
        q = q.filter(Asset.nom_ccos == service)
    if run:
        found_asset_ids = [
            row.asset_id for row in RunAssetStatus.query.filter(
                RunAssetStatus.run_id == run.id,
                RunAssetStatus.status == 'Encontrado'
            ).all()
        ]
        if not found_asset_ids:
            assets_scope = []
        else:
            assets_scope = q.filter(Asset.id.in_(found_asset_ids)).order_by(Asset.c_act.asc()).all()
    elif period:
        runs_in_period_q = InventoryRun.query.filter(InventoryRun.period_id == period.id)
        runs_in_period = runs_in_period_q.all()
        if service:
            svc_cf = str(service).casefold()
            runs_in_period = [
                r for r in runs_in_period
                if any(str(s).casefold() == svc_cf for s in run_scope_services(r))
            ]
        run_ids = [r.id for r in runs_in_period]
        scoped_assets = q.order_by(Asset.c_act.asc()).all()
        if not run_ids or not scoped_assets:
            assets_scope = []
        else:
            asset_ids = [a.id for a in scoped_assets]
            statuses = RunAssetStatus.query.filter(
                RunAssetStatus.run_id.in_(run_ids),
                RunAssetStatus.asset_id.in_(asset_ids)
            ).order_by(RunAssetStatus.id.desc()).all()
            latest_by_asset = {}
            for st in statuses:
                if st.asset_id not in latest_by_asset:
                    latest_by_asset[st.asset_id] = st.status
            allowed_ids = {aid for aid, st in latest_by_asset.items() if st == 'Encontrado'}
            assets_scope = [a for a in scoped_assets if a.id in allowed_ids]
    else:
        # Sin jornada: toma solo activos encontrados (escaneados/verificados como encontrados).
        assets_scope = q.filter(Asset.estado_inventario == 'Encontrado').order_by(Asset.c_act.asc()).all()
    return run, assets_scope, None


def normalize_inventory_status(value):
    txt = str(value or '').strip().upper()
    if txt == 'ENCONTRADO':
        return 'Encontrado'
    if txt == 'NO ENCONTRADO':
        return 'No encontrado'
    return 'Pendiente'


def build_reconciliation_rows(service=None, run_id=None, period_id=None):
    q = Asset.query
    run = None
    period = None
    if period_id:
        period = InventoryPeriod.query.get(period_id)
        if not period:
            return None, (jsonify({'error': 'Periodo no encontrado'}), 404)
    if run_id:
        run = InventoryRun.query.get(run_id)
        if not run:
            return None, (jsonify({'error': 'Jornada no encontrada'}), 404)
        if period and run.period_id != period.id:
            # Tolerancia ante desfasajes temporales de UI: si llega una combinacion
            # periodo/jornada invalida, usa la jornada como fuente de verdad.
            period = InventoryPeriod.query.get(run.period_id) if run.period_id else None
        q = apply_run_scope_filter(q, run)
    if service:
        q = q.filter(Asset.nom_ccos == service)
    assets = q.order_by(Asset.nom_ccos.asc(), Asset.c_act.asc()).all()

    status_map = {}
    if run and assets:
        statuses = RunAssetStatus.query.filter(
            RunAssetStatus.run_id == run.id,
            RunAssetStatus.asset_id.in_([a.id for a in assets])
        ).all()
        status_map = {s.asset_id: s.status for s in statuses}
    elif period and assets:
        runs_q = InventoryRun.query.filter(InventoryRun.period_id == period.id)
        runs_in_period = runs_q.all()
        if service:
            svc_cf = str(service).casefold()
            runs_in_period = [
                r for r in runs_in_period
                if any(str(s).casefold() == svc_cf for s in run_scope_services(r))
            ]
        run_ids = [r.id for r in runs_in_period]
        if run_ids:
            statuses = RunAssetStatus.query.filter(
                RunAssetStatus.run_id.in_(run_ids),
                RunAssetStatus.asset_id.in_([a.id for a in assets])
            ).order_by(RunAssetStatus.id.desc()).all()
            for st in statuses:
                if st.asset_id not in status_map:
                    status_map[st.asset_id] = st.status

    rows = []
    for a in assets:
        if run or period:
            status_value = status_map.get(a.id, 'Pendiente')
        else:
            status_value = a.estado_inventario
        rows.append({
            'C_ACT': a.c_act or '',
            'NOM': a.nom or '',
            'SERVICIO': a.nom_ccos or '',
            'UBICACION': a.des_ubi or '',
            'RESPONSABLE': a.nom_resp or '',
            'TIPO': classify_asset_group(a),
            'ESTADO_INVENTARIO': normalize_inventory_status(status_value),
            'FECHA_VERIFICACION': date_only(a.fecha_verificacion),
            'USUARIO_VERIFICADOR': a.usuario_verificador or '',
            'COSTO': to_number(a.costo),
            'SALDO': to_number(a.saldo),
        })
    return rows, None


def write_reconciliation_sheet(ws, title, rows):
    headers = [
        'CODIGO',
        'DESCRIPCION',
        'SERVICIO',
        'UBICACION',
        'RESPONSABLE',
        'TIPO',
        'ESTADO',
        'FECHA VERIFICACION',
        'USUARIO',
        'COSTO',
        'SALDO',
    ]
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws['A1'].font = Font(bold=True, size=13, color='0B4F6C')
    ws.append(headers)
    header_fill = PatternFill(fill_type='solid', start_color='EAF4FA', end_color='EAF4FA')
    header_font = Font(bold=True, color='0B4F6C')
    thin = Side(style='thin', color='D6E3EC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=2, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border

    for r in rows:
        ws.append([
            r['C_ACT'], r['NOM'], r['SERVICIO'], r['UBICACION'], r['RESPONSABLE'],
            r['TIPO'], r['ESTADO_INVENTARIO'], r['FECHA_VERIFICACION'], r['USUARIO_VERIFICADOR'],
            r['COSTO'], r['SALDO'],
        ])

    for i in range(3, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(i, c)
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(i, 10).number_format = '"$"#,##0'
        ws.cell(i, 11).number_format = '"$"#,##0'

    widths = [14, 36, 22, 24, 24, 18, 14, 16, 16, 14, 14]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = 'A3'


def excel_safe_sheet_name(base_name, used_names):
    txt = str(base_name or '').strip()
    if not txt:
        txt = 'SIN SERVICIO'
    for ch in ['\\', '/', '?', '*', '[', ']', ':']:
        txt = txt.replace(ch, ' ')
    txt = ' '.join(txt.split())
    if not txt:
        txt = 'SIN SERVICIO'
    txt = txt[:31]
    candidate = txt
    suffix = 2
    while candidate in used_names:
        base_trim = txt[: max(1, 31 - len(str(suffix)) - 1)]
        candidate = f'{base_trim}-{suffix}'
        suffix += 1
    used_names.add(candidate)
    return candidate


def fit_logo_to_a22_box(sheet, img, from_col=1, to_col=2, from_row=2, to_row=5, padding_px=8, shrink=0.88):
    # Aproximacion de tamaño de columnas/filas de Excel a pixeles.
    def col_px(col_index):
        letter = get_column_letter(col_index)
        width = sheet.column_dimensions[letter].width
        if width is None:
            width = 8.43
        return max(10, int(width * 7 + 5))

    def row_px(row_index):
        height = sheet.row_dimensions[row_index].height
        if height is None:
            height = 15
        return max(8, int(height * 96 / 72))

    target_w = sum(col_px(c) for c in range(from_col, to_col + 1)) - (padding_px * 2)
    target_h = sum(row_px(r) for r in range(from_row, to_row + 1)) - (padding_px * 2)
    if target_w <= 0 or target_h <= 0:
        return

    if img.width and img.height:
        scale = min(target_w / img.width, target_h / img.height) * shrink
        if scale > 0:
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
            offset_x = max(0, int((target_w - img.width) / 2) + padding_px)
            offset_y = max(0, int((target_h - img.height) / 2) + padding_px)

            marker = AnchorMarker(
                col=from_col - 1,
                colOff=pixels_to_EMU(offset_x),
                row=from_row - 1,
                rowOff=pixels_to_EMU(offset_y),
            )
            ext = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
            img.anchor = OneCellAnchor(_from=marker, ext=ext)


def build_dashboard_payload(service=None, run_id=None, period_id=None):
    q = Asset.query
    base_q = Asset.query
    run = None
    period = None
    period_runs = []
    if period_id:
        period = InventoryPeriod.query.get(period_id)
        if not period:
            return None, 'Periodo no encontrado'
        period_runs_q = InventoryRun.query.filter(InventoryRun.period_id == period.id)
        period_runs = period_runs_q.all()
    if run_id:
        run = InventoryRun.query.get(run_id)
        if not run:
            return None, 'Jornada no encontrada'
        if not period and run.period_id:
            period = InventoryPeriod.query.get(run.period_id)
        if period and run.period_id != period.id:
            return None, 'La jornada no pertenece al periodo seleccionado'
        q = apply_run_scope_filter(q, run)
    if service:
        q = q.filter(Asset.nom_ccos == service)
        base_q = base_q.filter(Asset.nom_ccos == service)
    elif period and not run:
        # En vista por periodo (sin jornada), limita a servicios que realmente
        # tuvieron jornadas en ese periodo para evitar mezclar servicios externos.
        services_in_period = sorted({
            svc
            for r in period_runs
            for svc in run_scope_services(r)
            if svc
        })
        if services_in_period:
            q = q.filter(Asset.nom_ccos.in_(services_in_period))
        else:
            q = q.filter(text('1=0'))

    base_assets = base_q.all()
    base_total_assets = len(base_assets)
    base_total_value = sum(asset_book_value(a) for a in base_assets)

    assets_scope = q.all()
    if not assets_scope:
        payload = {
            'kpis': summarize_status([]),
            'financial': {
                'total_value': 0,
                'found_value': 0,
                'not_found_value': 0,
                'pending_value': 0,
                'not_found_value_pct': 0,
                'critical_not_found_value': 0,
            },
            'by_service': [],
            'by_type': [],
            'by_area': [],
            'critical_not_found': [],
            'not_found_assets': [],
            'not_found_assets_total': 0,
            'not_found_assets_capped': False,
            'top_not_found_by_service_value': [],
            'top_not_found_by_type_value': [],
            'coverage': {
                'base_total_assets': base_total_assets,
                'base_total_value': round(base_total_value, 2),
                'scope_assets': 0,
                'scope_value': 0,
                'scope_assets_pct': 0,
                'scope_value_pct': 0,
                'base_not_in_scope_assets': max(base_total_assets, 0),
                'base_not_in_scope_value': round(base_total_value, 2),
            },
            'meta': {
                'run': run.to_dict() if run else None,
                'period': period.to_dict() if period else None,
                'service_filter': service or '',
                'generated_at': now_iso(),
                'generated_at_local': format_dt_local(now_iso()),
            }
        }
        payload['insights'] = build_management_insights(payload)
        return payload, None

    status_map = {}
    if run:
        statuses = RunAssetStatus.query.filter(
            RunAssetStatus.run_id == run.id,
            RunAssetStatus.asset_id.in_([a.id for a in assets_scope])
        ).all()
        status_map = {s.asset_id: s.status for s in statuses}
    elif period:
        run_ids = [r.id for r in period_runs]
        asset_ids = [a.id for a in assets_scope]
        if run_ids and asset_ids:
            statuses = RunAssetStatus.query.filter(
                RunAssetStatus.run_id.in_(run_ids),
                RunAssetStatus.asset_id.in_(asset_ids)
            ).order_by(RunAssetStatus.id.desc()).all()
            for st in statuses:
                if st.asset_id not in status_map:
                    status_map[st.asset_id] = st.status

    records = []
    critical_not_found = []
    not_found_assets = []
    for a in assets_scope:
        if run or period:
            status = status_map.get(a.id, 'Pendiente')
        else:
            status = a.estado_inventario
        if status not in {'Encontrado', 'No encontrado'}:
            status = 'Pendiente'
        value = asset_book_value(a)
        critical_info = classify_critical_asset(a)
        records.append({
            'asset_id': a.id,
            'code': a.c_act,
            'asset_name': a.nom or '',
            'service': a.nom_ccos or 'SIN SERVICIO',
            'type': classify_asset_group(a),
            'area': classify_area(a.nom_ccos),
            'status': status,
            'value': value,
            'is_critical': critical_info['is_critical'],
            'critical_score': critical_info['score'],
            'critical_reasons': critical_info['reasons'],
        })
        if status == 'No encontrado' and critical_info['is_critical']:
            critical_not_found.append({
                'code': a.c_act,
                'name': a.nom or '',
                'service': a.nom_ccos or 'SIN SERVICIO',
                'type': classify_asset_group(a),
                'value': value,
                'critical_score': critical_info['score'],
                'critical_reasons': critical_info['reasons'],
                'model': a.modelo or '',
                'serial': a.serie or '',
                'responsible': a.nom_resp or '',
                'location': a.des_ubi or '',
            })
        if status == 'No encontrado':
            not_found_assets.append({
                'code': a.c_act,
                'name': a.nom or '',
                'service': a.nom_ccos or 'SIN SERVICIO',
                'type': classify_asset_group(a),
                'value': value,
                'model': a.modelo or '',
                'serial': a.serie or '',
                'responsible': a.nom_resp or '',
                'location': a.des_ubi or '',
            })

    by_service_map = {}
    by_type_map = {}
    by_area_map = {}
    for r in records:
        by_service_map.setdefault(r['service'], []).append(r)
        by_type_map.setdefault(r['type'], []).append(r)
        by_area_map.setdefault(r['area'], []).append(r)

    by_service = [{
        'name': name,
        **summarize_status(items),
    } for name, items in by_service_map.items()]
    by_type = [{
        'name': name,
        **summarize_status(items),
    } for name, items in by_type_map.items()]
    by_area = [{
        'name': name,
        **summarize_status(items),
    } for name, items in by_area_map.items()]

    by_service.sort(key=lambda x: x['total'], reverse=True)
    by_type.sort(key=lambda x: x['total'], reverse=True)
    by_area.sort(key=lambda x: x['total'], reverse=True)

    total_value = sum(r['value'] for r in records)
    found_value = sum(r['value'] for r in records if r['status'] == 'Encontrado')
    not_found_value = sum(r['value'] for r in records if r['status'] == 'No encontrado')
    pending_value = max(total_value - found_value - not_found_value, 0)
    critical_not_found_value = sum(x['value'] for x in critical_not_found)
    not_found_value_pct = round((not_found_value / total_value) * 100, 2) if total_value else 0
    scope_assets_pct = round((len(records) / base_total_assets) * 100, 2) if base_total_assets else 0
    scope_value_pct = round((total_value / base_total_value) * 100, 2) if base_total_value else 0

    by_service_value_map = {}
    by_type_value_map = {}
    for r in records:
        if r['status'] != 'No encontrado':
            continue
        by_service_value_map[r['service']] = by_service_value_map.get(r['service'], 0) + r['value']
        by_type_value_map[r['type']] = by_type_value_map.get(r['type'], 0) + r['value']

    top_not_found_by_service_value = sorted(
        [{'name': k, 'not_found_value': v} for k, v in by_service_value_map.items()],
        key=lambda x: x['not_found_value'],
        reverse=True
    )
    top_not_found_by_type_value = sorted(
        [{'name': k, 'not_found_value': v} for k, v in by_type_value_map.items()],
        key=lambda x: x['not_found_value'],
        reverse=True
    )
    critical_not_found.sort(key=lambda x: (x['critical_score'], x['value']), reverse=True)
    not_found_assets.sort(key=lambda x: x['value'], reverse=True)
    not_found_assets_total = len(not_found_assets)
    not_found_assets_cap = 500
    not_found_assets = not_found_assets[:not_found_assets_cap]

    payload = {
        'kpis': summarize_status(records),
        'financial': {
            'total_value': round(total_value, 2),
            'found_value': round(found_value, 2),
            'not_found_value': round(not_found_value, 2),
            'pending_value': round(pending_value, 2),
            'not_found_value_pct': not_found_value_pct,
            'critical_not_found_value': round(critical_not_found_value, 2),
        },
        'by_service': by_service,
        'by_type': by_type,
        'by_area': by_area,
        'critical_not_found': critical_not_found,
        'not_found_assets': not_found_assets,
        'not_found_assets_total': not_found_assets_total,
        'not_found_assets_capped': not_found_assets_total > not_found_assets_cap,
        'top_not_found_by_service_value': top_not_found_by_service_value,
        'top_not_found_by_type_value': top_not_found_by_type_value,
        'coverage': {
            'base_total_assets': base_total_assets,
            'base_total_value': round(base_total_value, 2),
            'scope_assets': len(records),
            'scope_value': round(total_value, 2),
            'scope_assets_pct': scope_assets_pct,
            'scope_value_pct': scope_value_pct,
            'base_not_in_scope_assets': max(base_total_assets - len(records), 0),
            'base_not_in_scope_value': round(max(base_total_value - total_value, 0), 2),
        },
        'meta': {
            'run': run.to_dict() if run else None,
            'period': period.to_dict() if period else None,
            'service_filter': service or '',
            'generated_at': now_iso(),
            'generated_at_local': format_dt_local(now_iso()),
        }
    }
    payload['insights'] = build_management_insights(payload)
    return payload, None


