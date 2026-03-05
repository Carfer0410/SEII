from .accounting_documents_life import *
from ..core.foundation import _DEC_ZERO, _DEC_TWO, _YELLOW_FILL_DES

def write_headers_row(ws, row_idx, columns):
    for col_idx, col_name in enumerate(columns, start=1):
        c = ws.cell(row_idx, col_idx, col_name)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill(fill_type='solid', fgColor='0B4F6C')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = Border(
            left=Side(style='thin', color='BFD5E3'),
            right=Side(style='thin', color='BFD5E3'),
            top=Side(style='thin', color='BFD5E3'),
            bottom=Side(style='thin', color='BFD5E3'),
        )


@app.route('/reports/accounting_monthly_excel', methods=['GET'])
def report_accounting_monthly_excel():
    ensure_db()
    force_refresh = str(request.args.get('refresh', '')).strip().lower() in {'1', 'true', 'yes', 'si'}
    report_title = str(request.args.get('report_title') or '').strip()
    if not report_title:
        return jsonify({'error': 'Debes indicar el titulo del informe contable'}), 400
    generated_by = str(request.args.get('generated_by') or '').strip()
    if not generated_by:
        return jsonify({'error': 'Debes indicar el usuario que genera el informe'}), 400
    month, year = normalize_month_year(request.args.get('month'), request.args.get('year'))
    month_label = MONTH_LABELS_ES.get(month, str(month))
    period_label = f'{month_label} {year}'
    accounting_base = get_latest_accounting_base(month, year)
    if not accounting_base:
        return jsonify({'error': f'No hay base mensual cargada para {period_label}. Cargala en el modulo de Informes antes de generar.'}), 400
    selected_overrides = get_accounting_cost_overrides(month, year)
    overrides_signature = accounting_overrides_signature(selected_overrides)
    template_path = get_accounting_template_path()
    if not os.path.exists(template_path):
        return jsonify({'error': 'No se encontro la plantilla "INFORME CONTABILIDAD REF.xlsx"'}), 400
    template_mtime = int(os.path.getmtime(template_path))
    current_cache_key = (
        f"{ACCOUNTING_REPORT_ALGO_VERSION}:{template_mtime}:{month}:{year}:{report_title}:{generated_by}:"
        f"{overrides_signature}:base:{accounting_base.id}:{accounting_base.uploaded_at}"
    )

    with ACCOUNTING_CACHE_LOCK:
        cached_version = ACCOUNTING_REPORT_CACHE.get('version')
        cached_bytes = ACCOUNTING_REPORT_CACHE.get('bytes')
        cached_filename = ACCOUNTING_REPORT_CACHE.get('filename')

    if (not force_refresh) and cached_version == current_cache_key and cached_bytes:
        safe_period = sanitize_filename(period_label.replace(' ', '_'))
        base_filename = f'informe_conciliacion_activos_fijos_contabilidad_{safe_period}.xlsx'
        persist_accounting_report_file(
            cached_bytes,
            base_filename,
            period_label,
            month,
            year,
            report_title,
            period_id=None,
            accounting_base_id=accounting_base.id,
        )
        return send_file(
            BytesIO(cached_bytes),
            as_attachment=True,
            download_name=base_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    asset_rows = db.session.query(
        AccountingMonthlyBaseAsset.raw_row_json,
        AccountingMonthlyBaseAsset.c_act,
        AccountingMonthlyBaseAsset.c_fam,
        AccountingMonthlyBaseAsset.nom_fam,
        AccountingMonthlyBaseAsset.costo,
        AccountingMonthlyBaseAsset.saldo,
    ).filter_by(base_id=accounting_base.id).all()
    if not asset_rows:
        return jsonify({'error': f'La base mensual seleccionada ({period_label}) no tiene activos'}), 400
    overrides_summary = build_overrides_summary_for_rows(asset_rows, selected_overrides)

    def row_values(ws, row_idx):
        return [ws.cell(row_idx, c).value for c in range(1, ws.max_column + 1)]

    def non_empty_headers(ws, row_idx):
        values = row_values(ws, row_idx)
        return [str(v).strip() for v in values if str(v or '').strip()]

    def clear_sheet_data(ws, from_row):
        if ws.max_row >= from_row:
            ws.delete_rows(from_row, ws.max_row - from_row + 1)

    family_name_by_code = {}
    rows_by_code = {}
    base_rows = []
    all_codes_seen = set()

    wb = load_workbook(template_path)
    if 'BASE COMPLETA' not in wb.sheetnames or 'DESGLOSE' not in wb.sheetnames or 'INFORME' not in wb.sheetnames:
        return jsonify({'error': 'La plantilla no tiene las hojas requeridas: BASE COMPLETA, DESGLOSE, INFORME'}), 400

    ws_base = wb['BASE COMPLETA']
    ws_des = wb['DESGLOSE']
    ws_inf = wb['INFORME']
    ws_inf.column_dimensions['C'].width = max(18, ws_inf.column_dimensions['C'].width or 0)
    ws_inf.row_dimensions[3].height = max(55, ws_inf.row_dimensions[3].height or 0)
    logo_path = get_hospital_logo_path()
    if logo_path:
        try:
            logo_img = XLImage(logo_path)
            # Logo anclado y contenido en C3.
            fit_logo_to_a22_box(ws_inf, logo_img, from_col=3, to_col=3, from_row=3, to_row=3, padding_px=2, shrink=0.98)
            ws_inf.add_image(logo_img)
        except Exception:
            pass
    ws_inf.cell(3, 3, report_title)
    ws_inf.cell(49, 4, generated_by)
    d48_cell = ws_inf.cell(48, 4)
    d48_align = copy(d48_cell.alignment) if d48_cell.alignment else Alignment()
    d48_align.horizontal = 'center'
    d48_align.vertical = 'center'
    d48_align.wrap_text = True
    d48_cell.alignment = d48_align
    by_cell = ws_inf.cell(49, 4)
    by_align = copy(by_cell.alignment) if by_cell.alignment else Alignment()
    by_align.wrap_text = True
    by_align.shrink_to_fit = True
    by_align.horizontal = 'center'
    by_align.vertical = 'center'
    by_cell.alignment = by_align
    if len(generated_by) > 34:
        ws_inf.row_dimensions[49].height = max(28, ws_inf.row_dimensions[49].height or 0)

    template_headers = non_empty_headers(ws_base, 1)
    columns_order = template_headers[:] if template_headers else []
    seen_cols = set(columns_order)

    for raw_row_json, c_act, c_fam, nom_fam, costo, saldo in asset_rows:
        payload = {}
        if raw_row_json:
            try:
                maybe_payload = json.loads(raw_row_json)
                if isinstance(maybe_payload, dict):
                    payload = maybe_payload
            except Exception:
                payload = {}

        fam_code = normalize_family_code(payload.get('C_FAM') or c_fam)
        payload['C_FAM'] = fam_code
        payload['NOM_FAM'] = payload.get('NOM_FAM') or nom_fam or ''
        payload['COSTO'] = to_number(payload.get('COSTO') if payload.get('COSTO') is not None else costo)
        payload['SALDO'] = to_number(payload.get('SALDO') if payload.get('SALDO') is not None else saldo)
        payload['C_ACT'] = payload.get('C_ACT') or c_act

        for col in payload.keys():
            if col not in seen_cols:
                seen_cols.add(col)
                columns_order.append(col)

        base_rows.append(payload)
        all_codes_seen.add(fam_code)
        if fam_code and payload.get('NOM_FAM'):
            family_name_by_code[fam_code] = str(payload.get('NOM_FAM') or '')
        rows_by_code.setdefault(fam_code, []).append(payload)

    if 'COSTO' not in columns_order:
        columns_order.append('COSTO')

    report_rows_by_code = {}
    for fam_code, fam_rows in rows_by_code.items():
        code = normalize_family_code(fam_code)
        if not code or code in ACCOUNTING_EXCLUDED_FAMILIES:
            continue
        report_rows_by_code[code] = list(fam_rows)
    report_scope_assets_count = sum(len(rows) for rows in report_rows_by_code.values())
    excluded_in_scope = [code for code in report_rows_by_code.keys() if code in ACCOUNTING_EXCLUDED_FAMILIES]
    if STRICT_ACCOUNTING_VALIDATION and excluded_in_scope:
        return jsonify({'error': f'Validacion interna fallo: familias excluidas en alcance reportable ({", ".join(excluded_in_scope)})'}), 500

    catalog_names = load_family_catalog_names()
    configured_parent_codes = [code for code in ACCOUNTING_FAMILY_ORDER if len(code) == 4]
    parent_codes = []
    parent_codes_seen = set()

    for parent in configured_parent_codes:
        has_rows = any(code.startswith(parent) for code in report_rows_by_code.keys())
        if has_rows and parent not in parent_codes_seen:
            parent_codes.append(parent)
            parent_codes_seen.add(parent)

    dynamic_parent_codes = sorted({
        code[:4] for code in report_rows_by_code.keys()
        if len(code) >= 4 and code[:4].isdigit() and code[:4] not in parent_codes_seen
    })
    for parent in dynamic_parent_codes:
        parent_codes.append(parent)
        parent_codes_seen.add(parent)

    import time as _time
    _t0 = _time.perf_counter()

    clear_sheet_data(ws_base, 2)
    for col_idx, header in enumerate(columns_order, start=1):
        ws_base.cell(1, col_idx, header)

    base_rows_sorted = sorted(
        base_rows,
        key=lambda r: (normalize_family_code(r.get('C_FAM')), str(r.get('C_ACT') or ''))
    )
    row_idx = 2
    for row in base_rows_sorted:
        for col_idx, col_name in enumerate(columns_order, start=1):
            value = row.get(col_name)
            cell = ws_base.cell(row_idx, col_idx, value if value is not None else '')
            if col_name == 'COSTO':
                cell.number_format = '"$"#,##0.00'
        row_idx += 1

    app.logger.warning(f'[PERF] BASE COMPLETA escritura: {_time.perf_counter()-_t0:.2f}s ({len(base_rows_sorted)} filas)')
    _t1 = _time.perf_counter()

    clear_sheet_data(ws_des, 2)
    for col_idx, header in enumerate(columns_order, start=1):
        ws_des.cell(1, col_idx, header)

    des_row = 3
    cost_col = columns_order.index('COSTO') + 1
    cost_col_letter = get_column_letter(cost_col)
    des_total_report_scope = _DEC_ZERO
    des_detail_rows_written = 0
    assigned_codes = set()
    nc_total_row_des = None
    des_formula_cells = []
    family_total_refs = {}
    family_total_values = {}
    override_hit_counts = {code: 0 for code in selected_overrides.keys()}
    configured_children_by_parent = {
        parent: [code for code in ACCOUNTING_FAMILY_ORDER if len(code) > 4 and code.startswith(parent)]
        for parent in configured_parent_codes
    }

    for parent in parent_codes:
        family_codes_order = []
        if parent in report_rows_by_code:
            family_codes_order.append(parent)

        configured_children = configured_children_by_parent.get(parent, [])
        for child in configured_children:
            if child in report_rows_by_code and child not in family_codes_order:
                family_codes_order.append(child)

        dynamic_children = sorted(
            code for code in report_rows_by_code.keys()
            if len(code) > 4 and code.startswith(parent) and code not in family_codes_order
        )
        family_codes_order.extend(dynamic_children)
        assigned_codes.update(family_codes_order)

        rows = []
        for fam_code in family_codes_order:
            rows.extend(report_rows_by_code.get(fam_code, []))
        rows = sorted(rows, key=lambda r: (str(r.get('C_FAM') or ''), str(r.get('C_ACT') or '')))
        parent_name = (
            family_name_by_code.get(parent)
            or catalog_names.get(parent)
            or f'FAMILIA {parent}'
        )

        ws_des.cell(des_row, 2, parent)
        ws_des.cell(des_row, 3, parent_name)
        des_row += 2

        for col_idx, header in enumerate(columns_order, start=1):
            ws_des.cell(des_row, col_idx, header)
        des_row += 1

        subtotal = _DEC_ZERO
        parent_total_refs = []
        for fam_code in family_codes_order:
            fam_rows = sorted(
                report_rows_by_code.get(fam_code, []),
                key=lambda r: str(r.get('C_ACT') or '')
            )
            if not fam_rows:
                continue

            fam_subtotal = _DEC_ZERO
            detail_start_row = des_row
            for row in fam_rows:
                c_act_key = normalize_override_asset_code(row.get('C_ACT'))
                if c_act_key in override_hit_counts:
                    override_hit_counts[c_act_key] += 1
                for col_idx, col_name in enumerate(columns_order, start=1):
                    value = row.get(col_name)
                    if col_name == 'COSTO' and c_act_key in selected_overrides:
                        value = float(selected_overrides[c_act_key])
                    cell = ws_des.cell(des_row, col_idx, value if value is not None else '')
                    if col_name == 'COSTO':
                        cell.number_format = '"$"#,##0.00'
                        if c_act_key in selected_overrides:
                            cell.fill = _YELLOW_FILL_DES
                effective_cost = selected_overrides.get(c_act_key)
                if effective_cost is None:
                    effective_cost = to_decimal_amount(row.get('COSTO')).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
                fam_subtotal += effective_cost
                des_detail_rows_written += 1
                des_row += 1

            detail_end_row = des_row - 1
            fam_name = (
                family_name_by_code.get(fam_code)
                or catalog_names.get(fam_code)
                or f'FAMILIA {fam_code}'
            )
            family_total_d = fam_subtotal.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
            ws_des.cell(des_row, 3, f'TOTAL {fam_code} - {fam_name}').font = Font(bold=True)
            formula_range = f"${cost_col_letter}${detail_start_row}:${cost_col_letter}${detail_end_row}"
            family_total_cell = ws_des.cell(des_row, cost_col, f"=SUM({formula_range})")
            family_total_cell.font = Font(bold=True)
            family_total_cell.number_format = '"$"#,##0.00'
            family_total_refs[fam_code] = f'DESGLOSE!${cost_col_letter}${des_row}'
            family_total_values[fam_code] = family_total_d
            parent_total_refs.append(f'${cost_col_letter}${des_row}')
            des_row += 1
            subtotal += fam_subtotal

        parent_total_d = subtotal.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
        ws_des.cell(des_row, 3, f'TOTAL {parent} - {parent_name}').font = Font(bold=True)
        if parent_total_refs:
            total_cell = ws_des.cell(des_row, cost_col, f"=SUM({','.join(parent_total_refs)})")
        else:
            total_cell = ws_des.cell(des_row, cost_col, 0.0)
        total_cell.font = Font(bold=True)
        total_cell.number_format = '"$"#,##0.00'
        des_total_report_scope += subtotal
        des_row += 2

    app.logger.warning(f'[PERF] DESGLOSE familias asignadas: {_time.perf_counter()-_t1:.2f}s')
    _t2 = _time.perf_counter()

    unassigned_codes = sorted(code for code in report_rows_by_code.keys() if code not in assigned_codes)
    if unassigned_codes:
        ws_des.cell(des_row, 2, 'NC')
        ws_des.cell(des_row, 3, 'NO CLASIFICADAS / FUERA DE ESTRUCTURA')
        des_row += 2
        for col_idx, header in enumerate(columns_order, start=1):
            ws_des.cell(des_row, col_idx, header)
        des_row += 1

        nc_subtotal = _DEC_ZERO
        nc_total_refs = []
        for fam_code in unassigned_codes:
            fam_rows = sorted(report_rows_by_code.get(fam_code, []), key=lambda r: str(r.get('C_ACT') or ''))
            if not fam_rows:
                continue

            fam_subtotal = _DEC_ZERO
            detail_start_row = des_row
            for row in fam_rows:
                c_act_key = normalize_override_asset_code(row.get('C_ACT'))
                if c_act_key in override_hit_counts:
                    override_hit_counts[c_act_key] += 1
                for col_idx, col_name in enumerate(columns_order, start=1):
                    value = row.get(col_name)
                    if col_name == 'COSTO' and c_act_key in selected_overrides:
                        value = float(selected_overrides[c_act_key])
                    cell = ws_des.cell(des_row, col_idx, value if value is not None else '')
                    if col_name == 'COSTO':
                        cell.number_format = '"$"#,##0.00'
                        if c_act_key in selected_overrides:
                            cell.fill = _YELLOW_FILL_DES
                effective_cost = selected_overrides.get(c_act_key)
                if effective_cost is None:
                    effective_cost = to_decimal_amount(row.get('COSTO')).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
                fam_subtotal += effective_cost
                des_detail_rows_written += 1
                des_row += 1

            detail_end_row = des_row - 1
            fam_name = (
                family_name_by_code.get(fam_code)
                or catalog_names.get(fam_code)
                or f'FAMILIA {fam_code}'
            )
            family_total_d = fam_subtotal.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
            ws_des.cell(des_row, 3, f'TOTAL {fam_code} - {fam_name}').font = Font(bold=True)
            formula_range = f"${cost_col_letter}${detail_start_row}:${cost_col_letter}${detail_end_row}"
            family_total_cell = ws_des.cell(des_row, cost_col, f"=SUM({formula_range})")
            family_total_cell.font = Font(bold=True)
            family_total_cell.number_format = '"$"#,##0.00'
            family_total_refs[fam_code] = f'DESGLOSE!${cost_col_letter}${des_row}'
            family_total_values[fam_code] = family_total_d
            nc_total_refs.append(f'${cost_col_letter}${des_row}')
            des_row += 1
            nc_subtotal += fam_subtotal

        nc_total_d = nc_subtotal.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
        ws_des.cell(des_row, 3, 'TOTAL NC - NO CLASIFICADAS / FUERA DE ESTRUCTURA').font = Font(bold=True)
        if nc_total_refs:
            total_cell = ws_des.cell(des_row, cost_col, f"=SUM({','.join(nc_total_refs)})")
        else:
            total_cell = ws_des.cell(des_row, cost_col, 0.0)
        total_cell.font = Font(bold=True)
        total_cell.number_format = '"$"#,##0.00'
        des_total_report_scope += nc_subtotal
        nc_total_row_des = des_row
        des_row += 2

    expected_scope_codes = set(report_rows_by_code.keys())
    covered_scope_codes = assigned_codes.union(set(unassigned_codes))
    if STRICT_ACCOUNTING_VALIDATION and covered_scope_codes != expected_scope_codes:
        missing_codes = sorted(expected_scope_codes - covered_scope_codes)
        return jsonify({'error': f'Validacion interna fallo: familias sin cobertura en desglose ({", ".join(missing_codes)})'}), 500
    if STRICT_ACCOUNTING_VALIDATION and des_detail_rows_written != report_scope_assets_count:
        return jsonify({
            'error': (
                'Validacion interna fallo: filas de detalle en DESGLOSE '
                f'({des_detail_rows_written}) no coinciden con activos reportables ({report_scope_assets_count})'
            )
        }), 500
    duplicate_override_codes = sorted([code for code, hits in override_hit_counts.items() if hits > 1])
    if STRICT_ACCOUNTING_VALIDATION and duplicate_override_codes:
        return jsonify({
            'error': (
                'Validacion interna fallo: codigos override duplicados en el mes '
                f'({", ".join(duplicate_override_codes)})'
            )
        }), 500
    missing_override_codes = sorted([code for code, hits in override_hit_counts.items() if hits == 0])
    if missing_override_codes:
        app.logger.warning(
            '[ACCOUNTING] Overrides no encontrados para %s-%s: %s',
            year,
            str(month).zfill(2),
            ','.join(missing_override_codes),
        )
    app.logger.warning(f'[PERF] DESGLOSE NC + validaciones: {_time.perf_counter()-_t2:.2f}s')
    _t3 = _time.perf_counter()

    expected_scope_total = sum(
        (
            selected_overrides.get(normalize_override_asset_code(row.get('C_ACT')))
            or to_decimal_amount(row.get('COSTO')).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
        )
        for rows in report_rows_by_code.values()
        for row in rows
    ).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
    des_total_d = des_total_report_scope.quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
    if STRICT_ACCOUNTING_VALIDATION and abs(des_total_d - expected_scope_total) > Decimal('0.01'):
        return jsonify({
            'error': (
                'Validacion interna fallo: total DESGLOSE no coincide con total reportable '
                f'({des_total_d} vs {expected_scope_total})'
            )
        }), 500

    # INFORME: escribir exactamente sobre el bloque de la plantilla (filas 6..31)
    for r in range(6, 32):
        for c in range(3, 8):
            ws_inf.cell(r, c, None)
    # Quitar columnas de conciliacion contable del formato generado por sistema
    ws_inf.cell(5, 6, None)
    ws_inf.cell(5, 7, None)

    info_row = 6
    parent_rows_written = []

    def refs_for_prefix(prefix):
        return [
            ref
            for fam_code, ref in family_total_refs.items()
            if normalize_family_code(fam_code).startswith(prefix)
        ]

    def vals_for_prefix(prefix):
        return [
            family_total_values[fam_code]
            for fam_code in family_total_values
            if normalize_family_code(fam_code).startswith(prefix)
        ]

    parent_dec_totals = []
    for group in ACCOUNTING_REPORT_STRUCTURE:
        child_specs = []
        for child in group['children']:
            prefixes = child.get('source_prefixes') or [child.get('source_prefix')]
            formula_refs = []
            child_val_list = []
            for p in prefixes:
                if not p:
                    continue
                formula_refs.extend(refs_for_prefix(p))
                child_val_list.extend(vals_for_prefix(p))
            child_specs.append((child, formula_refs, child_val_list))

        parent_row = info_row
        ws_inf.cell(info_row, 3, group['parent_code'])
        ws_inf.cell(info_row, 4, group['parent_name'])
        ws_inf.cell(info_row, 3).font = Font(bold=True)
        ws_inf.cell(info_row, 4).font = Font(bold=True)
        parent_rows_written.append(parent_row)
        info_row += 1

        group_child_decs = []
        for child, formula_refs, child_val_list in child_specs:
            fallback_prefix = (child.get('source_prefixes') or [child.get('source_prefix')] or [''])[0]
            child_name = str(child.get('name') or '').strip() or catalog_names.get(fallback_prefix) or f"SUBFAMILIA {fallback_prefix}"
            ws_inf.cell(info_row, 3, child['report_code'])
            ws_inf.cell(info_row, 4, child_name)
            child_d = sum(child_val_list, _DEC_ZERO).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
            group_child_decs.append(child_d)
            if formula_refs:
                child_value_cell = ws_inf.cell(info_row, 5, f"=SUM({','.join(formula_refs)})")
            else:
                child_value_cell = ws_inf.cell(info_row, 5, 0.0)
            child_value_cell.number_format = '"$"#,##0.00'
            info_row += 1

        parent_d = sum(group_child_decs, _DEC_ZERO).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
        parent_dec_totals.append(parent_d)
        child_start_row = parent_row + 1
        child_end_row = info_row - 1
        if child_end_row >= child_start_row:
            parent_value_cell = ws_inf.cell(parent_row, 5, f"=SUM(E{child_start_row}:E{child_end_row})")
        else:
            parent_value_cell = ws_inf.cell(parent_row, 5, 0.0)
        parent_value_cell.number_format = '"$"#,##0.00'
        parent_value_cell.font = Font(bold=True)

    # Primer subtotal (solo familias inventariadas)
    for c in range(3, 8):
        ws_inf.cell(31, c, None)
    ws_inf.cell(31, 4, 'SUBTOTAL').font = Font(bold=True)
    subtotal_d = sum(parent_dec_totals, _DEC_ZERO).quantize(_DEC_TWO, rounding=ROUND_HALF_UP)
    if parent_rows_written:
        subtotal_refs = [f"E{r}" for r in parent_rows_written]
        subtotal_cell = ws_inf.cell(31, 5, f"=SUM({','.join(subtotal_refs)})")
    else:
        subtotal_cell = ws_inf.cell(31, 5, 0.0)
    subtotal_cell.font = Font(bold=True)
    subtotal_cell.number_format = '"$"#,##0.00'
    ws_inf.cell(31, 6, None)
    ws_inf.cell(31, 7, None)

    # Mantener estructura visual pero sin valores numericos en seccion contable fija
    for r in range(34, 44):
        for c in range(5, 8):
            ws_inf.cell(r, c, None)
    for r in [34, 35, 36, 38, 41]:
        ws_inf.cell(r, 5, '')

    # Limpieza defensiva: borrar valores/formulas de filas no requeridas,
    # aunque la plantilla cambie de posicion en futuras versiones.
    blocked_labels = {
        'TERRENO',
        'MUEBLES EN BODEGA',
        'EDIFICACIONES',
        'DEPRECIACION ACUMULADA',
        'TOTAL PROPIEDAD PLANTA Y EQUIPO',
    }
    blocked_codes = {'1605', '1635', '1640', '1685'}
    subtotal_rows = []
    for r in range(1, ws_inf.max_row + 1):
        code_txt = str(ws_inf.cell(r, 3).value or '').strip()
        label_txt = str(ws_inf.cell(r, 4).value or '').strip().upper()
        if label_txt == 'SUBTOTAL':
            subtotal_rows.append(r)
        if label_txt in blocked_labels or code_txt in blocked_codes:
            for c in range(5, 8):
                ws_inf.cell(r, c, None)

    # Mantener solo el primer SUBTOTAL (familias). Limpiar cualquier subtotal adicional.
    if subtotal_rows:
        keep_row = min(subtotal_rows)
        for r in subtotal_rows:
            if r != keep_row:
                for c in range(5, 8):
                    ws_inf.cell(r, c, None)

    # Forzar celdas sin formula en seccion contable fija.
    for r in [34, 35, 36, 38, 41]:
        ws_inf.cell(r, 5, '')

    # Sin proteccion de hojas: el usuario puede editar libremente.
    ws_des.protection.sheet = False
    ws_inf.protection.sheet = False
    ws_base.protection.sheet = False

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.calcMode = 'auto'

    app.logger.warning(f'[PERF] INFORME hoja procesada: {_time.perf_counter()-_t3:.2f}s')
    _t4 = _time.perf_counter()

    out = BytesIO()
    wb.save(out)
    content = out.getvalue()
    safe_period = sanitize_filename(period_label.replace(' ', '_'))
    filename = f'informe_conciliacion_activos_fijos_contabilidad_{safe_period}.xlsx'

    app.logger.warning(f'[PERF] wb.save (serializar xlsx): {_time.perf_counter()-_t4:.2f}s ({len(content)//1024} KB)')
    _t5 = _time.perf_counter()

    persist_accounting_report_file(
        content,
        filename,
        period_label,
        month,
        year,
        report_title,
        period_id=None,
        accounting_base_id=accounting_base.id,
        overrides_summary=overrides_summary.get('rows', []),
    )
    app.logger.warning(f'[PERF] persist_accounting_report_file: {_time.perf_counter()-_t5:.2f}s')
    app.logger.warning(f'[PERF] TOTAL generacion informe: {_time.perf_counter()-_t0:.2f}s')

    with ACCOUNTING_CACHE_LOCK:
        ACCOUNTING_REPORT_CACHE['version'] = current_cache_key
        ACCOUNTING_REPORT_CACHE['algo_version'] = ACCOUNTING_REPORT_ALGO_VERSION
        ACCOUNTING_REPORT_CACHE['bytes'] = content
        ACCOUNTING_REPORT_CACHE['filename'] = filename

    return send_file(
        BytesIO(content),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/reports/accounting_monthly_bases/upload', methods=['POST'])
def upload_accounting_monthly_base():
    ensure_db()
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Debes adjuntar el archivo base mensual'}), 400
    month, year = normalize_month_year(request.form.get('month'), request.form.get('year'))
    uploaded_by = str(request.form.get('uploaded_by') or '').strip()
    period_label = f"{MONTH_LABELS_ES.get(month, str(month))} {year}"

    source_name = sanitize_filename(os.path.basename(str(f.filename or '').strip() or f'base_{year}_{month:02d}.xlsx'))
    raw_content = f.read()
    if not raw_content:
        return jsonify({'error': 'El archivo base mensual esta vacio'}), 400

    try:
        if source_name.lower().endswith('.csv'):
            df = pd.read_csv(BytesIO(raw_content))
        else:
            df = pd.read_excel(BytesIO(raw_content))
        parsed_rows = extract_accounting_base_rows(df)
    except Exception as exc:
        return jsonify({'error': f'Error leyendo archivo base mensual: {exc}'}), 400

    if not parsed_rows:
        return jsonify({'error': 'La base mensual no contiene activos validos'}), 400

    unique_rows = {}
    for row in parsed_rows:
        unique_rows[row['c_act']] = row
    rows_to_insert = list(unique_rows.values())

    folder = accounting_base_folder(year, month)
    os.makedirs(folder, exist_ok=True)
    stamped_name = sanitize_filename(f"{os.path.splitext(source_name)[0]}_{now_local_dt().strftime('%Y%m%d%H%M%S')}{os.path.splitext(source_name)[1]}")
    file_path = os.path.join(folder, stamped_name)
    with open(file_path, 'wb') as out:
        out.write(raw_content)

    base_row = AccountingMonthlyBase(
        period_year=year,
        period_month=month,
        period_label=period_label,
        source_file_name=stamped_name,
        source_file_path=file_path,
        uploaded_by=uploaded_by,
        uploaded_at=now_iso(),
        asset_count=len(rows_to_insert),
        status='active',
    )
    db.session.add(base_row)
    db.session.flush()

    for row in rows_to_insert:
        db.session.add(AccountingMonthlyBaseAsset(
            base_id=base_row.id,
            c_act=row['c_act'],
            c_fam=row.get('c_fam'),
            nom_fam=row.get('nom_fam'),
            costo=row.get('costo'),
            saldo=row.get('saldo'),
            raw_row_json=row.get('raw_row_json'),
        ))
    db.session.commit()
    invalidate_accounting_report_cache()

    return jsonify({
        'base': base_row.to_dict(),
        'parsed_rows': len(parsed_rows),
        'deduplicated_rows': len(rows_to_insert),
    })


@app.route('/reports/accounting_monthly_bases', methods=['GET'])
def list_accounting_monthly_bases():
    ensure_db()
    month_raw = request.args.get('month')
    year_raw = request.args.get('year')
    q = AccountingMonthlyBase.query
    if str(month_raw or '').strip() and str(year_raw or '').strip():
        month, year = normalize_month_year(month_raw, year_raw)
        q = q.filter_by(period_month=month, period_year=year)
    rows = q.order_by(
        AccountingMonthlyBase.period_year.desc(),
        AccountingMonthlyBase.period_month.desc(),
        AccountingMonthlyBase.id.desc(),
    ).limit(300).all()
    return jsonify({'items': [r.to_dict() for r in rows]})


@app.route('/reports/accounting_monthly_bases/<int:base_id>/download', methods=['GET'])
def download_accounting_monthly_base(base_id):
    ensure_db()
    row = AccountingMonthlyBase.query.get(base_id)
    if not row:
        return jsonify({'error': 'Base mensual no encontrada'}), 404
    if not row.source_file_path or not os.path.exists(row.source_file_path):
        return jsonify({'error': 'El archivo de base mensual no existe en almacenamiento'}), 404
    return send_file(
        row.source_file_path,
        as_attachment=True,
        download_name=row.source_file_name or os.path.basename(row.source_file_path),
    )


@app.route('/reports/accounting_monthly_overrides_summary', methods=['GET'])
def accounting_monthly_overrides_summary():
    ensure_db()
    month, year = normalize_month_year(request.args.get('month'), request.args.get('year'))
    base_row = get_latest_accounting_base(month, year)
    if not base_row:
        return jsonify({'error': f'No hay base mensual cargada para {MONTH_LABELS_ES.get(month, month)} {year}'}), 404

    asset_rows = db.session.query(
        AccountingMonthlyBaseAsset.raw_row_json,
        AccountingMonthlyBaseAsset.c_act,
        AccountingMonthlyBaseAsset.c_fam,
        AccountingMonthlyBaseAsset.nom_fam,
        AccountingMonthlyBaseAsset.costo,
        AccountingMonthlyBaseAsset.saldo,
    ).filter_by(base_id=base_row.id).all()
    selected_overrides = get_accounting_cost_overrides(month, year)
    summary = build_overrides_summary_for_rows(asset_rows, selected_overrides)
    return jsonify({
        'period': {'month': month, 'year': year, 'label': base_row.period_label},
        'base': base_row.to_dict(),
        'summary': summary,
    })


@app.route('/reports/accounting_monthly_history', methods=['GET'])
def accounting_monthly_history():
    ensure_db()
    rows = GeneratedReport.query.filter_by(report_type='accounting_monthly').order_by(GeneratedReport.id.desc()).limit(200).all()
    base_ids = sorted({r.accounting_base_id for r in rows if r.accounting_base_id})
    base_map = {}
    if base_ids:
        for base_row in AccountingMonthlyBase.query.filter(AccountingMonthlyBase.id.in_(base_ids)).all():
            base_map[base_row.id] = base_row.to_dict()
    items = []
    for row in rows:
        payload = row.to_dict()
        payload['accounting_base'] = base_map.get(row.accounting_base_id)
        items.append(payload)
    return jsonify({'items': items})


@app.route('/reports/accounting_monthly_history/<int:report_id>/download', methods=['GET'])
def accounting_monthly_history_download(report_id):
    ensure_db()
    row = GeneratedReport.query.filter_by(id=report_id, report_type='accounting_monthly').first()
    if not row:
        return jsonify({'error': 'Informe no encontrado'}), 404
    if not row.file_path or not os.path.exists(row.file_path):
        return jsonify({'error': 'El archivo no existe en almacenamiento'}), 404
    return send_file(
        row.file_path,
        as_attachment=True,
        download_name=row.file_name or os.path.basename(row.file_path),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/reports/a22_history', methods=['GET'])
def a22_history():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    q = GeneratedReport.query.filter(
        GeneratedReport.report_type.in_(['a22_excel', 'a22_pdf'])
    )
    if period_id:
        q = q.filter(GeneratedReport.period_id == period_id)
    rows = q.order_by(GeneratedReport.id.desc()).limit(300).all()
    return jsonify({'items': [r.to_dict() for r in rows]})


@app.route('/reports/a22_history/<int:report_id>/download', methods=['GET'])
def a22_history_download(report_id):
    ensure_db()
    row = GeneratedReport.query.filter(
        GeneratedReport.id == report_id,
        GeneratedReport.report_type.in_(['a22_excel', 'a22_pdf'])
    ).first()
    if not row:
        return jsonify({'error': 'Informe A22 no encontrado'}), 404
    if not row.file_path or not os.path.exists(row.file_path):
        return jsonify({'error': 'El archivo no existe en almacenamiento'}), 404
    ext = os.path.splitext(row.file_name or row.file_path)[1].lower()
    mime = 'application/pdf' if ext == '.pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return send_file(
        row.file_path,
        as_attachment=True,
        download_name=row.file_name or os.path.basename(row.file_path),
        mimetype=mime
    )
                                                              

