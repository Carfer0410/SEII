from .inventory_periods_issues import *


def _iter_chunks(items, size=900):
    for i in range(0, len(items), size):
        yield items[i:i + size]


def count_new_assets_for_run(run):
    q = apply_run_scope_filter(Asset.query, run)
    asset_ids = [a.id for a in q.all()]
    if not asset_ids:
        return 0
    managed_ids = set()
    for chunk in _iter_chunks(asset_ids):
        rows = db.session.query(RunAssetStatus.asset_id).filter(
            RunAssetStatus.run_id == run.id,
            RunAssetStatus.asset_id.in_(chunk),
        ).all()
        managed_ids.update(int(aid) for (aid,) in rows)
    return sum(1 for aid in asset_ids if aid not in managed_ids)


@app.route('/assets')
def assets():
    ensure_db()
    service = request.args.get('service')
    run_id = request.args.get('run_id', type=int)
    q = Asset.query

    run = None
    if run_id:
        run = InventoryRun.query.get(run_id)
        if not run:
            return jsonify({'error': 'Jornada no encontrada'}), 404
        q = apply_run_scope_filter(q, run)

    if service:
        q = q.filter(Asset.nom_ccos == service)

    assets_list = q.limit(5000).all()
    items = [a.to_dict() for a in assets_list]

    ids = [a.id for a in assets_list]
    status_by_asset = {}
    if run and ids:
        statuses = RunAssetStatus.query.filter(
            RunAssetStatus.run_id == run.id,
            RunAssetStatus.asset_id.in_(ids)
        ).all()
        status_by_asset = {s.asset_id: s.status for s in statuses}

    for i, asset in enumerate(assets_list):
        items[i]['estado_jornada'] = status_by_asset.get(asset.id, '')
        items[i]['gestionado_jornada'] = asset.id in status_by_asset

    if assets_list:
        disposal_rows = AssetDisposal.query.filter(
            AssetDisposal.asset_id.in_([a.id for a in assets_list])
        ).all()
        disposal_by_asset = {d.asset_id: d.status for d in disposal_rows}
        for i, asset in enumerate(assets_list):
            items[i]['estado_baja'] = disposal_by_asset.get(asset.id, '')

    return jsonify({'assets': items})


@app.route('/assets/<int:asset_id>/classification', methods=['PATCH'])
def update_asset_classification(asset_id):
    ensure_db()
    asset = Asset.query.get(asset_id)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    data = request.get_json() or {}
    classification = (data.get('classification') or '').strip()
    disposal_reason = (data.get('disposal_reason') or '').strip()
    period_id = parse_int(data.get('period_id'))
    run_id = parse_int(data.get('run_id'))
    user = get_actor_username((data.get('user') or '').strip() or 'usuario_movil')
    if not classification:
        return jsonify({'error': 'Debe enviar la clasificacion'}), 400

    allowed = {
        'Pendiente verificacion',
        'En mantenimiento',
        'Prestado',
        'Activo de control',
        'Para baja',
        'Baja aprobada',
    }
    if classification not in allowed:
        return jsonify({'error': 'Clasificacion invalida'}), 400
    if classification in {'Para baja', 'Baja aprobada'} and not disposal_reason:
        return jsonify({'error': 'Debes registrar el motivo real de baja'}), 400

    now_iso_value = now_iso()
    resolved_period_id = None
    if run_id:
        run = InventoryRun.query.get(run_id)
        if run and run.period_id:
            resolved_period_id = run.period_id
    if resolved_period_id is None and period_id:
        period = InventoryPeriod.query.get(period_id)
        if not period:
            return jsonify({'error': 'Periodo no encontrado'}), 404
        resolved_period_id = period.id
    if resolved_period_id is None:
        resolved_period_id = get_or_create_default_period().id

    asset.estado_inventario = classification
    asset.fecha_verificacion = now_iso_value
    asset.usuario_verificador = user

    if classification in {'Para baja', 'Baja aprobada'}:
        disposal = AssetDisposal.query.filter_by(asset_id=asset.id).first()
        if not disposal:
            disposal = AssetDisposal(
                asset_id=asset.id,
                period_id=resolved_period_id,
                status='Pendiente baja',
                reason=disposal_reason,
                requested_by=user,
                requested_at=now_iso_value,
            )
            db.session.add(disposal)
        disposal.period_id = resolved_period_id
        if classification == 'Para baja':
            disposal.status = 'Pendiente baja'
            disposal.reason = disposal_reason
            disposal.requested_by = user
            disposal.requested_at = now_iso_value
            disposal.reviewed_by = None
            disposal.reviewed_at = None
            disposal.review_notes = None
        if classification == 'Baja aprobada':
            disposal.status = 'Aprobada para baja'
            disposal.reason = disposal_reason
            disposal.reviewed_by = user
            disposal.reviewed_at = now_iso_value
            disposal.review_notes = disposal.review_notes or 'Aprobada desde inventario'

    refresh_asset_type_cache(asset)
    db.session.commit()
    return jsonify({'ok': True, 'asset': asset.to_dict(), 'classification': classification})


@app.route('/assets/<int:asset_id>/service', methods=['PATCH'])
def update_asset_service(asset_id):
    ensure_db()
    asset = Asset.query.get(asset_id)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    data = request.get_json() or {}
    service = (data.get('service') or '').strip()
    user = get_actor_username((data.get('user') or '').strip() or 'usuario_movil')
    run_id = data.get('run_id')
    period_id = data.get('period_id')
    auto_transfer = parse_bool(data.get('auto_transfer'), default=False)
    sync_location = parse_bool(data.get('sync_location'), default=False)
    transfer_reason = (data.get('transfer_reason') or '').strip()
    transfer_notes = (data.get('transfer_notes') or '').strip()

    if not service:
        return jsonify({'error': 'Debe enviar el servicio destino'}), 400

    run = None
    if run_id is not None:
        run = InventoryRun.query.get(run_id)
        if not run:
            return jsonify({'error': 'Jornada no encontrada'}), 404
        if run.status != 'active':
            return jsonify({'error': 'La jornada ya esta cerrada'}), 400
        run_scope = run_scope_services(run)
        run_scope_cf = {s.casefold() for s in run_scope}
        if run_scope and service.casefold() not in run_scope_cf:
            return jsonify({'error': 'El servicio destino debe estar dentro del alcance de la jornada activa'}), 400

    now_iso_value = now_iso()
    old_service = str(asset.nom_ccos or '').strip()
    old_location = str(asset.des_ubi or '').strip()
    old_responsible = str(asset.nom_resp or '').strip()
    asset.nom_ccos = service
    if sync_location:
        asset.des_ubi = service
    asset.fecha_verificacion = now_iso_value
    asset.usuario_verificador = user

    transfer_payload = None
    if auto_transfer:
        resolved_period_id = None
        if run and run.period_id:
            resolved_period_id = run.period_id
        if resolved_period_id is None and period_id not in (None, ''):
            try:
                resolved_period_id = int(period_id)
            except Exception:
                resolved_period_id = None

        linked_issue = None
        if resolved_period_id:
            linked_issue = AssetIssue.query.filter(
                AssetIssue.asset_id == asset.id,
                AssetIssue.period_id == resolved_period_id,
                AssetIssue.issue_type.in_(['SCANNED_OTHER_SERVICE', 'LOCATION_REVIEW', 'RESPONSIBLE_REVIEW'])
            ).order_by(AssetIssue.id.desc()).first()

        justification = transfer_reason or (
            f"Traslado automatico por escaneo fuera de alcance de jornada. Servicio anterior: '{old_service or 'N/D'}'."
        )
        execution_notes = transfer_notes or (
            f"Regularizacion automatica desde modulo Inventario. "
            f"Ubicacion {'sincronizada' if sync_location else 'no sincronizada'} a '{service}'."
        )
        transfer_row = AssetTransferCase(
            issue_id=linked_issue.id if linked_issue else None,
            asset_id=asset.id,
            period_id=resolved_period_id,
            run_id=run.id if run else None,
            status='Ejecutado',
            origin_service=old_service,
            target_service=service,
            origin_responsible=old_responsible,
            target_responsible=str(asset.nom_resp or '').strip(),
            justification=justification,
            requested_by=user,
            requested_at=now_iso_value,
            approved_by=user,
            approved_at=now_iso_value,
            approval_notes='Aprobacion automatica por regularizacion operativa de inventario.',
            executed_by=user,
            executed_at=now_iso_value,
            execution_notes=execution_notes,
            created_at=now_iso_value,
            updated_at=now_iso_value,
        )
        db.session.add(transfer_row)
        db.session.flush()

        if resolved_period_id:
            open_related_issues = AssetIssue.query.filter(
                AssetIssue.asset_id == asset.id,
                AssetIssue.period_id == resolved_period_id,
                AssetIssue.issue_type.in_(['SCANNED_OTHER_SERVICE', 'LOCATION_REVIEW', 'RESPONSIBLE_REVIEW']),
                AssetIssue.status != 'Cerrado'
            ).all()
            for issue_row in open_related_issues:
                issue_row.status = 'Cerrado'
                issue_row.resolution_notes = append_text_note(
                    issue_row.resolution_notes,
                    f"Cierre automatico por traslado ejecutado a '{service}'.",
                )
                issue_row.updated_at = now_iso_value

        pdf_bytes = build_transfer_acta_pdf_bytes(transfer_row, asset, issue=linked_issue)
        timestamp = now_local_dt().strftime('%Y%m%d%H%M%S')
        public_name = f"acta_traslado_{clean_filename(asset.c_act)}_{timestamp}.pdf"
        storage_name = f"{clean_filename(os.path.splitext(public_name)[0])}_{timestamp}.pdf"
        file_path = os.path.join(DOCUMENTS_DIR, storage_name)
        with open(file_path, 'wb') as fp:
            fp.write(pdf_bytes)

        doc_row = DocumentRecord(
            link_type='asset',
            asset_id=asset.id,
            asset_code=asset.c_act or '',
            asset_name=asset.nom or '',
            document_type='Novedad',
            title=f'Acta de traslado {asset.c_act or ""}',
            description=(
                f"Traslado automatico por escaneo fuera de alcance. "
                f"Servicio: {old_service or 'N/D'} -> {service or 'N/D'}. "
                f"Ubicacion: {old_location or 'N/D'} -> {asset.des_ubi or 'N/D'}."
            ),
            doc_date=now_local_dt().strftime('%Y-%m-%d'),
            area_service=service,
            radicado=f'TR-{transfer_row.id:06d}',
            file_name=public_name,
            file_path=file_path,
            file_ext='.pdf',
            file_size=len(pdf_bytes),
            uploaded_by=user,
            uploaded_at=now_iso_value,
            status='active',
        )
        db.session.add(doc_row)
        db.session.flush()
        transfer_row.acta_doc_id = doc_row.id
        transfer_row.acta_file_path = file_path
        transfer_payload = transfer_row.to_dict()

    db.session.commit()
    return jsonify({
        'ok': True,
        'asset': asset.to_dict(),
        'old_service': old_service,
        'new_service': service,
        'old_location': old_location,
        'new_location': str(asset.des_ubi or '').strip(),
        'location_synced': bool(sync_location),
        'transfer': transfer_payload,
    })


@app.route('/scan', methods=['POST'])
def scan():
    ensure_db()
    data = request.get_json() or {}
    code = data.get('code')
    user = get_actor_username(data.get('user') or 'unknown')
    run_id = data.get('run_id')
    if not code:
        return jsonify({'error': 'No code provided'}), 400
    scanned_code = normalize_scan_code(code)
    try:
        asset, matched_by = get_asset_by_code(scanned_code)
    except Exception as exc:
        app.logger.exception('[SCAN] Error en get_asset_by_code para "%s": %s', scanned_code, exc)
        asset = get_asset_by_c_act_strict(scanned_code)
        matched_by = 'C_ACT' if asset else None
    if not asset:
        return jsonify({'found': False, 'scanned_code': scanned_code}), 200

    if run_id is None:
        return jsonify({'error': 'Debes iniciar una jornada activa para escanear'}), 400

    run = InventoryRun.query.get(run_id)
    if not run:
        return jsonify({'error': 'Jornada no encontrada'}), 404
    if run.status != 'active':
        return jsonify({'error': 'La jornada ya esta cerrada'}), 400
    run_scope = run_scope_services(run)
    run_service = str(run.service or '').strip()
    asset_service = str(asset.nom_ccos or '').strip()
    run_scope_cf = {s.casefold() for s in run_scope}
    if run_scope and asset_service.casefold() not in run_scope_cf:
        expected_label = ', '.join(run_scope[:3]) + (' ...' if len(run_scope) > 3 else '')
        return jsonify({
            'error': f'Escaneado fuera del alcance de la jornada. Base actual: {asset_service or "sin servicio"} | Alcance: {expected_label}',
            'code': 'SERVICE_MISMATCH',
            'expected_service': run_service or (run_scope[0] if run_scope else ''),
            'expected_services': run_scope,
            'expected_service_label': expected_label,
            'current_service': asset_service,
            'matched_by': matched_by or 'C_ACT',
            'scanned_code': scanned_code,
            'asset': asset.to_dict(),
            'run': run.to_dict(),
        }), 409

    run_status = RunAssetStatus.query.filter_by(run_id=run.id, asset_id=asset.id).first()
    if not run_status:
        run_status = RunAssetStatus(
            run_id=run.id,
            asset_id=asset.id,
            status='Encontrado',
            scanned_at=now_iso(),
            scanned_by=user,
        )
        db.session.add(run_status)
    else:
        run_status.status = 'Encontrado'
        run_status.scanned_at = now_iso()
        run_status.scanned_by = user

    asset.estado_inventario = 'Encontrado'
    asset.fecha_verificacion = now_iso()
    asset.usuario_verificador = user
    db.session.commit()
    return jsonify({
        'found': True,
        'asset': asset.to_dict(),
        'run_id': run.id if run else None,
        'matched_by': matched_by or 'C_ACT',
        'scanned_code': scanned_code,
    })


@app.route('/runs', methods=['GET'])
def list_runs():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    status = (request.args.get('status') or '').strip().lower()
    q = InventoryRun.query
    if period_id:
        q = q.filter(InventoryRun.period_id == period_id)
    if status in {'active', 'closed', 'cancelled'}:
        q = q.filter(InventoryRun.status == status)
    runs = q.order_by(InventoryRun.id.desc()).limit(300).all()
    run_ids = [r.id for r in runs]
    period_ids = sorted({r.period_id for r in runs if r.period_id})
    periods_map = {}
    if period_ids:
        periods_map = {p.id: p for p in InventoryPeriod.query.filter(InventoryPeriod.id.in_(period_ids)).all()}
    found_by_run = {}
    not_found_by_run = {}
    if run_ids:
        statuses = db.session.query(
            RunAssetStatus.run_id,
            RunAssetStatus.status,
            db.func.count(RunAssetStatus.id)
        ).filter(
            RunAssetStatus.run_id.in_(run_ids)
        ).group_by(
            RunAssetStatus.run_id,
            RunAssetStatus.status
        ).all()
        for run_id, status, count in statuses:
            if status == 'Encontrado':
                found_by_run[run_id] = int(count or 0)
            elif status == 'No encontrado':
                not_found_by_run[run_id] = int(count or 0)

    payload = []
    for r in runs:
        row = r.to_dict()
        period = periods_map.get(r.period_id)
        row['period_name'] = period.name if period else None
        row['period_status'] = period.status if period else None
        row['found'] = found_by_run.get(r.id, 0)
        row['not_found'] = not_found_by_run.get(r.id, 0)
        if r.status == 'closed':
            new_assets = count_new_assets_for_run(r)
            row['new_assets_in_scope'] = new_assets
            row['can_reopen'] = new_assets > 0
        else:
            row['new_assets_in_scope'] = 0
            row['can_reopen'] = False
        payload.append(row)
    return jsonify({'runs': payload})


@app.route('/disposals', methods=['GET'])
def list_disposals():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    rows = query_disposals(service=service, status=status, period_id=period_id)
    items = []
    for row in rows:
        items.append({
            'id': row['id'],
            'period_id': row.get('period_id'),
            'reason': row['reason'],
            'status': row['status'],
            'asset': {
                'C_ACT': row['code'],
                'NOM': row['name'],
                'NOM_CCOS': row['service'],
                'TIPO_ACTIVO': row['type'],
                'COSTO': row['cost'],
                'SALDO': row['saldo'],
                'FECHA_COMPRA': row['date'],
            }
        })
    return jsonify({'disposals': items})


@app.route('/disposals/export_excel', methods=['GET'])
def export_disposals_excel():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    type_exact = (request.args.get('type_exact') or '').strip()
    type_key = normalize_disposal_type_key(request.args.get('type'))
    rows = query_disposals(service=service, status=status, period_id=period_id)
    if type_exact:
        rows = [r for r in rows if str(r.get('type', '')).strip().upper() == type_exact.upper()]
    elif type_key:
        rows = [r for r in rows if normalize_disposal_type_key(r.get('type')) == type_key]
    if not rows:
        return jsonify({'error': 'No hay activos para exportar con ese filtro'}), 400

    is_control_report = (type_exact and 'CONTROL' in type_exact.upper()) or (type_key == 'CONTROL')
    logo_path = get_hospital_logo_path()
    wb = Workbook()
    ws = wb.active
    title = f'Bajas {type_exact}' if type_exact else (f'Bajas {type_key}' if type_key else 'Bajas - Todos los tipos')
    write_disposal_sheet(
        ws,
        title,
        rows,
        saldo_header='SALDO CONTABLE (NO DEPRECIABLE)' if is_control_report else 'SALDO POR DEPRECIAR',
        note_text=(
            'Nota: los activos de control no se deprecian; por politica contable su saldo contable suele coincidir con el costo inicial.'
            if is_control_report else None
        ),
    )
    add_logo_to_excel_sheet(ws, logo_path=logo_path)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    filename_type = (type_exact or type_key or 'todos').lower().replace(' ', '_')
    filename = clean_filename(f"bajas_{filename_type}.xlsx")
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/disposals/export_pdf', methods=['GET'])
def export_disposals_pdf():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    type_exact = (request.args.get('type_exact') or '').strip()
    type_key = normalize_disposal_type_key(request.args.get('type'))
    rows = query_disposals(service=service, status=status, period_id=period_id)
    if type_exact:
        rows = [r for r in rows if str(r.get('type', '')).strip().upper() == type_exact.upper()]
    elif type_key:
        rows = [r for r in rows if normalize_disposal_type_key(r.get('type')) == type_key]
    if not rows:
        return jsonify({'error': 'No hay activos para exportar con ese filtro'}), 400

    is_control_report = (type_exact and 'CONTROL' in type_exact.upper()) or (type_key == 'CONTROL')
    summary = summarize_disposals(rows)
    logo_path = get_hospital_logo_path()
    out = BytesIO()
    doc = SimpleDocTemplate(
        out, pagesize=letter, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=22 * mm, bottomMargin=14 * mm
    )
    styles = getSampleStyleSheet()
    story = []
    report_title = f"Bajas - {type_exact}" if type_exact else (f"Bajas - {type_key}" if type_key else "Bajas - Todos los tipos")
    append_pdf_header_with_logo(
        story,
        report_title,
        f"Generado: {now_local_dt().strftime('%Y-%m-%d %H:%M')} | Servicio: {service or 'TODOS'} | Estado: {status or 'TODOS'}",
        include_logo=False,
    )
    story.append(Paragraph(
        f"<b>Total activos:</b> {summary['count']} &nbsp;&nbsp; <b>Total costo inicial:</b> {money_text(summary['total_cost'])} "
        f"&nbsp;&nbsp; <b>{'Total saldo contable' if is_control_report else 'Total saldo por depreciar'}:</b> {money_text(summary['total_saldo'])}",
        styles['Normal']
    ))
    if is_control_report:
        story.append(Paragraph(
            '<b>Nota:</b> Los activos de control no se deprecian; por politica contable su saldo contable suele coincidir con el costo inicial.',
            ParagraphStyle('CtrlNote', parent=styles['Normal'], textColor=colors.HexColor('#9A5F00'), fontSize=9)
        ))
    story.append(Spacer(1, 8))

    table_data = [[
        pdf_cell('COD ACTIVO FIJO', styles, bold=True, align='CENTER'),
        pdf_cell('DESCRIPCION', styles, bold=True, align='CENTER'),
        pdf_cell('COSTO INICIAL', styles, bold=True, align='CENTER'),
        pdf_cell('SALDO POR DEPRECIAR', styles, bold=True, align='CENTER'),
        pdf_cell('FECHA ADQUISICION', styles, bold=True, align='CENTER'),
        pdf_cell('MOTIVO DE BAJA', styles, bold=True, align='CENTER'),
    ]]
    for r in rows:
        table_data.append([
            pdf_cell(r['code'], styles, align='CENTER'),
            pdf_cell(r['name'], styles),
            pdf_cell(money_text(r['cost']), styles, align='RIGHT'),
            pdf_cell(money_text(r['saldo']), styles, align='RIGHT'),
            pdf_cell(r['date'], styles, align='CENTER'),
            pdf_cell(r['reason'], styles),
        ])
    table = Table(table_data, colWidths=[24*mm, 60*mm, 24*mm, 24*mm, 24*mm, 42*mm], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAF4FA')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0B4F6C')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#C8D8E4')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FCFF')]),
    ]))
    story.append(table)
    page_header = make_pdf_page_header(logo_path)
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    out.seek(0)
    filename_type = (type_exact or type_key or 'todos').lower().replace(' ', '_')
    filename = clean_filename(f"bajas_{filename_type}.pdf")
    return send_file(out, as_attachment=True, download_name=filename, mimetype='application/pdf')


@app.route('/disposals/export_general_excel', methods=['GET'])
def export_disposals_general_excel():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    all_rows = query_disposals(service=service, status=status, period_id=period_id)
    rows = [r for r in all_rows if normalize_disposal_type_key(r.get('type')) != 'CONTROL']
    if not rows:
        return jsonify({'error': 'No hay activos para exportar en reporte general'}), 400

    logo_path = get_hospital_logo_path()
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = 'Resumen General'

    grouped = {k: [] for k in DISPOSAL_TYPE_KEYS if k != 'CONTROL'}
    for r in rows:
        key = normalize_disposal_type_key(r.get('type'))
        if key in grouped:
            grouped[key].append(r)

    ws_summary.append(['REPORTE GENERAL DE BAJAS (SIN CONTROL)'])
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws_summary['A1'].font = Font(bold=True, size=14, color='0B4F6C')
    ws_summary.append([
        f"Generado: {now_local_dt().strftime('%Y-%m-%d %H:%M')} | Servicio: {service or 'TODOS'} | Estado: {status or 'TODOS'}"
    ])
    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    ws_summary.append(['TIPO', 'CANTIDAD', 'TOTAL COSTO INICIAL', 'TOTAL SALDO POR DEPRECIAR', '% PARTICIPACION'])

    total_cost = sum(r['cost'] for r in rows)
    total_saldo = sum(r['saldo'] for r in rows)
    total_count = len(rows)
    for t in ['BIOMEDICO', 'MUEBLE Y ENSER', 'INDUSTRIAL', 'TECNOLOGICO']:
        sub = grouped.get(t, [])
        sub_cost = sum(r['cost'] for r in sub)
        sub_saldo = sum(r['saldo'] for r in sub)
        pct = round((len(sub) / total_count) * 100, 2) if total_count else 0
        ws_summary.append([t, len(sub), sub_cost, sub_saldo, pct])

    ws_summary.append(['TOTAL GENERAL', total_count, total_cost, total_saldo, 100 if total_count else 0])
    for c in ['A', 'B', 'C', 'D', 'E']:
        ws_summary.column_dimensions[c].width = [26, 12, 22, 24, 16][ord(c) - ord('A')]
    for row in ws_summary.iter_rows(min_row=3, max_row=ws_summary.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.alignment = Alignment(vertical='center', horizontal='center')
    for r in range(4, ws_summary.max_row + 1):
        ws_summary.cell(r, 3).number_format = '"$"#,##0'
        ws_summary.cell(r, 4).number_format = '"$"#,##0'
        ws_summary.cell(r, 5).number_format = '0.00"%"'

    for t in ['BIOMEDICO', 'MUEBLE Y ENSER', 'INDUSTRIAL', 'TECNOLOGICO']:
        ws = wb.create_sheet(title=t[:31])
        write_disposal_sheet(ws, f'Bajas {t}', grouped.get(t, []))
        add_logo_to_excel_sheet(ws, logo_path=logo_path)
    add_logo_to_excel_sheet(ws_summary, logo_path=logo_path)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    filename = clean_filename('bajas_general_sin_control.xlsx')
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/disposals/export_general_pdf', methods=['GET'])
def export_disposals_general_pdf():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    all_rows = query_disposals(service=service, status=status, period_id=period_id)
    rows = [r for r in all_rows if normalize_disposal_type_key(r.get('type')) != 'CONTROL']
    if not rows:
        return jsonify({'error': 'No hay activos para exportar en reporte general'}), 400

    grouped = {k: [] for k in ['BIOMEDICO', 'MUEBLE Y ENSER', 'INDUSTRIAL', 'TECNOLOGICO']}
    for r in rows:
        key = normalize_disposal_type_key(r.get('type'))
        if key in grouped:
            grouped[key].append(r)

    total = summarize_disposals(rows)
    logo_path = get_hospital_logo_path()
    out = BytesIO()
    doc = SimpleDocTemplate(
        out, pagesize=letter, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=22 * mm, bottomMargin=14 * mm
    )
    styles = getSampleStyleSheet()
    story = []
    append_pdf_header_with_logo(
        story,
        'Reporte General de Bajas',
        f"Generado: {now_local_dt().strftime('%Y-%m-%d %H:%M')} | Servicio: {service or 'TODOS'} | Estado: {status or 'TODOS'}",
        include_logo=False,
    )
    story.append(Paragraph(
        f"<b>Total activos:</b> {total['count']} &nbsp;&nbsp; <b>Total costo inicial:</b> {money_text(total['total_cost'])} "
        f"&nbsp;&nbsp; <b>Total saldo por depreciar:</b> {money_text(total['total_saldo'])}",
        styles['Normal']
    ))
    story.append(Spacer(1, 8))

    res = [['TIPO', 'CANTIDAD', 'TOTAL COSTO', 'TOTAL SALDO']]
    for t in ['BIOMEDICO', 'MUEBLE Y ENSER', 'INDUSTRIAL', 'TECNOLOGICO']:
        sub = grouped[t]
        sum_sub = summarize_disposals(sub)
        res.append([t, sum_sub['count'], money_text(sum_sub['total_cost']), money_text(sum_sub['total_saldo'])])
    summary_table = Table(res, colWidths=[48*mm, 26*mm, 48*mm, 48*mm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAF4FA')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#C8D8E4')),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10))

    chart_d = Drawing(170 * mm, 45 * mm)
    chart = HorizontalBarChart()
    chart.x = 22
    chart.y = 4
    chart.width = 132 * mm
    chart.height = 34 * mm
    chart.data = [[len(grouped['BIOMEDICO']), len(grouped['MUEBLE Y ENSER']), len(grouped['INDUSTRIAL']), len(grouped['TECNOLOGICO'])]]
    chart.categoryAxis.categoryNames = ['Biomedico', 'Mueble y enser', 'Industrial', 'Tecnologico']
    chart.valueAxis.valueMin = 0
    chart.bars[0].fillColor = colors.HexColor('#1E88E5')
    chart_d.add(chart)
    chart_d.add(String(6, 42 * mm, 'Distribucion de activos por tipo', fontSize=9, fillColor=colors.HexColor('#0B4F6C')))
    story.append(chart_d)

    for t in ['BIOMEDICO', 'MUEBLE Y ENSER', 'INDUSTRIAL', 'TECNOLOGICO']:
        sub = grouped[t]
        if not sub:
            continue
        story.append(PageBreak())
        story.append(Paragraph(f'Detalle {t}', ParagraphStyle(
            'Sec', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#0B4F6C')
        )))
        sum_sub = summarize_disposals(sub)
        story.append(Paragraph(
            f"<b>Activos:</b> {sum_sub['count']} &nbsp;&nbsp; <b>Costo:</b> {money_text(sum_sub['total_cost'])} "
            f"&nbsp;&nbsp; <b>Saldo:</b> {money_text(sum_sub['total_saldo'])}",
            styles['Normal']
        ))
        story.append(Spacer(1, 6))
        detail = [[
            pdf_cell('COD ACTIVO FIJO', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('DESCRIPCION', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('COSTO INICIAL', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('SALDO POR DEPRECIAR', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('FECHA ADQUISICION', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('MOTIVO DE BAJA', styles, bold=True, align='CENTER', size=7.5),
        ]]
        for r in sub:
            detail.append([
                pdf_cell(r['code'], styles, align='CENTER'),
                pdf_cell(r['name'], styles),
                pdf_cell(money_text(r['cost']), styles, align='RIGHT'),
                pdf_cell(money_text(r['saldo']), styles, align='RIGHT'),
                pdf_cell(r['date'], styles, align='CENTER'),
                pdf_cell(r['reason'], styles),
            ])
        t_detail = Table(detail, colWidths=[16*mm, 58*mm, 22*mm, 22*mm, 18*mm, 48*mm], repeatRows=1)
        t_detail.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F3F9FD')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D4E2EC')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(t_detail)

    page_header = make_pdf_page_header(logo_path)
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    out.seek(0)
    filename = clean_filename('bajas_general_sin_control.pdf')
    return send_file(out, as_attachment=True, download_name=filename, mimetype='application/pdf')


@app.route('/disposals/export_general_control_excel', methods=['GET'])
def export_disposals_general_control_excel():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    all_rows = query_disposals(service=service, status=status, period_id=period_id)
    rows = [r for r in all_rows if normalize_disposal_type_key(r.get('type')) == 'CONTROL']

    grouped = {}
    for r in rows:
        key = str(r.get('type') or 'CONTROL - OTROS').strip().upper()
        grouped.setdefault(key, []).append(r)

    logo_path = get_hospital_logo_path()
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = 'Resumen Control'
    ws_summary.append(['REPORTE GENERAL DE BAJAS - ACTIVOS DE CONTROL'])
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws_summary['A1'].font = Font(bold=True, size=14, color='9A5F00')
    ws_summary.append([
        f"Generado: {now_local_dt().strftime('%Y-%m-%d %H:%M')} | Servicio: {service or 'TODOS'} | Estado: {status or 'TODOS'}"
    ])
    ws_summary.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    ws_summary.append([
        'Nota: los activos de control no se deprecian; su saldo contable suele coincidir con el costo inicial.'
    ])
    ws_summary.merge_cells(start_row=3, start_column=1, end_row=3, end_column=5)
    ws_summary['A3'].font = Font(bold=True, color='9A5F00')
    ws_summary.append(['SUBTIPO CONTROL', 'CANTIDAD', 'TOTAL COSTO INICIAL', 'TOTAL SALDO CONTABLE', '% PARTICIPACION'])

    total_cost = sum(r['cost'] for r in rows)
    total_saldo = sum(r['saldo'] for r in rows)
    total_count = len(rows)
    for key in sorted(grouped.keys()):
        sub = grouped.get(key, [])
        sub_cost = sum(r['cost'] for r in sub)
        sub_saldo = sum(r['saldo'] for r in sub)
        pct = round((len(sub) / total_count) * 100, 2) if total_count else 0
        ws_summary.append([key, len(sub), sub_cost, sub_saldo, pct])

    ws_summary.append(['TOTAL GENERAL CONTROL', total_count, total_cost, total_saldo, 100 if total_count else 0])
    for c in ['A', 'B', 'C', 'D', 'E']:
        ws_summary.column_dimensions[c].width = [34, 12, 22, 24, 16][ord(c) - ord('A')]
    for row in ws_summary.iter_rows(min_row=4, max_row=ws_summary.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.alignment = Alignment(vertical='center', horizontal='center')
    for r in range(5, ws_summary.max_row + 1):
        ws_summary.cell(r, 3).number_format = '"$"#,##0'
        ws_summary.cell(r, 4).number_format = '"$"#,##0'
        ws_summary.cell(r, 5).number_format = '0.00"%"'

    for key in sorted(grouped.keys()):
        ws = wb.create_sheet(title=key[:31])
        write_disposal_sheet(
            ws,
            f'Bajas {key}',
            grouped.get(key, []),
            saldo_header='SALDO CONTABLE (NO DEPRECIABLE)',
            note_text='Nota: activo de control no depreciable (vida util 0 o marcado como no depreciable).',
        )
        add_logo_to_excel_sheet(ws, logo_path=logo_path)
    add_logo_to_excel_sheet(ws_summary, logo_path=logo_path)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    filename = clean_filename('bajas_general_control.xlsx')
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/disposals/export_general_control_pdf', methods=['GET'])
def export_disposals_general_control_pdf():
    ensure_db()
    service = request.args.get('service')
    status = request.args.get('status')
    period_id = request.args.get('period_id', type=int)
    all_rows = query_disposals(service=service, status=status, period_id=period_id)
    rows = [r for r in all_rows if normalize_disposal_type_key(r.get('type')) == 'CONTROL']

    grouped = {}
    for r in rows:
        key = str(r.get('type') or 'CONTROL - OTROS').strip().upper()
        grouped.setdefault(key, []).append(r)

    total = summarize_disposals(rows)
    logo_path = get_hospital_logo_path()
    out = BytesIO()
    doc = SimpleDocTemplate(
        out, pagesize=letter, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=22 * mm, bottomMargin=14 * mm
    )
    styles = getSampleStyleSheet()
    story = []
    append_pdf_header_with_logo(
        story,
        'Reporte General de Bajas - Activos de Control',
        f"Generado: {now_local_dt().strftime('%Y-%m-%d %H:%M')} | Servicio: {service or 'TODOS'} | Estado: {status or 'TODOS'}",
        include_logo=False,
    )
    story.append(Paragraph(
        f"<b>Total activos control:</b> {total['count']} &nbsp;&nbsp; <b>Total costo inicial:</b> {money_text(total['total_cost'])} "
        f"&nbsp;&nbsp; <b>Total saldo contable:</b> {money_text(total['total_saldo'])}",
        styles['Normal']
    ))
    control_pct = round((len(rows) / len(all_rows)) * 100, 2) if all_rows else 0
    story.append(Paragraph(
        f"<b>Participacion de control sobre bajas filtradas:</b> {control_pct}%",
        styles['Normal']
    ))
    story.append(Paragraph(
        '<b>Nota:</b> Los activos de control no se deprecian; por politica contable su saldo contable suele coincidir con el costo inicial.',
        ParagraphStyle('CtrlGenNote', parent=styles['Normal'], textColor=colors.HexColor('#9A5F00'), fontSize=9)
    ))
    story.append(Spacer(1, 8))

    res = [['SUBTIPO CONTROL', 'CANTIDAD', 'TOTAL COSTO', 'TOTAL SALDO CONTABLE']]
    for key in sorted(grouped.keys()):
        sub = grouped[key]
        sum_sub = summarize_disposals(sub)
        res.append([key, sum_sub['count'], money_text(sum_sub['total_cost']), money_text(sum_sub['total_saldo'])])
    summary_table = Table(res, colWidths=[62*mm, 22*mm, 44*mm, 44*mm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FFF4DE')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#E2C998')),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
    ]))
    story.append(summary_table)
    if not rows:
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            'No hay activos de control para los filtros seleccionados. Se genera reporte en blanco para control documental.',
            ParagraphStyle('CtrlEmpty', parent=styles['Normal'], textColor=colors.HexColor('#7A4A00'), fontSize=9)
        ))
    story.append(Spacer(1, 10))

    chart_keys = sorted(grouped.keys())
    chart_vals = [len(grouped[k]) for k in chart_keys]
    chart_d = Drawing(170 * mm, 45 * mm)
    chart = HorizontalBarChart()
    chart.x = 22
    chart.y = 4
    chart.width = 132 * mm
    chart.height = 34 * mm
    chart.data = [chart_vals] if chart_vals else [[0]]
    chart.categoryAxis.categoryNames = [k.replace('CONTROL - ', '').title() for k in chart_keys] if chart_keys else ['Sin datos']
    chart.valueAxis.valueMin = 0
    chart.bars[0].fillColor = colors.HexColor('#D97706')
    chart_d.add(chart)
    chart_d.add(String(6, 42 * mm, 'Distribucion de activos de control por subtipo', fontSize=9, fillColor=colors.HexColor('#9A5F00')))
    story.append(chart_d)

    for key in sorted(grouped.keys()):
        sub = grouped[key]
        if not sub:
            continue
        story.append(PageBreak())
        story.append(Paragraph(f'Detalle {key}', ParagraphStyle(
            'SecCtrl', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#9A5F00')
        )))
        sum_sub = summarize_disposals(sub)
        story.append(Paragraph(
            f"<b>Activos:</b> {sum_sub['count']} &nbsp;&nbsp; <b>Costo:</b> {money_text(sum_sub['total_cost'])} "
            f"&nbsp;&nbsp; <b>Saldo:</b> {money_text(sum_sub['total_saldo'])}",
            styles['Normal']
        ))
        story.append(Spacer(1, 6))
        detail = [[
            pdf_cell('COD ACTIVO FIJO', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('DESCRIPCION', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('COSTO INICIAL', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('SALDO POR DEPRECIAR', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('FECHA ADQUISICION', styles, bold=True, align='CENTER', size=7.5),
            pdf_cell('MOTIVO DE BAJA', styles, bold=True, align='CENTER', size=7.5),
        ]]
        for r in sub:
            detail.append([
                pdf_cell(r['code'], styles, align='CENTER'),
                pdf_cell(r['name'], styles),
                pdf_cell(money_text(r['cost']), styles, align='RIGHT'),
                pdf_cell(money_text(r['saldo']), styles, align='RIGHT'),
                pdf_cell(r['date'], styles, align='CENTER'),
                pdf_cell(r['reason'], styles),
            ])
        t_detail = Table(detail, colWidths=[16*mm, 58*mm, 22*mm, 22*mm, 18*mm, 48*mm], repeatRows=1)
        t_detail.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FFF7E8')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#E4D1A9')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(t_detail)

    page_header = make_pdf_page_header(logo_path)
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    out.seek(0)
    filename = clean_filename('bajas_general_control.pdf')
    return send_file(out, as_attachment=True, download_name=filename, mimetype='application/pdf')


@app.route('/maintenance/reclassify', methods=['POST'])
def reclassify_assets():
    ensure_db()
    data = request.get_json() or {}
    service = (data.get('service') or '').strip()
    only_disposals = bool(data.get('only_disposals'))

    q = Asset.query
    if service:
        q = q.filter(Asset.nom_ccos == service)
    if only_disposals:
        q = q.join(AssetDisposal, AssetDisposal.asset_id == Asset.id)

    assets = q.all()
    updated = 0
    for asset in assets:
        old = (asset.tipo_activo_cache or '').strip()
        new = refresh_asset_type_cache(asset)
        if old != new:
            updated += 1

    db.session.commit()
    return jsonify({
        'ok': True,
        'total_processed': len(assets),
        'updated': updated,
        'service': service or None,
        'only_disposals': only_disposals,
        'executed_at': now_iso(),
    })


@app.route('/disposals', methods=['POST'])
def create_disposal():
    ensure_db()
    data = request.get_json() or {}
    code = data.get('code')
    reason = (data.get('reason') or '').strip()
    requested_by = get_actor_username((data.get('requested_by') or '').strip() or 'unknown')
    period_id = parse_int(data.get('period_id'))
    if not code:
        return jsonify({'error': 'Debe enviar codigo de activo'}), 400
    if not period_id:
        return jsonify({'error': 'Debes seleccionar un periodo para registrar la baja'}), 400

    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404

    asset, _ = get_asset_by_code(code)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    disposal = AssetDisposal.query.filter_by(asset_id=asset.id).first()
    now_iso_value = now_iso()
    if not disposal:
        disposal = AssetDisposal(
            asset_id=asset.id,
            period_id=period.id,
            status='Pendiente baja',
            reason=reason,
            requested_by=requested_by,
            requested_at=now_iso_value,
        )
        db.session.add(disposal)
    else:
        disposal.period_id = period.id
        disposal.status = 'Pendiente baja'
        disposal.reason = reason or disposal.reason
        disposal.requested_by = requested_by
        disposal.requested_at = now_iso_value
        disposal.reviewed_by = None
        disposal.reviewed_at = None
        disposal.review_notes = None

    db.session.commit()
    return jsonify({'disposal': disposal.to_dict(asset=asset)})


@app.route('/disposals/<int:disposal_id>', methods=['PATCH'])
def update_disposal(disposal_id):
    ensure_db()
    disposal = AssetDisposal.query.get(disposal_id)
    if not disposal:
        return jsonify({'error': 'Registro de baja no encontrado'}), 404

    data = request.get_json() or {}
    new_status = (data.get('status') or '').strip()
    reason_raw = data.get('reason', None)
    review_notes = (data.get('review_notes') or '').strip() or None
    reviewed_by = get_actor_username((data.get('reviewed_by') or '').strip() or 'unknown')
    type_override = data.get('type_override', None)
    allowed = {'Pendiente baja', 'Aprobada para baja', 'Rechazada'}
    if new_status and new_status not in allowed:
        return jsonify({'error': 'Estado de baja invalido'}), 400
    if reason_raw is not None:
        reason_txt = str(reason_raw).strip()
        if not reason_txt:
            return jsonify({'error': 'El motivo de baja no puede quedar vacio'}), 400
        disposal.reason = reason_txt
    if new_status:
        disposal.status = new_status
        disposal.reviewed_by = reviewed_by
        disposal.reviewed_at = now_iso()
        disposal.review_notes = review_notes

    asset = Asset.query.get(disposal.asset_id)
    if type_override is not None:
        normalized_type = normalize_manual_disposal_type(type_override)
        if not normalized_type:
            return jsonify({
                'error': 'Tipo de reclasificacion invalido',
                'allowed_types': DISPOSAL_MANUAL_TYPE_OPTIONS,
            }), 400
        if not asset:
            return jsonify({'error': 'Activo no encontrado para reclasificar'}), 404
        asset.tipo_activo_cache = normalized_type

    db.session.commit()
    return jsonify({'disposal': disposal.to_dict(asset=asset)})


@app.route('/disposals/<int:disposal_id>', methods=['DELETE'])
def delete_disposal(disposal_id):
    ensure_db()
    disposal = AssetDisposal.query.get(disposal_id)
    if not disposal:
        return jsonify({'error': 'Registro de baja no encontrado'}), 404

    asset = Asset.query.get(disposal.asset_id)
    if asset:
        asset.estado_inventario = 'Pendiente verificacion'
        asset.fecha_verificacion = now_iso()
        asset.usuario_verificador = get_actor_username('usuario_movil')

    db.session.delete(disposal)
    db.session.commit()
    return jsonify({'ok': True, 'asset_id': disposal.asset_id})


@app.route('/dashboard/summary', methods=['GET'])
def dashboard_summary():
    ensure_db()
    service = request.args.get('service')
    run_id = request.args.get('run_id', type=int)
    period_id = request.args.get('period_id', type=int)
    payload, error = build_dashboard_payload(service=service, run_id=run_id, period_id=period_id)
    if error:
        return jsonify({'error': error}), 404
    return jsonify(payload)


@app.route('/dashboard/report_pdf', methods=['GET'])
def dashboard_report_pdf():
    ensure_db()
    service = request.args.get('service')
    run_id = request.args.get('run_id', type=int)
    period_id = request.args.get('period_id', type=int)
    payload, error = build_dashboard_payload(service=service, run_id=run_id, period_id=period_id)
    if error:
        return jsonify({'error': error}), 404

    out = BytesIO()
    logo_path = get_hospital_logo_path()
    doc = SimpleDocTemplate(
        out,
        pagesize=letter,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=16 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    brand_blue = colors.HexColor('#0A6FB3')
    brand_blue_dark = colors.HexColor('#07507F')
    brand_green = colors.HexColor('#1E9E57')
    brand_yellow = colors.HexColor('#F2C94C')
    brand_red = colors.HexColor('#C0392B')
    title_style = ParagraphStyle(
        'DashTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=21,
        textColor=brand_blue_dark,
        spaceAfter=6,
    )
    subtitle_style = ParagraphStyle(
        'DashSubTitle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#51606F'),
        leading=12,
        spaceAfter=4,
    )
    section_style = ParagraphStyle(
        'DashSection',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        textColor=brand_blue_dark,
        spaceBefore=6,
        spaceAfter=3,
    )
    cell_style = ParagraphStyle(
        'Cell',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
    )

    k = payload['kpis']
    meta = payload.get('meta', {})
    run_name = (meta.get('run') or {}).get('name') if meta.get('run') else 'Sin jornada'
    period_name = (meta.get('period') or {}).get('name') if meta.get('period') else 'Sin periodo'
    service_filter = meta.get('service_filter') or 'TODOS'
    generated_at = meta.get('generated_at_local') or format_dt_local(meta.get('generated_at') or now_iso())

    story = []
    hero = Table([[
        Paragraph(
            '<font color="white"><b>Dashboard Institucional de Inventario</b></font><br/>'
            '<font color="white">Hospital Francisco de Paula Santander E.S.E.</font>',
            ParagraphStyle(
                'HeroTitle',
                parent=styles['Normal'],
                fontName='Helvetica-Bold',
                fontSize=15,
                leading=18,
            )
        )
    ]], colWidths=[182 * mm])
    hero.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), brand_blue),
        ('BOX', (0, 0), (-1, -1), 1.0, brand_blue_dark),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(hero)
    story.append(Spacer(1, 5))
    story.append(Paragraph('Reporte Ejecutivo Dashboard Inventario', title_style))
    story.append(Paragraph(
        f'Fecha: {generated_at} &nbsp;&nbsp;|&nbsp;&nbsp; Periodo: {period_name} '
        f'&nbsp;&nbsp;|&nbsp;&nbsp; Jornada: {run_name} '
        f'&nbsp;&nbsp;|&nbsp;&nbsp; Servicio: {service_filter}',
        subtitle_style
    ))

    narrative = build_executive_narrative(payload)
    plan = build_executive_action_plan(payload)
    story.append(Paragraph('Objetivo General', section_style))
    story.append(Paragraph(narrative.get('objetivo_general', ''), subtitle_style))
    story.append(Paragraph('Objetivos Especificos', section_style))
    obj_lines = '<br/>'.join([f'- {x}' for x in narrative.get('objetivos_especificos', [])])
    story.append(Paragraph(obj_lines or 'Sin objetivos definidos.', subtitle_style))
    story.append(Paragraph('Resumen Ejecutivo', section_style))
    story.append(Paragraph(narrative.get('resumen', ''), subtitle_style))
    story.append(Paragraph('Interpretacion Contextual', section_style))
    int_lines = '<br/>'.join([f'- {x}' for x in narrative.get('interpretacion', [])])
    story.append(Paragraph(int_lines or 'Sin interpretacion disponible.', subtitle_style))
    story.append(Spacer(1, 3))

    risk_color = {
        'ALTO': '#B42318',
        'MEDIO': '#B26A00',
        'BAJO': '#0D7A52',
    }.get(plan.get('risk_level'), '#0D7A52')
    semaforo = Table([[
        Paragraph(f"<b>Semaforo de riesgo:</b> {plan.get('risk_level', 'N/D')}<br/>{plan.get('risk_reason', '')}", styles['Normal'])
    ]], colWidths=[168 * mm])
    semaforo.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#F8FBFF')),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.HexColor(risk_color)),
        ('BOX', (0, 0), (-1, -1), 1.0, colors.HexColor(risk_color)),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(semaforo)
    story.append(Spacer(1, 4))

    kpi_data = [[
        Paragraph('<b>Total activos</b><br/>{}'.format(k.get('total', 0)), styles['Normal']),
        Paragraph('<b>Encontrados</b><br/>{} ({}%)'.format(k.get('found', 0), k.get('found_pct', 0)), styles['Normal']),
        Paragraph('<b>No encontrados</b><br/>{} ({}%)'.format(k.get('not_found', 0), k.get('not_found_pct', 0)), styles['Normal']),
        Paragraph('<b>Pendientes</b><br/>{}'.format(k.get('pending', 0)), styles['Normal']),
    ]]
    kpi_table = Table(kpi_data, colWidths=[42 * mm, 42 * mm, 42 * mm, 42 * mm])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#E8F2FC')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#E9F7EF')),
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#FFF0F0')),
        ('BACKGROUND', (3, 0), (3, 0), colors.HexColor('#FFF7E6')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#D7E5EE')),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 4))

    coverage = payload.get('coverage', {})
    coverage_data = [[
        Paragraph('<b>Base total activos</b><br/>{}'.format(coverage.get('base_total_assets', 0)), styles['Normal']),
        Paragraph('<b>Activos en alcance periodo/jornada</b><br/>{} ({}%)'.format(
            coverage.get('scope_assets', 0), coverage.get('scope_assets_pct', 0)
        ), styles['Normal']),
        Paragraph('<b>Base fuera de alcance</b><br/>{}'.format(coverage.get('base_not_in_scope_assets', 0)), styles['Normal']),
        Paragraph('<b>Cobertura valor</b><br/>{} / {} ({}%)'.format(
            money_text(to_number(coverage.get('scope_value', 0))),
            money_text(to_number(coverage.get('base_total_value', 0))),
            coverage.get('scope_value_pct', 0),
        ), styles['Normal']),
    ]]
    coverage_table = Table(coverage_data, colWidths=[42 * mm, 42 * mm, 42 * mm, 42 * mm])
    coverage_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (3, 0), colors.HexColor('#E9F7EF')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#C7E5C9')),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#DDF0DF')),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(coverage_table)
    story.append(Spacer(1, 4))

    financial = payload.get('financial', {})
    financial_data = [[
        Paragraph('<b>Valor total inventario</b><br/>{}'.format(money_text(to_number(financial.get('total_value', 0)))), styles['Normal']),
        Paragraph('<b>Valor encontrado</b><br/>{}'.format(money_text(to_number(financial.get('found_value', 0)))), styles['Normal']),
        Paragraph('<b>Valor no encontrado</b><br/>{} ({}%)'.format(
            money_text(to_number(financial.get('not_found_value', 0))),
            financial.get('not_found_value_pct', 0)
        ), styles['Normal']),
        Paragraph('<b>Valor critico no encontrado</b><br/>{}'.format(
            money_text(to_number(financial.get('critical_not_found_value', 0)))
        ), styles['Normal']),
    ]]
    financial_table = Table(financial_data, colWidths=[42 * mm, 42 * mm, 42 * mm, 42 * mm])
    financial_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (3, 0), colors.HexColor('#FFF6E8')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#E6D0BC')),
        ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#F0E1D2')),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(financial_table)
    story.append(Spacer(1, 4))

    story.append(Paragraph('Distribucion del estado del inventario', section_style))
    pie_drawing = Drawing(170 * mm, 58 * mm)
    pie = Pie()
    pie.x = 18
    pie.y = 6
    pie.width = 70 * mm
    pie.height = 46 * mm
    pie.slices.strokeWidth = 0.5
    found_v = max(0, int(k.get('found', 0) or 0))
    not_found_v = max(0, int(k.get('not_found', 0) or 0))
    pending_v = max(0, int(k.get('pending', 0) or 0))
    total_v = found_v + not_found_v + pending_v
    pie.data = [found_v, not_found_v, pending_v] if total_v > 0 else [1]
    pie.labels = ['Encontrados', 'No encontrados', 'Pendientes'] if total_v > 0 else ['Sin datos']
    if total_v > 0:
        pie.slices[0].fillColor = brand_green
        pie.slices[1].fillColor = brand_red
        pie.slices[2].fillColor = brand_yellow
        pie.slices[1].popout = 2
    else:
        pie.slices[0].fillColor = colors.HexColor('#CBD5E1')
    pie_drawing.add(pie)
    pie_drawing.add(String(95 * mm, 40 * mm, f'Encontrados: {found_v} ({k.get("found_pct", 0)}%)', fontSize=9, fillColor=brand_green))
    pie_drawing.add(String(95 * mm, 30 * mm, f'No encontrados: {not_found_v} ({k.get("not_found_pct", 0)}%)', fontSize=9, fillColor=brand_red))
    pie_drawing.add(String(95 * mm, 20 * mm, f'Pendientes: {pending_v}', fontSize=9, fillColor=colors.HexColor('#9A6700')))
    story.append(pie_drawing)
    story.append(Spacer(1, 4))

    insights = payload.get('insights', [])
    story.append(Paragraph('Mensajes Clave para Alta Gerencia', section_style))
    if insights:
        bullets = ''.join([f'• {i}<br/>' for i in insights])
        story.append(Paragraph(bullets, subtitle_style))
    else:
        story.append(Paragraph('Sin hallazgos relevantes para este corte.', subtitle_style))
    story.append(Spacer(1, 3))

    def make_chart(title, rows, top_n=15):
        drawing = Drawing(520, 200)
        drawing.add(String(0, 186, title, fontName='Helvetica-Bold', fontSize=11, fillColor=brand_blue_dark))
        if not rows:
            drawing.add(String(0, 162, 'Sin datos', fontName='Helvetica', fontSize=9, fillColor=colors.HexColor('#7A8794')))
            return drawing

        selected = rows[:top_n]
        labels = [str(r.get('name', ''))[:46] for r in selected]
        values = [float(r.get('total', 0) or 0) for r in selected]

        chart = HorizontalBarChart()
        chart.x = 110
        chart.y = 10
        chart.height = 155
        chart.width = 390
        chart.data = [values]
        chart.categoryAxis.categoryNames = labels
        chart.categoryAxis.labels.fontName = 'Helvetica'
        chart.categoryAxis.labels.fontSize = 7
        chart.categoryAxis.labels.boxAnchor = 'e'
        chart.categoryAxis.labels.dx = -4
        chart.valueAxis.valueMin = 0
        chart.valueAxis.labels.fontSize = 7
        chart.valueAxis.visibleGrid = 1
        chart.valueAxis.gridStrokeColor = colors.HexColor('#DCE7EE')
        chart.bars[0].fillColor = brand_blue
        chart.barSpacing = 2
        chart.groupSpacing = 4
        palette = [brand_blue, colors.HexColor('#118AB2'), brand_green, colors.HexColor('#2FAE66'), brand_yellow]
        for i in range(len(values)):
            try:
                chart.bars[(0, i)].fillColor = palette[i % len(palette)]
            except Exception:
                pass
        drawing.add(chart)
        return drawing

    story.append(Paragraph('Visualizaciones', section_style))
    story.append(make_chart('Activos por servicio (Top 10)', payload.get('by_service', []), top_n=10))
    story.append(Spacer(1, 3))
    story.append(make_chart('Activos por tipo de equipo (Top 10)', payload.get('by_type', []), top_n=10))
    story.append(Spacer(1, 3))
    story.append(make_chart('Activos por área', payload.get('by_area', []), top_n=15))
    story.append(Spacer(1, 4))

    story.append(Paragraph('Activos Criticos y Costosos No Encontrados', section_style))
    critical_rows = payload.get('critical_not_found', [])
    if not critical_rows:
        story.append(Paragraph('No se identificaron activos criticos no encontrados en este corte.', subtitle_style))
    else:
        critical_data = [[
            Paragraph('<b>Codigo</b>', cell_style),
            Paragraph('<b>Activo</b>', cell_style),
            Paragraph('<b>Servicio</b>', cell_style),
            Paragraph('<b>Tipo</b>', cell_style),
            Paragraph('<b>Valor libro</b>', cell_style),
            Paragraph('<b>Criticidad</b>', cell_style),
            Paragraph('<b>Motivo</b>', cell_style),
        ]]
        for item in critical_rows:
            critical_data.append([
                Paragraph(str(item.get('code', '')), cell_style),
                Paragraph(str(item.get('name', '')), cell_style),
                Paragraph(str(item.get('service', '')), cell_style),
                Paragraph(str(item.get('type', '')), cell_style),
                Paragraph(money_text(to_number(item.get('value', 0))), cell_style),
                Paragraph(str(item.get('critical_score', 0)), cell_style),
                Paragraph(str(item.get('critical_reasons', '')), cell_style),
            ])
        critical_table = Table(
            critical_data,
            colWidths=[20 * mm, 46 * mm, 26 * mm, 28 * mm, 20 * mm, 16 * mm, 24 * mm],
            repeatRows=1
        )
        critical_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#7A1F1F')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (4, 1), (5, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#D6B5B5')),
            ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#E7CFCF')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#FFF7F7')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(critical_table)
    story.append(Spacer(1, 4))

    story.append(Paragraph('Anexo Operativo: Activos No Encontrados', section_style))
    not_found_rows = payload.get('not_found_assets', [])
    if not not_found_rows:
        story.append(Paragraph('No hay activos no encontrados para este corte.', subtitle_style))
    else:
        annex_cap = 120
        not_found_rows = not_found_rows[:annex_cap]
        if payload.get('not_found_assets_capped') or payload.get('not_found_assets_total', 0) > annex_cap:
            story.append(Paragraph(
                f"Se muestran los primeros {len(not_found_rows)} de {payload.get('not_found_assets_total', len(not_found_rows))} activos no encontrados. El detalle completo se gestiona en exportes operativos.",
                subtitle_style
            ))
        nf_data = [[
            Paragraph('<b>Codigo</b>', cell_style),
            Paragraph('<b>Activo</b>', cell_style),
            Paragraph('<b>Servicio</b>', cell_style),
            Paragraph('<b>Responsable</b>', cell_style),
            Paragraph('<b>Ubicacion</b>', cell_style),
            Paragraph('<b>Valor libro</b>', cell_style),
        ]]
        for item in not_found_rows:
            nf_data.append([
                Paragraph(str(item.get('code', '')), cell_style),
                Paragraph(str(item.get('name', '')), cell_style),
                Paragraph(str(item.get('service', '')), cell_style),
                Paragraph(str(item.get('responsible', '')), cell_style),
                Paragraph(str(item.get('location', '')), cell_style),
                Paragraph(money_text(to_number(item.get('value', 0))), cell_style),
            ])
        nf_table = Table(
            nf_data,
            colWidths=[20 * mm, 48 * mm, 36 * mm, 34 * mm, 36 * mm, 22 * mm],
            repeatRows=1,
        )
        nf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#922020')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (5, 1), (5, -1), 'RIGHT'),
            ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#D5B3B3')),
            ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#E9D7D7')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#FFF8F8')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(nf_table)
    story.append(Spacer(1, 4))

    def value_table(title, rows):
        story.append(Paragraph(title, section_style))
        if not rows:
            story.append(Paragraph('Sin datos para esta seccion.', subtitle_style))
            story.append(Spacer(1, 3))
            return
        data = [[
            Paragraph('<b>Nombre</b>', cell_style),
            Paragraph('<b>Valor no encontrado</b>', cell_style),
        ]]
        for r in rows:
            data.append([
                Paragraph(str(r.get('name', '')), cell_style),
                Paragraph(money_text(to_number(r.get('not_found_value', 0))), cell_style),
            ])
        table = Table(data, colWidths=[120 * mm, 60 * mm], repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#244B5A')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
            ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FBFD')]),
        ]))
        story.append(table)
        story.append(Spacer(1, 4))

    value_table('Impacto Economico de No Encontrados por Servicio', payload.get('top_not_found_by_service_value', []))
    value_table('Impacto Economico de No Encontrados por Tipo de Equipo', payload.get('top_not_found_by_type_value', []))
    story.append(PageBreak())

    def section_table(title, rows):
        story.append(Paragraph(title, section_style))
        if not rows:
            story.append(Paragraph('Sin datos para esta sección.', subtitle_style))
            story.append(Spacer(1, 3))
            return

        data = [[
            Paragraph('<b>Nombre</b>', cell_style),
            Paragraph('<b>Total</b>', cell_style),
            Paragraph('<b>Encontrados</b>', cell_style),
            Paragraph('<b>No encontrados</b>', cell_style),
            Paragraph('<b>Pendientes</b>', cell_style),
            Paragraph('<b>% Encontrados</b>', cell_style),
        ]]
        for r in rows:
            data.append([
                Paragraph(str(r.get('name', '')), cell_style),
                Paragraph(str(r.get('total', 0)), cell_style),
                Paragraph(str(r.get('found', 0)), cell_style),
                Paragraph(str(r.get('not_found', 0)), cell_style),
                Paragraph(str(r.get('pending', 0)), cell_style),
                Paragraph(str(r.get('found_pct', 0)), cell_style),
            ])

        table = Table(data, colWidths=[77 * mm, 16 * mm, 20 * mm, 25 * mm, 19 * mm, 22 * mm], repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#07507F')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
            ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FBFD')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        story.append(Spacer(1, 4))

    section_table('Detalle Completo por Servicio', payload.get('by_service', []))
    section_table('Detalle Completo por Tipo de Equipo', payload.get('by_type', []))
    section_table('Detalle completo por área', payload.get('by_area', []))

    story.append(PageBreak())
    story.append(Paragraph('Conclusion Final', section_style))
    story.append(Paragraph(build_executive_conclusion(payload), subtitle_style))

    page_header = make_pdf_page_header(
        logo_path=None,
        right_image_path=logo_path,
        right_width_mm=14,
        right_height_mm=14,
        right_top_mm=16,
    )
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    out.seek(0)
    base_name = run_name if run_name and run_name != 'Sin jornada' else service_filter
    safe_name = clean_filename(base_name)
    filename = f'dashboard_{safe_name}.pdf'
    return send_file(out, download_name=filename, as_attachment=True, mimetype='application/pdf')


