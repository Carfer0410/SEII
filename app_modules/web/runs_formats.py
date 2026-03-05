from .inventory_assets_disposals import *

@app.route('/dashboard/compare_periods', methods=['GET'])
def dashboard_compare_periods():
    ensure_db()
    period_a = request.args.get('period_a', type=int)
    period_b = request.args.get('period_b', type=int)
    service = (request.args.get('service') or '').strip() or None
    if not period_a or not period_b:
        return jsonify({'error': 'Debes indicar ambos periodos para comparar'}), 400

    payload_a, err_a = build_dashboard_payload(service=service, period_id=period_a)
    if err_a:
        return jsonify({'error': err_a}), 400
    payload_b, err_b = build_dashboard_payload(service=service, period_id=period_b)
    if err_b:
        return jsonify({'error': err_b}), 400

    ka = payload_a.get('kpis', {})
    kb = payload_b.get('kpis', {})
    pa = (payload_a.get('meta', {}).get('period') or {})
    pb = (payload_b.get('meta', {}).get('period') or {})
    response = {
        'period_a': {
            'id': pa.get('id'),
            'name': pa.get('name') or f'Periodo {period_a}',
            'total': ka.get('total', 0),
            'found': ka.get('found', 0),
            'not_found': ka.get('not_found', 0),
            'found_pct': ka.get('found_pct', 0),
            'not_found_pct': ka.get('not_found_pct', 0),
        },
        'period_b': {
            'id': pb.get('id'),
            'name': pb.get('name') or f'Periodo {period_b}',
            'total': kb.get('total', 0),
            'found': kb.get('found', 0),
            'not_found': kb.get('not_found', 0),
            'found_pct': kb.get('found_pct', 0),
            'not_found_pct': kb.get('not_found_pct', 0),
        },
        'delta': {
            'total': kb.get('total', 0) - ka.get('total', 0),
            'found': kb.get('found', 0) - ka.get('found', 0),
            'not_found': kb.get('not_found', 0) - ka.get('not_found', 0),
            'found_pct': round((kb.get('found_pct', 0) - ka.get('found_pct', 0)), 2),
            'not_found_pct': round((kb.get('not_found_pct', 0) - ka.get('not_found_pct', 0)), 2),
        },
        'service_filter': service or '',
        'generated_at': now_iso(),
    }
    return jsonify(response)


@app.route('/runs', methods=['POST'])
def create_run():
    ensure_db()
    data = request.get_json() or {}
    name = (data.get('name') or '').strip()
    services = normalize_service_scope(data.get('services'))
    if not services:
        legacy_service = (data.get('service') or '').strip()
        services = normalize_service_scope([legacy_service] if legacy_service else [])
    service = services[0] if services else None
    period_id = data.get('period_id')
    try:
        period_id = int(period_id) if period_id not in (None, '') else None
    except Exception:
        period_id = None
    created_by = (data.get('created_by') or '').strip() or 'unknown'
    if not name:
        return jsonify({'error': 'Debe indicar nombre de jornada'}), 400
    if not services:
        return jsonify({'error': 'Debes seleccionar al menos un servicio para la jornada'}), 400
    if not period_id:
        return jsonify({'error': 'Debes seleccionar el periodo de inventario'}), 400
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404
    if period.status != 'open':
        return jsonify({'error': 'El periodo seleccionado esta cerrado o anulado'}), 400

    # Regla operativa: un servicio solo puede tener una jornada por periodo.
    # Si se requiere continuar por activos nuevos, se debe reabrir la jornada existente.
    requested_services_map = {str(s or '').strip().casefold(): str(s or '').strip() for s in services if str(s or '').strip()}
    scoped_runs = InventoryRun.query.filter(
        InventoryRun.period_id == period.id,
        InventoryRun.status.in_(['active', 'closed'])
    ).all()
    blocked_services = {}
    for scoped_run in scoped_runs:
        closed_scope_cf = {str(s or '').strip().casefold() for s in run_scope_services(scoped_run)}
        for svc_cf, svc_label in requested_services_map.items():
            if svc_cf in closed_scope_cf and svc_label not in blocked_services:
                blocked_services[svc_label] = scoped_run.name or f'ID {scoped_run.id}'

    if blocked_services:
        blocked_detail = ', '.join(
            [f'{svc} (jornada "{run_name}")' for svc, run_name in blocked_services.items()]
        )
        return jsonify({
            'error': (
                'No puedes crear la jornada porque estos servicios ya tienen jornada '
                f'en este periodo: {blocked_detail}. Si hay activos nuevos, reabre la jornada cerrada.'
            )
        }), 400

    run = InventoryRun(
        name=name,
        period_id=period.id,
        service=service,
        service_scope_json=json.dumps(services, ensure_ascii=False),
        status='active',
        started_at=now_iso(),
        created_by=created_by,
    )
    db.session.add(run)
    db.session.commit()
    row = run.to_dict()
    row['period_name'] = period.name
    row['period_status'] = period.status
    return jsonify({'run': row})


@app.route('/runs/<int:run_id>/summary', methods=['GET'])
def run_summary(run_id):
    ensure_db()
    run, err = get_run_or_404(run_id)
    if err:
        return err

    q = Asset.query
    q = apply_run_scope_filter(q, run)
    total = q.count()

    found = RunAssetStatus.query.filter_by(run_id=run.id, status='Encontrado').count()
    not_found = RunAssetStatus.query.filter_by(run_id=run.id, status='No encontrado').count()
    pending = max(total - found - not_found, 0)
    return jsonify({
        'run': run.to_dict(),
        'summary': {
            'total': total,
            'found': found,
            'not_found': not_found,
            'pending': pending,
        }
    })


def _iter_chunks(items, size=900):
    for i in range(0, len(items), size):
        yield items[i:i + size]


def _run_status_rows_for_assets(run_id, asset_ids):
    if not asset_ids:
        return []
    rows = []
    for chunk in _iter_chunks(asset_ids):
        rows.extend(
            RunAssetStatus.query.filter(
                RunAssetStatus.run_id == run_id,
                RunAssetStatus.asset_id.in_(chunk),
            ).all()
        )
    return rows


def build_run_coverage_summary(run):
    q = apply_run_scope_filter(Asset.query, run)
    assets_scope = q.order_by(Asset.c_act.asc()).all()
    asset_ids = [a.id for a in assets_scope]
    total = len(asset_ids)
    found = 0
    not_found = 0
    if asset_ids:
        statuses = _run_status_rows_for_assets(run.id, asset_ids)
        for st in statuses:
            if st.status == 'Encontrado':
                found += 1
            elif st.status == 'No encontrado':
                not_found += 1
    pending = max(total - found - not_found, 0)
    missing = max(total - found, 0)
    found_pct = round((found / total) * 100.0, 2) if total > 0 else 0.0
    return {
        'total': total,
        'found': found,
        'not_found': not_found,
        'pending': pending,
        'missing': missing,
        'found_pct': found_pct,
        'assets_scope': assets_scope,
    }


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


def build_clearance_validation(period_id, run_id):
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return None, None, None, None, (jsonify({'error': 'Periodo no encontrado'}), 404)
    run = InventoryRun.query.get(run_id)
    if not run:
        return None, None, None, None, (jsonify({'error': 'Jornada no encontrada'}), 404)
    if run.period_id != period.id:
        return None, None, None, None, (jsonify({'error': 'La jornada no pertenece al periodo seleccionado'}), 400)
    if run.status != 'closed':
        return None, None, None, None, (jsonify({'error': 'Solo puedes generar paz y salvo con jornadas cerradas'}), 400)

    summary = build_run_coverage_summary(run)
    if summary['total'] <= 0:
        return None, None, None, None, (jsonify({'error': 'La jornada no tiene activos en alcance para emitir paz y salvo'}), 400)

    allowed = summary['found'] == summary['total']
    if allowed:
        reason = 'Cumplimiento 100%: todos los activos del alcance fueron encontrados.'
    else:
        reason = (
            f'No se puede generar paz y salvo. Faltan {summary["missing"]} activos para completar el 100%. '
            f'No encontrados: {summary["not_found"]}. Pendientes: {summary["pending"]}.'
        )
    return period, run, summary, {'allowed': allowed, 'message': reason}, None


@app.route('/paz_y_salvo/validate', methods=['GET'])
def validate_paz_y_salvo():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    run_id = request.args.get('run_id', type=int)
    if not period_id:
        return jsonify({'error': 'Debes seleccionar el periodo para validar paz y salvo'}), 400
    if not run_id:
        return jsonify({'error': 'Debes seleccionar la jornada para validar paz y salvo'}), 400

    period, run, summary, validation, err = build_clearance_validation(period_id, run_id)
    if err:
        return err
    return jsonify({
        'period': period.to_dict(),
        'run': run.to_dict(),
        'summary': {
            'total': summary['total'],
            'found': summary['found'],
            'not_found': summary['not_found'],
            'pending': summary['pending'],
            'missing': summary['missing'],
            'found_pct': summary['found_pct'],
        },
        'allowed': validation['allowed'],
        'message': validation['message'],
    })


@app.route('/paz_y_salvo/generate', methods=['POST'])
def generate_paz_y_salvo_pdf():
    ensure_db()
    data = request.get_json() or {}
    period_id = data.get('period_id')
    run_id = data.get('run_id')
    try:
        period_id = int(period_id) if period_id not in (None, '') else None
    except Exception:
        period_id = None
    try:
        run_id = int(run_id) if run_id not in (None, '') else None
    except Exception:
        run_id = None
    if not period_id:
        return jsonify({'error': 'Debes seleccionar el periodo para generar paz y salvo'}), 400
    if not run_id:
        return jsonify({'error': 'Debes seleccionar la jornada para generar paz y salvo'}), 400

    outgoing_responsible = (data.get('outgoing_responsible') or '').strip()
    incoming_responsible = (data.get('incoming_responsible') or '').strip()
    issued_by = (data.get('issued_by') or '').strip() or 'Responsable activos fijos'
    observations = (data.get('observations') or '').strip()
    report_date = (data.get('report_date') or '').strip()
    if report_date:
        try:
            datetime.strptime(report_date, '%Y-%m-%d')
        except Exception:
            return jsonify({'error': 'Fecha invalida. Usa formato YYYY-MM-DD'}), 400
    selected_date = report_date or now_local_dt().strftime('%Y-%m-%d')

    if not outgoing_responsible:
        return jsonify({'error': 'Debes indicar el responsable saliente'}), 400
    if not incoming_responsible:
        return jsonify({'error': 'Debes indicar el responsable entrante'}), 400

    period, run, summary, validation, err = build_clearance_validation(period_id, run_id)
    if err:
        return err
    if not validation['allowed']:
        return jsonify({
            'error': validation['message'],
            'summary': {
                'total': summary['total'],
                'found': summary['found'],
                'not_found': summary['not_found'],
                'pending': summary['pending'],
                'missing': summary['missing'],
                'found_pct': summary['found_pct'],
            }
        }), 400

    assets_scope = summary['assets_scope']
    services = run_scope_services(run)
    services_label = ', '.join(services) if services else (run.service or '')
    out = BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=letter,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=22 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'pys_title',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=13,
        textColor=colors.HexColor('#0B4F6C'),
    )
    normal = ParagraphStyle('pys_normal', parent=styles['Normal'], fontSize=8.2, leading=10)
    centered = ParagraphStyle('pys_centered', parent=normal, alignment=1)

    story = []
    story.append(Paragraph('PAZ Y SALVO DE ACTIVOS FIJOS', title_style))
    story.append(Spacer(1, 5))
    meta = [
        [Paragraph('<b>Periodo</b>', normal), Paragraph(period.name or f'Periodo {period.id}', normal)],
        [Paragraph('<b>Jornada</b>', normal), Paragraph(run.name or f'Jornada {run.id}', normal)],
        [Paragraph('<b>Servicio(s)</b>', normal), Paragraph(services_label or 'N/D', normal)],
        [Paragraph('<b>Fecha emision</b>', normal), Paragraph(selected_date, normal)],
        [Paragraph('<b>Responsable saliente</b>', normal), Paragraph(outgoing_responsible, normal)],
        [Paragraph('<b>Responsable entrante</b>', normal), Paragraph(incoming_responsible, normal)],
        [Paragraph('<b>Total activos en alcance</b>', normal), Paragraph(str(summary['total']), normal)],
        [Paragraph('<b>Total activos encontrados</b>', normal), Paragraph(str(summary['found']), normal)],
        [Paragraph('<b>Cumplimiento</b>', normal), Paragraph(f"{summary['found_pct']}%", normal)],
    ]
    meta_table = Table(meta, colWidths=[54 * mm, 124 * mm])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EAF4FA')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))
    statement = (
        f'Se certifica que la jornada "{run.name}" del periodo "{period.name}" '
        f'alcanzó cumplimiento del 100% en inventario de activos fijos. '
        f'En consecuencia, se emite paz y salvo para el cambio de responsable '
        f'del servicio {services_label or "N/D"}.'
    )
    story.append(Paragraph(statement, normal))
    if observations:
        story.append(Spacer(1, 4))
        story.append(Paragraph(f"<b>Observaciones:</b> {escape(observations)}", normal))
    story.append(Spacer(1, 7))

    rows = [[
        Paragraph('<b>N</b>', normal),
        Paragraph('<b>COD ACTIVO</b>', normal),
        Paragraph('<b>DESCRIPCION</b>', normal),
        Paragraph('<b>SERVICIO</b>', normal),
        Paragraph('<b>UBICACION</b>', normal),
    ]]
    for i, asset in enumerate(assets_scope, start=1):
        rows.append([
            Paragraph(str(i), normal),
            Paragraph(str(asset.c_act or ''), normal),
            Paragraph(str(asset.nom or ''), normal),
            Paragraph(str(asset.nom_ccos or ''), normal),
            Paragraph(str(asset.des_ubi or ''), normal),
        ])
    assets_table = Table(
        rows,
        colWidths=[10 * mm, 24 * mm, 64 * mm, 42 * mm, 38 * mm],
        repeatRows=1,
    )
    assets_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0B4F6C')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FBFD')]),
    ]))
    story.append(assets_table)
    story.append(Spacer(1, 18))

    signatures = Table([
        [
            Paragraph('__________________________', centered),
            Paragraph('__________________________', centered),
            Paragraph('__________________________', centered),
        ],
        [
            Paragraph(outgoing_responsible, centered),
            Paragraph(incoming_responsible, centered),
            Paragraph(issued_by, centered),
        ],
        [
            Paragraph('RESPONSABLE SALIENTE', centered),
            Paragraph('RESPONSABLE ENTRANTE', centered),
            Paragraph('EMITE CERTIFICACION', centered),
        ],
    ], colWidths=[58 * mm, 58 * mm, 58 * mm])
    signatures.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    story.append(signatures)

    page_header = make_pdf_page_header(get_hospital_logo_path())
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    content = out.getvalue()

    timestamp = now_local_dt().strftime('%Y%m%d%H%M%S')
    public_name = clean_filename(
        f"paz_y_salvo_{period.name}_{run.name}_{selected_date}.pdf"
    )
    storage_name = f"{clean_filename(os.path.splitext(public_name)[0])}_{timestamp}.pdf"
    file_path = os.path.join(DOCUMENTS_DIR, storage_name)
    with open(file_path, 'wb') as fp:
        fp.write(content)

    doc_row = DocumentRecord(
        link_type='general',
        document_type='Certificacion',
        title=f'Paz y salvo activos fijos - {run.name}',
        description=(
            f'Paz y salvo por cambio de responsable. '
            f'Periodo: {period.name}. Jornada: {run.name}. Cumplimiento: 100%.'
        ),
        doc_date=selected_date,
        area_service=services_label or (run.service or ''),
        radicado=f'PYS-{period.id}-{run.id}-{timestamp}',
        file_name=public_name,
        file_path=file_path,
        file_ext='.pdf',
        file_size=len(content),
        uploaded_by=issued_by,
        uploaded_at=now_iso(),
        status='active',
    )
    db.session.add(doc_row)
    db.session.commit()

    return send_file(BytesIO(content), download_name=public_name, as_attachment=True, mimetype='application/pdf')


@app.route('/runs/<int:run_id>/close', methods=['POST'])
def close_run(run_id):
    ensure_db()
    run, err = get_run_or_404(run_id)
    if err:
        return err
    if run.status == 'cancelled':
        return jsonify({'error': 'La jornada esta anulada'}), 400
    if run.status != 'active':
        return jsonify({'error': 'La jornada ya esta cerrada'}), 400

    data = request.get_json() or {}
    user = (data.get('user') or '').strip() or 'system_close'
    now_iso_value = now_iso()

    q = Asset.query
    q = apply_run_scope_filter(q, run)
    assets_scope = q.all()
    asset_ids = [a.id for a in assets_scope]

    existing_statuses = _run_status_rows_for_assets(run.id, asset_ids)
    existing_map = {s.asset_id: s for s in existing_statuses}

    created_not_found = 0
    for asset in assets_scope:
        if asset.id in existing_map:
            continue
        db.session.add(RunAssetStatus(
            run_id=run.id,
            asset_id=asset.id,
            status='No encontrado',
            scanned_at=now_iso_value,
            scanned_by=user,
        ))
        asset.estado_inventario = 'No encontrado'
        asset.fecha_verificacion = now_iso_value
        asset.usuario_verificador = user
        created_not_found += 1

    run.status = 'closed'
    run.closed_at = now_iso_value
    db.session.commit()

    found = RunAssetStatus.query.filter_by(run_id=run.id, status='Encontrado').count()
    not_found = RunAssetStatus.query.filter_by(run_id=run.id, status='No encontrado').count()
    return jsonify({
        'run': run.to_dict(),
        'summary': {
            'total': len(assets_scope),
            'found': found,
            'not_found': not_found,
            'pending': 0,
        },
        'auto_marked_not_found': created_not_found,
    })


@app.route('/runs/<int:run_id>/cancel', methods=['POST'])
def cancel_run(run_id):
    ensure_db()
    run, err = get_run_or_404(run_id)
    if err:
        return err
    if run.status == 'cancelled':
        return jsonify({'error': 'La jornada ya esta anulada'}), 400

    data = request.get_json() or {}
    reason = (data.get('reason') or '').strip()
    user = (data.get('user') or '').strip() or 'usuario_movil'
    if not reason:
        return jsonify({'error': 'Debes indicar el motivo de anulacion de la jornada'}), 400

    has_scan_trace = RunAssetStatus.query.filter_by(run_id=run.id).count() > 0
    if has_scan_trace:
        return jsonify({'error': 'No puedes anular la jornada porque ya tiene trazabilidad de escaneo'}), 400

    run.status = 'cancelled'
    run.closed_at = now_iso()
    run.cancelled_at = now_iso()
    run.cancelled_by = user
    run.cancel_reason = reason
    db.session.commit()
    return jsonify({'run': run.to_dict()})


@app.route('/runs/<int:run_id>/reopen', methods=['POST'])
def reopen_run(run_id):
    ensure_db()
    run, err = get_run_or_404(run_id)
    if err:
        return err
    if run.status == 'cancelled':
        return jsonify({'error': 'La jornada esta anulada y no puede reabrirse'}), 400
    if run.status == 'active':
        return jsonify({'error': 'La jornada ya esta activa'}), 400
    if run.status != 'closed':
        return jsonify({'error': 'Solo puedes reabrir jornadas cerradas'}), 400

    period = InventoryPeriod.query.get(run.period_id) if run.period_id else None
    if not period:
        return jsonify({'error': 'La jornada no tiene periodo asociado'}), 400
    if period.status != 'open':
        return jsonify({'error': 'Solo puedes reabrir jornadas en periodos abiertos'}), 400

    new_assets = count_new_assets_for_run(run)
    if new_assets <= 0:
        return jsonify({
            'error': 'No hay activos nuevos en el alcance de la jornada para justificar reapertura'
        }), 400

    run.status = 'active'
    run.closed_at = None
    db.session.commit()
    return jsonify({
        'run': run.to_dict(),
        'new_assets_in_scope': new_assets,
        'message': f'Jornada reabierta correctamente. Activos nuevos en alcance: {new_assets}.',
    })


@app.route('/export', methods=['GET', 'POST'])
def export():
    ensure_db()
    payload = request.get_json(silent=True) if request.method == 'POST' else None
    source = payload or request.args
    service = (source.get('service') or '').strip()
    run_id = source.get('run_id')
    try:
        run_id = int(run_id) if run_id not in (None, '') else None
    except Exception:
        run_id = None
    period_id = source.get('period_id')
    try:
        period_id = int(period_id) if period_id not in (None, '') else None
    except Exception:
        period_id = None
    receiver = (source.get('receiver') or '').strip()
    observation = (source.get('observation') or '').strip()
    report_date = (source.get('report_date') or '').strip()
    warehouse_lead = (source.get('warehouse_lead') or '').strip()
    assets_manager = (source.get('assets_manager') or '').strip()
    per_asset_observations = source.get('per_asset_observations') or {}
    if not isinstance(per_asset_observations, dict):
        per_asset_observations = {}
    if not period_id:
        return jsonify({'error': 'Debes seleccionar el periodo para generar A22'}), 400
    if not run_id:
        return jsonify({'error': 'Debes seleccionar la jornada del periodo para generar A22'}), 400
    if not warehouse_lead:
        return jsonify({'error': 'Lider de almacen es obligatorio'}), 400
    if not assets_manager:
        return jsonify({'error': 'Responsable de activos fijos es obligatorio'}), 400
    if report_date:
        try:
            datetime.strptime(report_date, '%Y-%m-%d')
        except Exception:
            return jsonify({'error': 'Fecha invalida. Usa formato YYYY-MM-DD'}), 400
    selected_date = report_date or now_local_dt().strftime('%Y-%m-%d')

    run, assets_scope, err = get_a22_scope(service=service or None, run_id=run_id, period_id=period_id)
    if err:
        return err
    if not assets_scope:
        return jsonify({'error': 'No hay activos encontrados para generar A22 con ese filtro'}), 400
    if not run or not run.service:
        return jsonify({'error': 'La jornada seleccionada no tiene centro de costo asociado'}), 400

    if not os.path.exists(TEMPLATE_A22_PATH):
        return jsonify({'error': 'No existe la plantilla formato a22.xlsx'}), 400

    wb = load_workbook(TEMPLATE_A22_PATH)
    ws = wb[wb.sheetnames[0]]

    assets_scope = sort_assets_for_a22(assets_scope)
    selected_service = run.service
    selected_receiver = receiver or assets_scope[0].nom_resp or ''
    work_area = classify_area(selected_service).upper()

    data_rows = [r for r in range(13, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, (int, float))]
    capacity = len(data_rows)
    if capacity == 0:
        return jsonify({'error': 'La plantilla no tiene filas de detalle configuradas'}), 400

    data_start = min(data_rows)
    data_end = max(data_rows)

    logo_path = next((p for p in A22_LOGO_CANDIDATES if os.path.exists(p)), None)
    logo_bytes = None
    if logo_path is None:
        try:
            with zipfile.ZipFile(TEMPLATE_A22_PATH, 'r') as zf:
                media_files = [n for n in zf.namelist() if n.startswith('xl/media/')]
                if media_files:
                    logo_bytes = zf.read(media_files[0])
        except Exception:
            logo_bytes = None

    def apply_header_and_signature(sheet):
        sheet.cell(6, 6).value = selected_date
        sheet.cell(7, 6).value = selected_service
        sheet.cell(8, 6).value = selected_receiver
        sheet.cell(9, 6).value = f'LIDER {selected_service}'
        sheet.cell(10, 6).value = work_area

        # Reemplaza nombre fijo en firmas por el responsable seleccionado.
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if isinstance(cell.value, str) and 'LYDA MARTINEZ' in cell.value.upper():
                    cell.value = selected_receiver

        # Texto legal dinámico (evita depender de fórmulas rotas por estructura).
        legal_text = (
            f'Una vez culminado el proceso de inventario general en el área de {selected_service} '
            f'se entrega a {selected_receiver} responsable de dicha área, un documento detallado '
            f'que incluye todos los activos fijos asignados, clasificados y verificados. '
            f'A partir de la entrega formal de este documento, el responsable del área asume la '
            f'obligación de velar por la custodia y el buen estado de cada activo listado. '
            f'En caso de pérdida, daño o cualquier irregularidad que afecte los activos bajo su '
            f'supervisión, el responsable deberá presentar una justificación oportuna y detallada, '
            f'y asumir las consecuencias correspondientes. Esta medida busca asegurar la transparencia '
            f'y la adecuada gestión de los recursos de la institución.'
        )
        sheet.cell(288, 1).value = legal_text
        legal_align = copy(sheet.cell(288, 1).alignment) if sheet.cell(288, 1).alignment else Alignment()
        legal_align.wrap_text = True
        legal_align.vertical = 'top'
        sheet.cell(288, 1).alignment = legal_align

        # Firma del responsable del área dinámica.
        sheet.cell(295, 2).value = warehouse_lead
        sheet.cell(295, 6).value = selected_receiver
        sheet.cell(295, 8).value = assets_manager

        # Ajuste visual de líneas de firma para que no se vean pegadas.
        # Primero limpia cualquier merge que se cruce con H:L en filas de firma.
        target_rows = (294, 295, 296)
        target_col_start = 8   # H
        target_col_end = 12    # L
        for m in list(sheet.merged_cells.ranges):
            min_col, min_row, max_col, max_row = m.bounds
            intersects_rows = not (max_row < min(target_rows) or min_row > max(target_rows))
            intersects_cols = not (max_col < target_col_start or min_col > target_col_end)
            if intersects_rows and intersects_cols:
                sheet.unmerge_cells(str(m))

        # Amplía horizontalmente el bloque de "Responsable de activos fijos" a H:L.
        for r in target_rows:
            sheet.merge_cells(start_row=r, start_column=8, end_row=r, end_column=12)

        # Líneas: las dos primeras estándar y la tercera más larga por el bloque más ancho.
        sheet.cell(294, 2).value = '______________________________'
        sheet.cell(294, 6).value = '______________________________'
        sheet.cell(294, 8).value = '_______________________________________________'

        for r, c in [(294, 2), (294, 6), (294, 8)]:
            align = copy(sheet.cell(r, c).alignment) if sheet.cell(r, c).alignment else Alignment()
            align.horizontal = 'center'
            align.vertical = 'center'
            sheet.cell(r, c).alignment = align

        # Asegura texto exacto bajo la tercera firma.
        sheet.cell(296, 2).value = 'LIDER DE ALMACEN'
        sheet.cell(296, 6).value = 'RESPONSABLE DE AREA'
        sheet.cell(296, 8).value = 'RESPONSABLE DE ACTIVOS FIJOS'

        # Centrado de nombres/cargos de firma para mejor presentación.
        for r, c in [(295, 2), (295, 6), (295, 8), (296, 2), (296, 6), (296, 8)]:
            align = copy(sheet.cell(r, c).alignment) if sheet.cell(r, c).alignment else Alignment()
            align.horizontal = 'center'
            align.vertical = 'center'
            align.wrap_text = True
            sheet.cell(r, c).alignment = align

        if logo_path is not None:
            try:
                img = XLImage(logo_path)
                img.anchor = 'A2'
                fit_logo_to_a22_box(sheet, img, from_col=1, to_col=2, from_row=2, to_row=5)
                sheet.add_image(img)
            except Exception:
                pass
        elif logo_bytes is not None:
            try:
                image_stream = BytesIO(logo_bytes)
                pil_img = PILImage.open(image_stream)
                img = XLImage(pil_img)
                img.anchor = 'A2'
                fit_logo_to_a22_box(sheet, img, from_col=1, to_col=2, from_row=2, to_row=5)
                sheet.add_image(img)
            except Exception:
                pass

    chunks = [assets_scope[i:i + capacity] for i in range(0, len(assets_scope), capacity)]
    template_ws = ws
    template_ws.title = 'A22 1'

    for chunk_index, chunk_assets in enumerate(chunks, start=1):
        if chunk_index == 1:
            ws_chunk = template_ws
        else:
            ws_chunk = wb.copy_worksheet(template_ws)
            ws_chunk.title = f'A22 {chunk_index}'

        apply_header_and_signature(ws_chunk)

        for row_idx in data_rows:
            for col in [2, 3, 6, 7, 8, 9, 10]:
                ws_chunk.cell(row_idx, col).value = None

        for idx, asset in enumerate(chunk_assets):
            row_idx = data_rows[idx]
            ws_chunk.cell(row_idx, 2).value = asset.c_act
            ws_chunk.cell(row_idx, 3).value = asset.nom or ''
            ws_chunk.cell(row_idx, 6).value = asset.des_ubi or ''
            asset_obs = (per_asset_observations.get(str(asset.c_act)) or '').strip()
            ws_chunk.cell(row_idx, 7).value = asset_obs or observation or (asset.observacion_inventario or '')
            ws_chunk.cell(row_idx, 8).value = classify_asset_group(asset)
            ws_chunk.cell(row_idx, 9).value = reference_serial(asset)
            ws_chunk.cell(row_idx, 10).value = asset.modelo or ''

            # Ajuste visual para textos largos.
            text_candidates = [
                ws_chunk.cell(row_idx, 3).value or '',
                ws_chunk.cell(row_idx, 6).value or '',
                ws_chunk.cell(row_idx, 7).value or '',
                ws_chunk.cell(row_idx, 9).value or '',
                ws_chunk.cell(row_idx, 10).value or '',
            ]
            max_len = max(len(str(x)) for x in text_candidates)
            lines = max(1, min(8, (max_len // 20) + 1))
            ws_chunk.row_dimensions[row_idx].height = max(24, 14 * lines)
            for col in [3, 6, 7, 8, 9, 10]:
                c = ws_chunk.cell(row_idx, col)
                base_align = copy(c.alignment) if c.alignment else Alignment()
                base_align.wrap_text = True
                base_align.vertical = 'top'
                c.alignment = base_align

        # Mantiene formato y celdas combinadas: solo oculta filas sobrantes del bloque de activos.
        used = len(chunk_assets)
        for i, row_idx in enumerate(data_rows):
            ws_chunk.row_dimensions[row_idx].hidden = i >= used

    out = BytesIO()
    wb.save(out)
    base_name = run.name if run else selected_service
    safe_name = clean_filename(base_name)
    filename = f'a22_inventario_{safe_name}.xlsx'
    content = out.getvalue()
    date_parts = selected_date.split('-') if selected_date else []
    year_num = int(date_parts[0]) if len(date_parts) == 3 and date_parts[0].isdigit() else now_local_dt().year
    month_num = int(date_parts[1]) if len(date_parts) == 3 and date_parts[1].isdigit() else now_local_dt().month
    period_label = f'{selected_service} - {selected_date}'
    persist_generated_report_file(
        content=content,
        report_type='a22_excel',
        title='Acta A22 - Excel',
        period_label=period_label,
        period_id=period_id,
        file_name=filename,
        folder_group='a22',
        year=year_num,
        month=month_num,
    )
    return send_file(BytesIO(content), download_name=filename, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/export_a22_pdf', methods=['GET', 'POST'])
def export_a22_pdf():
    ensure_db()
    payload = request.get_json(silent=True) if request.method == 'POST' else None
    source = payload or request.args
    service = (source.get('service') or '').strip()
    run_id = source.get('run_id')
    try:
        run_id = int(run_id) if run_id not in (None, '') else None
    except Exception:
        run_id = None
    period_id = source.get('period_id')
    try:
        period_id = int(period_id) if period_id not in (None, '') else None
    except Exception:
        period_id = None
    receiver = (source.get('receiver') or '').strip()
    observation = (source.get('observation') or '').strip()
    report_date = (source.get('report_date') or '').strip()
    warehouse_lead = (source.get('warehouse_lead') or '').strip()
    assets_manager = (source.get('assets_manager') or '').strip()
    per_asset_observations = source.get('per_asset_observations') or {}
    if not isinstance(per_asset_observations, dict):
        per_asset_observations = {}
    if not period_id:
        return jsonify({'error': 'Debes seleccionar el periodo para generar A22'}), 400
    if not run_id:
        return jsonify({'error': 'Debes seleccionar la jornada del periodo para generar A22'}), 400
    if not warehouse_lead:
        return jsonify({'error': 'Lider de almacen es obligatorio'}), 400
    if not assets_manager:
        return jsonify({'error': 'Responsable de activos fijos es obligatorio'}), 400
    if report_date:
        try:
            datetime.strptime(report_date, '%Y-%m-%d')
        except Exception:
            return jsonify({'error': 'Fecha invalida. Usa formato YYYY-MM-DD'}), 400
    selected_date = report_date or now_local_dt().strftime('%Y-%m-%d')

    run, assets_scope, err = get_a22_scope(service=service or None, run_id=run_id, period_id=period_id)
    if err:
        return err
    if not assets_scope:
        return jsonify({'error': 'No hay activos encontrados para generar A22 con ese filtro'}), 400
    if not run or not run.service:
        return jsonify({'error': 'La jornada seleccionada no tiene centro de costo asociado'}), 400

    assets_scope = sort_assets_for_a22(assets_scope)
    selected_service = run.service
    selected_receiver = receiver or assets_scope[0].nom_resp or ''
    work_area = classify_area(selected_service).upper()
    logo_path = get_hospital_logo_path()
    cod_path = get_codificacion_path()

    out = BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=letter,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=22 * mm,
        bottomMargin=10 * mm,
    )
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle('a22h', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#0B4F6C'))
    normal = ParagraphStyle('a22n', parent=styles['Normal'], fontSize=8, leading=10)
    centered = ParagraphStyle('a22c', parent=normal, alignment=1)
    legal_style = ParagraphStyle('a22legal', parent=styles['Normal'], fontSize=8, leading=11, textColor=colors.HexColor('#1F2937'))

    story = []
    story.append(Paragraph('FORMATO A22 - INVENTARIO GENERAL DE ACTIVOS FIJOS', header_style))
    story.append(Spacer(1, 4))
    meta_data = [
        [Paragraph('<b>Fecha</b>', normal), Paragraph(selected_date, normal)],
        [Paragraph('<b>Centro de costo</b>', normal), Paragraph(selected_service, normal)],
        [Paragraph('<b>Responsable centro de costo</b>', normal), Paragraph(selected_receiver, normal)],
        [Paragraph('<b>Cargo</b>', normal), Paragraph(f'LIDER {selected_service}', normal)],
        [Paragraph('<b>Area de trabajo</b>', normal), Paragraph(work_area, normal)],
        [Paragraph('<b>Cantidad activos entregados</b>', normal), Paragraph(str(len(assets_scope)), normal)],
    ]
    meta_table = Table(meta_data, colWidths=[48 * mm, 130 * mm])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EAF4FA')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 6))

    data = [[
        Paragraph('<b>N°</b>', normal),
        Paragraph('<b>CODIGO ACTIVO FIJO</b>', normal),
        Paragraph('<b>DESCRIPCION ACTIVO FIJO</b>', normal),
        Paragraph('<b>UBICACION</b>', normal),
        Paragraph('<b>OBSERVACION</b>', normal),
        Paragraph('<b>TIPO ACTIVO</b>', normal),
        Paragraph('<b>REFERENCIA/SERIAL</b>', normal),
        Paragraph('<b>MODELO</b>', normal),
    ]]
    for i, asset in enumerate(assets_scope, start=1):
        asset_obs = (per_asset_observations.get(str(asset.c_act)) or '').strip()
        data.append([
            Paragraph(str(i), normal),
            Paragraph(str(asset.c_act or ''), normal),
            Paragraph(str(asset.nom or ''), normal),
            Paragraph(str(asset.des_ubi or ''), normal),
            Paragraph(asset_obs or observation or str(asset.observacion_inventario or ''), normal),
            Paragraph(classify_asset_group(asset), normal),
            Paragraph(reference_serial(asset), normal),
            Paragraph(str(asset.modelo or ''), normal),
        ])
    table = Table(data, colWidths=[10 * mm, 24 * mm, 48 * mm, 32 * mm, 28 * mm, 22 * mm, 30 * mm, 20 * mm], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0B4F6C')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD5E3')),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D7E5EE')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FBFD')]),
    ]))
    story.append(table)

    legal_text = (
        f'Una vez culminado el proceso de inventario general en el area de {selected_service} '
        f'se entrega a {selected_receiver} responsable de dicha area, un documento detallado '
        f'que incluye todos los activos fijos asignados, clasificados y verificados. '
        f'A partir de la entrega formal de este documento, el responsable del area asume la '
        f'obligacion de velar por la custodia y el buen estado de cada activo listado. '
        f'En caso de perdida, dano o cualquier irregularidad que afecte los activos bajo su '
        f'supervision, el responsable debera presentar una justificacion oportuna y detallada, '
        f'y asumir las consecuencias correspondientes. Esta medida busca asegurar la transparencia '
        f'y la adecuada gestion de los recursos de la institucion.'
    )
    story.append(Spacer(1, 8))
    story.append(Paragraph(legal_text, legal_style))
    story.append(Spacer(1, 28))

    sign_table = Table([
        [
            Paragraph('______________________________', centered),
            Paragraph('______________________________', centered),
            Paragraph('______________________________', centered),
        ],
        [
            Paragraph(warehouse_lead, centered),
            Paragraph(selected_receiver, centered),
            Paragraph(assets_manager, centered),
        ],
        [
            Paragraph('LIDER DE ALMACEN', centered),
            Paragraph('RESPONSABLE DE AREA', centered),
            Paragraph('RESPONSABLE DE ACTIVOS FIJOS', centered),
        ],
    ], colWidths=[60 * mm, 60 * mm, 60 * mm])
    sign_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(sign_table)
    page_header = make_pdf_page_header(logo_path, right_image_path=cod_path)
    doc.build(story, onFirstPage=page_header, onLaterPages=page_header)
    content = out.getvalue()
    safe_name = str(selected_service).replace(' ', '_').replace('/', '_')
    filename = clean_filename(f'a22_inventario_{safe_name}.pdf')
    date_parts = selected_date.split('-') if selected_date else []
    year_num = int(date_parts[0]) if len(date_parts) == 3 and date_parts[0].isdigit() else now_local_dt().year
    month_num = int(date_parts[1]) if len(date_parts) == 3 and date_parts[1].isdigit() else now_local_dt().month
    period_label = f'{selected_service} - {selected_date}'
    persist_generated_report_file(
        content=content,
        report_type='a22_pdf',
        title='Acta A22 - PDF',
        period_label=period_label,
        period_id=period_id,
        file_name=filename,
        folder_group='a22',
        year=year_num,
        month=month_num,
    )
    return send_file(BytesIO(content), download_name=filename, as_attachment=True, mimetype='application/pdf')


@app.route('/reconciliation/export_found', methods=['GET'])
def reconciliation_export_found():
    ensure_db()
    service = (request.args.get('service') or '').strip()
    run_id = request.args.get('run_id', type=int)
    period_id = request.args.get('period_id', type=int)
    if not run_id and not period_id:
        return jsonify({'error': 'Debes seleccionar periodo o jornada para exportar'}), 400
    rows, err = build_reconciliation_rows(service=service, run_id=run_id, period_id=period_id)
    if err:
        return err
    found_rows = [r for r in rows if r['ESTADO_INVENTARIO'] == 'Encontrado']
    if not found_rows:
        return jsonify({'error': 'No hay activos encontrados para exportar'}), 400

    grouped = {}
    for row in found_rows:
        svc = str(row.get('SERVICIO') or '').strip() or 'SIN SERVICIO'
        grouped.setdefault(svc, []).append(row)

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()
    for svc in sorted(grouped.keys(), key=lambda x: x.casefold()):
        sheet_name = excel_safe_sheet_name(svc, used_names)
        ws = wb.create_sheet(title=sheet_name)
        svc_rows = sorted(grouped[svc], key=lambda r: str(r.get('C_ACT') or ''))
        title = f'Base depurada - Encontrados ({svc})'
        write_reconciliation_sheet(ws, title, svc_rows)
        add_logo_to_excel_sheet(ws, logo_path=get_hospital_logo_path())

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    name = clean_filename(f"base_depurada_encontrados_{(service or 'todos').replace(' ', '_')}.xlsx")
    return send_file(out, as_attachment=True, download_name=name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/reconciliation/export_not_found', methods=['GET'])
def reconciliation_export_not_found():
    ensure_db()
    service = (request.args.get('service') or '').strip()
    run_id = request.args.get('run_id', type=int)
    period_id = request.args.get('period_id', type=int)
    if not run_id and not period_id:
        return jsonify({'error': 'Debes seleccionar periodo o jornada para exportar'}), 400
    rows, err = build_reconciliation_rows(service=service, run_id=run_id, period_id=period_id)
    if err:
        return err
    not_found_rows = [r for r in rows if r['ESTADO_INVENTARIO'] == 'No encontrado']

    grouped = {}
    for row in not_found_rows:
        svc = str(row.get('SERVICIO') or '').strip() or 'SIN SERVICIO'
        grouped.setdefault(svc, []).append(row)

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()
    if grouped:
        for svc in sorted(grouped.keys(), key=lambda x: x.casefold()):
            sheet_name = excel_safe_sheet_name(svc, used_names)
            ws = wb.create_sheet(title=sheet_name)
            svc_rows = sorted(grouped[svc], key=lambda r: str(r.get('C_ACT') or ''))
            title = f'Listado no encontrados ({svc})'
            write_reconciliation_sheet(ws, title, svc_rows)
            add_logo_to_excel_sheet(ws, logo_path=get_hospital_logo_path())
    else:
        ws = wb.create_sheet(title=excel_safe_sheet_name('SIN_NO_ENCONTRADOS', used_names))
        write_reconciliation_sheet(ws, 'Listado no encontrados (sin registros)', [])
        add_logo_to_excel_sheet(ws, logo_path=get_hospital_logo_path())

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    name = clean_filename(f"listado_no_encontrados_{(service or 'todos').replace(' ', '_')}.xlsx")
    return send_file(out, as_attachment=True, download_name=name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/reconciliation/export_consolidated', methods=['GET'])
def reconciliation_export_consolidated():
    ensure_db()
    service = (request.args.get('service') or '').strip()
    run_id = request.args.get('run_id', type=int)
    period_id = request.args.get('period_id', type=int)
    if not run_id and not period_id:
        return jsonify({'error': 'Debes seleccionar periodo o jornada para exportar'}), 400
    rows, err = build_reconciliation_rows(service=service, run_id=run_id, period_id=period_id)
    if err:
        return err
    if not rows:
        return jsonify({'error': 'No hay activos para exportar'}), 400

    found_rows = [r for r in rows if r['ESTADO_INVENTARIO'] == 'Encontrado']
    not_found_rows = [r for r in rows if r['ESTADO_INVENTARIO'] == 'No encontrado']
    pending_rows = [r for r in rows if r['ESTADO_INVENTARIO'] == 'Pendiente']

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = 'Resumen'
    ws_summary.append(['CONSOLIDADO FINAL DE INVENTARIO'])
    ws_summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws_summary['A1'].font = Font(bold=True, size=14, color='0B4F6C')
    ws_summary.append([f"Servicio: {service or 'TODOS'}", f"Fecha: {now_local_dt().strftime('%Y-%m-%d %H:%M')}"])
    ws_summary.append(['Estado', 'Cantidad', 'Total costo', 'Total saldo'])
    ws_summary.append(['Encontrado', len(found_rows), sum(x['COSTO'] for x in found_rows), sum(x['SALDO'] for x in found_rows)])
    ws_summary.append(['No encontrado', len(not_found_rows), sum(x['COSTO'] for x in not_found_rows), sum(x['SALDO'] for x in not_found_rows)])
    ws_summary.append(['Pendiente', len(pending_rows), sum(x['COSTO'] for x in pending_rows), sum(x['SALDO'] for x in pending_rows)])
    ws_summary.append(['TOTAL', len(rows), sum(x['COSTO'] for x in rows), sum(x['SALDO'] for x in rows)])
    for col in ['A', 'B', 'C', 'D']:
        ws_summary.column_dimensions[col].width = [22, 12, 18, 18][ord(col) - ord('A')]
    for r in range(4, ws_summary.max_row + 1):
        ws_summary.cell(r, 3).number_format = '"$"#,##0'
        ws_summary.cell(r, 4).number_format = '"$"#,##0'
    add_logo_to_excel_sheet(ws_summary, logo_path=get_hospital_logo_path())

    ws_found = wb.create_sheet('Encontrados')
    write_reconciliation_sheet(ws_found, 'Activos encontrados', found_rows)
    add_logo_to_excel_sheet(ws_found, logo_path=get_hospital_logo_path())

    ws_not = wb.create_sheet('No encontrados')
    write_reconciliation_sheet(ws_not, 'Activos no encontrados', not_found_rows)
    add_logo_to_excel_sheet(ws_not, logo_path=get_hospital_logo_path())

    ws_pending = wb.create_sheet('Pendientes')
    write_reconciliation_sheet(ws_pending, 'Activos pendientes', pending_rows)
    add_logo_to_excel_sheet(ws_pending, logo_path=get_hospital_logo_path())

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    name = clean_filename(f"consolidado_final_inventario_{(service or 'todos').replace(' ', '_')}.xlsx")
    return send_file(out, as_attachment=True, download_name=name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


