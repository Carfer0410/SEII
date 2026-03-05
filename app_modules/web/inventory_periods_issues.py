from .pages_imports import *

@app.route('/services')
def services():
    ensure_db()
    raw_services = db.session.query(Asset.nom_ccos).all()
    cleaned = set()
    for row in raw_services:
        value = row[0]
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        if is_excluded_service_name(text):
            continue
        cleaned.add(text)
    services = sorted(cleaned, key=lambda x: x.casefold())
    return jsonify({'services': services, 'total': len(services)})


@app.route('/responsibles')
def responsibles():
    ensure_db()
    scoped = str(request.args.get('scoped') or '').strip().lower() in {'1', 'true', 'si', 'yes'}
    period_id = request.args.get('period_id', type=int)
    run_id = request.args.get('run_id', type=int)
    service = (request.args.get('service') or '').strip()

    q = Asset.query
    if scoped:
        run = None
        if run_id:
            run = InventoryRun.query.get(run_id)
            if not run:
                return jsonify({'error': 'Jornada no encontrada'}), 404
            if period_id and run.period_id != period_id:
                return jsonify({'error': 'La jornada no pertenece al periodo seleccionado'}), 400
        if run:
            q = apply_run_scope_filter(q, run)
        elif period_id:
            runs = InventoryRun.query.filter(InventoryRun.period_id == period_id).all()
            scope_services = set()
            for r in runs:
                for s in run_scope_services(r):
                    s_clean = str(s or '').strip()
                    if s_clean:
                        scope_services.add(s_clean)
            if scope_services:
                q = q.filter(Asset.nom_ccos.in_(sorted(scope_services)))

    if scoped and service:
        q = q.filter(Asset.nom_ccos == service)

    raw = q.with_entities(Asset.nom_resp).all()
    cleaned = set()
    for row in raw:
        value = row[0]
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        cleaned.add(text)
    items = sorted(cleaned, key=lambda x: x.casefold())
    return jsonify({'responsibles': items, 'total': len(items)})


@app.route('/periods', methods=['GET'])
def list_periods():
    ensure_db()
    status = (request.args.get('status') or '').strip().lower()
    q = InventoryPeriod.query
    if status in {'open', 'closed', 'cancelled'}:
        q = q.filter(InventoryPeriod.status == status)
    periods = q.order_by(InventoryPeriod.id.desc()).all()
    return jsonify({'periods': [p.to_dict() for p in periods]})


@app.route('/periods', methods=['POST'])
def create_period():
    ensure_db()
    data = request.get_json() or {}
    name = (data.get('name') or '').strip()
    period_type = (data.get('period_type') or 'semestral').strip().lower()
    start_date = (data.get('start_date') or '').strip() or None
    end_date = (data.get('end_date') or '').strip() or None
    notes = (data.get('notes') or '').strip() or None

    if not name:
        return jsonify({'error': 'Nombre del periodo es obligatorio'}), 400
    if period_type not in {'semestral', 'aleatorio', 'historico'}:
        return jsonify({'error': 'Tipo de periodo invalido'}), 400
    if InventoryPeriod.query.filter(db.func.upper(InventoryPeriod.name) == name.upper()).first():
        return jsonify({'error': 'Ya existe un periodo con ese nombre'}), 400

    period = InventoryPeriod(
        name=name,
        period_type=period_type,
        start_date=start_date,
        end_date=end_date,
        status='open',
        notes=notes,
        created_at=now_iso(),
    )
    db.session.add(period)
    db.session.commit()
    return jsonify({'period': period.to_dict()})


@app.route('/periods/<int:period_id>/close', methods=['POST'])
def close_period(period_id):
    ensure_db()
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404
    if period.status == 'cancelled':
        return jsonify({'error': 'No puedes cerrar un periodo anulado'}), 400
    total_runs = InventoryRun.query.filter_by(period_id=period.id).count()
    if total_runs <= 0:
        return jsonify({'error': 'No puedes cerrar un periodo sin jornadas registradas'}), 400
    active_runs = InventoryRun.query.filter_by(period_id=period.id, status='active').count()
    if active_runs:
        return jsonify({'error': 'No puedes cerrar el periodo con jornadas activas'}), 400
    period.status = 'closed'
    db.session.commit()
    return jsonify({'period': period.to_dict()})


@app.route('/periods/<int:period_id>/cancel', methods=['POST'])
def cancel_period(period_id):
    ensure_db()
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404
    if period.status == 'cancelled':
        return jsonify({'error': 'El periodo ya esta anulado'}), 400

    data = request.get_json() or {}
    reason = (data.get('reason') or '').strip()
    user = (data.get('user') or '').strip() or 'usuario_movil'
    if not reason:
        return jsonify({'error': 'Debes indicar el motivo de anulacion del periodo'}), 400

    runs = InventoryRun.query.filter_by(period_id=period.id).all()
    if any((r.status or '').strip().lower() == 'active' for r in runs):
        return jsonify({'error': 'No puedes anular el periodo porque tiene jornadas activas'}), 400

    run_ids = [r.id for r in runs]
    has_scan_trace = False
    if run_ids:
        has_scan_trace = db.session.query(RunAssetStatus.id).filter(
            RunAssetStatus.run_id.in_(run_ids)
        ).first() is not None
    if has_scan_trace:
        return jsonify({'error': 'No puedes anular el periodo porque ya tiene trazabilidad de escaneo'}), 400

    has_issues = AssetIssue.query.filter_by(period_id=period.id).count() > 0
    if has_issues:
        return jsonify({'error': 'No puedes anular el periodo porque tiene novedades registradas'}), 400

    has_disposals = AssetDisposal.query.filter_by(period_id=period.id).count() > 0
    if has_disposals:
        return jsonify({'error': 'No puedes anular el periodo porque tiene bajas asociadas'}), 400

    for run in runs:
        run.status = 'cancelled'
        run.closed_at = run.closed_at or now_iso()
        run.cancelled_at = now_iso()
        run.cancelled_by = user
        run.cancel_reason = f'Anulada por anulacion de periodo: {reason}'

    period.status = 'cancelled'
    period.cancelled_at = now_iso()
    period.cancelled_by = user
    period.cancel_reason = reason
    db.session.commit()
    return jsonify({'period': period.to_dict(), 'cancelled_runs': len(runs)})


@app.route('/periods/<int:period_id>/service_coverage', methods=['GET'])
def period_service_coverage(period_id):
    ensure_db()
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404

    raw_services = db.session.query(Asset.nom_ccos).all()
    base_services = sorted({
        str(row[0]).strip()
        for row in raw_services
        if row[0] is not None and str(row[0]).strip() and not is_excluded_service_name(str(row[0]).strip())
    }, key=lambda x: x.casefold())

    runs = InventoryRun.query.filter_by(period_id=period_id).all()
    run_ids = [r.id for r in runs]
    run_service_map = {r.id: run_scope_services(r) for r in runs}

    done_services = set()
    if run_ids:
        status_rows = db.session.query(RunAssetStatus.run_id).filter(
            RunAssetStatus.run_id.in_(run_ids)
        ).distinct().all()
        for row in status_rows:
            for svc in run_service_map.get(row[0], []):
                if svc and not is_excluded_service_name(svc):
                    done_services.add(svc)

    # Si no hay registros de escaneo aun, considera servicios con jornada cerrada como gestionados.
    if not done_services:
        for r in runs:
            for svc in run_scope_services(r):
                if svc and (not is_excluded_service_name(svc)) and r.status == 'closed':
                    done_services.add(svc)

    pending_services = [s for s in base_services if s not in done_services]
    total_services = len(base_services)
    done_count = len(done_services)
    pending_count = len(pending_services)
    done_pct = round((done_count / total_services) * 100, 2) if total_services else 0.0
    pending_pct = round((pending_count / total_services) * 100, 2) if total_services else 0.0

    workload_rows = db.session.query(
        Asset.nom_ccos,
        db.func.count(Asset.id),
        db.func.coalesce(db.func.sum(Asset.costo), 0),
    ).filter(
        Asset.nom_ccos.isnot(None)
    ).group_by(
        Asset.nom_ccos
    ).all()
    workload_map = {}
    for svc, cnt, total_cost in workload_rows:
        svc_name = str(svc or '').strip()
        if not svc_name or is_excluded_service_name(svc_name):
            continue
        workload_map[svc_name] = {
            'asset_count': int(cnt or 0),
            'total_cost': float(total_cost or 0),
        }

    service_rows = []
    for svc in base_services:
        # Obtener todos los activos de este servicio
        assets = Asset.query.filter(Asset.nom_ccos == svc).all()
        asset_ids = [a.id for a in assets]
        total_assets = len(asset_ids)
        # Buscar el último status de cada activo en las jornadas de este periodo
        found_count = 0
        if asset_ids and run_ids:
            # Buscar el último status registrado para cada activo en este periodo
            subq = db.session.query(
                RunAssetStatus.asset_id,
                db.func.max(RunAssetStatus.id).label('max_id')
            ).filter(
                RunAssetStatus.run_id.in_(run_ids),
                RunAssetStatus.asset_id.in_(asset_ids)
            ).group_by(RunAssetStatus.asset_id).subquery()
            latest_statuses = db.session.query(RunAssetStatus).join(
                subq, (RunAssetStatus.asset_id == subq.c.asset_id) & (RunAssetStatus.id == subq.c.max_id)
            ).all()
            found_count = sum(1 for s in latest_statuses if s.status == 'Encontrado')
        status_pct = round((found_count / total_assets) * 100, 2) if total_assets else 0

        # Determinar el estado: 'Inventariado', 'En proceso' o 'Pendiente'
        # Un servicio solo es 'Inventariado' si todas sus jornadas están cerradas y está en done_services
        # Si tiene una jornada activa, debe decir 'En proceso' aunque esté en done_services
        in_active_run = any(
            r.status == 'active' and svc in run_scope_services(r)
            for r in runs
        )
        if in_active_run:
            status_label = 'En proceso'
        elif svc in done_services:
            status_label = 'Inventariado'
        else:
            status_label = 'Pendiente'

        service_rows.append({
            'service': svc,
            'status': status_label,
            'status_pct': status_pct,
            'asset_count': int(workload_map.get(svc, {}).get('asset_count', 0)),
            'total_cost': float(workload_map.get(svc, {}).get('total_cost', 0)),
        })

    # Recomendaciones operativas por carga (mayor cantidad de activos primero)
    pending_ranked = sorted(
        [r for r in service_rows if r['status'] == 'Pendiente'],
        key=lambda x: (x.get('asset_count', 0), x.get('total_cost', 0)),
        reverse=True
    )

    recommendations = []
    for idx, row in enumerate(pending_ranked[:5], start=1):
        recommendations.append({
            'priority': idx,
            'service': row['service'],
            'reason': f"Alta carga operativa ({row.get('asset_count', 0)} activos).",
        })

    def cluster_key(service_name):
        txt = str(service_name or '').upper().strip()
        if 'URGEN' in txt:
            return 'URGENCIAS'
        if 'HOSPITAL' in txt:
            return 'HOSPITALIZACION'
        if 'CIRUG' in txt or 'QUIROF' in txt:
            return 'CIRUGIA'
        return txt.split(' ')[0] if txt else 'OTROS'

    cluster_counts = {}
    for row in pending_ranked:
        key = cluster_key(row['service'])
        cluster_counts[key] = cluster_counts.get(key, 0) + 1

    grouped_tips = []
    for key, count in sorted(cluster_counts.items(), key=lambda x: x[1], reverse=True):
        if count >= 2:
            grouped_tips.append(
                f"{key}: {count} subservicios pendientes. Recomendado ejecutarlos el mismo dia con jornadas separadas por subservicio."
            )

    return jsonify({
        'period': period.to_dict(),
        'summary': {
            'total_services': total_services,
            'done_services': done_count,
            'pending_services': pending_count,
            'done_pct': done_pct,
            'pending_pct': pending_pct,
        },
        'services': service_rows,
        'recommendations': recommendations,
        'grouped_tips': grouped_tips,
    })


@app.route('/periods/<int:period_id>/closed_services', methods=['GET'])
def period_closed_services(period_id):
    ensure_db()
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404

    closed_runs = InventoryRun.query.filter_by(period_id=period_id, status='closed').all()
    service_map = {}
    for run in closed_runs:
        run_name = str(run.name or '').strip() or f'Jornada {run.id}'
        for svc in run_scope_services(run):
            svc_name = str(svc or '').strip()
            if not svc_name:
                continue
            if svc_name not in service_map:
                service_map[svc_name] = []
            service_map[svc_name].append({
                'run_id': run.id,
                'run_name': run_name,
                'closed_at': run.closed_at,
                'closed_at_local': format_dt_local(run.closed_at),
            })

    items = []
    for svc_name in sorted(service_map.keys(), key=lambda x: x.casefold()):
        runs = sorted(service_map[svc_name], key=lambda x: x.get('run_id') or 0)
        items.append({
            'service': svc_name,
            'runs': runs,
            'last_run_name': runs[-1].get('run_name') if runs else '',
            'last_closed_at': runs[-1].get('closed_at') if runs else '',
            'last_closed_at_local': runs[-1].get('closed_at_local') if runs else '',
        })

    return jsonify({
        'period': period.to_dict(),
        'total_closed_services': len(items),
        'items': items,
    })


def detect_asset_issues_for_period(period_id, analyze_base=False):
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return None, 'Periodo no encontrado'

    runs = InventoryRun.query.filter_by(period_id=period_id).all()
    run_ids = [r.id for r in runs]
    run_by_id = {r.id: r for r in runs}
    run_services = sorted({
        svc
        for r in runs
        for svc in run_scope_services(r)
        if svc
    })

    q_assets = Asset.query
    if analyze_base:
        assets = q_assets.all()
    else:
        if not run_services:
            assets = []
        else:
            assets = q_assets.filter(Asset.nom_ccos.in_(run_services)).all()
    assets_by_id = {a.id: a for a in assets}

    latest_status_map = {}
    status_rows = []
    if run_ids and assets:
        status_rows = RunAssetStatus.query.filter(
            RunAssetStatus.run_id.in_(run_ids),
            RunAssetStatus.asset_id.in_([a.id for a in assets])
        ).order_by(RunAssetStatus.id.desc()).all()
        for s in status_rows:
            if s.asset_id not in latest_status_map:
                latest_status_map[s.asset_id] = s

    duplicate_codes = {}
    rows_dup = db.session.query(Asset.c_act, db.func.count(Asset.id)).group_by(Asset.c_act).having(db.func.count(Asset.id) > 1).all()
    for code, count in rows_dup:
        duplicate_codes[str(code or '').strip()] = int(count or 0)

    duplicate_intelligent = {}
    for a in assets:
        payload = asset_raw_payload(a)
        c_int = str(payload.get('CODINTELIGENTE') or '').strip()
        if c_int:
            duplicate_intelligent[c_int] = duplicate_intelligent.get(c_int, 0) + 1

    disposal_by_asset = {d.asset_id: d for d in AssetDisposal.query.all()}

    now_iso_value = now_iso()
    AssetIssue.query.filter_by(period_id=period_id, source='auto').delete()
    db.session.flush()

    created = 0
    for a in assets:
        latest = latest_status_map.get(a.id)
        status = normalize_inventory_status(latest.status if latest else a.estado_inventario)
        value = asset_book_value(a)
        critical = classify_critical_asset(a)

        def add_issue(issue_type, severity, title, description, run_id=None):
            nonlocal created
            db.session.add(AssetIssue(
                issue_type=issue_type,
                title=title,
                severity=severity,
                status='Nuevo',
                source='auto',
                period_id=period_id,
                run_id=run_id,
                asset_id=a.id,
                service=a.nom_ccos or '',
                detected_value=value,
                description=description,
                created_at=now_iso_value,
                updated_at=now_iso_value,
            ))
            created += 1

        if status == 'No encontrado' and critical.get('is_critical'):
            add_issue(
                'NOT_FOUND_CRITICAL',
                'Alta',
                'Activo critico no encontrado',
                f"Estado='{status}' | Criticidad='{critical.get('critical_reasons')}' | Valor aprox={money_text(value)}."
            )
        if status == 'No encontrado' and value >= 20_000_000:
            add_issue(
                'NOT_FOUND_HIGH_VALUE',
                'Alta',
                'Activo no encontrado de alto valor',
                f"Estado='{status}' | Valor contable aprox={money_text(value)}."
            )

        if not str(a.serie or '').strip() and not str(a.ref or '').strip():
            add_issue(
                'MISSING_SERIAL_REF',
                'Media',
                'Activo sin serial ni referencia',
                f"SERIE='{str(a.serie or '').strip() or 'vacio'}' | REF='{str(a.ref or '').strip() or 'vacio'}'."
            )
        if not str(a.modelo or '').strip() and not str(a.nom_marca or '').strip():
            add_issue(
                'MISSING_MODEL_BRAND',
                'Baja',
                'Activo sin marca y modelo',
                f"MARCA='{str(a.nom_marca or '').strip() or 'vacio'}' | MODELO='{str(a.modelo or '').strip() or 'vacio'}'."
            )
        if not str(a.nom_resp or '').strip() or not str(a.des_ubi or '').strip():
            add_issue(
                'MISSING_CUSTODY_DATA',
                'Media',
                'Activo con datos de custodia incompletos',
                f"RESPONSABLE='{str(a.nom_resp or '').strip() or 'vacio'}' | UBICACION='{str(a.des_ubi or '').strip() or 'vacio'}'."
            )
        if to_number(a.costo) <= 0 or to_number(a.saldo) < 0:
            add_issue(
                'INVALID_FINANCIAL_VALUES',
                'Alta',
                'Valores financieros inconsistentes',
                f"Costo={to_number(a.costo)} | Saldo={to_number(a.saldo)}."
            )
        dep_no = is_non_depreciable(a.deprecia)
        vida_zero = is_zero_useful_life(a.vida_util)
        is_control_asset = (classify_asset_group(a) == 'CONTROL')
        if (not is_control_asset) and ((dep_no and (not vida_zero)) or ((not dep_no) and vida_zero)):
            add_issue(
                'DEPRECIATION_INCONSISTENT',
                'Media',
                'Inconsistencia entre deprecia y vida util',
                f"DEPRECIA='{a.deprecia or ''}' | VIDA_UTIL='{a.vida_util or ''}'."
            )
        if run_ids and status == 'Pendiente':
            add_issue(
                'PENDING_UNSCANNED',
                'Media',
                'Activo pendiente sin escaneo',
                f"Estado inventario='{status}' sin verificacion en jornada del periodo."
            )

        if str(a.c_act or '').strip() in duplicate_codes:
            add_issue(
                'DUPLICATE_CODE',
                'Alta',
                'Codigo de activo duplicado',
                f"Existen {duplicate_codes[str(a.c_act or '').strip()]} registros con el mismo codigo."
            )
        else:
            payload = asset_raw_payload(a)
            c_int = str(payload.get('CODINTELIGENTE') or '').strip()
            if c_int and duplicate_intelligent.get(c_int, 0) > 1:
                add_issue(
                    'DUPLICATE_CODE',
                    'Media',
                    'Codificacion inteligente repetida',
                    f"CODINTELIGENTE '{c_int}' repetido en {duplicate_intelligent.get(c_int, 0)} activos."
                )

        disp = disposal_by_asset.get(a.id)
        if disp and str(disp.status or '').strip().lower() in {'pendiente baja', 'en analisis', 'pendiente'}:
            sev = 'Alta' if value >= 10_000_000 else 'Media'
            add_issue(
                'CANDIDATE_DISPOSAL',
                sev,
                'Activo con baja pendiente',
                f"Estado baja='{disp.status}' | Valor aprox={money_text(value)}."
            )

    # Revisiones por escaneo en servicio distinto
    for s in status_rows:
        if normalize_inventory_status(s.status) != 'Encontrado':
            continue
        run = run_by_id.get(s.run_id)
        a = assets_by_id.get(s.asset_id)
        if not run or not a:
            continue
        run_scope = run_scope_services(run)
        run_service = str(run.service or '').strip()
        asset_service = str(a.nom_ccos or '').strip()
        scope_cf = {x.casefold() for x in run_scope}
        if run_scope and asset_service and asset_service.casefold() not in scope_cf:
            run_label = ', '.join(run_scope[:3]) + (' ...' if len(run_scope) > 3 else '')
            for issue_type, title in [
                ('SCANNED_OTHER_SERVICE', 'Escaneado en servicio distinto'),
                ('LOCATION_REVIEW', 'Revision de ubicacion requerida'),
                ('RESPONSIBLE_REVIEW', 'Revision de responsable requerida'),
            ]:
                db.session.add(AssetIssue(
                    issue_type=issue_type,
                    title=title,
                    severity='Media',
                    status='Nuevo',
                    source='auto',
                    period_id=period_id,
                    run_id=run.id,
                    asset_id=a.id,
                    service=run_label or run_service,
                    detected_value=asset_book_value(a),
                    description=f"Escaneado en alcance '{run_label or run_service}' pero base actual indica '{asset_service}'.",
                    created_at=now_iso_value,
                    updated_at=now_iso_value,
                ))
                created += 1

    db.session.commit()
    return {'created': created}, None


@app.route('/issues/scan', methods=['POST'])
def issues_scan():
    ensure_db()
    data = request.get_json() or {}
    period_id = data.get('period_id')
    try:
        period_id = int(period_id)
    except Exception:
        period_id = None
    if not period_id:
        return jsonify({'error': 'Periodo es obligatorio'}), 400
    analyze_base = parse_bool(data.get('analyze_base'), default=False)
    result, err = detect_asset_issues_for_period(period_id, analyze_base=analyze_base)
    if err:
        return jsonify({'error': err}), 400
    return jsonify({'ok': True, **result})


@app.route('/issues', methods=['GET'])
def issues_list():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    status = (request.args.get('status') or '').strip()
    severity = (request.args.get('severity') or '').strip()
    issue_type = (request.args.get('issue_type') or '').strip()

    q = AssetIssue.query
    if period_id:
        q = q.filter(AssetIssue.period_id == period_id)
    if status:
        q = q.filter(AssetIssue.status == status)
    if severity:
        q = q.filter(AssetIssue.severity == severity)
    if issue_type:
        q = q.filter(AssetIssue.issue_type == issue_type)

    rows = q.order_by(AssetIssue.severity.asc(), AssetIssue.id.desc()).all()
    items = [r.to_dict() for r in rows]
    total_value_risk = sum(to_number(x.get('detected_value')) for x in items if x.get('status') != 'Cerrado')

    by_status = {}
    by_severity = {}
    for x in items:
        by_status[x['status']] = by_status.get(x['status'], 0) + 1
        by_severity[x['severity']] = by_severity.get(x['severity'], 0) + 1

    return jsonify({
        'items': items,
        'summary': {
            'total': len(items),
            'open': sum(1 for x in items if x.get('status') != 'Cerrado'),
            'value_risk': round(total_value_risk, 2),
            'by_status': by_status,
            'by_severity': by_severity,
        },
        'meta': {
            'statuses': ISSUE_STATUSES,
            'severities': ISSUE_SEVERITIES,
            'issue_types': [{'key': k, 'label': v} for k, v in ISSUE_TYPE_LABELS.items()],
        }
    })


@app.route('/issues/<int:issue_id>', methods=['PATCH'])
def issues_update(issue_id):
    ensure_db()
    row = AssetIssue.query.get(issue_id)
    if not row:
        return jsonify({'error': 'Novedad no encontrada'}), 404
    data = request.get_json() or {}
    status = str(data.get('status') or '').strip()
    assigned_to = str(data.get('assigned_to') or '').strip()
    due_date = str(data.get('due_date') or '').strip()
    resolution_notes = str(data.get('resolution_notes') or '').strip()
    severity = str(data.get('severity') or '').strip()

    if status and status in ISSUE_STATUSES:
        row.status = status
    if severity and severity in ISSUE_SEVERITIES:
        row.severity = severity
    row.assigned_to = assigned_to or row.assigned_to
    row.due_date = due_date or row.due_date
    row.resolution_notes = resolution_notes or row.resolution_notes
    row.updated_at = now_iso()
    db.session.commit()
    return jsonify({'item': row.to_dict()})


def append_text_note(base_text, extra_text):
    base = str(base_text or '').strip()
    extra = str(extra_text or '').strip()
    if not extra:
        return base
    return f"{base} | {extra}" if base else extra


def build_transfer_acta_pdf_bytes(transfer_row, asset, issue=None):
    out = BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=letter,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=18 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'transfer_title',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#0B4F6C'),
    )
    text_style = ParagraphStyle('transfer_text', parent=styles['Normal'], fontSize=9, leading=11)

    story = []
    story.append(Paragraph('Acta de traslado de activo fijo', title_style))
    story.append(Spacer(1, 6))

    header_data = [
        ['Consecutivo', f'TR-{transfer_row.id:06d}'],
        ['Fecha ejecucion', format_dt_local(transfer_row.executed_at) or format_dt_local(now_iso())],
        ['Activo', f"{asset.c_act or ''} - {asset.nom or ''}"],
        ['Servicio origen', transfer_row.origin_service or 'No definido'],
        ['Servicio destino', transfer_row.target_service or 'No definido'],
        ['Responsable origen', transfer_row.origin_responsible or 'No definido'],
        ['Responsable destino', transfer_row.target_responsible or 'No definido'],
        ['Solicitado por', transfer_row.requested_by or 'No definido'],
        ['Aprobado por', transfer_row.approved_by or 'No definido'],
        ['Ejecutado por', transfer_row.executed_by or 'No definido'],
    ]
    header = Table(header_data, colWidths=[55 * mm, 125 * mm])
    header.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#C9DCE8')),
        ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#DCE8F0')),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F2F8FD')),
    ]))
    story.append(header)
    story.append(Spacer(1, 8))

    story.append(Paragraph('<b>Justificacion del traslado</b>', text_style))
    story.append(Paragraph(escape(transfer_row.justification or 'Sin justificacion registrada.'), text_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph('<b>Observaciones de ejecucion</b>', text_style))
    story.append(Paragraph(escape(transfer_row.execution_notes or 'Sin observaciones adicionales.'), text_style))
    if issue:
        story.append(Spacer(1, 6))
        story.append(Paragraph('<b>Novedad origen</b>', text_style))
        story.append(Paragraph(escape(f"{issue.title or ''}: {issue.description or ''}"), text_style))

    story.append(Spacer(1, 12))
    sign = Table([
        ['Firma entrega', 'Firma recibe', 'Vo.Bo Activos Fijos'],
        ['\n\n\n', '\n\n\n', '\n\n\n'],
    ], colWidths=[60 * mm, 60 * mm, 60 * mm])
    sign.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor('#DCE8F0')),
        ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#E7EEF4')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F7FBFD')),
    ]))
    story.append(sign)

    doc.build(
        story,
        onFirstPage=make_pdf_page_header(get_hospital_logo_path()),
        onLaterPages=make_pdf_page_header(get_hospital_logo_path()),
    )
    out.seek(0)
    return out.read()


@app.route('/transfers', methods=['GET'])
def transfers_list():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    status = str(request.args.get('status') or '').strip()
    asset_code = str(request.args.get('asset_code') or '').strip()

    q = AssetTransferCase.query
    if period_id:
        q = q.filter(AssetTransferCase.period_id == period_id)
    if status and status in TRANSFER_STATUSES:
        q = q.filter(AssetTransferCase.status == status)
    if asset_code:
        asset = get_asset_by_c_act_strict(asset_code)
        if not asset:
            return jsonify({'items': []})
        q = q.filter(AssetTransferCase.asset_id == asset.id)

    rows = q.order_by(AssetTransferCase.id.desc()).limit(500).all()
    return jsonify({
        'items': [r.to_dict() for r in rows],
        'meta': {'statuses': TRANSFER_STATUSES},
    })


@app.route('/transfers/from_issue', methods=['POST'])
def transfers_create_from_issue():
    ensure_db()
    data = request.get_json() or {}
    issue_id = data.get('issue_id')
    try:
        issue_id = int(issue_id)
    except Exception:
        issue_id = None
    if not issue_id:
        return jsonify({'error': 'Debes indicar la novedad origen'}), 400

    issue = AssetIssue.query.get(issue_id)
    if not issue:
        return jsonify({'error': 'Novedad no encontrada'}), 404
    if not issue.asset_id:
        return jsonify({'error': 'La novedad no tiene activo asociado'}), 400
    if issue.issue_type not in {'SCANNED_OTHER_SERVICE', 'LOCATION_REVIEW', 'RESPONSIBLE_REVIEW'}:
        return jsonify({'error': 'Esta novedad no aplica para traslado'}), 400

    asset = Asset.query.get(issue.asset_id)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    target_service = str(data.get('target_service') or '').strip()
    target_responsible = str(data.get('target_responsible') or '').strip()
    requested_by = str(data.get('requested_by') or '').strip() or 'coordinador_activos'
    justification = str(data.get('justification') or '').strip() or (issue.description or '')

    if not target_service:
        return jsonify({'error': 'Debes indicar el servicio destino'}), 400

    existing = AssetTransferCase.query.filter(
        AssetTransferCase.asset_id == issue.asset_id,
        AssetTransferCase.period_id == issue.period_id,
        AssetTransferCase.status.in_(['Pendiente aprobacion', 'Aprobado'])
    ).order_by(AssetTransferCase.id.desc()).first()
    if existing:
        return jsonify({'item': existing.to_dict(), 'existing': True})

    now_value = now_iso()
    row = AssetTransferCase(
        issue_id=issue.id,
        asset_id=issue.asset_id,
        period_id=issue.period_id,
        run_id=issue.run_id,
        status='Pendiente aprobacion',
        origin_service=str(asset.nom_ccos or '').strip(),
        target_service=target_service,
        origin_responsible=str(asset.nom_resp or '').strip(),
        target_responsible=target_responsible,
        justification=justification,
        requested_by=requested_by,
        requested_at=now_value,
        created_at=now_value,
        updated_at=now_value,
    )
    db.session.add(row)
    if issue.status == 'Nuevo':
        issue.status = 'Escalado'
    issue.updated_at = now_value
    db.session.commit()
    return jsonify({'item': row.to_dict(), 'existing': False})


@app.route('/transfers/<int:transfer_id>/approve', methods=['PATCH'])
def transfers_approve(transfer_id):
    ensure_db()
    row = AssetTransferCase.query.get(transfer_id)
    if not row:
        return jsonify({'error': 'Caso de traslado no encontrado'}), 404
    if row.status == 'Ejecutado':
        return jsonify({'error': 'El caso ya fue ejecutado'}), 400

    data = request.get_json() or {}
    decision = str(data.get('decision') or 'approve').strip().lower()
    approved_by = str(data.get('approved_by') or '').strip() or 'jefe_activos'
    approval_notes = str(data.get('approval_notes') or '').strip()

    now_value = now_iso()
    if decision == 'reject':
        row.status = 'Rechazado'
    else:
        row.status = 'Aprobado'
    row.approved_by = approved_by
    row.approved_at = now_value
    row.approval_notes = approval_notes
    row.updated_at = now_value

    issue = AssetIssue.query.get(row.issue_id) if row.issue_id else None
    if issue:
        issue.status = 'En analisis' if row.status == 'Aprobado' else 'Cerrado'
        issue.resolution_notes = append_text_note(issue.resolution_notes, approval_notes)
        issue.updated_at = now_value

    db.session.commit()
    return jsonify({'item': row.to_dict()})


@app.route('/transfers/<int:transfer_id>/execute', methods=['PATCH'])
def transfers_execute(transfer_id):
    ensure_db()
    row = AssetTransferCase.query.get(transfer_id)
    if not row:
        return jsonify({'error': 'Caso de traslado no encontrado'}), 404
    if row.status != 'Aprobado':
        return jsonify({'error': 'Solo se pueden ejecutar casos aprobados'}), 400
    if not str(row.target_service or '').strip():
        return jsonify({'error': 'El caso no tiene servicio destino'}), 400

    asset = Asset.query.get(row.asset_id)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    data = request.get_json() or {}
    executed_by = str(data.get('executed_by') or '').strip() or 'equipo_activos'
    execution_notes = str(data.get('execution_notes') or '').strip()
    now_value = now_iso()

    previous_service = str(asset.nom_ccos or '').strip()
    previous_responsible = str(asset.nom_resp or '').strip()
    asset.nom_ccos = str(row.target_service or '').strip()
    if str(row.target_responsible or '').strip():
        asset.nom_resp = str(row.target_responsible or '').strip()
    asset.observacion_inventario = append_text_note(
        asset.observacion_inventario,
        f"Traslado ejecutado {format_dt_local(now_value)}: {previous_service or 'N/D'} -> {asset.nom_ccos or 'N/D'}",
    )
    asset.fecha_verificacion = now_value
    asset.usuario_verificador = executed_by

    row.status = 'Ejecutado'
    row.executed_by = executed_by
    row.executed_at = now_value
    row.execution_notes = execution_notes
    row.updated_at = now_value

    issue = AssetIssue.query.get(row.issue_id) if row.issue_id else None
    if issue:
        issue.status = 'Cerrado'
        issue.resolution_notes = append_text_note(
            issue.resolution_notes,
            f"Traslado ejecutado a '{asset.nom_ccos or ''}'.",
        )
        issue.updated_at = now_value

    sibling_issues = AssetIssue.query.filter(
        AssetIssue.asset_id == asset.id,
        AssetIssue.period_id == row.period_id,
        AssetIssue.issue_type.in_(['SCANNED_OTHER_SERVICE', 'LOCATION_REVIEW', 'RESPONSIBLE_REVIEW']),
        AssetIssue.status != 'Cerrado'
    ).all()
    for sibling in sibling_issues:
        sibling.status = 'Cerrado'
        sibling.resolution_notes = append_text_note(
            sibling.resolution_notes,
            f"Cierre automatico por traslado ejecutado hacia '{asset.nom_ccos or ''}'.",
        )
        sibling.updated_at = now_value

    pdf_bytes = build_transfer_acta_pdf_bytes(row, asset, issue=issue)
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
        description=f'Traslado ejecutado de {previous_service or "N/D"} a {asset.nom_ccos or "N/D"}. '
                    f'Responsable anterior: {previous_responsible or "N/D"}. Responsable actual: {asset.nom_resp or "N/D"}.',
        doc_date=now_local_dt().strftime('%Y-%m-%d'),
        area_service=asset.nom_ccos or '',
        radicado=f'TR-{row.id:06d}',
        file_name=public_name,
        file_path=file_path,
        file_ext='.pdf',
        file_size=len(pdf_bytes),
        uploaded_by=executed_by,
        uploaded_at=now_value,
        status='active',
    )
    db.session.add(doc_row)
    db.session.flush()

    row.acta_doc_id = doc_row.id
    row.acta_file_path = file_path
    db.session.commit()
    return jsonify({'item': row.to_dict()})


@app.route('/transfers/<int:transfer_id>/acta', methods=['GET'])
def transfers_acta_download(transfer_id):
    ensure_db()
    row = AssetTransferCase.query.get(transfer_id)
    if not row:
        return jsonify({'error': 'Caso de traslado no encontrado'}), 404
    if not row.acta_doc_id:
        return jsonify({'error': 'El caso aun no tiene acta'}), 404
    doc_row = DocumentRecord.query.filter_by(id=row.acta_doc_id, status='active').first()
    if not doc_row or not doc_row.file_path or not os.path.exists(doc_row.file_path):
        return jsonify({'error': 'Acta no disponible en almacenamiento'}), 404
    return send_file(
        doc_row.file_path,
        as_attachment=True,
        download_name=doc_row.file_name or os.path.basename(doc_row.file_path),
        mimetype='application/pdf'
    )


@app.route('/issues/report_pdf', methods=['GET'])
def issues_report_pdf():
    ensure_db()
    period_id = request.args.get('period_id', type=int)
    if not period_id:
        return jsonify({'error': 'Periodo es obligatorio'}), 400
    period = InventoryPeriod.query.get(period_id)
    if not period:
        return jsonify({'error': 'Periodo no encontrado'}), 404

    rows = AssetIssue.query.filter_by(period_id=period_id).order_by(AssetIssue.id.desc()).all()
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=letter, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=18 * mm, bottomMargin=12 * mm)
    styles = getSampleStyleSheet()
    brand_blue = colors.HexColor('#0A6FB3')
    brand_blue_dark = colors.HexColor('#07507F')
    brand_green = colors.HexColor('#1E9E57')
    brand_yellow = colors.HexColor('#F2C94C')
    brand_red = colors.HexColor('#C0392B')
    brand_blue = colors.HexColor('#0A6FB3')
    brand_blue_dark = colors.HexColor('#07507F')
    brand_green = colors.HexColor('#1E9E57')
    brand_yellow = colors.HexColor('#F2C94C')
    brand_red = colors.HexColor('#C0392B')
    brand_blue = colors.HexColor('#0A6FB3')
    brand_blue_dark = colors.HexColor('#07507F')
    brand_green = colors.HexColor('#1E9E57')
    brand_yellow = colors.HexColor('#F2C94C')
    brand_red = colors.HexColor('#C0392B')
    title_style = ParagraphStyle('it', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#0B4F6C'))
    normal = ParagraphStyle('in', parent=styles['Normal'], fontSize=8, leading=10)
    story = []
    story.append(Paragraph(f'Informe de novedades y saneamiento - {period.name}', title_style))
    story.append(Spacer(1, 6))
    summary_data = [
        ['Total novedades', str(len(rows))],
        ['Abiertas', str(sum(1 for r in rows if r.status != 'Cerrado'))],
        ['Valor en riesgo', money_text(sum(to_number(r.detected_value) for r in rows if r.status != 'Cerrado'))],
    ]
    t = Table(summary_data, colWidths=[60 * mm, 120 * mm])
    t.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#C9DCE8')),
        ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#DCE8F0')),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F2F8FD')),
    ]))
    story.append(t)
    story.append(Spacer(1, 8))

    data = [[
        Paragraph('<b>ID</b>', normal),
        Paragraph('<b>TIPO</b>', normal),
        Paragraph('<b>ACTIVO</b>', normal),
        Paragraph('<b>SERVICIO</b>', normal),
        Paragraph('<b>SEVERIDAD</b>', normal),
        Paragraph('<b>ESTADO</b>', normal),
        Paragraph('<b>ASIGNADO</b>', normal),
    ]]
    for r in rows[:400]:
        info = r.to_dict()
        data.append([
            Paragraph(str(info.get('id', '')), normal),
            Paragraph(str(info.get('issue_type_label', '')), normal),
            Paragraph(f"{info.get('asset_code', '')} - {info.get('asset_name', '')}", normal),
            Paragraph(str(info.get('service', '') or ''), normal),
            Paragraph(str(info.get('severity', '')), normal),
            Paragraph(str(info.get('status', '')), normal),
            Paragraph(str(info.get('assigned_to', '') or ''), normal),
        ])
    tb = Table(data, colWidths=[10 * mm, 36 * mm, 56 * mm, 34 * mm, 18 * mm, 20 * mm, 24 * mm], repeatRows=1)
    tb.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0B4F6C')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#C9DCE8')),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#DCE8F0')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FBFD')]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(tb)
    doc.build(story, onFirstPage=make_pdf_page_header(get_hospital_logo_path()), onLaterPages=make_pdf_page_header(get_hospital_logo_path()))
    out.seek(0)
    name = f"novedades_saneamiento_{period.name.replace(' ', '_')}.pdf"
    return send_file(out, as_attachment=True, download_name=name, mimetype='application/pdf')


