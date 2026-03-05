from .accounting_common import *
from .accounting_common import _pick_first_value

def build_asset_life_sheet_payload(asset, matched_by='C_ACT'):
    payload = asset_raw_payload(asset)
    codigo_inteligente = _pick_first_value(payload, [
        'CODINTELIGENTE', 'CODIGO_INTELIGENTE', 'COD_INTELIGENTE', 'CODIGO INTELIGENTE'
    ])
    subtipo_codigo = _pick_first_value(payload, [
        'SUBTIPO', 'SUBTIPO_ACTIVO', 'COD_SUBTIPO', 'COD_SUBTIPO_ACTIVO'
    ])
    subtipo_nombre = _pick_first_value(payload, [
        'DES_SUBTIAC', 'DESC_SUBTIAC', 'SUBTIPO_ACTIVO', 'DESC_SUBTIPO_ACTIVO'
    ]) or (asset.desc_subtiac or '')

    data = {
        'codigo': asset.c_act or '',
        'codigo_inteligente': codigo_inteligente or (asset.codigo_inteligente or ''),
        'descripcion_activo': asset.nom or '',
        'familia_codigo': asset.c_fam or '',
        'familia_nombre': asset.nom_fam or '',
        'tipo_codigo': asset.c_tiac or '',
        'tipo_nombre': asset.desc_tiac or '',
        'subtipo_codigo': subtipo_codigo or (asset.subtipo_codigo or ''),
        'subtipo_nombre': subtipo_nombre,
        'marca': asset.nom_marca or _pick_first_value(payload, ['MARCA']),
        'modelo': asset.modelo or '',
        'serial_referencia': asset.serie or asset.ref or '',
        'color': _pick_first_value(payload, ['COLOR', 'COLORES']) or (asset.color or ''),
        'nit_proveedor': _pick_first_value(payload, ['NIT_PROVEEDOR', 'NIT PROVEEDOR', 'NIT']) or (asset.nit_proveedor or ''),
        'proveedor': _pick_first_value(payload, ['PROVEEDOR', 'DESCRIPCION_PROVEEDOR', 'DESCRIPCION DEL PROVEEDOR']) or (asset.desc_proveedor or ''),
        'fecha_incorporacion': date_only(asset.fecha_compra),
        'forma_adquisicion': _pick_first_value(payload, ['FORMA_ADQUISICION', 'FORMA DE ADQUISICION', 'ADQUISICION']) or (asset.forma_adquisicion or ''),
        'en_garantia': _pick_first_value(payload, ['EN_GARANTIA', 'GARANTIA']) or (asset.en_garantia or 'No'),
        'entidad': _pick_first_value(payload, ['ENTIDAD', 'ENTIDAD_GARANTIA']) or (asset.entidad_garantia or ''),
        'garantia_desde': _pick_first_value(payload, ['GARANTIA_DESDE', 'DESDE']) or (asset.garantia_desde or ''),
        'garantia_hasta': _pick_first_value(payload, ['GARANTIA_HASTA', 'HASTA']) or (asset.garantia_hasta or ''),
        'estado': asset.est or '',
        'condicion': _pick_first_value(payload, ['CONDICION']) or (asset.estado_inventario or ''),
        'metodo_deprec': asset.deprecia or '',
        'costo_activo': round(to_number(asset.costo), 2),
        'saldo': round(to_number(asset.saldo), 2),
        'total_activo': round(to_number(asset.costo), 2),
        'responsable': asset.nom_resp or '',
        'ubicacion': asset.des_ubi or '',
        'centro_costo': _pick_first_value(payload, ['C_CCOS', 'CENTRO_COSTO', 'COD_CENTRO_COSTO']) or (asset.centro_costo_code or ''),
        'servicio': asset.nom_ccos or '',
        'agencia': _pick_first_value(payload, ['AGENCIA']) or (asset.agencia or ''),
        'area': classify_area(asset.nom_ccos),
        'observaciones': asset.observacion_inventario or _pick_first_value(payload, ['OBSERVACIONES', 'OBSERVACION']) or '',
        'matched_by': matched_by or 'C_ACT',
        'fecha_generacion': format_dt_local(now_iso()),
    }
    return data


@app.route('/asset_life_sheet', methods=['GET'])
def asset_life_sheet():
    ensure_db()
    code = (request.args.get('code') or '').strip()
    if not code:
        return jsonify({'error': 'Debes indicar el codigo del activo'}), 400
    allow_barcode = parse_bool(request.args.get('allow_barcode'), False)
    asset = get_asset_by_c_act_strict(code)
    matched_by = 'C_ACT'
    if (not asset) and allow_barcode:
        asset, matched_by = get_asset_by_code(code)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404
    return jsonify({'item': build_asset_life_sheet_payload(asset, matched_by)})


@app.route('/asset_life_sheet/quick_lookup', methods=['POST'])
def asset_life_sheet_quick_lookup():
    ensure_db()
    payload = request.get_json(silent=True) or {}
    category_key = str(payload.get('category_key') or '').strip()
    service_hint = str(payload.get('service_hint') or '').strip()
    location_hint = str(payload.get('location_hint') or '').strip()
    subtype_text = str(payload.get('subtype_text') or '').strip()
    query_text = str(payload.get('query_text') or '').strip()
    technical_text = str(payload.get('technical_text') or '').strip()
    limit = parse_int(payload.get('limit'), default=30) or 30

    if not any([category_key, service_hint, location_hint, subtype_text, query_text, technical_text]):
        return jsonify({'error': 'Debes indicar al menos un criterio de busqueda'}), 400

    candidates, pool_size = build_quick_lookup_candidates(
        rule_key=category_key,
        service_hint=service_hint,
        location_hint=location_hint,
        query_text=query_text,
        technical_text=technical_text,
        subtype_text=subtype_text,
        limit=limit,
    )
    rule = get_assist_rule(category_key) if category_key else None
    return jsonify({
        'analysis': {
            'category_key': category_key,
            'category_label': (rule.get('label') if rule else ''),
            'service_hint': service_hint,
            'location_hint': location_hint,
            'subtype_text': subtype_text,
            'query_text': query_text,
            'technical_text': technical_text,
            'not_found_only': True,
            'not_found_pool_size': pool_size,
            'returned_candidates': len(candidates),
        },
        'candidates': candidates,
    })


@app.route('/asset_life_sheet/pdf', methods=['GET'])
def asset_life_sheet_pdf():
    ensure_db()
    code = (request.args.get('code') or '').strip()
    if not code:
        return jsonify({'error': 'Debes indicar el codigo del activo'}), 400
    allow_barcode = parse_bool(request.args.get('allow_barcode'), False)
    asset = get_asset_by_c_act_strict(code)
    matched_by = 'C_ACT'
    if (not asset) and allow_barcode:
        asset, matched_by = get_asset_by_code(code)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404

    item = build_asset_life_sheet_payload(asset, matched_by)
    out = BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=letter,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'LifeTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        textColor=colors.HexColor('#0A5C8D'),
        alignment=1,
        spaceAfter=6,
    )
    small_style = ParagraphStyle(
        'LifeSmall',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#405569'),
        alignment=1,
        spaceAfter=6,
    )
    label_style = ParagraphStyle(
        'LifeLabel',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#0F4E75'),
        fontName='Helvetica-Bold',
    )
    value_style = ParagraphStyle(
        'LifeValue',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#12212F'),
    )

    def row(label, value):
        return [
            Paragraph(escape(str(label or '')), label_style),
            Paragraph(escape(str(value or '-')), value_style),
        ]

    story = [
        Paragraph('HOJA DE VIDA DE ACTIVOS', title_style),
        Paragraph(
            f"Generado: {escape(item['fecha_generacion'])} | Codigo consultado: {escape(str(code))}",
            small_style
        ),
    ]

    table_rows = [
        row('Codigo', item['codigo']),
        row('Codigo inteligente', item['codigo_inteligente']),
        row('Descripcion activo fijo', item['descripcion_activo']),
        row('Familia', f"{item['familia_codigo']} - {item['familia_nombre']}"),
        row('Tipo de activo', f"{item['tipo_codigo']} - {item['tipo_nombre']}"),
        row('Subtipo de activo', f"{item['subtipo_codigo']} - {item['subtipo_nombre']}"),
        row('Marca / Modelo', f"{item['marca']} / {item['modelo']}"),
        row('No. serial o referencia', item['serial_referencia']),
        row('Color', item['color']),
        row('NIT proveedor', item['nit_proveedor']),
        row('Descripcion del proveedor', item['proveedor']),
        row('Fecha incorporacion', item['fecha_incorporacion']),
        row('Forma de adquisicion', item['forma_adquisicion']),
        row('En garantia', item['en_garantia']),
        row('Entidad garantia', item['entidad']),
        row('Garantia desde / hasta', f"{item['garantia_desde']} / {item['garantia_hasta']}"),
        row('Estado', item['estado']),
        row('Condicion', item['condicion']),
        row('Metodo depreciacion', item['metodo_deprec']),
        row('Costo del activo', money_text(item['costo_activo'])),
        row('Saldo', money_text(item['saldo'])),
        row('Total activo', money_text(item['total_activo'])),
        row('Responsable', item['responsable']),
        row('Ubicacion', item['ubicacion']),
        row('Centro de costo', item['centro_costo']),
        row('Servicio', item['servicio']),
        row('Agencia', item['agencia']),
        row('Area', item['area']),
        row('Observaciones', item['observaciones']),
    ]
    detail = Table(table_rows, colWidths=[52 * mm, 128 * mm])
    detail.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F7FBFF')),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.HexColor('#EEF5FC'), colors.HexColor('#FFFFFF')]),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#BFD3E3')),
        ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#D3E1ED')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(detail)
    doc.build(story)
    out.seek(0)
    filename = clean_filename(f"hoja_vida_{asset.c_act}.pdf")
    return send_file(out, as_attachment=True, download_name=filename, mimetype='application/pdf')


def allowed_document_extension(file_name):
    ext = os.path.splitext(str(file_name or ''))[1].lower()
    allowed = {
        '.pdf', '.xlsx', '.xls', '.csv',
        '.png', '.jpg', '.jpeg', '.webp',
        '.doc', '.docx', '.txt',
        '.ppt', '.pptx',
    }
    return ext in allowed


def document_mimetype_by_ext(ext):
    ext = str(ext or '').lower()
    mapping = {
        '.pdf': 'application/pdf',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.xls': 'application/vnd.ms-excel',
        '.csv': 'text/csv',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.webp': 'image/webp',
        '.doc': 'application/msword',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.ppt': 'application/vnd.ms-powerpoint',
        '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        '.txt': 'text/plain',
    }
    return mapping.get(ext, 'application/octet-stream')


def get_document_preview_mode(ext):
    ext = str(ext or '').lower()
    direct_preview = {'.pdf', '.png', '.jpg', '.jpeg', '.webp', '.txt', '.csv'}
    office_to_pdf_preview = {'.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'}
    if ext in direct_preview:
        return 'direct'
    if ext in office_to_pdf_preview:
        return 'office_to_pdf'
    return 'unsupported'


def _document_preview_cache_dir():
    return os.path.join(DOCUMENTS_DIR, '_preview_cache')


def _find_libreoffice_command():
    candidates = [
        shutil.which('soffice'),
        shutil.which('libreoffice'),
        os.path.join('C:\\', 'Program Files', 'LibreOffice', 'program', 'soffice.exe'),
        os.path.join('C:\\', 'Program Files (x86)', 'LibreOffice', 'program', 'soffice.exe'),
    ]
    for cmd in candidates:
        if cmd and os.path.exists(cmd):
            return cmd
    return None


def _build_preview_pdf_path(row):
    cache_dir = _document_preview_cache_dir()
    os.makedirs(cache_dir, exist_ok=True)
    base_name = clean_filename(os.path.splitext(row.file_name or f'documento_{row.id}')[0]) or f'documento_{row.id}'
    fingerprint = f"{int(os.path.getmtime(row.file_path))}_{int(row.file_size or 0)}"
    return os.path.join(cache_dir, f'{base_name}_{row.id}_{fingerprint}.pdf')


def _convert_document_to_pdf_for_preview(row):
    cmd = _find_libreoffice_command()
    if not cmd:
        raise RuntimeError('No se encontro LibreOffice en el servidor para convertir documentos Office a PDF.')

    cache_dir = _document_preview_cache_dir()
    os.makedirs(cache_dir, exist_ok=True)
    final_pdf_path = _build_preview_pdf_path(row)
    if os.path.exists(final_pdf_path):
        return final_pdf_path

    source_ext = os.path.splitext(str(row.file_path or ''))[1].lower()
    staged_name = f'doc_preview_{row.id}_{int(os.path.getmtime(row.file_path))}{source_ext}'
    try:
        with tempfile.TemporaryDirectory(dir=cache_dir) as tmp_dir:
            staged_input = os.path.join(tmp_dir, staged_name)
            shutil.copy2(row.file_path, staged_input)
            subprocess.run(
                [cmd, '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir, staged_input],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=120,
            )
            converted_pdf = os.path.join(tmp_dir, f"{os.path.splitext(staged_name)[0]}.pdf")
            if not os.path.exists(converted_pdf):
                raise RuntimeError('LibreOffice no devolvio un archivo PDF para la vista previa.')
            shutil.copy2(converted_pdf, final_pdf_path)
            return final_pdf_path
    except subprocess.TimeoutExpired:
        raise RuntimeError('La conversion a PDF excedio el tiempo limite (120s).')
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or '').strip()
        stdout = (exc.stdout or '').strip()
        detail = stderr or stdout or 'error desconocido'
        raise RuntimeError(f'No fue posible convertir el documento a PDF: {detail}')


def _document_preview_info_html(row, title, message):
    safe_title = escape(str(title or 'Vista previa'))
    safe_message = escape(str(message or 'No fue posible cargar la vista previa.'))
    download_url = f"/documents/{int(row.id)}/download"
    file_label = escape(str(row.file_name or os.path.basename(row.file_path or 'archivo')))
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{safe_title}</title>
  <style>
    body {{
      margin: 0;
      font-family: Segoe UI, Arial, sans-serif;
      background: #f4f8fc;
      color: #123;
      display: grid;
      place-items: center;
      min-height: 100vh;
      padding: 16px;
    }}
    .card {{
      width: min(680px, 100%);
      background: #fff;
      border: 1px solid #d7e4ef;
      border-radius: 12px;
      padding: 18px;
      box-shadow: 0 8px 20px rgba(11, 36, 57, 0.12);
    }}
    h2 {{
      margin: 0 0 10px;
      color: #0f4e75;
      font-size: 20px;
    }}
    p {{
      margin: 0 0 12px;
      color: #304f63;
      line-height: 1.45;
      font-size: 14px;
    }}
    a {{
      display: inline-flex;
      align-items: center;
      height: 34px;
      padding: 0 12px;
      border-radius: 8px;
      background: #0a7ea4;
      color: #fff;
      text-decoration: none;
      font-weight: 700;
      font-size: 13px;
    }}
    .name {{
      margin-top: 10px;
      font-size: 12px;
      color: #607587;
      word-break: break-all;
    }}
  </style>
</head>
<body>
  <div class="card">
    <h2>{safe_title}</h2>
    <p>{safe_message}</p>
    <a href="{download_url}">Descargar archivo</a>
    <div class="name">{file_label}</div>
  </div>
</body>
</html>"""


@app.route('/assets/find_by_code', methods=['GET'])
def assets_find_by_code():
    ensure_db()
    code = (request.args.get('code') or '').strip()
    if not code:
        return jsonify({'error': 'Debes indicar codigo de activo'}), 400
    asset = get_asset_by_c_act_strict(code)
    if not asset:
        return jsonify({'error': 'Activo no encontrado'}), 404
    return jsonify({
        'asset': {
            'id': asset.id,
            'code': asset.c_act or '',
            'name': asset.nom or '',
            'service': asset.nom_ccos or '',
            'location': asset.des_ubi or '',
        }
    })


@app.route('/documents/types', methods=['GET'])
def documents_types():
    return jsonify({'types': DOCUMENT_TYPE_OPTIONS})


@app.route('/documents', methods=['GET'])
def documents_list():
    ensure_db()
    q = DocumentRecord.query
    search = (request.args.get('search') or '').strip()
    link_type = (request.args.get('link_type') or '').strip().lower()
    document_type = (request.args.get('document_type') or '').strip()
    area_service = (request.args.get('area_service') or '').strip()
    status = (request.args.get('status') or 'active').strip().lower()

    if status in {'active', 'archived'}:
        q = q.filter(DocumentRecord.status == status)
    elif status == 'all':
        pass
    else:
        q = q.filter(DocumentRecord.status == 'active')

    if link_type in {'asset', 'general'}:
        q = q.filter(DocumentRecord.link_type == link_type)
    if document_type:
        q = q.filter(DocumentRecord.document_type == document_type)
    if area_service:
        q = q.filter(DocumentRecord.area_service == area_service)
    if search:
        term = f'%{search}%'
        q = q.filter(
            (DocumentRecord.title.like(term)) |
            (DocumentRecord.asset_code.like(term)) |
            (DocumentRecord.radicado.like(term)) |
            (DocumentRecord.file_name.like(term))
        )

    rows = q.order_by(DocumentRecord.id.desc()).limit(800).all()
    return jsonify({'items': [r.to_dict() for r in rows]})


@app.route('/documents', methods=['POST'])
def documents_create():
    ensure_db()
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Debes adjuntar un archivo'}), 400

    raw_file_name = str(f.filename or '').strip()
    if not raw_file_name:
        return jsonify({'error': 'Nombre de archivo invalido'}), 400
    if not allowed_document_extension(raw_file_name):
        return jsonify({'error': 'Tipo de archivo no permitido'}), 400

    file_ext = os.path.splitext(raw_file_name)[1].lower()
    f.seek(0, os.SEEK_END)
    file_size = int(f.tell() or 0)
    f.seek(0)
    if file_size <= 0:
        return jsonify({'error': 'El archivo esta vacio'}), 400
    if file_size > 30 * 1024 * 1024:
        return jsonify({'error': 'Archivo supera el limite de 30 MB'}), 400

    link_type = (request.form.get('link_type') or 'general').strip().lower()
    if link_type not in {'asset', 'general'}:
        return jsonify({'error': 'Tipo de vinculacion invalido'}), 400
    document_type = (request.form.get('document_type') or '').strip()
    title = (request.form.get('title') or '').strip()
    description = (request.form.get('description') or '').strip()
    doc_date = (request.form.get('doc_date') or '').strip()
    area_service = (request.form.get('area_service') or '').strip()
    radicado = (request.form.get('radicado') or '').strip()
    uploaded_by = get_actor_username((request.form.get('uploaded_by') or '').strip() or 'usuario_movil')

    if not document_type:
        return jsonify({'error': 'Debes seleccionar tipo de documento'}), 400
    if not title:
        return jsonify({'error': 'Debes indicar el titulo/asunto'}), 400

    asset_id = None
    asset_code = ''
    asset_name = ''
    if link_type == 'asset':
        code = (request.form.get('asset_code') or '').strip()
        if not code:
            return jsonify({'error': 'Debes indicar codigo activo'}), 400
        asset = get_asset_by_c_act_strict(code)
        if not asset:
            return jsonify({'error': 'Codigo activo no existe en la base'}), 400
        asset_id = asset.id
        asset_code = asset.c_act or ''
        asset_name = asset.nom or ''

    stamped_name = f"{clean_filename(os.path.splitext(raw_file_name)[0])}_{now_local_dt().strftime('%Y%m%d%H%M%S')}{file_ext}"
    file_path = os.path.join(DOCUMENTS_DIR, stamped_name)
    f.save(file_path)

    row = DocumentRecord(
        link_type=link_type,
        asset_id=asset_id,
        asset_code=asset_code,
        asset_name=asset_name,
        document_type=document_type,
        title=title,
        description=description,
        doc_date=doc_date,
        area_service=area_service,
        radicado=radicado,
        file_name=raw_file_name,
        file_path=file_path,
        file_ext=file_ext,
        file_size=file_size,
        uploaded_by=uploaded_by,
        uploaded_at=now_iso(),
        status='active',
    )
    db.session.add(row)
    db.session.commit()
    return jsonify({'item': row.to_dict()})


@app.route('/documents/<int:doc_id>/download', methods=['GET'])
def documents_download(doc_id):
    ensure_db()
    row = DocumentRecord.query.filter_by(id=doc_id, status='active').first()
    if not row:
        return jsonify({'error': 'Documento no encontrado'}), 404
    if not row.file_path or not os.path.exists(row.file_path):
        return jsonify({'error': 'El archivo no existe en almacenamiento'}), 404
    mime = document_mimetype_by_ext(row.file_ext)
    return send_file(
        row.file_path,
        as_attachment=True,
        download_name=row.file_name or os.path.basename(row.file_path),
        mimetype=mime
    )


@app.route('/documents/<int:doc_id>/preview', methods=['GET'])
def documents_preview(doc_id):
    ensure_db()
    row = DocumentRecord.query.filter_by(id=doc_id, status='active').first()
    if not row:
        return jsonify({'error': 'Documento no encontrado'}), 404
    if not row.file_path or not os.path.exists(row.file_path):
        return jsonify({'error': 'El archivo no existe en almacenamiento'}), 404

    ext = str(row.file_ext or '').lower()
    mode = get_document_preview_mode(ext)

    if mode == 'direct':
        mime = document_mimetype_by_ext(ext)
        return send_file(
            row.file_path,
            as_attachment=False,
            download_name=row.file_name or os.path.basename(row.file_path),
            mimetype=mime
        )

    if mode == 'office_to_pdf':
        try:
            preview_pdf_path = _convert_document_to_pdf_for_preview(row)
        except RuntimeError as exc:
            html = _document_preview_info_html(
                row,
                'Vista previa no disponible',
                f'No fue posible generar vista previa para este Office en el servidor. {str(exc)}'
            )
            return app.response_class(html, mimetype='text/html')
        pdf_name = f"{os.path.splitext(row.file_name or os.path.basename(row.file_path))[0]}.pdf"
        return send_file(
            preview_pdf_path,
            as_attachment=False,
            download_name=pdf_name,
            mimetype='application/pdf'
        )

    html = _document_preview_info_html(
        row,
        'Vista previa no disponible',
        'Este tipo de archivo no se puede renderizar directamente en el navegador.'
    )
    return app.response_class(html, mimetype='text/html')


@app.route('/documents/<int:doc_id>', methods=['PATCH'])
def documents_update(doc_id):
    ensure_db()
    row = DocumentRecord.query.filter_by(id=doc_id, status='active').first()
    if not row:
        return jsonify({'error': 'Documento no encontrado'}), 404
    data = request.get_json() or {}
    title = (data.get('title') or '').strip()
    description = (data.get('description') or '').strip()
    doc_date = (data.get('doc_date') or '').strip()
    area_service = (data.get('area_service') or '').strip()
    radicado = (data.get('radicado') or '').strip()
    document_type = (data.get('document_type') or '').strip()
    if title:
        row.title = title
    if document_type:
        row.document_type = document_type
    row.description = description
    row.doc_date = doc_date
    row.area_service = area_service
    row.radicado = radicado
    db.session.commit()
    return jsonify({'item': row.to_dict()})


@app.route('/documents/<int:doc_id>/archive', methods=['POST'])
def documents_archive(doc_id):
    ensure_db()
    row = DocumentRecord.query.filter_by(id=doc_id, status='active').first()
    if not row:
        return jsonify({'error': 'Documento no encontrado'}), 404
    row.status = 'archived'
    db.session.commit()
    return jsonify({'ok': True, 'id': doc_id})


