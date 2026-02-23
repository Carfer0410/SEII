Sistema de Inventario - MVP

Descripcion
-----------
Pequeno MVP para inventario de activos: importa Excel/CSV, lista activos por servicio, modo de escaneo (pistola USB o camara) que marca activos como "Encontrado" y exporta A22 en Excel.

Requisitos
----------
- Python 3.10+
- pip
- (Opcional) wkhtmltopdf para exportar PDF desde HTML

Instalacion rapida
------------------
Crear entorno y instalar dependencias:

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Ejecutar la aplicacion (desarrollo):

```bash
set FLASK_APP=app.py
set FLASK_ENV=development
flask run
```

Uso basico
----------
- Abrir http://127.0.0.1:5000 en el navegador.
- Importar el Excel/CSV generado por el sistema (debe contener la columna `C_ACT`).
- Seleccionar servicio y usar el campo de escaneo o la camara para marcar activos.
- Exportar A22 desde la interfaz.

Notas
-----
- Primera version MVP para pruebas locales.
- Para produccion se recomienda desplegar con Gunicorn/uWSGI y configurar SSL y backup de la base de datos.
