# Modularizacion de la aplicacion

## Objetivo

Separar `app.py` (monolito) en modulos por responsabilidad, sin cambiar comportamiento funcional.

## Estructura actual

- `app.py`: punto de entrada minimo.
- `app_modules/bootstrap.py`: cargador principal de modulos web.
- `app_modules/core/foundation.py`: configuracion base, constantes, modelos, utilidades comunes.
- `app_modules/web/pages_imports.py`: pantallas base, importaciones y endpoints de entrada.
- `app_modules/web/inventory_periods_issues.py`: periodos, cobertura, incidencias y traslados.
- `app_modules/web/inventory_assets_disposals.py`: activos, escaneo, jornadas, bajas y dashboard operativo.
- `app_modules/web/runs_formats.py`: formatos, exportaciones y flujos de cierre/reapertura.
- `app_modules/web/accounting_common.py`: utilidades y base de conciliacion contable.
- `app_modules/web/accounting_documents_life.py`: hoja de vida y gestion documental.
- `app_modules/web/accounting_reports.py`: informe mensual de conciliacion e historicos.

## Regla de importacion

La cadena actual es:

1. `foundation` -> base tecnica y de datos.
2. `pages_imports` -> pagina/rutas iniciales.
3. `inventory_periods_issues` -> periodos e incidencias.
4. `inventory_assets_disposals` -> activos y bajas.
5. `runs_formats` -> formatos/reportes.
6. `accounting_common` -> base contable.
7. `accounting_documents_life` -> hoja de vida y documentos.
8. `accounting_reports` -> conciliacion mensual.

Esto conserva el registro de rutas Flask en el mismo orden logico del sistema.

## Nota importante de rutas

`BASE_DIR` se resuelve a la raiz del proyecto desde `app_modules/core/foundation.py`.
Con esto:

- `assets.db` se mantiene en la raiz.
- `generated_reports` y `generated_documents` se mantienen en la raiz.
- no se mezclan datos operativos dentro de `app_modules/`.
