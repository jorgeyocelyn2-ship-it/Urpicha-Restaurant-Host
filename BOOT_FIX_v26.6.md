# v26.6 — FIX de arranque Render

Causa exacta del fallo:
La base SQLite persistente ya tenía `order_settings` creada por una versión anterior con la columna `manual_open`, mientras que `orders_app.py` intentaba insertar/usar `manual_open_date` y `manual_closed`. SQLite no modifica una tabla existente con `CREATE TABLE IF NOT EXISTS`, por lo que Gunicorn terminaba con:
`sqlite3.OperationalError: table order_settings has no column named manual_open_date`.

Corrección:
- Se añadió una migración automática y compatible dentro de `init_db()`.
- Detecta las columnas existentes de `order_settings`.
- Añade `manual_open_date`, `manual_closed`, `updated_at` o `closing_time` si faltan.
- Conserva la hora de cierre existente.
- Si una versión anterior tenía `manual_open=1`, lo convierte en reapertura manual para el día actual.
- No elimina ni recrea la base de datos, por lo que conserva los pedidos existentes.
- Se verificó la migración contra una base SQLite simulando exactamente el esquema antiguo.
- `app.py` y `orders_app.py` pasan validación de sintaxis Python.

Las rutas `/admin/pedidos`, `/admin/pedidos/configuracion`, `/admin/pedidos/manual`, `/admin/pedidos/reabrir` y `/admin/pedidos/cerrar-manual` permanecen en el sistema.
