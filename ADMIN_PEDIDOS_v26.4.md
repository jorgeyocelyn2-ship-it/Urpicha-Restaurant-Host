# v26.4 — Administración de pedidos

Se añadieron:
- Configuración de la hora normal de cierre en `HH:MM`.
- Reapertura manual de pedidos sin modificar la hora de cierre.
- Cierre manual nuevamente, también sin modificar la hora configurada.
- Registro de pedidos manuales desde el panel administrativo.
- Persistencia de la configuración en SQLite mediante `order_settings`.
- La reapertura manual tiene prioridad sobre la hora normal hasta que el administrador la desactive.

Rutas:
- `/admin/pedidos/configuracion`
- `/admin/pedidos/manual`
- `/admin/pedidos/reabrir` (POST)
- `/admin/pedidos/cerrar-manual` (POST)
