# v26.5 — Corrección completa del panel de pedidos

- `/admin/pedidos` funciona como entrada del panel real de pedidos.
- Se añadieron botones visibles en el panel: configurar hora, abrir pedidos ahora e ingresar pedido manual.
- `/admin/pedidos/configuracion`: cambia la hora normal de cierre y permite abrir/cerrar manualmente.
- `/admin/pedidos/manual`: registra pedidos directamente en la misma tabla de pedidos.
- La hora de cierre y el estado manual se guardan en SQLite.
- La reapertura manual solo aplica a hoy y no cambia la hora configurada.
- Al comenzar otro día vuelve a regir automáticamente la hora normal.
- Se eliminaron las rutas duplicadas que estaban en `app.py` y estaban interfiriendo con el sistema real de pedidos.
