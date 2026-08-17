# v26.7 — Menú al reabrir + marca Urpicha

- Al reabrir pedidos, el público muestra claramente que están abiertos manualmente.
- Se evita mostrar el contador 00:00 cuando la reapertura manual está activa.
- Si el menú de HOY quedó vacío tras una migración/reinicio, se recupera el menú predeterminado del día para que la reapertura no deje el formulario sin platos.
- Se añadió un botón Actualizar en el menú público.
- Las páginas dinámicas envían `Cache-Control: no-store`.
- El logo proporcionado se usa como marca de agua discreta en la esquina inferior derecha.
- La marca se oculta al imprimir.
- Se aplica tanto al sistema de pedidos como al scanner administrativo.
- No se elimina la base de datos ni los pedidos existentes.
