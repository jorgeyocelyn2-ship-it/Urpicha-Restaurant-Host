# Urpicha 20.3 — DNI para TALMA

- Solo el enlace de empresa `talma` solicita DNI.
- DNI obligatorio: exactamente 8 dígitos.
- El DNI se guarda en la tabla `orders`.
- Para TALMA, el DNI es el identificador único por fecha; evita duplicados aunque cambie el nombre.
- El portal `/talma` permite filtrar por DNI.
- El Excel del portal incluye DNI y resumen por DNI.
- Empresas distintas de TALMA no muestran ni exigen DNI.
- La migración agrega la columna `dni` automáticamente a bases existentes.
