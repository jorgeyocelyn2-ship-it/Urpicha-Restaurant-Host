# v20.4 — Corrección de menú por fecha

Correcciones:
- Guarda el menú de forma atómica por fecha.
- Acepta YYYY-MM-DD, DD/MM/YYYY y DD-MM-YYYY.
- Limpia espacios correctamente.
- Permite platos separados por saltos de línea o punto y coma.
- Verifica inmediatamente en SQLite que las entradas y fondos guardados coincidan con lo enviado.
- Redirige al mismo día y vuelve a leer el menú desde la base de datos.
- Muestra cuántas entradas y fondos quedaron guardados.
- Corrige el error de regex que impedía normalizar espacios correctamente.
