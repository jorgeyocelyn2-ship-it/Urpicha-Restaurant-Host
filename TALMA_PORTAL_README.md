# Urpicha v19.8 — Portal TALMA persistente

Esta versión parte del proyecto v19.7.

## Objetivo
- Portal privado `/talma`.
- Contraseña independiente mediante `TALMA_PASSWORD`.
- No indexar `/talma` en buscadores.
- Persistencia de los datos de TALMA en el directorio `datos/`.
- Preparado para usar Persistent Disk de Render.

## Render
Para persistencia real en Render, agrega un Persistent Disk al Web Service:
- Mount Path: `/opt/render/project/src/datos`

No dependas del sistema de archivos efímero del contenedor para conservar datos.

## Contraseña
Configura en Render:
TALMA_PASSWORD=Talma2026

Se recomienda cambiarla antes de producción.

## Nota
El archivo incluye la base del portal y la estructura persistente. La ruta final debe integrarse con las rutas Flask existentes del proyecto para que el portal consulte los pedidos TALMA almacenados por la aplicación.
