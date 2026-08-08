# Publicar esta versión en Render

## Importante: usar Docker

Los scanners TALMA y POLICÍA necesitan **Tesseract OCR**, que es un paquete del sistema operativo. Por eso esta versión incluye un `Dockerfile` que instala:

- `tesseract-ocr`
- `tesseract-ocr-spa`
- las dependencias Python del proyecto

En Render, el servicio debe usar **Runtime / Language: Docker** para que el scanner funcione en producción.

## Configuración prevista

El `render.yaml` ya incluye:

- Health check: `/health`
- Disco persistente: `/opt/render/project/src/datos`
- `DATA_DIR=/opt/render/project/src/datos`
- `SECRET_KEY` generada por Render
- un solo worker Gunicorn para evitar duplicar el servidor interno y mantener SQLite simple

Variables que debes revisar en Render:

- `RESTAURANT_NAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`
- `ORDER_DEADLINE`
- `DATA_DIR=/opt/render/project/src/datos`

## Si tu servicio actual está como Python

El Dockerfile no se usará mientras el servicio continúe con runtime Python. Debes cambiar/sincronizar el servicio a Docker o crear un Web Service Docker conectado al mismo repositorio.

Después del despliegue comprueba:

1. `/health` responde `ok`.
2. `/admin` abre el inicio de sesión.
3. Al entrar aparecen las pestañas Pedidos, TALMA y POLICÍA.
4. `/admin/scanner-diagnostico` debe indicar que Tesseract fue encontrado y mostrar `spa` entre los idiomas.
5. Prueba una captura TALMA y otra POLICÍA antes de usarlo en producción.

## Dominio personalizado

El dominio personalizado se configura en Render sobre el Web Service que esté ejecutando esta versión. Si sustituyes el servicio anterior por uno nuevo, tendrás que asociar el dominio al servicio nuevo.


## Corrección v19.4
Las plantillas TALMA/POLICÍA ahora tienen respaldo embebido en `scanner_templates.py`.
Aunque Render no encuentre la carpeta `templates`, los paneles `/admin/talma/` y `/admin/policia/` pueden cargar.
También existe `/admin/scanner-diagnostico` para comprobar plantillas y Tesseract.
