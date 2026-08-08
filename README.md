# Urpicha Restaurante - panel administrador integrado

Esta versión reúne en el mismo sitio web los tres apartados del administrador:

1. **Pedidos**: menú diario, empresas, enlaces privados, pedidos, CSV/Excel y cupones del sistema web.
2. **TALMA**: scanner OCR de la planilla TALMA, calibración de columnas, revisión, guardado, PDF de cuponera, Excel mensual y actualización de Excel.
3. **POLICÍA**: scanner OCR de la planilla PNP, calibración, revisión con tildes/ñ, PDF de cuponera, Excel mensual y actualización de Excel.

Después de iniciar sesión en `/admin`, el panel muestra las pestañas **Pedidos · TALMA · POLICÍA**. Los scanners usan la misma sesión de administrador; no quedan expuestos públicamente.

## Ejecutar en Windows

1. Instala Python 3.
2. Instala Tesseract OCR. Puedes abrir `INSTALAR_OCR_WINDOWS.bat` y durante la instalación incluir el idioma español.
3. Ejecuta `INICIAR.bat`.
4. Abre `http://localhost:8080/admin`.

## Archivos principales

- `app.py`: aplicación frontal, scanner TALMA/POLICÍA y proxy al sistema de pedidos.
- `orders_app.py`: sistema original de pedidos.
- `templates/` y `static/`: interfaz de scanners.
- `requirements.txt`: dependencias Python.
- `Dockerfile`: instala Python, Tesseract y español para Render.
- `render.yaml`: configuración de Render con disco persistente.

## Datos persistentes

El sistema sigue usando `DATA_DIR` para la base SQLite de pedidos. Además crea:

- `DATA_DIR/scanners/datos_talma.json`
- `DATA_DIR/scanners/datos_policia.json`
- `DATA_DIR/scanners/uploads/`
- `DATA_DIR/scanners/outputs/`

Así los datos del scanner pueden vivir en el mismo disco persistente de Render.

Versión: 19.7 - actualización de Excel histórico desde Panel Pedidos + orden alfabético por persona
