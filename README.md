# Sistema de pedidos de almuerzos

Aplicación web local para que los empleados de distintas empresas elijan su almuerzo desde un enlace privado y el restaurante reciba todos los pedidos en un panel.

## Incluye

- Enlace privado por empresa.
- Nombre, área, entrada, plato de fondo, modalidad y observación.
- Bloqueo de pedidos duplicados por nombre, empresa y fecha.
- Menú diario editable.
- Resumen de cantidades para cocina.
- Filtro por fecha y empresa.
- Exportación CSV compatible con Excel.
- Cupones imprimibles o guardables como PDF.
- Base de datos SQLite automática.

## Cómo probarlo en Windows

1. Instale **Python 3.11 o superior** desde python.org.
2. Durante la instalación, active **Add Python to PATH**.
3. Descomprima esta carpeta.
4. Haga doble clic en `INICIAR.bat`.
5. Abra `http://localhost:8080/admin`.
6. Contraseña inicial: `cambiar123`.

## Primer uso

1. Entre al panel administrador.
2. Cambie el menú del día y guárdelo.
3. Cree una empresa.
4. Copie el enlace privado que aparece en la tabla.
5. Envíe ese enlace a los empleados.
6. Los pedidos aparecerán en el panel.

## Cambiar nombre del restaurante y contraseña

Abra `config.json` con Bloc de notas y cambie:

```json
{
  "restaurant_name": "Nombre del restaurante",
  "admin_password": "una-clave-segura",
  "secret_key": "una-clave-muy-larga-y-dificil-de-adivinar",
  "order_deadline": "10:30"
}
```

Reinicie `INICIAR.bat` después de guardar.

## Uso desde otros celulares o computadoras de la misma red Wi-Fi

1. En la PC del restaurante, abra CMD y ejecute `ipconfig`.
2. Busque la dirección IPv4, por ejemplo `192.168.1.50`.
3. En otro dispositivo de la misma red, abra `http://192.168.1.50:8080`.
4. En el panel, los enlaces de empresa deben usar esa IP en lugar de `localhost`.

Puede ser necesario permitir Python en el Firewall de Windows.

## Publicarlo en internet

Esta versión es funcional para pruebas, red local y operación pequeña. Para acceso público por internet se recomienda desplegarla detrás de HTTPS, cambiar contraseña y clave secreta, configurar copias de seguridad y usar un proveedor de hosting.

## Datos

La base se guarda en `datos/pedidos.db`. Para hacer una copia de seguridad, copie ese archivo con el programa cerrado.

## Versión preparada para hosting

Para publicarla en Render, use `README_HOSTING_RENDER.md`. Esta versión admite las variables de entorno `RESTAURANT_NAME`, `ADMIN_PASSWORD`, `SECRET_KEY`, `ORDER_DEADLINE`, `DATA_DIR` y `PORT`.
