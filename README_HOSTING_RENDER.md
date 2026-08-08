# Publicar el sistema en Render con dominio y HTTPS

Esta carpeta ya está preparada para Render.

## Antes de subirla

No suba a internet una base de datos local con pedidos reales. La carpeta `datos` debe estar vacía.

## 1. Crear el repositorio en GitHub

1. Ingrese a GitHub y cree un repositorio privado.
2. Seleccione **Add file > Upload files**.
3. Suba todos los archivos de esta carpeta, incluido `render.yaml`.
4. Confirme con **Commit changes**.

## 2. Crear el servicio en Render

1. Ingrese a Render usando su cuenta de GitHub.
2. Abra **Blueprints** y seleccione **New Blueprint Instance**.
3. Seleccione el repositorio que acaba de crear.
4. Render leerá `render.yaml`.
5. Cuando lo solicite, complete:
   - `RESTAURANT_NAME`: nombre de su restaurante.
   - `ADMIN_PASSWORD`: una contraseña larga y privada.
6. Confirme el despliegue.

El archivo configura automáticamente:

- Python.
- Puerto público proporcionado por Render.
- HTTPS en la dirección `.onrender.com`.
- Disco persistente para `datos/pedidos.db`.
- Clave secreta aleatoria.
- Revisión de estado en `/health`.

## 3. Primera prueba

Cuando el servicio figure como **Live**, abra la URL indicada por Render, por ejemplo:

`https://pedidos-almuerzos.onrender.com/admin`

Entre con la contraseña configurada en `ADMIN_PASSWORD`.

## 4. Conectar un dominio propio

1. En Render, abra el servicio.
2. Entre a **Settings > Custom Domains**.
3. Pulse **Add Custom Domain**.
4. Escriba su dominio o subdominio, por ejemplo `pedidos.mirestaurante.com`.
5. Render mostrará los registros DNS que debe crear donde compró el dominio.
6. Cree esos registros y vuelva a Render.
7. Pulse **Verify**.

Render emitirá el certificado TLS y redirigirá automáticamente de HTTP a HTTPS.

## Copia de seguridad

La base está en `/opt/render/project/src/datos/pedidos.db`. El disco persistente conserva los pedidos entre reinicios y despliegues. Mantenga además una copia periódica del archivo SQLite.
