# Citas Pro — guía de despliegue

Sistema de reservas para peluquerías, estética, uñas y masajes. Multi-negocio: cada negocio tiene su panel, su enlace de reservas (`?n=su-negocio`) y su suscripción prepago (14 días de prueba, luego 30 días por pago con myPOS).

## 1. Backend (Google Sheet + Apps Script)
1. Crea un Google Sheet nuevo llamado `citas pro`.
2. Extensiones → Apps Script. Pega todo `Codigo.gs`.
3. Ajustes del proyecto → marca "Mostrar el archivo de manifiesto appsscript.json" y pega el contenido de `appsscript.json` (zona horaria Europe/Madrid).
4. Ejecuta la función `setup` una vez y acepta los permisos (crea las pestañas).
5. Implementar → Nueva implementación → Aplicación web → Ejecutar como: yo · Acceso: cualquier usuario. Copia la URL que acaba en `/exec`.
6. Cada cambio posterior en `Codigo.gs` necesita Implementar → Gestionar implementaciones → Nueva versión.

## 2. Frontend (GitHub Pages)
1. En `index.html`, sustituye `PEGA_AQUI_LA_URL_EXEC` por tu URL `/exec`.
2. Sube a un repo (por ejemplo `puntofibra/citas`): `index.html`, `sw.js`, `manifest.webmanifest`, `icon-192.png`, `icon-512.png`, `maskable-512.png`.
3. Settings → Pages → rama main, carpeta raíz.

## 3. Tu acceso de administrador
- Abre la web → Entrar → "He olvidado mi contraseña" con `futurmovil.com@gmail.com`. Te llega un código por correo y fijas tu contraseña.
- Ya dentro, botón **Authenticator → Configurar** y escanea el QR. Desde ese momento entras dejando el correo vacío y escribiendo el código de 6 cifras.
- **Llave NFC:** mantén pulsado "Entrar" 3 segundos. Para que funcione hay que pegar en `Codigo.gs` (bloque "LLAVE NFC") las funciones `llaveSesion` y `llaveEstado` de Bonificaciones; devuelven `{estado:'ok', token: nuevaSesion('ADMIN')}` cuando la llave es válida. Si usan `UrlFetchApp`, añade el permiso `https://www.googleapis.com/auth/script.external_request` en `appsscript.json`.

## 4. Cómo cobras
Cada negocio ve en "Plan" el botón de pago `https://mypos.com/@futurmovil/19`. Cuando veas el cobro, en tu panel admin pulsas **+30 d** en ese negocio (si le quedaban días, se suman). Para cambiar el precio edita `PRECIO_MES` en `Codigo.gs`.

## Notas
- Las contraseñas no viajan en claro: el navegador envía un SHA-256 y el servidor lo vuelve a cifrar con sal.
- Los clientes finales no ven ninguna mención a Google.
- Imágenes de logo y portada: por enlace https (sin subir archivos).
