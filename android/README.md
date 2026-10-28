# Sanitas Proveedor — app Android (uso interno, fuera de tiendas)

Envuelve el panel `https://sntspe.com/proveedor` y agrega un **servicio en primer plano** que:
- envía la ubicación cada 15 s aunque la app esté minimizada o el celular bloqueado;
- muestra una notificación fija "Compartiendo tu ubicación" mientras el doctor está conectado;
- avisa con notificación sonora cuando llega una oferta (revisa el sistema cada 15 s, sin Firebase).

El doctor inicia sesión igual que en el PWA (usuario, clave, Google Authenticator).

## 1. Generar el APK firmado — SIN Android Studio (recomendado): desde Cloud Shell
Requiere solo tu proyecto de Google Cloud (el mismo de sofi-apis).
1. Abrir **Cloud Shell**, subir y descomprimir el zip: `unzip sanitas-proveedor-android.zip -d android && cd android`
2. **Una sola vez**: `./construir_apk.sh configurar`
   - Crea la llave de firma y la guarda en Secret Manager.
   - **Descarga `sanitas.jks` (menú ⋮ de Cloud Shell → Download) y guarda ese archivo y la clave que muestra la pantalla.**
     Sin la misma llave no podrás publicar actualizaciones que se instalen sobre la versión anterior.
3. **Cada vez que quieras un APK**: `./construir_apk.sh construir` (tarda unos 10 minutos; la primera vez más).
   Al terminar deja `sanitas-proveedor.apk` en esa carpeta.
4. Si algo falla, el comando muestra el enlace al registro de Cloud Build; copia el error y envíalo para corregirlo.

### Alternativa con Android Studio
Abrir la carpeta, **Build → Generate Signed App Bundle / APK → APK**, crear el `.jks` y elegir *release*.

## 2. Publicarlo para tus proveedores
1. Copiar `sanitas-proveedor.apk` a `static/app/sanitas-proveedor.apk` en el proyecto `sofi-apis` y hacer `./deploy.sh`.
   (Verifica que tu `main.py` sirva `/static/app/…` como archivo público.)
2. El PWA ya muestra la tarjeta **"📱 App para Android"** con el botón de descarga a quien lo abre desde Android.
3. Cada doctor: descargar → abrir el archivo → permitir "instalar apps de esta fuente" (solo la primera vez) → Instalar.

## 3. Primer uso en cada celular
1. Abrir la app, iniciar sesión.
2. Permitir **Ubicación** ("Mientras se usa la app" basta: el servicio mantiene la ubicación en segundo plano).
3. Permitir **Notificaciones**.
4. Aceptar **"Sin restricciones de batería"**. En Xiaomi/Redmi, Huawei, Oppo, Vivo y Samsung conviene además:
   Ajustes → Apps → Sanitas Proveedor → Batería → *Sin restricciones*, y activar *Inicio automático* (Xiaomi/Oppo/Vivo).
5. Activar el interruptor de disponibilidad. Debe aparecer la notificación fija de ubicación.

## 4. Publicar una versión nueva
1. En `app/build.gradle` subir `versionCode` (+1) y cambiar `versionName` (ej. `1.1`).
2. Volver a ejecutar `./construir_apk.sh construir` (usa la misma llave) y reemplazar `static/app/sanitas-proveedor.apk`.
3. Poner el mismo `versionName` en la variable de Cloud Run `MAD_APK_VERSION` (`gcloud run services update sofi-apis --update-env-vars MAD_APK_VERSION=1.1 …`).
   A los doctores con la app les aparece "Hay una versión nueva de la app → Descargar e instalar".
Los cambios del panel web (pantallas, textos) **no** requieren un APK nuevo: se actualizan solos.

## Notas
- iPhone: esta app es solo Android. Para iPhone se necesitaría TestFlight/App Store privada de Apple.
- La app solo carga páginas de tu dominio; enlaces externos (WhatsApp, llamadas) se abren fuera.
- Si el doctor cierra la sesión o lo desconecta el sistema, el servicio se detiene y avisa.
