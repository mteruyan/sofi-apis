#!/usr/bin/env bash
# Genera el APK de Sanitas Proveedor usando Google Cloud (sin Android Studio).
#   ./construir_apk.sh configurar   → UNA sola vez: crea la llave de firma y el bucket
#   ./construir_apk.sh construir    → cada vez que quieras un APK (tarda ~10 min)
set -euo pipefail
cd "$(dirname "$0")"

PROJECT="$(gcloud config get-value project 2>/dev/null)"
NUM="$(gcloud projects describe "$PROJECT" --format='value(projectNumber)')"
BUCKET="gs://${PROJECT}-apk"

case "${1:-}" in
  configurar)
    gcloud services enable cloudbuild.googleapis.com secretmanager.googleapis.com
    gsutil ls "$BUCKET" >/dev/null 2>&1 || gsutil mb -l us-central1 "$BUCKET"

    if gcloud secrets describe android-keystore >/dev/null 2>&1; then
      echo "Ya existe la llave de firma (android-keystore). No se vuelve a crear."
    else
      PASS="$(head -c 24 /dev/urandom | base64 | tr -dc 'A-Za-z0-9' | head -c 20)"
      keytool -genkeypair -v -keystore sanitas.jks -alias sanitas -keyalg RSA -keysize 2048 -validity 10000 \
        -storepass "$PASS" -keypass "$PASS" -dname "CN=Sanitas Proveedor, O=Sanitas Peru, C=PE"
      gcloud secrets create android-keystore --data-file=sanitas.jks
      printf %s "$PASS" | gcloud secrets create android-keystore-pass --data-file=-
      echo
      echo "=================================================================="
      echo "IMPORTANTE: descarga AHORA el archivo sanitas.jks (menú ⋮ de Cloud Shell → Download)"
      echo "y guárdalo junto con esta clave en un lugar seguro:   $PASS"
      echo "Sin ellos no podrás publicar actualizaciones de la app."
      echo "(También quedan guardados en Secret Manager del proyecto.)"
      echo "=================================================================="
    fi

    # Permisos para que Cloud Build lea la llave y guarde el APK (según el proyecto, usa una de estas dos cuentas)
    for SA in "${NUM}@cloudbuild.gserviceaccount.com" "${NUM}-compute@developer.gserviceaccount.com"; do
      for S in android-keystore android-keystore-pass; do
        gcloud secrets add-iam-policy-binding "$S" --member="serviceAccount:$SA" \
          --role=roles/secretmanager.secretAccessor >/dev/null 2>&1 || true
      done
      gsutil iam ch "serviceAccount:$SA:objectAdmin" "$BUCKET" >/dev/null 2>&1 || true
    done
    echo "Listo. Ahora ejecuta:  ./construir_apk.sh construir"
    ;;

  construir)
    gcloud builds submit . --config cloudbuild.yaml --substitutions="_BUCKET=${BUCKET}${IMAGE:+,_IMAGE=$IMAGE}"
    gsutil cp "${BUCKET}/sanitas-proveedor.apk" ./sanitas-proveedor.apk
    echo
    echo "APK listo: $(pwd)/sanitas-proveedor.apk"
    echo "Cópialo a  static/app/sanitas-proveedor.apk  en tu proyecto sofi-apis y haz ./deploy.sh"
    ;;

  *)
    echo "Uso: $0 configurar | construir"; exit 1;;
esac
