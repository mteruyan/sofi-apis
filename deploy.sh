#!/usr/bin/env bash
set -euo pipefail

gcloud builds submit --tag gcr.io/sanitas-recetas-prod/sofi-apis && \
gcloud run deploy sofi-apis \
  --image gcr.io/sanitas-recetas-prod/sofi-apis \
  --region us-central1 \
  --allow-unauthenticated \
  --add-cloudsql-instances=sanitas-recetas-prod:us-central1:sofi-whatsapp
