#!/bin/bash
# Prueba variantes del formato del DID sobre el endpoint /agents/{did}
# Reemplaza TOKEN por un access_token fresco de webhook@sanitas

TOKEN="eyJhbGciOiJIUzUxMiJ9.eyJyb2xlc05hbWVzIjpbIkFQSS1DSEFUVElHTyJdLCJzdWIiOiJ3ZWJob29rQHNhbml0YXMiLCJyb2xlIjpbIjE2Il0sImlkQ2xpZW50IjoxODYyLCJjYW1wYWluZ3MiOlsiODY3NSIsIjg3MDYiLCI4NzA3Il0sInByb2ZpbGUiOiIxIiwiaWRVc3VhcmlvIjo1NDgyNSwic2Vzc2lvbklkIjoiMTkyMWYxMWY2MDY5NDllNzk1ZTQ4OGZmOTNiMTcxZDIiLCJwZXJtaXNzaW9uc05hbWUiOltdLCJpZFVzZXJBcGkiOjU1MDI3LCJpZHNlc2lvbiI6NDY1OTcxNzgsImVuYWJsZSI6dHJ1ZSwibmFtZSI6IndlYmhvb2tAc2FuaXRhcyIsImV4cCI6MTc4NzEzNzAwMiwiaWF0IjoxNzg3MDkzODAyfQ.oC5eIwcQfHq_mQOFTsb2GzenYjggwj4845ZR4pV3U8Cnh04LbIKdy9wiVkSm98hyhzHsa746ljNFxXZv6OWdxg"
BASE="https://channels.chattigo.com/webhook/chattigo/agents"

DIDS=(
  "51955973936"
  "+51955973936"
  "955973936"
  "0051955973936"
)

for DID in "${DIDS[@]}"; do
  echo "--- Probando did=${DID} ---"
  curl -s -o /tmp/resp.json -w "Status: %{http_code}\n" \
    "${BASE}/${DID}" \
    -H "Authorization: Bearer ${TOKEN}"
  cat /tmp/resp.json
  echo -e "\n"
done
