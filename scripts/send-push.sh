#!/usr/bin/env bash
# Manda una notificacion push (Firebase Cloud Messaging) a todos los dispositivos activados en
# pushTokens, usando una cuenta de servicio de Firebase (no el SDK cliente de la app). Pensado
# para correr desde una rutina de Claude Code (RemoteTrigger), no desde la app.
#
# Variables de entorno requeridas:
#   FIREBASE_SERVICE_ACCOUNT_JSON  — contenido completo del JSON de la cuenta de servicio
#                                    (Firebase Console → Project settings → Service accounts →
#                                    Generate new private key).
# Argumentos: send-push.sh "<titulo>" "<cuerpo>"
set -euo pipefail

TITULO="${1:?Falta el titulo}"
CUERPO="${2:?Falta el cuerpo}"

if [ -z "${FIREBASE_SERVICE_ACCOUNT_JSON:-}" ]; then
  echo "Falta la variable de entorno FIREBASE_SERVICE_ACCOUNT_JSON" >&2
  exit 1
fi

PROJECT_ID=$(echo "$FIREBASE_SERVICE_ACCOUNT_JSON" | jq -r .project_id)
CLIENT_EMAIL=$(echo "$FIREBASE_SERVICE_ACCOUNT_JSON" | jq -r .client_email)
PRIVATE_KEY=$(echo "$FIREBASE_SERVICE_ACCOUNT_JSON" | jq -r .private_key)

b64url() { openssl base64 -A | tr '+/' '-_' | tr -d '='; }

NOW=$(date +%s)
EXP=$((NOW + 3600))
HEADER=$(printf '{"alg":"RS256","typ":"JWT"}' | b64url)
CLAIMS=$(printf '{"iss":"%s","scope":"https://www.googleapis.com/auth/firebase.messaging https://www.googleapis.com/auth/datastore","aud":"https://oauth2.googleapis.com/token","iat":%s,"exp":%s}' "$CLIENT_EMAIL" "$NOW" "$EXP" | b64url)
UNSIGNED="$HEADER.$CLAIMS"
SIGNATURE=$(printf '%s' "$UNSIGNED" | openssl dgst -sha256 -sign <(printf '%s' "$PRIVATE_KEY") | b64url)
JWT="$UNSIGNED.$SIGNATURE"

ACCESS_TOKEN=$(curl -s -X POST https://oauth2.googleapis.com/token \
  -d "grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=$JWT" \
  | jq -r .access_token)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
  echo "No se pudo obtener el access token" >&2
  exit 1
fi

TOKENS_JSON=$(curl -s -H "Authorization: Bearer $ACCESS_TOKEN" \
  "https://firestore.googleapis.com/v1/projects/$PROJECT_ID/databases/(default)/documents/pushTokens")

TOKENS=$(echo "$TOKENS_JSON" | jq -r '.documents[]?.fields.token.stringValue // empty')

if [ -z "$TOKENS" ]; then
  echo "No hay dispositivos con notificaciones activadas (pushTokens vacio)."
  exit 0
fi

ENVIADOS=0
FALLIDOS=0
while IFS= read -r TOKEN; do
  [ -z "$TOKEN" ] && continue
  BODY=$(jq -n --arg t "$TOKEN" --arg titulo "$TITULO" --arg cuerpo "$CUERPO" --arg link "https://gaston1505.github.io/Aeon-stock-app/" \
    '{message: {token: $t, notification: {title: $titulo, body: $cuerpo}, webpush: {fcm_options: {link: $link}}}}')
  RESP=$(curl -s -o /dev/null -w "%{http_code}" -X POST \
    "https://fcm.googleapis.com/v1/projects/$PROJECT_ID/messages:send" \
    -H "Authorization: Bearer $ACCESS_TOKEN" -H "Content-Type: application/json" \
    -d "$BODY")
  if [ "$RESP" = "200" ]; then
    ENVIADOS=$((ENVIADOS + 1))
  else
    FALLIDOS=$((FALLIDOS + 1))
    echo "Fallo enviando a token $TOKEN (HTTP $RESP)" >&2
  fi
done <<< "$TOKENS"

echo "Push enviado: $ENVIADOS ok, $FALLIDOS fallidos."
