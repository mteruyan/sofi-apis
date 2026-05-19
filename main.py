import json
import os
import pytz
import psycopg2
from datetime import datetime
import functions_framework

# ─────────────────────────────────────────────
# CONEXIÓN
# ─────────────────────────────────────────────
def get_connection():
    return psycopg2.connect(
        dbname="sofi_whatsapp",
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        host=f"/cloudsql/{os.environ['CLOUD_SQL_CONNECTION_NAME']}"
    )

# ─────────────────────────────────────────────
# HEADERS
# ─────────────────────────────────────────────
HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
}

# ─────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────
@functions_framework.http
def api(request):
    if request.method == "OPTIONS":
        return ("", 204, HEADERS)

    path = request.path.rstrip("/")

    if path == "/bitacora":
        return handle_bitacora(request)

    return (json.dumps({"error": f"Ruta no encontrada: {path}"}), 404, HEADERS)


# ─────────────────────────────────────────────
# API: /bitacora
# ─────────────────────────────────────────────
def handle_bitacora(request):
    data = request.get_json(silent=True) or {}
    message_id = data.get("message_id", "")

    if not message_id:
        return (json.dumps({"error": "El campo message_id es obligatorio"}), 400, HEADERS)

    try:
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            INSERT INTO sofi_bitacora (
                message_id, chat_id, chat_type,
                channel_identifier, channel_type, user_name,
                campaign_id, user_identifier, first_user_input,
                last_user_input, id_agent, username_agent, fecha
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            message_id,
            data.get("chat_id", ""),
            data.get("chat_type", ""),
            data.get("channel_identifier", ""),
            data.get("channel_type", ""),
            data.get("user_name", ""),
            data.get("campaign_id", ""),
            data.get("user_identifier", ""),
            data.get("first_user_input", ""),
            data.get("last_user_input", ""),
            data.get("id_agent", ""),
            data.get("username_agent", ""),
            fecha
        ))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "Linea insertada"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

