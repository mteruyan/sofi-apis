import json
import json
import json
import os
import pytz
import psycopg2
from datetime import datetime, timedelta
import functions_framework
import re
import requests
from difflib import SequenceMatcher
import jellyfish

DANACONNECT_URL  = "https://appserv.danaconnect.com/api/1.0/rest/conversation/517078/start/data"
DANACONNECT_AUTH = os.environ["DANACONNECT_AUTH"]
WHATSAPP_URL     = "wa.me/51955973936"

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
# VALIDACIÓN API KEY
# ─────────────────────────────────────────────
def validar_api_key(request):
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("ApiKey "):
        return False
    key = auth.split(" ", 1)[1]
    return key == os.environ.get("API_KEY", "")

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
    
    if not validar_api_key(request):
        return (json.dumps({"error": "API Key invalida"}), 401, HEADERS)

    path = request.path.rstrip("/")

    if path == "/bitacora":
        return handle_bitacora(request)
    elif path == "/update_bitacora":
        return handle_update_bitacora(request)
    elif path == "/whatsapp_busca_doi":
        return handle_whatsapp_busca_doi(request)
    elif path == "/busca_doi_completo":
        return handle_busca_doi_completo(request)
    elif path == "/update_doi":
        return handle_update_doi(request)
    elif path == "/update_escoge":
        return handle_update_escoge(request)
    elif path == "/update_doi_direccion":
        return handle_update_doi_direccion(request)
    elif path == "/update_doi_distrito":
        return handle_update_doi_distrito(request)
    elif path == "/listar_distritos":
        return handle_listar_distritos(request)
    elif path == "/trae_escoge":
        return handle_trae_escoge(request)
    elif path == "/inserta_bitacora_ivr":
        return handle_inserta_bitacora_ivr(request)
    elif path == "/lead_medicamentos":
        return handle_lead_medicamentos(request)
    elif path == "/lead_sctr":
        return handle_lead_sctr(request)
    elif path == "/lead_consultas":
        return handle_lead_consultas(request)
    elif path == "/busca_doi_potestativo":
        return handle_busca_doi_potestativo(request)
    elif path == "/busca_doi_potestativo_carta":
        return handle_busca_doi_potestativo_carta(request)
    elif path == "/envia_sms_leads_sctr":
        return handle_envia_sms_leads_sctr(request)
    elif path == "/envia_sms_leads_medicamentos":
        return handle_envia_sms_leads_medicamentos(request)
    elif path == "/envia_sms_leads_consultas":
        return handle_envia_sms_leads_consultas(request)
    elif path == "/envia_sctr_vencidos":
        return handle_envia_sctr_vencidos(request)
    elif path == "/recordatorio_sctr":
        return handle_recordatorio_sctr(request)
    elif path == "/busca_ultimo_registro":
        return handle_busca_ultimo_registro(request)
    elif path == "/consulta_ruc":
        return handle_consulta_ruc(request)
    elif path == "/busca_actividad":
        return handle_busca_actividad(request)
    elif path == "/busca_carta_rechazada":
        return handle_busca_carta_rechazada(request)
    elif path == "/update_cg_rechazados":
        return handle_update_cg_rechazados(request)
    elif path == "/envia_sms_leads_sctr":
        return handle_envia_sms_leads_sctr(request)
    elif path == "/envia_sms_leads_medicamentos":
        return handle_envia_sms_leads_medicamentos(request)
    elif path == "/envia_sms_leads_consultas":
        return handle_envia_sms_leads_consultas(request)

    return (json.dumps({"error": f"Ruta no encontrada: {path}"}), 404, HEADERS)

# ─────────────────────────────────────────────
# UTILIDAD: oculta texto dejando solo primera letra
# ─────────────────────────────────────────────
def convierte(texto):
    palabras = texto.split(" ")
    resultado = []
    for palabra in palabras:
        if len(palabra) > 0:
            resultado.append(palabra[0] + "x" * (len(palabra) - 1))
    return " ".join(resultado)

# ─────────────────────────────────────────────
# API: /bitacora
# ─────────────────────────────────────────────
def handle_bitacora(request):
    #print("RAW BODY:", request.get_data(as_text=True))
    #print("CONTENT-TYPE:", request.content_type)
    
    #data = request.get_json(force=True, silent=True) or {}
    #print("DATA PARSEADA:", data)

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

# ─────────────────────────────────────────────
# API: /update_bitacora (UPDATE)
# ─────────────────────────────────────────────
def handle_update_bitacora(request):
    data       = request.get_json(silent=True) or {}
    message_id = data.get("message_id", "")
    chat_id    = data.get("chat_id", "")

    if not message_id:
        return (json.dumps({"error": "El campo message_id es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            UPDATE sofi_bitacora
            SET username_agent = %s
            WHERE chat_id = %s
        """, (
            data.get("username_agent", ""),
            chat_id
        ))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "Linea actualizada"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# API: /whatsapp_busca_doi (SELECT)
# ─────────────────────────────────────────────
def handle_whatsapp_busca_doi(request):
    data             = request.get_json(silent=True) or {}
    celular_consulta = data.get("celular_consulta", "")

    if not celular_consulta:
        return (json.dumps({"error": "El campo celular_consulta es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            SELECT doi FROM sofi_ivr
            WHERE user_identifier = %s
              AND opcion NOT IN ('SCTR hablar con agente', 'SCTR frecuentes preguntas')
            ORDER BY fecha DESC
            LIMIT 1
        """, (celular_consulta,))

        row = cur.fetchone()
        cur.close()
        conn.close()

        if row:
            doi      = row[0].title() if row[0] else ""
            contador = 1
        else:
            doi      = ""
            contador = 0

        return (json.dumps({"doi": doi, "contador": str(contador)}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /busca_doi_completo (SELECT)
# ─────────────────────────────────────────────
def handle_busca_doi_completo(request):
    data             = request.get_json(silent=True) or {}
    doi              = data.get("doi", "")
    celular_consulta = data.get("celular_consulta", "")

    nombres=""; apellidos=""; direccion1=""; distrito1=""
    direccion2=""; distrito2=""; direccion3=""; distrito3=""
    nro_contacto1=""; nro_contacto2=""; email=""; escoge=""
    mensaje=""; contador=0

    try:
        conn = get_connection()
        cur  = conn.cursor()

        if doi:
            # Con doi: busca directo en sofi_usuarios
            cur.execute("""
                SELECT doi, nombres, apellidos,
                       direccion1, distrito1, direccion2, distrito2,
                       direccion3, distrito3, nro_contacto1, nro_contacto2,
                       email, escoge
                FROM sofi_usuarios WHERE doi = %s
            """, (doi,))
            row = cur.fetchone()
            if row:
                doi           = row[0]  or ""
                nombres       = row[1].upper()  if row[1]  else ""
                apellidos     = row[2].upper()  if row[2]  else ""
                direccion1    = row[3].upper()  if row[3]  else ""
                distrito1     = row[4].upper()  if row[4]  else ""
                direccion2    = row[5].upper()  if row[5]  else ""
                distrito2     = row[6].upper()  if row[6]  else ""
                direccion3    = row[7].upper()  if row[7]  else ""
                distrito3     = row[8].upper()  if row[8]  else ""
                nro_contacto1 = row[9]  or ""
                nro_contacto2 = row[10] or ""
                email         = row[11] or ""
                escoge        = row[12] or ""
                mensaje       = "DOI encontrado"
                contador      = 1

        else:
            # Sin doi: busca doi en sofi_ivr por celular
            cur.execute("""
                SELECT doi FROM sofi_ivr
                WHERE user_identifier = %s
                  AND opcion NOT IN ('SCTR hablar con agente', 'SCTR frecuentes preguntas')
                ORDER BY fecha DESC LIMIT 1
            """, (celular_consulta,))
            row_ivr = cur.fetchone()
            doi = row_ivr[0].title() if row_ivr and row_ivr[0] else ""

            if doi:
                cur.execute("""
                    SELECT doi, nombres, apellidos,
                           direccion1, distrito1, direccion2, distrito2,
                           direccion3, distrito3, nro_contacto1, nro_contacto2,
                           email, escoge
                    FROM sofi_usuarios WHERE doi = %s
                """, (doi,))
                row = cur.fetchone()
                if row:
                    doi           = row[0]  or ""
                    nombres       = row[1].upper()  if row[1]  else ""
                    apellidos     = row[2].upper()  if row[2]  else ""
                    direccion1    = row[3].upper()  if row[3]  else ""
                    distrito1     = row[4].upper()  if row[4]  else ""
                    direccion2    = row[5].upper()  if row[5]  else ""
                    distrito2     = row[6].upper()  if row[6]  else ""
                    direccion3    = row[7].upper()  if row[7]  else ""
                    distrito3     = row[8].upper()  if row[8]  else ""
                    nro_contacto1 = row[9]  or ""
                    nro_contacto2 = row[10] or ""
                    email         = row[11] or ""
                    escoge        = row[12] or ""
                    mensaje       = "DOI encontrado"
                    contador      = 2

        # Si celular no coincide con contactos, ocultar direcciones
        if celular_consulta not in nro_contacto1 and celular_consulta not in nro_contacto2:
            if direccion1: direccion1 = convierte(direccion1)
            if distrito1:  distrito1  = convierte(distrito1)
            if direccion2: direccion2 = convierte(direccion2)
            if distrito2:  distrito2  = convierte(distrito2)
            if direccion3: direccion3 = convierte(direccion3)
            if distrito3:  distrito3  = convierte(distrito3)

        cur.close()
        conn.close()

        # Armar mensaje con direcciones
        mensaje_dir  = ""
        contador_dir = 0
        if direccion1 and distrito1:
            contador_dir += 1
            mensaje_dir  += f"1.- {direccion1}, {distrito1}\n\n"
        if direccion2 and distrito2:
            contador_dir += 1
            mensaje_dir  += f"2.- {direccion2}, {distrito2}\n\n"
        if direccion3 and distrito3:
            contador_dir += 1
            mensaje_dir  += f"3.- {direccion3}, {distrito3}\n"

        return (json.dumps({
            "mensaje":              mensaje_dir,
            "doi":                  doi,
            "nombres":              nombres,
            "apellidos":            apellidos,
            "direccion1":           direccion1,
            "distrito1":            distrito1,
            "direccion2":           direccion2,
            "distrito2":            distrito2,
            "direccion3":           direccion3,
            "distrito3":            distrito3,
            "nro_contacto1":        nro_contacto1,
            "nro_contacto2":        nro_contacto2,
            "email":                email,
            "escoge":               escoge,
            "contador":             str(contador),
            "contador_direcciones": contador_dir,
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /update_doi (INSERT si no existe)
# ─────────────────────────────────────────────
def handle_update_doi(request):
    data = request.get_json(silent=True) or {}
    doi  = data.get("doi", "")

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    nombres       = data.get("nombres", "").strip().upper()
    apellidos     = data.get("apellidos", "").strip().upper()
    direccion1    = data.get("direccion1", "").strip().upper()
    distrito1     = data.get("distrito1", "").strip().upper()
    nro_contacto1 = data.get("nro_contacto1", "")
    nro_contacto2 = data.get("nro_contacto2", "")
    email         = data.get("email", "").strip().lower()

    try:
        conn = get_connection()
        cur  = conn.cursor()

        # Verifica si el doi ya existe
        cur.execute("SELECT COUNT(*) FROM sofi_usuarios WHERE doi = %s", (doi,))
        contador = cur.fetchone()[0]

        if contador == 0:
            cur.execute("""
                INSERT INTO sofi_usuarios (
                    doi, nombres, apellidos, direccion1, distrito1,
                    nro_contacto1, nro_contacto2, email, escoge
                ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,'1')
            """, (
                doi, nombres, apellidos, direccion1, distrito1,
                nro_contacto1, nro_contacto2, email
            ))
            conn.commit()
            mensaje = "DOI insertado"
        else:
            mensaje = "DOI actualizado"

        cur.close()
        conn.close()
        return (json.dumps({"mensaje": mensaje}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /update_escoge (UPDATE)
# ─────────────────────────────────────────────
def handle_update_escoge(request):
    data   = request.get_json(silent=True) or {}
    doi    = data.get("doi", "")
    escoge = data.get("escoge", "")

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            UPDATE sofi_usuarios
            SET escoge = %s
            WHERE doi = %s
        """, (escoge, doi))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "DOI actualizado"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /update_doi_direccion (UPDATE)
# ─────────────────────────────────────────────
def handle_update_doi_direccion(request):
    data          = request.get_json(silent=True) or {}
    doi           = data.get("doi", "")
    direccion     = data.get("direccion", "").upper()
    distrito      = data.get("distrito", "").upper()
    direccion_nro = data.get("direccion_nro", 0)

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    if direccion_nro not in [1, 2, 3]:
        return (json.dumps({"error": "El campo direccion_nro debe ser 1, 2 o 3"}), 400, HEADERS)

    campos = {
        1: ("direccion1", "distrito1"),
        2: ("direccion2", "distrito2"),
        3: ("direccion3", "distrito3"),
    }
    col_dir, col_dis = campos[direccion_nro]

    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute(f"""
            UPDATE sofi_usuarios
            SET {col_dir} = %s, {col_dis} = %s
            WHERE doi = %s
        """, (direccion, distrito, doi))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "DOI actualizado"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /update_doi_distrito (UPDATE por nombre o ubigeo)
# ─────────────────────────────────────────────
def handle_update_doi_distrito(request):
    data                   = request.get_json(silent=True) or {}
    doi                    = data.get("doi", "")
    nombre_ubigeo_distrito = data.get("nombre_ubigeo_distrito", "").strip()
    va_nombre_o_ubigeo     = data.get("va_nombre_o_ubigeo", "")
    direccion_nro          = data.get("direccion_nro", 0)

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    if va_nombre_o_ubigeo not in ["nombre", "ubigeo"]:
        return (json.dumps({"error": "va_nombre_o_ubigeo debe ser 'nombre' o 'ubigeo'"}), 400, HEADERS)

    if direccion_nro not in [1, 2, 3]:
        return (json.dumps({"error": "direccion_nro debe ser 1, 2 o 3"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()

        # Busca en sofi_distritos por nombre o ubigeo
        if va_nombre_o_ubigeo == "nombre":
            cur.execute("""
                SELECT ubigeo, departamento, provincia, distrito
                FROM sofi_distritos
                WHERE distrito = %s
            """, (nombre_ubigeo_distrito,))
        else:
            cur.execute("""
                SELECT ubigeo, departamento, provincia, distrito
                FROM sofi_distritos
                WHERE ubigeo = %s
            """, (nombre_ubigeo_distrito,))

        row      = cur.fetchone()
        contador = 0
        mensaje  = "Ocurrió un error"

        if row:
            contador     = 1
            linea_grabar = f"{row[3]}-{row[2]}-{row[1]}"

            col_dis = {1: "distrito1", 2: "distrito2", 3: "distrito3"}[direccion_nro]

            cur.execute(f"""
                UPDATE sofi_usuarios
                SET {col_dis} = %s
                WHERE doi = %s
            """, (linea_grabar, doi))
            conn.commit()
            mensaje = "Se actualizo el dato de manera correcta"

        cur.close()
        conn.close()

        return (json.dumps({
            "doi":      doi,
            "mensaje":  mensaje,
            "contador": str(contador),
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /listar_distritos (SELECT)
# ─────────────────────────────────────────────
def handle_listar_distritos(request):
    data     = request.get_json(silent=True) or {}
    distrito = data.get("distrito", "").strip()

    if not distrito:
        return (json.dumps({"error": "El campo distrito es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            SELECT ubigeo, departamento, provincia, distrito
            FROM sofi_distritos
            WHERE distrito LIKE %s
        """, (f"%{distrito}%",))

        rows     = cur.fetchall()
        contador = len(rows)
        imprime  = ""

        for row in rows:
            imprime += f"*Ubigeo:* {row[0]} - {row[1]}-{row[2]}-{row[3]}\n"

        cur.close()
        conn.close()

        return (json.dumps({
            "mensaje":  imprime,
            "contador": str(contador),
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

    # ─────────────────────────────────────────────
# API: /trae_escoge (SELECT)
# ─────────────────────────────────────────────
def handle_trae_escoge(request):
    data             = request.get_json(silent=True) or {}
    doi              = data.get("doi", "")
    celular_consulta = data.get("celular_consulta", "")

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()

        nombres=""; apellidos=""; direccion1=""; distrito1=""
        direccion2=""; distrito2=""; direccion3=""; distrito3=""
        nro_contacto1=""; nro_contacto2=""; email=""; escoge=""
        direccion=""; distrito=""; mensaje=""; contador=0

        cur.execute("""
            SELECT nombres, apellidos,
                   direccion1, distrito1, direccion2, distrito2,
                   direccion3, distrito3, nro_contacto1, nro_contacto2,
                   email, escoge
            FROM sofi_usuarios WHERE doi = %s
        """, (doi,))
        row = cur.fetchone()

        if row:
            nombres       = row[0].title()  if row[0]  else ""
            apellidos     = row[1].title()  if row[1]  else ""
            direccion1    = row[2].title()  if row[2]  else ""
            distrito1     = row[3].title()  if row[3]  else ""
            direccion2    = row[4].title()  if row[4]  else ""
            distrito2     = row[5].title()  if row[5]  else ""
            direccion3    = row[6].title()  if row[6]  else ""
            distrito3     = row[7].title()  if row[7]  else ""
            nro_contacto1 = row[8]  or ""
            nro_contacto2 = row[9]  or ""
            email         = row[10] or ""
            escoge        = str(row[11]) if row[11] else ""
            mensaje       = "DOI encontrado"
            contador      = 1

        # Selecciona dirección según escoge
        if escoge in ["1", "", None]:
            direccion = direccion1
            distrito  = distrito1
        elif escoge == "2":
            direccion = direccion2
            distrito  = distrito2
        elif escoge == "3":
            direccion = direccion3
            distrito  = distrito3

        # Si celular no coincide, ocultar dirección y enmascarar contacto
        if celular_consulta not in nro_contacto1 and celular_consulta not in nro_contacto2:
            if direccion: direccion = convierte(direccion)
            if distrito:  distrito  = convierte(distrito)

            if len(nro_contacto1) >= 7:
                nro_contacto1 = nro_contacto1[:5] + "x" * (len(nro_contacto1) - 5)

        # Busca último acceso en sofi_ivr
        ultimo_acceso = ""
        cur.execute("""
            SELECT fecha FROM sofi_ivr
            WHERE doi = %s
              AND opcion = 'Medicamentos-Coord de receta'
            ORDER BY fecha DESC
            LIMIT 1
        """, (doi,))
        row_ivr = cur.fetchone()
        if row_ivr:
            ultimo_acceso = str(row_ivr[0])

        cur.close()
        conn.close()

        return (json.dumps({
            "doi":           doi,
            "nombres":       nombres,
            "apellidos":     apellidos,
            "direccion":     direccion,
            "distrito":      distrito,
            "nro_contacto1": nro_contacto1,
            "nro_contacto2": nro_contacto2,
            "email":         email,
            "mensaje":       mensaje,
            "contador":      str(contador),
            "ultimo_acceso": ultimo_acceso,
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /inserta_bitacora_ivr (INSERT + Storage + Danaconnect)
# ─────────────────────────────────────────────
def handle_inserta_bitacora_ivr(request):
    data               = request.get_json(silent=True) or {}
    message_id         = data.get("message_id", "")
    chat_id            = data.get("chat_id", "")
    channel_identifier = data.get("channel_identifier", "")
    user_name          = data.get("user_name", "")
    campaign_id        = data.get("campaign_id", "")
    user_identifier    = data.get("user_identifier", "")
    last_user_input    = data.get("first_user_input", "")
    opcion             = data.get("opcion", "")
    doi                = data.get("doi", "")
    nombres            = data.get("nombres", "")
    apellidos          = data.get("apellidos", "")
    direccion          = data.get("direccion", "")
    distrito           = data.get("distrito", "")
    contacto           = data.get("contacto", "")
    url                = data.get("url", "")

    if not message_id:
        return (json.dumps({"error": "El campo message_id es obligatorio"}), 400, HEADERS)

    try:
        conn = get_connection()
        cur  = conn.cursor()

        # ── 1. Busca datos del usuario por doi ───────────────────────────────
        nro_contacto1 = ""; download_name = ""
        if doi:
            cur.execute("""
                SELECT nombres, apellidos,
                       direccion1, distrito1, direccion2, distrito2,
                       direccion3, distrito3, nro_contacto1, nro_contacto2,
                       email, escoge
                FROM sofi_usuarios WHERE doi = %s
            """, (doi,))
            row = cur.fetchone()
            if row:
                nombres       = row[0].title() if row[0] else nombres
                apellidos     = row[1].title() if row[1] else apellidos
                escoge        = str(row[11])   if row[11] else "1"
                nro_contacto1 = row[8] or ""

                if escoge in ["1", ""]:
                    direccion = row[2].title() if row[2] else ""
                    distrito  = row[3].title() if row[3] else ""
                elif escoge == "2":
                    direccion = row[4].title() if row[4] else ""
                    distrito  = row[5].title() if row[5] else ""
                elif escoge == "3":
                    direccion = row[6].title() if row[6] else ""
                    distrito  = row[7].title() if row[7] else ""

        # ── 2. Descarga archivo y sube a Cloud Storage ───────────────────────
        try:
            url_publica = url
            if url and opcion != "Consulta Asesor":
                import urllib.request
                from google.cloud import storage
                from urllib.parse import urlparse
                import requests as req_download

                download_name = os.path.basename(urlparse(url).path)

                response = req_download.get(url, timeout=30)

                if response.status_code != 200:
                    return (json.dumps({
                        "error": f"Error descargando archivo: status={response.status_code}, detalle={response.text[:200]}"
                    }), 500, HEADERS)

                file_data = response.content
                content_type = response.headers.get("Content-Type", "application/octet-stream")

                # Descarga el archivo en memoria
                #with urllib.request.urlopen(url) as response:
                #    file_data = response.read()

                # Sube a Cloud Storage
                storage_client = storage.Client()
                bucket = storage_client.bucket("sanitas-recetas-recetas")
                blob = bucket.blob(download_name)
                blob.upload_from_string(file_data, content_type=content_type)
                blob.make_public()

                url_publica = f"https://storage.googleapis.com/sanitas-recetas-recetas/{download_name}"

                # Sube a Cloud Storage
                #storage_client = storage.Client()
                #bucket         = storage_client.bucket("sanitas-recetas-recetas")
                #blob           = bucket.blob(download_name)
                #blob.upload_from_string(file_data)
                #blob.make_public()

                #url_publica = f"https://storage.googleapis.com/sanitas-recetas-recetas/{download_name}"

        except Exception as e:
            return (json.dumps({"error": f"Error en Storage: {str(e)}"}), 500, HEADERS)

        # ── 3. INSERT en sofi_ivr ─────────────────────────────────────────────
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima)

        cur.execute("""
            INSERT INTO sofi_ivr (
                message_id, chat_id, channel_identifier, user_name,
                campaign_id, user_identifier, last_user_input, fecha,
                opcion, doi, nombres, apellidos, direccion, distrito,
                contacto, url
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            message_id, chat_id, channel_identifier, user_name,
            campaign_id, user_identifier, last_user_input, fecha,
            opcion, doi, nombres, apellidos, direccion, distrito,
            contacto, url_publica
        ))

        # ── 4. UPDATE en sofi_bitacora ────────────────────────────────────────
        cur.execute("""
            UPDATE sofi_bitacora
            SET username_agent = %s
            WHERE chat_id = %s
        """, (opcion, chat_id))

        conn.commit()
        cur.close()
        conn.close()

        # ── 5. Llama a Danaconnect si hay dirección y no es Consulta Asesor ──
        if direccion and opcion != "Consulta Asesor":
            import requests as req
            payload = json.dumps({
                "DOI":             doi,
                "NOMBRES":         nombres,
                "APELLIDOS":       apellidos,
                "EMAIL":           "mteruya@sanitasperu.com",
                "DIRECCION":       direccion,
                "DISTRITO":        distrito,
                "CONTACTO":        nro_contacto1,
                "URL":             url_publica,
                "USER_IDENTIFIER": user_identifier,
                "ID_CHAT":         chat_id
            })
            req.post(
                "https://appserv.danaconnect.com/api/1.0/rest/conversation/553927/start/data",
                data=payload,
                headers={
                    "Content-Type":  "application/json",
                    "Authorization": DANACONNECT_AUTH
                },
                timeout=10
            )

        return (json.dumps({"mensaje": "Linea insertada"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /lead-medicamentos
# ─────────────────────────────────────────────
def handle_lead_medicamentos(request):
    data = request.get_json(silent=True) or {}
    nombre  = data.get("nombre", "")
    celular = data.get("celular", "")
    if not nombre:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)
    try:
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima).strftime('%Y-%m-%d %H:%M:%S')
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            INSERT INTO sofi_leads_medicamentos (nombre, celular, fecha, enviado)
            VALUES (%s, %s, %s, 'N')
        """, (nombre, celular, fecha))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "Registro ingresado OK"}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /lead-sctr
# ─────────────────────────────────────────────
def handle_lead_sctr(request):
    data = request.get_json(silent=True) or {}
    nombre  = data.get("nombre", "")
    celular = data.get("celular", "")
    if not nombre:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)
    try:
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima).strftime('%Y-%m-%d %H:%M:%S')
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            INSERT INTO sofi_leads_sctr (nombre, celular, fecha, enviado)
            VALUES (%s, %s, %s, 'N')
        """, (nombre, celular, fecha))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "Registro ingresado OK"}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /lead-consultas
# ─────────────────────────────────────────────
def handle_lead_consultas(request):
    data = request.get_json(silent=True) or {}
    nombre  = data.get("nombre", "")
    celular = data.get("celular", "")
    if not nombre:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)
    try:
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima).strftime('%Y-%m-%d %H:%M:%S')
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            INSERT INTO sofi_leads_consultas (nombre, celular, fecha, enviado)
            VALUES (%s, %s, %s, 'N')
        """, (nombre, celular, fecha))
        conn.commit()
        cur.close()
        conn.close()
        return (json.dumps({"mensaje": "Registro ingresado OK"}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /busca-doi-potestativo
# ─────────────────────────────────────────────
def handle_busca_doi_potestativo(request):
    data = request.get_json(silent=True) or {}
    doi = data.get("doi", "")
    if not doi:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)
    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT * FROM sofi_potestativos WHERE nro_doi = %s", (doi,))
        col_names = [desc[0] for desc in cur.description]
        row = cur.fetchone()
        cur.close()
        conn.close()

        if not row:
            return (json.dumps({"mensaje": "No se encontró el DOI", "contador": 0}), 200, HEADERS)

        fila = dict(zip(col_names, row))

        tipo_doi       = str(fila.get("tipo_doi", "")).title()
        nro_doi        = str(fila.get("nro_doi", "")).title()
        sexo           = str(fila.get("sexo", "")).title()
        nombre         = str(fila.get("nombre", "")).title()
        plan           = str(fila.get("plan", "")).title()
        tarifa_actual  = str(fila.get("tarifa_actual", ""))
        tarifa_nueva   = str(fila.get("tarifa_nueva", ""))
        fec_nacimiento = str(fila.get("fec_nacimiento", ""))
        edad_2025      = fila.get("edad_2025", "")
        edad_2026      = fila.get("edad_2026", "")
        incremento     = fila.get("incremento", "")
        rango_2025     = fila.get("rango_2025", "")
        rango_2026     = fila.get("rango_2026", "")
        cambio         = fila.get("cambio", "")
        incre_nro      = fila.get("incre_nro", "")
        contador       = 1
        mensaje        = ""

        if cambio == "N":
            mensaje = (
                f"¡Hola {nombre}! 👋💙\n\n"
                f"Te explicamos el ajuste en tu tarifa que aplicará desde el 1ro de enero del 2026 "
                f"del plan {plan}, que pasará de S/ {tarifa_actual} a S/ {tarifa_nueva}. "
                f"(Descarga tu carta desde el menú)\n\n"
                f"📌 Importante\nVemos que tu edad no cambió de rango (sigues en el rango etario de "
                f"{rango_2026}). Sin embargo, hubo un incremento general del {incremento} en el plan.\n\n"
                f"📝 NOTA\nLa edad que se toma por año, es la que tiene el cliente el 1ro de enero del 2026.\n\n"
                f"💡 El ajuste en las tarifas responde a factores como el incremento en los costos "
                f"de los servicios de salud y la inflación médica."
            )

        if cambio == "S":
            mensaje = (
                f"¡Hola {nombre}! 👋💙\n\n"
                f"Te explicamos el ajuste en tu tarifa que aplicará desde el 1ro de enero del 2026 "
                f"del plan {plan}, que pasará de S/ {tarifa_actual} a S/ {tarifa_nueva}. "
                f"(Descarga tu carta desde el menú)\n\n"
                f"📌 Importante\nVemos que tu edad cambió de rango (estabas en el rango etario "
                f"{rango_2025} y cambiaste a {rango_2026}). Además, hubo un incremento general "
                f"del {incremento} en el plan.\n\n"
                f"📝 NOTA\nLa edad que se toma por año, es la que tiene el cliente el 1ro de enero del 2026.\n\n"
                f"💡 El ajuste en las tarifas responde a factores como el incremento en los costos "
                f"de los servicios de salud y la inflación médica.\n\n"
                f"En caso tengas alguna duda, puedes escribir a sanitasenlinea@sanitasperu.com "
                f"indicando tu DNI para que puedan revisar tu caso y darte mayor detalle. "
                f"¡Así te aseguras de que todo esté en orden! 😉"
            )

        return (json.dumps({
            "mensaje":        mensaje,
            "contador":       contador,
            "tipo_doi":       tipo_doi,
            "nro_doi":        nro_doi,
            "sexo":           sexo,
            "nombre":         nombre,
            "plan":           plan,
            "tarifa_actual":  tarifa_actual,
            "tarifa_nueva":   tarifa_nueva,
            "fec_nacimiento": fec_nacimiento,
            "edad_2025":      edad_2025,
            "edad_2026":      edad_2026,
            "incremento":     incremento,
            "rango_2025":     rango_2025,
            "rango_2026":     rango_2026,
            "cambio":         cambio,
            "incre_nro":      incre_nro,
        }, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /busca-doi-potestativo-carta
# ─────────────────────────────────────────────
def handle_busca_doi_potestativo_carta(request):
    data = request.get_json(silent=True) or {}
    doi = data.get("doi", "")
    if not doi:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)
    try:
        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT * FROM sofi_potesta_cartas WHERE nro_doi = %s", (doi,))
        col_names = [desc[0] for desc in cur.description]
        row = cur.fetchone()
        cur.close()
        conn.close()

        if not row:
            return (json.dumps({"mensaje": "No se encontró el DOI", "contador": 0}), 200, HEADERS)

        fila = dict(zip(col_names, row))

        tipo_doi       = str(fila.get("tipo_doi", "")).title()
        nro_doi        = str(fila.get("nro_doi", "")).title()
        sexo           = str(fila.get("sexo", "")).title()
        nombre         = str(fila.get("nombre", "")).title()
        plan           = str(fila.get("plan", "")).title()
        tarifa_actual  = str(fila.get("tarifa_actual", ""))
        tarifa_nueva   = str(fila.get("tarifa_nueva", ""))
        fec_nacimiento = str(fila.get("fec_nacimiento", ""))
        edad_2025      = fila.get("edad_2025", "")
        edad_2026      = fila.get("edad_2026", "")
        incremento     = fila.get("incremento", "")
        rango_2025     = fila.get("rango_2025", "")
        rango_2026     = fila.get("rango_2026", "")
        cambio         = fila.get("cambio", "")
        incre_nro      = fila.get("incre_nro", "")
        contrato       = fila.get("contrato", "")
        archivo        = fila.get("archivo", "")
        email          = fila.get("email", "")
        fec_envio      = str(fila.get("fec_envio", ""))
        estatus        = fila.get("estatus", "")
        contador       = 1
        mensaje        = ""

        return (json.dumps({
            "mensaje":        mensaje,
            "contador":       contador,
            "tipo_doi":       tipo_doi,
            "nro_doi":        nro_doi,
            "sexo":           sexo,
            "nombre":         nombre,
            "plan":           plan,
            "tarifa_actual":  tarifa_actual,
            "tarifa_nueva":   tarifa_nueva,
            "fec_nacimiento": fec_nacimiento,
            "edad_2025":      edad_2025,
            "edad_2026":      edad_2026,
            "incremento":     incremento,
            "rango_2025":     rango_2025,
            "rango_2026":     rango_2026,
            "cambio":         cambio,
            "incre_nro":      incre_nro,
            "contrato":       contrato,
            "archivo":        archivo,
            "email":          email,
            "fec_envio":      fec_envio,
            "estatus":        estatus,
        }, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# TAREA: /envia-sms-leads-sctr
# ─────────────────────────────────────────────
def sms_aviso(celular, descripcion):
    lima  = pytz.timezone("America/Lima")
    fecha = datetime.now(lima).strftime("%Y-%m-%d %H:%M:%S")
    conn  = get_connection()
    cur   = conn.cursor()
    cur.execute("""
        INSERT INTO sofi_sms_avisos (celular, descripcion, fecha)
        VALUES (%s, %s, %s)
    """, (celular, descripcion, fecha))
    conn.commit()
    cur.close()
    conn.close()


def envia_sms_lead_sctr(numero, nombre, id_lead):
    nombre_limpio = re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre, flags=re.UNICODE
    )
    mensaje = f"Sanitas SCTR: Hola {nombre_limpio}, ya estamos listos para ayudarte con tu consulta SCTR. Haz clic aqui"
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(
            WHATSAPP_URL,
            json=payload, headers=headers, timeout=30
        )
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")
        return

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("UPDATE sofi_leads_sctr SET enviado='S' WHERE id=%s", (id_lead,))
    conn.commit()
    cur.close()
    conn.close()

    sms_aviso(numero, "lead sctr")


def handle_envia_sms_leads_sctr(request):
    try:
        lima       = pytz.timezone("America/Lima")
        ahora      = datetime.now(lima)
        dia_semana = ahora.isoweekday()  # 1=Lunes ... 7=Domingo

        if dia_semana < 1 or dia_semana > 6:
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT id, nombre, celular FROM sofi_leads_sctr WHERE enviado='N'")
        leads = cur.fetchall()
        cur.close()
        conn.close()

        if not leads:
            return (json.dumps({"mensaje": "No hay leads pendientes"}), 200, HEADERS)

        for id_lead, nombre, celular in leads:
            numero = f"+{celular}"
            envia_sms_lead_sctr(numero, nombre, id_lead)

        return (json.dumps({"mensaje": f"Enviado: {len(leads)} leads procesados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# TAREA: /envia-sms-leads-medicamentos
# ─────────────────────────────────────────────
def envia_sms_lead_medicamentos(numero, nombre, id_lead):
    nombre_limpio = re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre, flags=re.UNICODE
    )
    mensaje = f"Sanitas Medicamentos: Hola {nombre_limpio}, ya estamos listos para ayudarte con tus medicamentos. Haz clic aqui"
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(
            DANACONNECT_URL,
            json=payload, headers=headers, timeout=30
        )
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")
        return

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("UPDATE sofi_leads_medicamentos SET enviado='S' WHERE id=%s", (id_lead,))
    conn.commit()
    cur.close()
    conn.close()

    sms_aviso(numero, "lead medicamentos")


def handle_envia_sms_leads_medicamentos(request):
    try:
        lima       = pytz.timezone("America/Lima")
        ahora      = datetime.now(lima)
        dia_semana = ahora.isoweekday()  # 1=Lunes ... 7=Domingo

        if dia_semana < 1 or dia_semana > 6:
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT id, nombre, celular FROM sofi_leads_medicamentos WHERE enviado='N'")
        leads = cur.fetchall()
        cur.close()
        conn.close()

        if not leads:
            return (json.dumps({"mensaje": "No hay leads pendientes"}), 200, HEADERS)

        for id_lead, nombre, celular in leads:
            numero = f"+{celular}"
            envia_sms_lead_medicamentos(numero, nombre, id_lead)

        return (json.dumps({"mensaje": f"Enviado: {len(leads)} leads procesados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# TAREA: /envia-sms-leads-consultas
# ─────────────────────────────────────────────
def envia_sms_lead_consultas(numero, nombre, id_lead):
    nombre_limpio = re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre, flags=re.UNICODE
    )
    mensaje = f"Sanitas Consultas: Hola {nombre_limpio}, ya estamos listos para ayudarte con tus consultas. Haz clic aqui"
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(
            DANACONNECT_URL,
            json=payload, headers=headers, timeout=30
        )
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")
        return

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("UPDATE sofi_leads_consultas SET enviado='S' WHERE id=%s", (id_lead,))
    conn.commit()
    cur.close()
    conn.close()

    sms_aviso(numero, "lead consultas")


def handle_envia_sms_leads_consultas(request):
    try:
        lima       = pytz.timezone("America/Lima")
        ahora      = datetime.now(lima)
        dia_semana = ahora.isoweekday()  # 1=Lunes ... 7=Domingo

        if dia_semana < 1 or dia_semana > 6:
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT id, nombre, celular FROM sofi_leads_consultas WHERE enviado='N'")
        leads = cur.fetchall()
        cur.close()
        conn.close()

        if not leads:
            return (json.dumps({"mensaje": "No hay leads pendientes"}), 200, HEADERS)

        for id_lead, nombre, celular in leads:
            numero = f"+{celular}"
            envia_sms_lead_consultas(numero, nombre, id_lead)

        return (json.dumps({"mensaje": f"Enviado: {len(leads)} leads procesados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# TAREA: /envia-sctr-vencidos
# ─────────────────────────────────────────────
def envia_sms_sctr_vencidos(numero, ruc):
    numero  = f"+{numero}"
    mensaje = f"Sanitas SCTR: (RUC {ruc}) tu SCTR vence hoy. Evita las multas. Renueva aqui"
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(
            DANACONNECT_URL,
            json=payload, headers=headers, timeout=30
        )
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")
        return

    sms_aviso(numero, "sctr vencidos")

# ─────────────────────────────────────────────
# TAREA: /recordatorio-sctr
# Helper: Danaconnect
# ─────────────────────────────────────────────

def enviar_danaconnect(numero, mensaje):
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")


# ─────────────────────────────────────────────
# Helper: Sincroniza sctr_empresa -> sctr_ruc_celular
# ─────────────────────────────────────────────
def sincronizar_ruc_celular(cur):
    cur.execute("""
        INSERT INTO sctr_ruc_celular (ruc, celular)
        SELECT t1.ruc, t1.celular
        FROM sctr_empresa t1
        LEFT JOIN sctr_ruc_celular t2
            ON t1.ruc = t2.ruc AND t1.celular = t2.celular
        WHERE t2.ruc IS NULL
    """)


# ─────────────────────────────────────────────
# Helper: Obtiene RUCs vencidos en una fecha
# ─────────────────────────────────────────────
def obtener_rucs_vencidos(cur, fecha):
    cur.execute("""
        SELECT DISTINCT nro_doc
        FROM sctr_emisiones
        WHERE fin_vigencia = %s
          AND fin_vigencia IS NOT NULL
          AND nro_doc IS NOT NULL
    """, (fecha,))
    return [row[0] for row in cur.fetchall()]


# ─────────────────────────────────────────────
# Helper: Obtiene celulares filtrando agentes
# ─────────────────────────────────────────────
def obtener_celulares_sin_agente(cur, rucs, fechas_filtro):
    cur.execute("""
        SELECT src.ruc, src.celular
        FROM sctr_ruc_celular src
        WHERE src.ruc = ANY(%s)
          AND src.celular NOT IN (
              SELECT DISTINCT user_identifier
              FROM sofi_ivr
              WHERE DATE(fecha) = ANY(%s)
                AND opcion = 'SCTR hablar con agente'
          )
    """, (rucs, fechas_filtro))
    return cur.fetchall()


# ─────────────────────────────────────────────
# Helper generico: envia recordatorio SCTR
# ─────────────────────────────────────────────
def envia_recordatorio_sctr(fecha_vencimiento, fechas_filtro, mensaje_template, aviso_descripcion):
    conn = get_connection()
    cur  = conn.cursor()

    sincronizar_ruc_celular(cur)
    conn.commit()

    rucs = obtener_rucs_vencidos(cur, fecha_vencimiento)
    if not rucs:
        cur.close()
        conn.close()
        return 0

    resultados = obtener_celulares_sin_agente(cur, rucs, fechas_filtro)
    cur.close()
    conn.close()

    for ruc, celular in resultados:
        numero  = f"+{celular}"
        mensaje = mensaje_template.format(ruc=ruc)
        enviar_danaconnect(numero, mensaje)
        sms_aviso(numero, aviso_descripcion)

    return len(resultados)


# ─────────────────────────────────────────────
# Recordatorio 1: vence HOY (11:25 - 11:35)
# ─────────────────────────────────────────────
def envia_sctr_vencidos_recordatorio1(hoy):
    return envia_recordatorio_sctr(
        fecha_vencimiento = hoy,
        fechas_filtro     = [hoy],
        mensaje_template  = "Sanitas SCTR: (RUC {ruc}) que no se te pase tu compra. Renueva aqui",
        aviso_descripcion = "sctr vencidos recordatorio 1"
    )


# ─────────────────────────────────────────────
# Recordatorio 2: vence HOY (14:55 - 15:05, lun-vie)
# ─────────────────────────────────────────────
def envia_sctr_vencidos_recordatorio2(hoy):
    return envia_recordatorio_sctr(
        fecha_vencimiento = hoy,
        fechas_filtro     = [hoy],
        mensaje_template  = "Sanitas SCTR: El tiempo corre y aun no renuevas. (RUC {ruc}). haz clic",
        aviso_descripcion = "sctr vencidos recordatorio 2"
    )


# ─────────────────────────────────────────────
# Recordatorio 3: venció AYER (08:25 - 08:35)
# ─────────────────────────────────────────────
def envia_sctr_vencidos_recordatorio3(hoy, ayer):
    return envia_recordatorio_sctr(
        fecha_vencimiento = ayer,
        fechas_filtro     = [hoy, ayer],
        mensaje_template  = "Sanitas SCTR: Tu SCTR vencio y aun no renuevas. (RUC {ruc}). compra aqui",
        aviso_descripcion = "sctr vencidos recordatorio -1"
    )

# ─────────────────────────────────────────────
# TAREA: /busca-ultimo-registro
# ─────────────────────────────────────────────
def envia_sms_alerta(numero, minutos, linea):
    mensaje = f"MONITOREO SOFI: Hace {minutos} minutos fue la última conversación. {linea}"
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  "bit.ly/sofisoporte"
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar alerta a {numero}: {e}")


def handle_busca_ultimo_registro(request):
    try:
        lima       = pytz.timezone("America/Lima")
        ahora      = datetime.now(lima)
        dia_semana = ahora.isoweekday()
        hora_actual = ahora.strftime("%H:%M")

        if not (1 <= dia_semana <= 6 and "08:15" <= hora_actual <= "20:00"):
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            SELECT fecha, user_name, user_identifier, opcion
            FROM sofi_ivr
            WHERE opcion IN (
                'Medicamentos-Coord de receta',
                'SCTR hablar con agente',
                'Medicamentos Estado de Pedido'
            )
            ORDER BY fecha DESC
            LIMIT 1
        """)
        col_names = [desc[0] for desc in cur.description]
        row = cur.fetchone()
        cur.close()
        conn.close()

        if not row:
            return (json.dumps({"mensaje": "Sin registros"}), 200, HEADERS)

        fila    = dict(zip(col_names, row))
        fecha_bd = fila["fecha"]
        nombre  = fila["user_name"]
        celular = fila["user_identifier"]
        opcion  = fila["opcion"]

        linea   = f"{fecha_bd} {celular} {nombre} {opcion}"

        # Calcular diferencia en minutos
        if fecha_bd.tzinfo is None:
            fecha_bd = lima.localize(fecha_bd)
        minutos = int((ahora - fecha_bd).total_seconds() / 60)

        if minutos > 20:
            for numero in ["+51997855645", "+51975448483", "+51971116968", "+51987680154"]:
                envia_sms_alerta(numero, minutos, linea)

        return (json.dumps({"mensaje": minutos}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /consulta-ruc
# ─────────────────────────────────────────────
APISPERU_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJlbWFpbCI6Im10ZXJ1eWFuQGdtYWlsLmNvbSJ9.bloeaPJiZxGFIaWj2PXMMktf8BBdiOO340Ki6YCYHkg"

def handle_consulta_ruc(request):
    data   = request.get_json(silent=True) or {}
    ruc    = data.get("ruc", "")
    celular_input = data.get("celular", "")

    if not ruc:
        return (json.dumps({"error": "El campo ruc es obligatorio"}), 400, HEADERS)

    try:
        lima  = pytz.timezone("America/Lima")
        ahora = datetime.now(lima).strftime("%Y-%m-%d %H:%M:%S")

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("SELECT * FROM sctr_empresa WHERE ruc = %s", (ruc,))
        col_names = [desc[0] for desc in cur.description]
        row = cur.fetchone()

        if not row:
            # RUC no existe: consultar API externa
            try:
                resp = requests.get(
                    f"https://dniruc.apisperu.com/api/v1/ruc/{ruc}?token={APISPERU_TOKEN}",
                    timeout=10
                )
                obj = resp.json()
            except Exception as e:
                cur.close()
                conn.close()
                return (json.dumps({"error": f"Error consultando API RUC: {str(e)}"}), 500, HEADERS)

            razon_social = obj.get("razonSocial", "")
            direccion    = obj.get("direccion", "")
            distrito     = obj.get("distrito", "")
            provincia    = obj.get("provincia", "")
            departamento = obj.get("departamento", "")
            celular      = celular_input
            correo       = ""
            actividad    = ""
            fecha_creacion   = ahora
            fecha_ult_acceso = ahora

            if razon_social:
                cur.execute("""
                    INSERT INTO sctr_empresa
                        (ruc, razon_social, celular, direccion, departamento, provincia, distrito, fecha_creacion, fecha_ult_acceso)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (ruc, razon_social, celular_input, direccion, departamento, provincia, distrito, ahora, ahora))
                conn.commit()

        else:
            # RUC existe: actualizar fecha y celular si cambió
            fila         = dict(zip(col_names, row))
            razon_social = fila.get("razon_social", "")
            celular      = fila.get("celular", "")
            correo       = fila.get("correo", "")
            direccion    = fila.get("direccion", "")
            departamento = fila.get("departamento", "")
            provincia    = fila.get("provincia", "")
            distrito     = fila.get("distrito", "")
            fecha_creacion   = str(fila.get("fecha_creacion", ""))
            actividad        = fila.get("actividad", "")
            fecha_ult_acceso = str(fila.get("fecha_ult_acceso", ""))

            if celular != celular_input:
                cur.execute("""
                    UPDATE sctr_empresa
                    SET celular = %s, fecha_ult_acceso = %s
                    WHERE ruc = %s
                """, (celular_input, ahora, ruc))
                celular = celular_input
            else:
                cur.execute("""
                    UPDATE sctr_empresa
                    SET fecha_ult_acceso = %s
                    WHERE ruc = %s
                """, (ahora, ruc))

            conn.commit()

        cur.close()
        conn.close()

        tiene_actividad = "S" if actividad else "N"

        return (json.dumps({
            "ruc":              ruc,
            "razon_social":     razon_social,
            "celular":          celular_input,
            "correo":           correo,
            "direccion":        direccion,
            "departamento":     departamento,
            "provincia":        provincia,
            "distrito":         distrito,
            "fecha_creacion":   fecha_creacion,
            "actividad":        actividad,
            "fecha_ult_acceso": fecha_ult_acceso,
            "tiene_actividad":  tiene_actividad,
        }, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /similar-texto
# ─────────────────────────────────────────────
def normalizar_texto(texto):
    texto = texto.lower()
    trans = str.maketrans("áéíóúäëïöüñ", "aeiouaeioun")
    texto = texto.translate(trans)
    texto = re.sub(r'[^a-z0-9 ]', '', texto)
    return texto.strip()


def similitud_avanzada(input_texto, texto_bd):
    a = normalizar_texto(input_texto)
    b = normalizar_texto(texto_bd)

    if not a or not b:
        return 0.0

    # 1) similar_text (equivalente con SequenceMatcher)
    porcentaje_similar = SequenceMatcher(None, a, b).ratio() * 100

    # 2) Levenshtein
    lev = levenshtein(a, b)
    max_long = max(len(a), len(b))
    porcentaje_levenshtein = 100 - (lev / max_long * 100) if max_long > 0 else 100

    # 3) Metaphone
    met_a = jellyfish.metaphone(a)
    met_b = jellyfish.metaphone(b)
    porcentaje_metaphone = SequenceMatcher(None, met_a, met_b).ratio() * 100

    # Peso combinado
    score = (
        (porcentaje_similar      * 0.50) +
        (porcentaje_levenshtein  * 0.30) +
        (porcentaje_metaphone    * 0.20)
    )

    # Bonus si la palabra aparece dentro del texto
    if a in b:
        score += 20

    return round(score, 2)


def levenshtein(s1, s2):
    if len(s1) < len(s2):
        return levenshtein(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev = list(range(len(s2) + 1))
    for i, c1 in enumerate(s1):
        curr = [i + 1]
        for j, c2 in enumerate(s2):
            curr.append(min(prev[j + 1] + 1, curr[j] + 1, prev[j] + (c1 != c2)))
        prev = curr
    return prev[-1]


def handle_busca_actividad(request):
    data   = request.get_json(silent=True) or {}
    tipo   = data.get("tipo", "")
    palabra = data.get("palabra", "")
    ruc    = data.get("ruc", "")
    selec  = data.get("selec", "")

    # Limpiar espacios y saltos de línea
    palabra = re.sub(r'\s+', ' ', palabra.replace('\r', ' ').replace('\n', ' ')).strip()

    if not palabra:
        return (json.dumps({"data_1": "", "mensaje": "Error dato invalido"}), 400, HEADERS)

    try:
        tabla = "sctr_ciiu_sunat" if tipo == "sunat" else "sctr_actividad"

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute(f"SELECT id, descripcion FROM {tabla}")
        filas = cur.fetchall()

        resultados = []
        for id_row, descripcion in filas:
            descripcion_bd = descripcion[16:] if len(descripcion) > 16 else descripcion
            score = similitud_avanzada(palabra, descripcion_bd)
            resultados.append({
                "id":          id_row,
                "descripcion": descripcion,
                "similaridad": score
            })

        # Ordenar por similitud descendente
        resultados.sort(key=lambda x: x["similaridad"], reverse=True)
        top7 = resultados[:7]

        mensaje = "Lista OK"
        imprime = ""

        for i, item in enumerate(top7, start=1):
            temp = item["descripcion"]

            if selec != "":
                selec_int = int(selec) if str(selec).isdigit() else 0

                if selec_int == i:
                    cur.execute(
                        "UPDATE sctr_empresa SET actividad = %s WHERE ruc = %s",
                        (temp, ruc)
                    )
                    conn.commit()
                    mensaje = f"{ruc} ha sido actualizado"

                if selec_int == 8:
                    cur.execute(
                        "UPDATE sctr_empresa SET actividad = %s WHERE ruc = %s",
                        ("096090 - 096090 OTRAS ACTIVIDADES DE SERVICIOS PERSONALES N.C.P.", ruc)
                    )
                    conn.commit()
                    mensaje = f"{ruc} ha sido actualizado"

            imprime += f"{i}. {temp}\n\n"

        cur.close()
        conn.close()

        return (json.dumps({"data_1": imprime, "mensaje": mensaje}, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /busca-carta-rechazada
# ─────────────────────────────────────────────
def safe_read_csv(url, tries=3, delay=1):
    import time
    while tries > 0:
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                response.encoding = 'utf-8'
                return response.text.splitlines()
        except Exception:
            pass
        time.sleep(delay)
        tries -= 1
    return None


def validar_consulta(celular, documento_intentado):
    if not celular:
        return {"status": True, "msg": "Consulta especial permitida"}

    lima  = pytz.timezone("America/Lima")
    ahora = datetime.now(lima)
    ahora_str = ahora.strftime("%Y-%m-%d %H:%M:%S")

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("SELECT * FROM sofi_control_consultas WHERE celular = %s", (celular,))
    col_names = [desc[0] for desc in cur.description]
    row = cur.fetchone()

    if not row:
        cur.execute("""
            INSERT INTO sofi_control_consultas (celular, documento, fecha_creacion, intentos, bloqueado, fecha_ult_acceso)
            VALUES (%s, %s, %s, 1, 0, %s)
        """, (celular, documento_intentado, ahora_str, ahora_str))
        conn.commit()
        cur.close()
        conn.close()
        return {"status": True, "msg": "Consulta permitida por primera vez"}

    registro = dict(zip(col_names, row))

    # Verificar si está bloqueado
    if registro["bloqueado"] == 1:
        fecha_bloqueo    = registro["fecha_ult_acceso"]
        if hasattr(fecha_bloqueo, 'tzinfo') and fecha_bloqueo.tzinfo is None:
            fecha_bloqueo = lima.localize(fecha_bloqueo)
        fecha_desbloqueo = fecha_bloqueo + timedelta(days=7)
        if ahora < fecha_desbloqueo:
            cur.close()
            conn.close()
            return {
                "status": False,
                "msg": f"Celular bloqueado temporalmente. Intenta después de: {fecha_desbloqueo.strftime('%Y-%m-%d %H:%M:%S')}"
            }
        else:
            cur.execute("UPDATE sofi_control_consultas SET bloqueado = 0, intentos = 0 WHERE celular = %s", (celular,))
            conn.commit()
            registro["bloqueado"] = 0
            registro["intentos"]  = 0

    # Validar si es el mismo día
    fecha_ult_acceso = registro["fecha_ult_acceso"]
    if hasattr(fecha_ult_acceso, 'date'):
        fecha_ult_date = fecha_ult_acceso.date()
    else:
        fecha_ult_date = datetime.strptime(str(fecha_ult_acceso), "%Y-%m-%d %H:%M:%S").date()

    if fecha_ult_date == ahora.date():
        if registro["documento"] != documento_intentado:
            intentos = registro["intentos"] + 1
            if intentos > 3:
                cur.execute("""
                    UPDATE sofi_control_consultas
                    SET bloqueado = 1, fecha_ult_acceso = %s
                    WHERE celular = %s
                """, (ahora_str, celular))
                conn.commit()
                cur.close()
                conn.close()
                return {"status": False, "msg": "Has excedido el número de intentos. Celular bloqueado por 7 días."}
            cur.execute("""
                UPDATE sofi_control_consultas
                SET intentos = %s, fecha_ult_acceso = %s
                WHERE celular = %s
            """, (intentos, ahora_str, celular))
            conn.commit()
            cur.close()
            conn.close()
            return {"status": False, "msg": "Solo puedes consultar un documento por día."}

        cur.execute("UPDATE sofi_control_consultas SET fecha_ult_acceso = %s WHERE celular = %s", (ahora_str, celular))
        conn.commit()
        cur.close()
        conn.close()
        return {"status": True, "msg": "Consulta permitida (mismo documento del día)"}

    # Nuevo día: resetear intentos
    cur.execute("""
        UPDATE sofi_control_consultas
        SET documento = %s, intentos = 1, fecha_ult_acceso = %s
        WHERE celular = %s
    """, (documento_intentado, ahora_str, celular))
    conn.commit()
    cur.close()
    conn.close()
    return {"status": True, "msg": "Consulta permitida (nuevo día)"}


def handle_busca_carta_rechazada(request):
    data    = request.get_json(silent=True) or {}
    doi     = data.get("doi", "")
    celular = data.get("celular", "")

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    try:
        resultado_validacion = validar_consulta(celular, doi)
        status = resultado_validacion["status"]
        msg    = resultado_validacion["msg"]

        id_           = ""; buscador = "vacio"; sol = ""; tipo_carta = ""
        afiliado      = ""; prog = ""; ipress = ""; plan = ""; cobertura = ""
        fecha_prog    = ""; estado = ""; monto = ""; usuario = ""; fecha_registro = ""
        tipo_proceso  = ""; tipo_solicitud = ""; dni = ""; estado_f_i = ""
        observacion   = ""; tipo_observ = ""; autonomia = ""; nro_celular = ""
        medico_evaluador = ""; marca = ""; ia_obs = str(status)

        if status:
            url   = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=1727191727&single=true&output=csv&t={int(datetime.now().timestamp())}"
            lines = safe_read_csv(url)

            if lines:
                import csv
                last5000 = lines[max(0, len(lines) - 5000):]
                reader   = csv.reader(last5000)

                for row in reader:
                    if not row:
                        continue
                    columna_documento = row[4].strip() if len(row) > 4 else ""
                    if doi in columna_documento:
                        id_              = row[0].strip()  if len(row) > 0  else ""
                        buscador         = row[1].strip()  if len(row) > 1  else ""
                        sol              = row[2].strip()  if len(row) > 2  else ""
                        tipo_carta       = row[3].strip()  if len(row) > 3  else ""
                        afiliado         = row[4].strip()  if len(row) > 4  else ""
                        prog             = row[5].strip()  if len(row) > 5  else ""
                        ipress           = row[6].strip()  if len(row) > 6  else ""
                        plan             = row[7].strip()  if len(row) > 7  else ""
                        cobertura        = row[8].strip()  if len(row) > 8  else ""
                        fecha_prog       = row[9].strip()  if len(row) > 9  else ""
                        estado           = row[10].strip() if len(row) > 10 else ""
                        monto            = row[11].strip() if len(row) > 11 else ""
                        usuario          = row[12].strip() if len(row) > 12 else ""
                        fecha_registro   = row[13].strip() if len(row) > 13 else ""
                        tipo_proceso     = row[14].strip() if len(row) > 14 else ""
                        tipo_solicitud   = row[15].strip() if len(row) > 15 else ""
                        dni              = row[16].strip() if len(row) > 16 else ""
                        estado_intermedio= row[17].strip() if len(row) > 17 else ""
                        estado_final     = row[18].strip() if len(row) > 18 else ""
                        observacion      = row[19].strip() if len(row) > 19 else ""
                        tipo_observ      = row[20].strip() if len(row) > 20 else ""
                        autonomia        = row[21].strip() if len(row) > 21 else ""
                        nro_celular      = row[25].strip() if len(row) > 25 else ""
                        medico_evaluador = row[26].strip() if len(row) > 26 else ""
                        marca            = row[33].strip() if len(row) > 33 else ""
                        ia_obs           = row[35].strip() if len(row) > 35 else ""

                        # Determinar estado final
                        if not estado_final and not estado_intermedio:
                            estado_f_i = "EN PROCESO"
                        elif estado_final:
                            estado_f_i = estado_final
                        else:
                            estado_f_i = estado_intermedio

                        if not observacion:
                            observacion = "Este campo por el momento esta vacío"

        if buscador == "vacio":
            observacion = "No existe información."
            if not status:
                observacion = msg

        return (json.dumps({
            "id":               id_,
            "buscador":         buscador,
            "sol":              sol,
            "tipo_carta":       tipo_carta,
            "afiliado":         afiliado,
            "prog":             prog,
            "ipress":           ipress,
            "plan":             plan,
            "cobertura":        cobertura,
            "fecha_prog":       fecha_prog,
            "estado":           estado,
            "monto":            monto,
            "usuario":          usuario,
            "fecha_registro":   fecha_registro,
            "tipo_proceso":     tipo_proceso,
            "tipo_solicitud":   tipo_solicitud,
            "dni":              dni,
            "estado_final":     estado_f_i,
            "observacion":      observacion,
            "tipo_observ":      tipo_observ,
            "autonomia":        autonomia,
            "nro_celular":      nro_celular,
            "medico_evaluador": medico_evaluador,
            "marca":            marca,
            "ia_obs":           ia_obs,
        }, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# API: /update-cg-rechazados
# ─────────────────────────────────────────────
def handle_update_cg_rechazados(request):
    data    = request.get_json(silent=True) or {}
    doi     = data.get("doi", "")
    campo   = data.get("campo", "")

    if not doi:
        return (json.dumps({"mensaje": "Error dato invalido"}), 400, HEADERS)

    try:
        lima  = pytz.timezone("America/Lima")
        fecha = datetime.now(lima).strftime("%Y-%m-%d %H:%M:%S")

        query  = ""
        mensaje = ""

        if campo == "ingreso_wsp":
            query   = "UPDATE sofi_cg_rechazados SET ingreso_wsp='S', fecha_ingreso_wsp=%s WHERE nombre LIKE %s"
            mensaje = f"Se actualizo el campo ingreso de whatsapp en: {doi}"

        if campo == "hablar_asesor":
            query   = "UPDATE sofi_cg_rechazados SET hablar_asesor='S', fecha_ingreso_wsp=%s WHERE nombre LIKE %s"
            mensaje = f"Se actualizo el campo hablar con asesor en: {doi}"

        if query:
            conn = get_connection()
            cur  = conn.cursor()
            cur.execute(query, (fecha, f"%{doi}"))
            conn.commit()
            cur.close()
            conn.close()

        return (json.dumps({"mensaje": mensaje}, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# HELPER GENERICO: envia SMS leads
# Usado por: /envia-sms-leads-sctr,
#            /envia-sms-leads-medicamentos,
#            /envia-sms-leads-consultas
# ─────────────────────────────────────────────
def envia_sms_leads_generico(tabla, mensaje_template, aviso_descripcion):
    lima       = pytz.timezone("America/Lima")
    ahora      = datetime.now(lima)
    dia_semana = ahora.isoweekday()  # 1=Lunes ... 7=Domingo

    if not (1 <= dia_semana <= 6):
        return "Fuera del horario"

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute(f"SELECT id, nombre, celular FROM {tabla} WHERE enviado='N'")
    leads = cur.fetchall()
    cur.close()
    conn.close()

    if not leads:
        return "No hay leads pendientes"

    for id_lead, nombre, celular in leads:
        numero = f"+{celular}"
        nombre_limpio = re.sub(
            r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
            r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
            '', nombre, flags=re.UNICODE
        )
        mensaje = mensaje_template.format(nombre=nombre_limpio)
        payload = {
            "NUM":  f" {numero}",
            "TEXT": mensaje,
            "URL":  WHATSAPP_URL
        }
        headers = {
            "Content-Type":  "application/json",
            "Authorization": DANACONNECT_AUTH
        }
        try:
            response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
            response.raise_for_status()
        except Exception as e:
            print(f"Error al enviar SMS a {numero}: {e}")
            continue

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute(f"UPDATE {tabla} SET enviado='S' WHERE id=%s", (id_lead,))
        conn.commit()
        cur.close()
        conn.close()

        sms_aviso(numero, aviso_descripcion)

    return f"Enviado: {len(leads)} leads procesados"


# ─────────────────────────────────────────────
# API: /envia-sms-leads-sctr
# ─────────────────────────────────────────────
def handle_envia_sms_leads_sctr(request):
    try:
        mensaje = envia_sms_leads_generico(
            tabla             = "sofi_leads_sctr",
            mensaje_template  = "Sanitas SCTR: Hola {nombre}, ya estamos listos para ayudarte con tu consulta SCTR. Haz clic aqui",
            aviso_descripcion = "lead sctr"
        )
        return (json.dumps({"mensaje": mensaje}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# API: /envia-sms-leads-medicamentos
# ─────────────────────────────────────────────
def handle_envia_sms_leads_medicamentos(request):
    try:
        mensaje = envia_sms_leads_generico(
            tabla             = "sofi_leads_medicamentos",
            mensaje_template  = "Sanitas Medicamentos: Hola {nombre}, ya estamos listos para ayudarte con tus medicamentos. Haz clic aqui",
            aviso_descripcion = "lead medicamentos"
        )
        return (json.dumps({"mensaje": mensaje}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# API: /envia-sms-leads-consultas
# ─────────────────────────────────────────────
def handle_envia_sms_leads_consultas(request):
    try:
        mensaje = envia_sms_leads_generico(
            tabla             = "sofi_leads_consultas",
            mensaje_template  = "Sanitas Consultas: Hola {nombre}, ya estamos listos para ayudarte con tus consultas. Haz clic aqui",
            aviso_descripcion = "lead consultas"
        )
        return (json.dumps({"mensaje": mensaje}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# TAREA: /recordatorio-sctr
# ─────────────────────────────────────────────
def lee_excel_rechazados():
    import csv
    import time

    lima    = pytz.timezone("America/Lima")
    ahora   = datetime.now(lima)
    hoy_str = ahora.strftime("%Y-%m-%d %H:%M:%S")

    # Si es lunes, tomar anteayer; si no, tomar ayer
    dias_atras = 2 if ahora.isoweekday() == 1 else 1
    ayer = (ahora - timedelta(days=dias_atras)).strftime("%d/%m/%Y")

    url = f"https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=577147525&single=true&output=csv&t={int(time.time())}"

    lines = safe_read_csv(url)
    if not lines:
        print("No se pudo leer el archivo, posiblemente está en actualización.")
        return "Error lectura CSV"

    last1000 = lines[max(0, len(lines) - 1000):]
    reader   = csv.reader(last1000)

    contador             = 0
    contador_con_celular = 0
    rows_to_insert       = []

    for row in reader:
        if not row:
            continue

        fecha_linea = row[23].strip() if len(row) > 23 else ""
        row_2  = row[2].strip()  if len(row) > 2  else ""
        row_4  = row[4].strip()  if len(row) > 4  else ""
        row_21 = row[21].strip() if len(row) > 21 else ""
        row_23 = row[23].strip() if len(row) > 23 else ""

        if ayer in fecha_linea:
            # Convertir fecha DD/MM/YYYY → YYYY-MM-DD
            try:
                fecha_rechazo = datetime.strptime(row_23, "%d/%m/%Y").strftime("%Y-%m-%d")
            except ValueError:
                fecha_rechazo = row_23

            rows_to_insert.append((
                row_2,          # sol
                fecha_rechazo,  # fecha_rechazo
                row_21,         # celular
                row_4,          # nombre
                hoy_str,        # fecha
                "N",            # ingreso_wsp
                None,           # fecha_ingreso_wsp
                "N"             # hablar_asesor
            ))

            # Enviar SMS si tiene celular válido
            if len(row_21) == 9 and row_21[0] == "9" and row_21 != "999999999":
                nombre = row_4[:64] if len(row_4) > 64 else row_4
                envia_sms_cartas_rechazadas(row_21, nombre)
                contador_con_celular += 1

            contador += 1

    # Insertar registros en base de datos
    if rows_to_insert:
        conn = get_connection()
        cur  = conn.cursor()
        cur.executemany("""
            INSERT INTO sofi_cg_rechazados
                (sol, fecha_rechazo, celular, nombre, fecha, ingreso_wsp, fecha_ingreso_wsp, hablar_asesor)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, rows_to_insert)
        conn.commit()
        cur.close()
        conn.close()

    return f"Total: {contador} | Tienen celular: {contador_con_celular}"


def envia_sms_sctr_vencidos(numero, ruc):
    numero  = f"+{numero}"
    mensaje = f"Sanitas SCTR: (RUC {ruc}) tu SCTR vence hoy. Evita las multas. Renueva aqui"
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "sctr vencidos")


def handle_envia_sctr_vencidos(request):
    try:
        lima = pytz.timezone("America/Lima")
        hoy  = datetime.now(lima).strftime("%Y-%m-%d")

        conn = get_connection()
        cur  = conn.cursor()

        sincronizar_ruc_celular(cur)
        conn.commit()

        rucs = obtener_rucs_vencidos(cur, hoy)
        if not rucs:
            cur.close()
            conn.close()
            return (json.dumps({"mensaje": "No hay RUC con emisión vencida"}), 200, HEADERS)

        cur.execute("""
            SELECT ruc, celular FROM sctr_ruc_celular
            WHERE ruc = ANY(%s)
        """, (rucs,))
        resultados = cur.fetchall()
        cur.close()
        conn.close()

        for ruc, celular in resultados:
            envia_sms_sctr_vencidos(celular, ruc)

        return (json.dumps({"mensaje": f"Todo OK: {len(resultados)} mensajes enviados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


def handle_recordatorio_sctr(request):
    try:
        lima        = pytz.timezone("America/Lima")
        ahora       = datetime.now(lima)
        dia_semana  = ahora.isoweekday()
        hora_actual = ahora.strftime("%H:%M")
        hoy         = ahora.strftime("%Y-%m-%d")
        ayer        = (ahora - timedelta(days=1)).strftime("%Y-%m-%d")

        ejecutados = []

        # Recordatorio 3: vencidos ayer (08:25 - 08:35, lun-sab)
        if 1 <= dia_semana <= 6 and "08:25" <= hora_actual <= "08:35":
            cuenta = envia_sctr_vencidos_recordatorio3(hoy, ayer)
            ejecutados.append(f"recordatorio3: {cuenta} enviados")

        # Recordatorio 1: vencen hoy (11:25 - 11:35, lun-sab)
        if 1 <= dia_semana <= 6 and "11:25" <= hora_actual <= "11:35":
            cuenta = envia_sctr_vencidos_recordatorio1(hoy)
            ejecutados.append(f"recordatorio1: {cuenta} enviados")

        # Recordatorio 2: vencen hoy (14:55 - 15:05, lun-vie)
        if 1 <= dia_semana <= 5 and "14:55" <= hora_actual <= "15:05":
            cuenta = envia_sctr_vencidos_recordatorio2(hoy)
            ejecutados.append(f"recordatorio2: {cuenta} enviados")

        # Lee excel rechazados (12:10 - 12:20, lun-sab)
        if 1 <= dia_semana <= 6 and "12:10" <= hora_actual <= "12:20":
            resultado = lee_excel_rechazados()
            ejecutados.append(f"excel_rechazados: {resultado}")

        if not ejecutados:
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        return (json.dumps({"mensaje": ", ".join(ejecutados)}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)