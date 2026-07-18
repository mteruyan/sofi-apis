import json
import os
import time
import pytz
import psycopg2
from psycopg2.extras import execute_values
from datetime import datetime, timedelta
import functions_framework
import re
import requests
from difflib import SequenceMatcher
import jellyfish
import csv
import io
import logging
import openpyxl
import hashlib
import base64
import traceback
from functools import wraps


from flask import send_from_directory
from sctr_forms import handle_sctr_form, handle_sctr_submit, handle_sctr_crear_grupo
from sctr_whatsapp import handle_sctr_resumen_whatsapp, handle_sctr_excel_whatsapp, handle_sctr_borrar_empleado, handle_sctr_excel_web
from admision_forms import handle_admision_sctr_nuevo, handle_admision_sctr_continuidad, handle_admision_eps_nuevo
from agente_consulta import (handle_consulta_agente, handle_agente_kb_upsert, handle_agente_kb_listar, handle_agente_kb_form_web, handle_agente_kb_perfil)

# ─────────────────────────────────────────────
# VARIABLES
# ─────────────────────────────────────────────
LIMA_TZ = pytz.timezone("America/Lima")
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

    # ── REDIRECT RAÍZ ──────────────────────────────
    if request.path == "/" or request.path == "":
        from flask import redirect
        return redirect("https://www.sanitasperu.com", code=301)
    # ───────────────────────────────────────────────

    if not request.path.startswith("/sctr/") and not request.path.startswith("/admision/") \
        and not request.path.startswith("/agente_kb/") and not request.path.startswith("/campania/") \
        and not request.path.startswith("/c/") and not request.path.startswith("/static/") \
        and not validar_api_key(request):
        return (json.dumps({"error": "API Key invalida"}), 401, HEADERS)

    path = request.path.rstrip("/")

    if path == "/bitacora":
        return handle_bitacora(request)
    elif path == "/update_bitacora":
        return handle_update_bitacora(request)
    elif path == "/whatsapp_busca_doi":
        return handle_whatsapp_busca_doi(request)
    elif path == "/consulta_agente":
        return handle_consulta_agente(request)
    elif path == "/agente_kb/upsert" and request.method == "POST":
        return handle_agente_kb_upsert(request)
    elif path == "/agente_kb/listar":
        return handle_agente_kb_listar(request)
    elif path == "/agente_kb/form" and request.method == "GET":
        return handle_agente_kb_form_web(request)
    elif path == "/agente_kb/perfil" and request.method == "GET":
        return handle_agente_kb_perfil(request)
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
    elif path == "/lead_pases":
        return handle_lead_pases(request)
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
    elif path == "/envia_sms_leads_pases":
        return handle_envia_sms_leads_pases(request)
    elif path == "/envia_sctr_vencidos":
        return handle_envia_sctr_vencidos(request)
    elif path == "/envia_sctr_vencidos_envio_wsp":
        return handle_envia_sctr_vencidos_envio_wsp(request)
    elif path == "/recordatorio_sctr":
        return handle_recordatorio_sctr(request)
    elif path == "/busca_ultimo_registro":
        return handle_busca_ultimo_registro(request)
    elif path == "/consulta_ruc":
        return handle_consulta_ruc(request)
    elif path == "/consulta_ruc2":
        return handle_consulta_ruc2(request)
    elif path == "/busca_actividad":
        return handle_busca_actividad(request)
    elif path == "/busca_carta_rechazada":
        return handle_busca_carta_rechazada(request)
    elif path == "/busca_carta_rechazada_tabla":
        return handle_busca_carta_rechazada_tabla(request)
    elif path == "/sync_cartas_rechazadas":
        return handle_sync_cartas_rechazadas(request)
    elif path == "/sync_retencion_sms":
        return handle_sync_retencion_sms(request)
    elif path == "/diagnostico_csv":
        return handle_diagnostico_csv_v2(request)
    elif path == "/diagnostico_headers":
        return handle_diagnostico_headers(request)
    elif path == "/diagnostico_filas":
        return handle_diagnostico_filas(request)
    elif path == "/update_cg_rechazados":
        return handle_update_cg_rechazados(request)
    elif path == "/recordatorio_sctr_3":
        return handle_recordatorio_sctr_3(request)
    elif path == "/recordatorio_sctr_1":
        return handle_recordatorio_sctr_1(request)
    elif path == "/recordatorio_sctr_2":
        return handle_recordatorio_sctr_2(request)
    elif path == "/excel_rechazados":
        return handle_excel_rechazados(request)
    elif path == "/sctr/crear_grupo":
        return handle_sctr_crear_grupo(request)
    elif path == "/sctr/whatsapp/resumen" and request.method == "POST":
        return handle_sctr_resumen_whatsapp(request)
    elif path == "/sctr/whatsapp/excel" and request.method == "POST":
        return handle_sctr_excel_whatsapp(request)
    elif path == "/sctr/whatsapp/borrar" and request.method == "POST":
        return handle_sctr_borrar_empleado(request)
    elif path == "/sctr/excel-web":
        return handle_sctr_excel_web(request)
    elif path == "/campania/cargar-excel" and request.method == "POST":
        return handle_cargar_excel_campania(request)
    elif path == "/campania/generar-sms" and request.method == "POST":
        return handle_generar_sms_campania(request)
    elif path == "/campania/mantenimiento" and request.method in ("GET", "POST"):
        return handle_mantenimiento_campania(request)
    elif path == "/campania/reporte-clicks" and request.method == "GET":
        return handle_reporte_clicks_campania(request)
    elif path == "/campania/cargar-excel-web" and request.method in ("GET", "POST"):
        return handle_cargar_excel_web(request)
    elif path == "/campania/generar-sms-web" and request.method in ("GET", "POST"):
        return handle_generar_sms_web(request)
    elif path == "/campania/generar-sms-csv" and request.method == "GET":
        return handle_generar_sms_csv(request)
    elif path.startswith("/sctr/") and request.method == "GET" and len(path.split("/")) == 3:
        codigo = path.split("/")[2]
        return handle_sctr_form(request, codigo)
    elif path.startswith("/sctr/") and path.endswith("/submit") and request.method == "POST":
        codigo = path.split("/")[2]
        return handle_sctr_submit(request, codigo)
    elif path == "/admision/sctr_nuevo" and request.method == "POST":
        return handle_admision_sctr_nuevo(request)
    elif path == "/admision/sctr_continuidad" and request.method == "POST":
        return handle_admision_sctr_continuidad(request)
    elif path == "/admision/eps_nuevo" and request.method == "POST":
        return handle_admision_eps_nuevo(request)
    elif path == "/admision/form" and request.method == "GET":
        return handle_admision_form_web(request)
    elif path.startswith("/c/"):
        code = path.split("/c/")[1]
        return handle_redirect_campania(code, request)
    elif path.startswith("/static/"):
        filename = path.split("/static/")[1]
        return handle_static_campania(filename)

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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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

        if row and row[0]:
            doi      = row[0].title()
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
    doi           = data.get("doi", "")
    direccion     = data.get("direccion", "").upper()
    distrito      = data.get("distrito", "").upper()
    direccion_nro = data.get("direccion_nro", 0)

    # Conversión antes de las validaciones
    try:
        direccion_nro = int(direccion_nro)
    except (ValueError, TypeError):
        return (json.dumps({"error": "El campo direccion_nro debe ser un número"}), 400, HEADERS)

    # Validaciones
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
    data = request.get_json(force=True, silent=True) or {}
    doi                    = data.get("doi", "")
    nombre_ubigeo_distrito = data.get("nombre_ubigeo_distrito", "").strip().upper()
    va_nombre_o_ubigeo     = data.get("va_nombre_o_ubigeo", "")
    direccion_nro          = data.get("direccion_nro", 0)

    # Conversión antes de las validaciones
    try:
        direccion_nro = int(direccion_nro)
    except (ValueError, TypeError):
        return (json.dumps({"error": "El campo direccion_nro debe ser un número"}), 400, HEADERS)

    # Validaciones
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
    data = request.get_json(force=True, silent=True) or {}
    distrito = data.get("distrito", "").strip().upper()

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
    data = request.get_json(force=True, silent=True) or {}
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
            nombres       = row[0].upper()  if row[0]  else ""
            apellidos     = row[1].upper()  if row[1]  else ""
            direccion1    = row[2].upper()  if row[2]  else ""
            distrito1     = row[3].upper()  if row[3]  else ""
            direccion2    = row[4].upper()  if row[4]  else ""
            distrito2     = row[5].upper()  if row[5]  else ""
            direccion3    = row[6].upper()  if row[6]  else ""
            distrito3     = row[7].upper()  if row[7]  else ""
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
    data = request.get_json(force=True, silent=True) or {}
    message_id         = data.get("message_id", "")
    chat_id            = data.get("chat_id", "")
    channel_identifier = data.get("channel_identifier", "")
    user_name          = data.get("user_name", "")
    campaign_id        = data.get("campaign_id", "")
    user_identifier    = data.get("user_identifier", "")
    last_user_input    = data.get("first_user_input", "")
    opcion             = data.get("opcion", "")
    doi                = data.get("doi", "")
    nombres            = data.get("nombres", "").upper()
    apellidos          = data.get("apellidos", "").upper()
    direccion          = data.get("direccion", "").upper()
    distrito           = data.get("distrito", "").upper()
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
                nombres       = row[0].upper() if row[0] else nombres
                apellidos     = row[1].upper() if row[1] else apellidos
                escoge        = str(row[11])   if row[11] else "1"
                nro_contacto1 = row[8] or ""

                if escoge in ["1", ""]:
                    direccion = row[2].upper() if row[2] else ""
                    distrito  = row[3].upper() if row[3] else ""
                elif escoge == "2":
                    direccion = row[4].upper() if row[4] else ""
                    distrito  = row[5].upper() if row[5] else ""
                elif escoge == "3":
                    direccion = row[6].upper() if row[6] else ""
                    distrito  = row[7].upper() if row[7] else ""

        # ── 2. Descarga archivo y sube a Cloud Storage ───────────────────────
        try:
            url_publica = url
            #if url and opcion != "Consulta Asesor":
            if url and opcion not in ["Consulta Asesor", "Pase de atencion clinicas"]:
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
                "EMAIL":           "sofimedicamentos@sanitasperu.com",
                "DIRECCION":       direccion,
                "DISTRITO":        distrito,
                "CONTACTO":        nro_contacto1,
                "URL":             url_publica,
                "USER_IDENTIFIER": user_identifier,
                "ID_CHAT":         chat_id
            })
            req.post(
                "https://appserv.danaconnect.com/api/1.0/rest/conversation/586630/start/data",
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
# API: /lead-pases-atencion
# ─────────────────────────────────────────────
def handle_lead_pases(request):
    data = request.get_json(force=True, silent=True) or {}
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
            INSERT INTO sofi_leads_pases (nombre, celular, fecha, enviado)
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
# TAREA: /envia_sms_lead_pases
# ─────────────────────────────────────────────
def envia_sms_lead_pases(numero, nombre, id_lead):
    nombre_limpio = re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre, flags=re.UNICODE
    )
    mensaje = f"Sanitas Pases: Hola {nombre_limpio}, ya estamos listos para ayudarte con tu pases de atencion. Haz clic aqui"
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
    cur.execute("UPDATE sofi_leads_pases SET enviado='S' WHERE id=%s", (id_lead,))
    conn.commit()
    cur.close()
    conn.close()

    sms_aviso(numero, "Pase de atencion")
    
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
              WHERE DATE(fecha) = ANY(%s::date[])
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

APISPERU_TOKEN = os.environ["APISPERU_TOKEN"]

def handle_consulta_ruc(request):
    data = request.get_json(force=True, silent=True) or {}
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
# API: /consulta-ruc_2
# ─────────────────────────────────────────────
def handle_consulta_ruc2(request):
    data = request.get_json(force=True, silent=True) or {}
    ruc           = data.get("ruc", "").strip()
    celular_input = data.get("celular", "").strip()

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
            try:
                resp = requests.get(
                    f"https://dniruc.apisperu.com/api/v1/ruc/{ruc}?token={APISPERU_TOKEN}",
                    timeout=10
                )
                obj = resp.json()
            except Exception as e:
                cur.close(); conn.close()
                return (json.dumps({"error": f"Error consultando API RUC: {str(e)}"}), 500, HEADERS)

            razon_social = obj.get("razonSocial", "")
            direccion    = obj.get("direccion", "")
            distrito     = obj.get("distrito", "")
            provincia    = obj.get("provincia", "")
            departamento = obj.get("departamento", "")
            correo       = ""
            actividad    = ""

            if razon_social:
                cur.execute("""
                    INSERT INTO sctr_empresa
                        (ruc, razon_social, celular, direccion, departamento, provincia, distrito, fecha_creacion, fecha_ult_acceso)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (ruc, razon_social, celular_input, direccion, departamento, provincia, distrito, ahora, ahora))
                conn.commit()

        else:
            fila         = dict(zip(col_names, row))
            razon_social = fila.get("razon_social", "")
            correo       = fila.get("correo", "")
            direccion    = fila.get("direccion", "")
            departamento = fila.get("departamento", "")
            provincia    = fila.get("provincia", "")
            distrito     = fila.get("distrito", "")
            actividad    = fila.get("actividad", "")

            if fila.get("celular", "") != celular_input:
                cur.execute("""
                    UPDATE sctr_empresa SET celular = %s, fecha_ult_acceso = %s WHERE ruc = %s
                """, (celular_input, ahora, ruc))
            else:
                cur.execute("""
                    UPDATE sctr_empresa SET fecha_ult_acceso = %s WHERE ruc = %s
                """, (ahora, ruc))
            conn.commit()

        # ── GRUPO Y LINK ──────────────────────────────────────────
        cur.execute("""
            SELECT codigo, comprador_telefono
            FROM sctr_grupos
            WHERE ruc = %s AND comprador_telefono = %s
            ORDER BY id DESC
            LIMIT 1
        """, (ruc, celular_input))
        grupo_existente = cur.fetchone()

        if grupo_existente and (grupo_existente[1] or "")[-9:] == celular_input[-9:]:
            codigo      = grupo_existente[0]
            grupo_nuevo = False
        else:
            import random, string
            codigo = ''.join(random.choices(string.ascii_uppercase + string.digits, k=7))
            cur.execute("""
                INSERT INTO sctr_grupos (codigo, empresa, ruc, comprador_telefono)
                VALUES (%s, %s, %s, %s)
            """, (codigo, razon_social.strip().upper(), ruc, celular_input))
            conn.commit()
            grupo_nuevo = True

        cur.close()
        conn.close()

        base_url = os.environ.get("BASE_URL", "https://sntspe.com")

        return (json.dumps({
            "ruc":             ruc,
            "razon_social":    razon_social,
            "celular":         celular_input,
            "correo":          correo,
            "direccion":       direccion,
            "departamento":    departamento,
            "provincia":       provincia,
            "distrito":        distrito,
            "actividad":       actividad,
            "tiene_actividad": "S" if actividad else "N",
            "codigo":          codigo,
            "link":            f"{base_url}/sctr/{codigo}",
            "grupo_nuevo":     grupo_nuevo,
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
    data = request.get_json(force=True, silent=True) or {}
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
# API: /envia-sms-leads-pases
# ─────────────────────────────────────────────
def handle_envia_sms_leads_pases(request):
    try:
        mensaje = envia_sms_leads_generico(
            tabla             = "sofi_leads_pases",
            mensaje_template  = "Sanitas Pases: Hola {nombre}, ya estamos listos para ayudarte con tu pase de atencion. Haz clic aqui",
            aviso_descripcion = "lead pases"
        )
        return (json.dumps({"mensaje": mensaje}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ─────────────────────────────────────────────
# ENVIO: Cartas Rechazadas
# ─────────────────────────────────────────────	
def envia_sms_cartas_rechazadas(numero, nombre):
    numero  = f"+51{numero}"
    mensaje = f"Sanitas: {nombre}, tu carta fue rechazada. +info en Consultas Generales / Estado de Carta de Garantia"
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "carta rechazada")

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

# ─────────────────────────────────────────────
# TAREA 1: Recordatorio 3 - Vencidos ayer
# Cloud Scheduler: 08:25, lun-sab (1-6)
# ─────────────────────────────────────────────
def handle_recordatorio_sctr_3(request):
    try:
        lima        = pytz.timezone("America/Lima")
        ahora       = datetime.now(lima)
        hoy         = ahora.strftime("%Y-%m-%d")
        ayer        = (ahora - timedelta(days=1)).strftime("%Y-%m-%d")

        cuenta = envia_sctr_vencidos_recordatorio3(hoy, ayer)
        return (json.dumps({"mensaje": f"recordatorio3: {cuenta} enviados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# TAREA 2: Recordatorio 1 - Vencen hoy
# Cloud Scheduler: 11:25, lun-sab (1-6)
# ─────────────────────────────────────────────
def handle_recordatorio_sctr_1(request):
    try:
        lima        = pytz.timezone("America/Lima")
        ahora       = datetime.now(lima)
        hoy         = ahora.strftime("%Y-%m-%d")

        cuenta = envia_sctr_vencidos_recordatorio1(hoy)
        return (json.dumps({"mensaje": f"recordatorio1: {cuenta} enviados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# TAREA 3: Recordatorio 2 - Vencen hoy (tarde)
# Cloud Scheduler: 14:55, lun-vie (1-5)
# ─────────────────────────────────────────────
def handle_recordatorio_sctr_2(request):
    try:
        lima        = pytz.timezone("America/Lima")
        ahora       = datetime.now(lima)
        hoy         = ahora.strftime("%Y-%m-%d")

        cuenta = envia_sctr_vencidos_recordatorio2(hoy)
        return (json.dumps({"mensaje": f"recordatorio2: {cuenta} enviados"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ─────────────────────────────────────────────
# TAREA 4: Lee Excel Rechazados
# Cloud Scheduler: 12:10, lun-sab (1-6)
# ─────────────────────────────────────────────
def handle_excel_rechazados(request):
    try:
        resultado = lee_excel_rechazados()
        return (json.dumps({"mensaje": f"excel_rechazados: {resultado}"}), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ────────────────────────────────────────────────────────
# Descarga el archivo excel y lo almacena en base de datos
# ────────────────────────────────────────────────────────
def handle_sync_cartas_rechazadas(request):
    try:
        print("PASO 1: Iniciando descarga del CSV")
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=1727191727&single=true&output=csv"
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        print(f"PASO 2: CSV descargado, tamaño: {len(resp.content)} bytes")

        try:
            contenido_corregido = resp.content.decode('utf-8').encode('latin-1').decode('utf-8')
        except (UnicodeDecodeError, UnicodeEncodeError):
            print("PASO 3-FALLBACK: No se pudo corregir doble-encoding, usando texto original")
            contenido_corregido = resp.content.decode('utf-8')

        #contenido_corregido = resp.content.decode('utf-8').encode('latin-1').decode('utf-8')
        print("PASO 3: Encoding corregido")

        reader = csv.reader(io.StringIO(contenido_corregido))
        next(reader)
        print("PASO 4: Header saltado, comenzando a procesar filas")

        valores = []
        for row in reader:
            if not row or len(row) <= 4 or not row[0].strip():
                continue
            valores.append((
                row[0]  if len(row)>0  else "",
                row[1]  if len(row)>1  else "",
                row[2]  if len(row)>2  else "",
                row[3]  if len(row)>3  else "",
                row[4]  if len(row)>4  else "",
                row[5]  if len(row)>5  else "",
                row[6]  if len(row)>6  else "",
                row[7]  if len(row)>7  else "",
                row[8]  if len(row)>8  else "",
                row[9]  if len(row)>9  else "",
                row[10] if len(row)>10 else "",
                row[11] if len(row)>11 else "",
                row[12] if len(row)>12 else "",
                row[13] if len(row)>13 else "",
                row[14] if len(row)>14 else "",
                row[15] if len(row)>15 else "",
                row[16] if len(row)>16 else "",
                row[17] if len(row)>17 else "",
                row[18] if len(row)>18 else "",
                row[19] if len(row)>19 else "",
                row[20] if len(row)>20 else "",
                row[21] if len(row)>21 else "",
                row[22] if len(row)>22 else "",
                row[27] if len(row)>27 else "",
                row[37] if len(row)>37 else "",
                "",
            ))

        print(f"PASO 5: Filas procesadas, total: {len(valores)}")

        valores = valores[-5000:]
        print(f"PASO 6: Recorte a últimas 5000: {len(valores)}")

        conn = get_connection()
        cur  = conn.cursor()
        print("PASO 7: Conexión a BD establecida")

        cur.execute("TRUNCATE TABLE cartas_rechazadas")
        print("PASO 8: Tabla truncada")

        execute_values(cur, """
            INSERT INTO cartas_rechazadas
            (id_origen, buscador, sol, tipo_carta, afiliado, prog, ipress, plan,
             cobertura, fecha_prog, estado, monto, usuario, fecha_registro,
             tipo_proceso, tipo_solicitud, dni, estado_intermedio, estado_final,
             observacion, tipo_observ, autonomia, nro_celular, medico_evaluador,
             marca, ia_obs)
            VALUES %s
        """, valores, page_size=500)
        print("PASO 9: Insert completado")

        conn.commit()
        cur.close(); conn.close()
        print("PASO 10: Commit y cierre de conexión OK")

        return (json.dumps({"ok": True, "insertados": len(valores)}), 200, HEADERS)

    except Exception as e:
        import traceback
        print(f"ERROR en sync_cartas_rechazadas: {str(e)}")
        print(traceback.format_exc())
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)

def handle_busca_carta_rechazada_tabla(request):
    data = request.get_json(force=True, silent=True) or {}
    doi     = data.get("doi", "")
    celular = data.get("celular", "")

    if not doi:
        return (json.dumps({"error": "El campo doi es obligatorio"}), 400, HEADERS)

    try:
        resultado_validacion = validar_consulta(celular, doi)
        if not resultado_validacion["status"]:
            return (json.dumps({
                "buscador":    "vacio",
                "observacion": resultado_validacion["msg"]
            }, ensure_ascii=False), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()
        cur.execute("""
            SELECT id_origen, buscador, sol, tipo_carta, afiliado, prog, ipress, plan,
                   cobertura, fecha_prog, estado, monto, usuario, fecha_registro,
                   tipo_proceso, tipo_solicitud, dni, estado_intermedio, estado_final,
                   observacion, tipo_observ, autonomia, nro_celular, medico_evaluador,
                   marca, ia_obs
            FROM cartas_rechazadas
            WHERE afiliado LIKE %s
            ORDER BY id_origen DESC
            LIMIT 1
        """, (f"%{doi}%",))
        row = cur.fetchone()
        cur.close(); conn.close()

        if not row:
            return (json.dumps({
                "buscador":    "vacio",
                "observacion": "No existe información."
            }, ensure_ascii=False), 200, HEADERS)

        cols = ["id","buscador","sol","tipo_carta","afiliado","prog","ipress","plan",
                "cobertura","fecha_prog","estado","monto","usuario","fecha_registro",
                "tipo_proceso","tipo_solicitud","dni","estado_intermedio","estado_final",
                "observacion","tipo_observ","autonomia","nro_celular","medico_evaluador",
                "marca","ia_obs"]
        resultado = dict(zip(cols, row))

        estado_final     = resultado.get("estado_final", "")
        estado_intermedio = resultado.get("estado_intermedio", "")
        if not estado_final and not estado_intermedio:
            resultado["estado_final"] = "EN PROCESO"
        elif not estado_final:
            resultado["estado_final"] = estado_intermedio

        if not resultado.get("observacion"):
            resultado["observacion"] = "Este campo por el momento esta vacío"

        return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

#borrame
def handle_diagnostico_csv_v2(request):
    try:
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=1727191727&single=true&output=csv"
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()

        reader   = csv.reader(io.StringIO(resp.text))
        all_rows = list(reader)
        rows     = all_rows[max(0, len(all_rows) - 5000):]

        columnas_idx = {
            "id_origen": 0, "buscador": 1, "sol": 2, "tipo_carta": 3,
            "afiliado": 4, "prog": 5, "ipress": 6, "plan": 7,
            "cobertura": 8, "fecha_prog": 9, "estado": 10, "monto": 11,
            "usuario": 12, "fecha_registro": 13, "tipo_proceso": 14,
            "tipo_solicitud": 15, "dni": 16, "estado_intermedio": 17,
            "estado_final": 18, "observacion": 19, "tipo_observ": 20,
            "autonomia": 21, "nro_celular": 22, "medico_evaluador": 27,
            "marca": 37
        }

        max_longitudes = {campo: 0 for campo in columnas_idx}
        ejemplo_max     = {campo: "" for campo in columnas_idx}

        for row in rows:
            if not row or len(row) <= 4:
                continue
            for campo, idx in columnas_idx.items():
                valor = row[idx].strip() if len(row) > idx else ""
                if len(valor) > max_longitudes[campo]:
                    max_longitudes[campo] = len(valor)
                    ejemplo_max[campo]    = valor[:120]

        resultado = {
            campo: {"max_longitud": max_longitudes[campo], "ejemplo": ejemplo_max[campo]}
            for campo in columnas_idx
        }

        return (json.dumps({"ok": True, "resultado": resultado}, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)


#borrame
def handle_diagnostico_filas(request):
    try:
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=1727191727&single=true&output=csv"
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()

        reader   = csv.reader(io.StringIO(resp.text))
        all_rows = list(reader)

        header = all_rows[0]
        muestras = {}
        # Tomamos filas 1, 2, 3 (cerca del inicio) y unas del medio/final para comparar
        indices_muestra = [1, 2, 3, 1500, 1501, 3000, 3001, len(all_rows)-2, len(all_rows)-1]

        for idx in indices_muestra:
            if idx < len(all_rows):
                row = all_rows[idx]
                muestras[f"fila_{idx}"] = {
                    "num_columnas": len(row),
                    "datos": {str(i): v for i, v in enumerate(row)}
                }

        return (json.dumps({
            "ok": True,
            "total_filas": len(all_rows),
            "num_columnas_header": len(header),
            "muestras": muestras
        }, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)

def handle_sync_retencion_sms(request):
    try:
        print("PASO 1: Iniciando descarga del CSV de retención")
        link_compartido = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT6_Qoa5V9oLRGS3YvoTkl3SLHEgLWLAbAnEXMN9JlZnMaumZWavfDwoZuSQ5odrsaJKhDoqoNXTvFD/pub?gid=0&single=true&output=csv"
        resp = requests.get(f"{link_compartido}&t={int(time.time())}", timeout=30)
        resp.raise_for_status()
        print(f"PASO 2: CSV descargado, tamaño: {len(resp.content)} bytes")
        try:
            contenido_corregido = resp.content.decode('utf-8').encode('latin-1').decode('utf-8')
        except (UnicodeDecodeError, UnicodeEncodeError):
            print("PASO 3-FALLBACK: No se pudo corregir doble-encoding, usando texto original")
            contenido_corregido = resp.content.decode('utf-8')
        print("PASO 3: Encoding corregido")
        reader = csv.reader(io.StringIO(contenido_corregido))
        next(reader)
        print("PASO 4: Header saltado, comenzando a procesar filas")

        def val(row, i):
            return row[i].strip() if len(row) > i and row[i] else ""

        ahora = datetime.now(LIMA_TZ).replace(tzinfo=None)
        filas_por_documento = {}  # n_documento -> tupla (la última aparición gana)
        duplicados = 0
        for row in reader:
            n_documento = val(row, 8)
            if not row or len(row) <= 8 or not n_documento:
                continue
            if n_documento in filas_por_documento:
                duplicados += 1
            filas_por_documento[n_documento] = (
                n_documento,        # n_documento (llave de upsert)
                val(row, 0),        # fecha_envio_sms
                val(row, 1),        # fecha_solicitud
                val(row, 3),        # n_contrato
                val(row, 4),        # n_familia
                val(row, 5),        # n_usuario
                val(row, 6),        # nombres_apellidos
                val(row, 7),        # tipo_documento
                val(row, 10),       # empresa_contratante
                val(row, 11),       # parentesco
                val(row, 12),       # tipo_contrato
                val(row, 13),       # vigencia_exclusion
                val(row, 14),       # solicitante
                val(row, 15),       # motivo_core
                val(row, 16),       # tipo_plan
                val(row, 17),       # plan_categoria
                val(row, 18),       # importe
                val(row, 19),       # inicio_contrato
                val(row, 20),       # antiguedad
                val(row, 21),       # antiguedad_rango
                val(row, 22),       # sexo
                val(row, 23),       # fecha_nac
                val(row, 24),       # edad
                val(row, 25),       # rango_etareo
                val(row, 27),       # ejecutivo
                val(row, 28),       # telefono1
                val(row, 29),       # telefono2
                val(row, 30),       # movil
                val(row, 35),       # telefono1_ok
                val(row, 36),       # telefono2_ok
                val(row, 37),       # telefono3_ok
                val(row, 38),       # correo
                val(row, 39),       # correo_ok
                val(row, 40),       # mes_renovacion
                val(row, 41),       # numero_usuarios_actuales
                val(row, 42),       # llave
                val(row, 43),       # primer_nombre
                val(row, 47),       # tipo_envio
                ahora,              # actualizado_at
            )
        print(f"PASO 5: Filas procesadas, total: {len(filas_por_documento)} ({duplicados} documentos duplicados resueltos)")
        valores = list(filas_por_documento.values())[-5000:]
        print(f"PASO 6: Recorte a últimas 5000: {len(valores)}")
        conn = get_connection()
        cur  = conn.cursor()
        print("PASO 7: Conexión a BD establecida")
        execute_values(cur, """
            INSERT INTO retencion_sms
            (n_documento, fecha_envio_sms, fecha_solicitud, n_contrato, n_familia,
             n_usuario, nombres_apellidos, tipo_documento, empresa_contratante,
             parentesco, tipo_contrato, vigencia_exclusion, solicitante, motivo_core,
             tipo_plan, plan_categoria, importe, inicio_contrato, antiguedad,
             antiguedad_rango, sexo, fecha_nac, edad, rango_etareo, ejecutivo,
             telefono1, telefono2, movil, telefono1_ok, telefono2_ok, telefono3_ok,
             correo, correo_ok, mes_renovacion, numero_usuarios_actuales, llave,
             primer_nombre, tipo_envio, actualizado_at)
            VALUES %s
            ON CONFLICT (n_documento) DO UPDATE SET
                fecha_envio_sms = EXCLUDED.fecha_envio_sms,
                fecha_solicitud = EXCLUDED.fecha_solicitud,
                n_contrato = EXCLUDED.n_contrato,
                n_familia = EXCLUDED.n_familia,
                n_usuario = EXCLUDED.n_usuario,
                nombres_apellidos = EXCLUDED.nombres_apellidos,
                tipo_documento = EXCLUDED.tipo_documento,
                empresa_contratante = EXCLUDED.empresa_contratante,
                parentesco = EXCLUDED.parentesco,
                tipo_contrato = EXCLUDED.tipo_contrato,
                vigencia_exclusion = EXCLUDED.vigencia_exclusion,
                solicitante = EXCLUDED.solicitante,
                motivo_core = EXCLUDED.motivo_core,
                tipo_plan = EXCLUDED.tipo_plan,
                plan_categoria = EXCLUDED.plan_categoria,
                importe = EXCLUDED.importe,
                inicio_contrato = EXCLUDED.inicio_contrato,
                antiguedad = EXCLUDED.antiguedad,
                antiguedad_rango = EXCLUDED.antiguedad_rango,
                sexo = EXCLUDED.sexo,
                fecha_nac = EXCLUDED.fecha_nac,
                edad = EXCLUDED.edad,
                rango_etareo = EXCLUDED.rango_etareo,
                ejecutivo = EXCLUDED.ejecutivo,
                telefono1 = EXCLUDED.telefono1,
                telefono2 = EXCLUDED.telefono2,
                movil = EXCLUDED.movil,
                telefono1_ok = EXCLUDED.telefono1_ok,
                telefono2_ok = EXCLUDED.telefono2_ok,
                telefono3_ok = EXCLUDED.telefono3_ok,
                correo = EXCLUDED.correo,
                correo_ok = EXCLUDED.correo_ok,
                mes_renovacion = EXCLUDED.mes_renovacion,
                numero_usuarios_actuales = EXCLUDED.numero_usuarios_actuales,
                llave = EXCLUDED.llave,
                primer_nombre = EXCLUDED.primer_nombre,
                tipo_envio = EXCLUDED.tipo_envio,
                actualizado_at = EXCLUDED.actualizado_at
        """, valores, page_size=500)
        print("PASO 8: Upsert completado (no se tocó sms_enviado de filas existentes)")
        conn.commit()
        cur.close(); conn.close()
        print("PASO 9: Commit y cierre de conexión OK")
        return (json.dumps({"ok": True, "procesados": len(valores)}), 200, HEADERS)
    except Exception as e:
        import traceback
        print(f"ERROR en sync_retencion_sms: {str(e)}")
        print(traceback.format_exc())
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)

# ────────────────────────────────────────────────────────
# Pagina web donde de solicita pase de atencion clinicas
# ────────────────────────────────────────────────────────
def handle_admision_form_web(request):
    html = """<!DOCTYPE html>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Formulario de Admisión - Sanitas</title>
<style>
  :root{
    --azul:#003876;
    --azul-claro:#0057b8;
    --gris:#f4f6f9;
    --borde:#d7dde5;
    --rojo:#c0392b;
    --verde:#1e8e3e;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    font-family:'Segoe UI',Arial,sans-serif;
    background:var(--gris);
    color:#222;
  }
  header{
    background:var(--azul);
    color:#fff;
    padding:18px 20px;
    text-align:center;
  }
  header h1{
    margin:0;
    font-size:20px;
    font-weight:600;
  }
  header p{
    margin:4px 0 0;
    font-size:13px;
    opacity:.85;
  }
  .container{
    max-width:640px;
    margin:0 auto;
    padding:16px;
  }
  .tabs{
    display:flex;
    background:#fff;
    border-radius:10px;
    overflow:hidden;
    box-shadow:0 1px 4px rgba(0,0,0,.08);
    margin-bottom:16px;
  }
  .tab{
    flex:1;
    padding:12px 6px;
    text-align:center;
    font-size:13px;
    font-weight:600;
    cursor:pointer;
    color:#667;
    border-bottom:3px solid transparent;
    transition:.15s;
  }
  .tab.active{
    color:var(--azul);
    border-bottom-color:var(--azul-claro);
    background:#eef3fa;
  }
  .card{
    background:#fff;
    border-radius:10px;
    padding:18px;
    box-shadow:0 1px 4px rgba(0,0,0,.08);
  }
  fieldset{
    border:1px solid var(--borde);
    border-radius:8px;
    margin:0 0 16px;
    padding:12px 14px 14px;
  }
  legend{
    font-size:13px;
    font-weight:700;
    color:var(--azul);
    padding:0 6px;
  }
  label{
    display:block;
    font-size:13px;
    font-weight:600;
    margin:10px 0 4px;
    color:#333;
  }
  label.req::after{
    content:" *";
    color:var(--rojo);
  }
  input[type=text],
  input[type=tel],
  input[type=email],
  input[type=datetime-local],
  input[type=date],
  select,
  textarea{
    width:100%;
    padding:10px 12px;
    border:1px solid var(--borde);
    border-radius:6px;
    font-size:14px;
    background:#fff;
  }
  textarea{
    resize:vertical;
    min-height:70px;
  }
  .row2{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:10px;
  }
  @media(max-width:480px){
    .row2{grid-template-columns:1fr;}
  }
  .radio-group{
    display:flex;
    gap:14px;
    margin-top:6px;
  }
  .radio-group label{
    display:flex;
    align-items:center;
    gap:6px;
    font-weight:500;
    margin:0;
  }
  button.submit{
    width:100%;
    background:var(--azul-claro);
    color:#fff;
    border:none;
    padding:14px;
    font-size:15px;
    font-weight:700;
    border-radius:8px;
    cursor:pointer;
    margin-top:6px;
  }
  button.submit:disabled{
    background:#9bb6d6;
    cursor:not-allowed;
  }
  .msg{
    margin-top:14px;
    padding:12px 14px;
    border-radius:8px;
    font-size:14px;
    display:none;
  }
  .msg.ok{
    display:block;
    background:#e7f6ec;
    color:var(--verde);
    border:1px solid #b9e3c4;
  }
  .msg.error{
    display:block;
    background:#fbe9e7;
    color:var(--rojo);
    border:1px solid #f0c2bb;
  }
  .hint{
    font-size:12px;
    color:#888;
    margin-top:3px;
  }
  .form-section{display:none;}
  .form-section.active{display:block;}
</style>
</head>
<body>

<header>
  <h1>Formulario de Admisión</h1>
  <p>Sanitas Perú · Registro de pacientes</p>
</header>

<div class="container">

  <div class="tabs">
    <div class="tab active" data-tipo="SCTR_NUEVO">SCTR Nuevo</div>
    <div class="tab" data-tipo="SCTR_CONTINUIDAD">SCTR Continuidad</div>
    <div class="tab" data-tipo="EPS_NUEVO">EPS Nuevo</div>
  </div>

  <div class="card">
    <form id="admForm" autocomplete="off">

      <!-- ───────────── Campos comunes (siempre visibles) ───────────── -->
      <fieldset>
        <legend>Datos de la clínica</legend>

        <label class="req">Clínica</label>
        <input type="text" id="clinica" required>

        <div class="row2">
          <div>
            <label>Teléfono de la clínica</label>
            <input type="tel" id="telefono_clinica">
          </div>
          <div>
            <label>Correo de la clínica</label>
            <input type="email" id="correo_clinica">
          </div>
        </div>

        <label class="req">Nombre del admisionista</label>
        <input type="text" id="nombre_admisionista" required>

        <label class="req">Celular de registro</label>
        <input type="tel" id="celular_registro" required placeholder="9XXXXXXXX">
      </fieldset>

      <fieldset>
        <legend>Datos del paciente</legend>

        <label class="req">DNI del paciente</label>
        <input type="text" id="dni_paciente" required maxlength="15">

        <label class="req">Nombre del paciente</label>
        <input type="text" id="nombre_paciente" required>

        <label>Teléfono del paciente</label>
        <input type="tel" id="telefono_paciente">
      </fieldset>

      <!-- ───────────── SCTR NUEVO ───────────── -->
      <div class="form-section active" data-tipo="SCTR_NUEVO">
        <fieldset>
          <legend>Datos de la empresa</legend>
          <div class="row2">
            <div>
              <label>RUC de la empresa</label>
              <input type="text" id="sn_ruc_empresa" maxlength="11">
            </div>
            <div>
              <label>Razón social</label>
              <input type="text" id="sn_razon_social">
            </div>
          </div>

          <label>¿Cuenta con formato SCTR firmado?</label>
          <div class="radio-group">
            <label><input type="radio" name="sn_formato_sctr_firmado" value="SI"> Sí</label>
            <label><input type="radio" name="sn_formato_sctr_firmado" value="NO"> No</label>
          </div>
        </fieldset>

        <fieldset>
          <legend>Datos del accidente</legend>
          <label>Fecha y hora del accidente</label>
          <input type="datetime-local" id="sn_fecha_hora_accidente">

          <div class="row2">
            <div>
              <label>Horario de trabajo</label>
              <input type="text" id="sn_horario_trabajo" placeholder="Ej: 8am - 5pm">
            </div>
            <div>
              <label>¿Realizaba horas extras?</label>
              <div class="radio-group">
                <label><input type="radio" name="sn_hace_horas_extras" value="SI"> Sí</label>
                <label><input type="radio" name="sn_hace_horas_extras" value="NO"> No</label>
              </div>
            </div>
          </div>

          <div class="row2">
            <div>
              <label>Cargo del paciente</label>
              <input type="text" id="sn_cargo_paciente">
            </div>
            <div>
              <label>Área de trabajo</label>
              <input type="text" id="sn_area_trabajo">
            </div>
          </div>

          <label>Dirección del accidente</label>
          <input type="text" id="sn_direccion_accidente">

          <label>Implementos de seguridad utilizados</label>
          <input type="text" id="sn_implementos_seguridad">

          <label>Detalle del accidente</label>
          <textarea id="sn_detalle_accidente" placeholder="Describa cómo ocurrió el accidente"></textarea>
        </fieldset>
      </div>

      <!-- ───────────── SCTR CONTINUIDAD ───────────── -->
      <div class="form-section" data-tipo="SCTR_CONTINUIDAD">
        <fieldset>
          <legend>Datos de la empresa</legend>
          <div class="row2">
            <div>
              <label>RUC de la empresa</label>
              <input type="text" id="sc_ruc_empresa" maxlength="11">
            </div>
            <div>
              <label>Razón social</label>
              <input type="text" id="sc_razon_social">
            </div>
          </div>

          <label>¿Cuenta con formato SCTR firmado?</label>
          <div class="radio-group">
            <label><input type="radio" name="sc_formato_sctr_firmado" value="SI"> Sí</label>
            <label><input type="radio" name="sc_formato_sctr_firmado" value="NO"> No</label>
          </div>
        </fieldset>

        <fieldset>
          <legend>Datos de la atención</legend>
          <div class="row2">
            <div>
              <label>Fecha y hora del accidente</label>
              <input type="datetime-local" id="sc_fecha_hora_accidente">
            </div>
            <div>
              <label>Fecha de última atención</label>
              <input type="date" id="sc_fecha_ultima_atencion">
            </div>
          </div>

          <div class="row2">
            <div>
              <label>Especialidad</label>
              <input type="text" id="sc_especialidad">
            </div>
            <div>
              <label>Lateralidad</label>
              <select id="sc_lateralidad">
                <option value="">-- Seleccionar --</option>
                <option value="IZQUIERDA">Izquierda</option>
                <option value="DERECHA">Derecha</option>
                <option value="BILATERAL">Bilateral</option>
                <option value="NO APLICA">No aplica</option>
              </select>
            </div>
          </div>

          <label>Diagnóstico</label>
          <textarea id="sc_diagnostico"></textarea>
        </fieldset>
      </div>

      <!-- ───────────── EPS NUEVO ───────────── -->
      <div class="form-section" data-tipo="EPS_NUEVO">
        <fieldset>
          <legend>Datos de la atención</legend>
          <label>Celular del paciente</label>
          <input type="tel" id="ep_celular_paciente" placeholder="9XXXXXXXX">

          <label>Especialidad</label>
          <input type="text" id="ep_especialidad">
        </fieldset>
      </div>

      <button type="submit" class="submit" id="btnSubmit">Enviar formulario</button>

      <div class="msg" id="msgBox"></div>
    </form>
  </div>
</div>

<script>
  // ──────────────────────────────────────────────
  // Ajustar esta URL base según el deploy de sofi-apis
  // ──────────────────────────────────────────────
  const API_BASE = window.location.origin; // o reemplazar por la URL fija del Cloud Run

  const ENDPOINTS = {
    SCTR_NUEVO:        "/admision/sctr_nuevo",
    SCTR_CONTINUIDAD:  "/admision/sctr_continuidad",
    EPS_NUEVO:         "/admision/eps_nuevo"
  };

  let tipoActual = "SCTR_NUEVO";

  const tabs = document.querySelectorAll(".tab");
  const sections = document.querySelectorAll(".form-section");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => t.classList.remove("active"));
      sections.forEach(s => s.classList.remove("active"));

      tab.classList.add("active");
      tipoActual = tab.dataset.tipo;
      document.querySelector(`.form-section[data-tipo="${tipoActual}"]`).classList.add("active");

      document.getElementById("msgBox").className = "msg";
      document.getElementById("msgBox").style.display = "none";
    });
  });

  function val(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
  }

  function radioVal(name) {
    const el = document.querySelector(`input[name="${name}"]:checked`);
    return el ? el.value : "";
  }

  function buildPayload(tipo) {
    const base = {
      clinica:              val("clinica"),
      telefono_clinica:     val("telefono_clinica"),
      correo_clinica:       val("correo_clinica"),
      nombre_admisionista:  val("nombre_admisionista"),
      celular_registro:     val("celular_registro"),
      dni_paciente:         val("dni_paciente"),
      nombre_paciente:      val("nombre_paciente"),
      telefono_paciente:    val("telefono_paciente"),
    };

    if (tipo === "SCTR_NUEVO") {
      return {
        ...base,
        ruc_empresa:            val("sn_ruc_empresa"),
        razon_social:           val("sn_razon_social"),
        formato_sctr_firmado:   radioVal("sn_formato_sctr_firmado"),
        fecha_hora_accidente:   val("sn_fecha_hora_accidente"),
        horario_trabajo:        val("sn_horario_trabajo"),
        hace_horas_extras:      radioVal("sn_hace_horas_extras"),
        cargo_paciente:         val("sn_cargo_paciente"),
        area_trabajo:           val("sn_area_trabajo"),
        direccion_accidente:    val("sn_direccion_accidente"),
        implementos_seguridad:  val("sn_implementos_seguridad"),
        detalle_accidente:      val("sn_detalle_accidente"),
      };
    }

    if (tipo === "SCTR_CONTINUIDAD") {
      return {
        ...base,
        ruc_empresa:            val("sc_ruc_empresa"),
        razon_social:           val("sc_razon_social"),
        formato_sctr_firmado:   radioVal("sc_formato_sctr_firmado"),
        fecha_hora_accidente:   val("sc_fecha_hora_accidente"),
        fecha_ultima_atencion:  val("sc_fecha_ultima_atencion"),
        especialidad:           val("sc_especialidad"),
        diagnostico:            val("sc_diagnostico"),
        lateralidad:            val("sc_lateralidad"),
      };
    }

    if (tipo === "EPS_NUEVO") {
      return {
        ...base,
        celular_paciente: val("ep_celular_paciente"),
        especialidad:     val("ep_especialidad"),
      };
    }
  }

  document.getElementById("admForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const msgBox = document.getElementById("msgBox");
    const btn = document.getElementById("btnSubmit");

    // Validación básica de campos requeridos comunes
    const requeridos = ["clinica", "nombre_admisionista", "celular_registro", "dni_paciente", "nombre_paciente"];
    for (const id of requeridos) {
      if (!val(id)) {
        msgBox.textContent = "Por favor complete todos los campos obligatorios (*).";
        msgBox.className = "msg error";
        msgBox.style.display = "block";
        document.getElementById(id).focus();
        return;
      }
    }

    const payload = buildPayload(tipoActual);
    const endpoint = ENDPOINTS[tipoActual];

    btn.disabled = true;
    btn.textContent = "Enviando...";
    msgBox.style.display = "none";

    try {
      const resp = await fetch(API_BASE + endpoint + "?origen=web", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const data = await resp.json();

      if (resp.ok && data.ok) {
        msgBox.textContent = `✓ ${data.mensaje || "Formulario registrado correctamente."} (N° ${data.formulario_id})`;
        msgBox.className = "msg ok";
        document.getElementById("admForm").reset();
      } else {
        msgBox.textContent = "Error: " + (data.error || "No se pudo registrar el formulario.");
        msgBox.className = "msg error";
      }
    } catch (err) {
      msgBox.textContent = "Error de conexión. Intente nuevamente.";
      msgBox.className = "msg error";
    } finally {
      msgBox.style.display = "block";
      btn.disabled = false;
      btn.textContent = "Enviar formulario";
    }
  });
</script>

</body>
</html>

"""
    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})

# ────────────────────────────────────────────────────────
# Envio de mensaje de whatsap a SCTR vencidos por importe
# ────────────────────────────────────────────────────────
def handle_envia_sctr_vencidos_envio_wsp(request):
    try:
        lima = pytz.timezone("America/Lima")
        hoy  = datetime.now(lima).strftime("%Y-%m-%d")
        conn = get_connection()
        cur  = conn.cursor()
        sincronizar_ruc_celular(cur)
        conn.commit()
        rucs = obtener_rucs_vencidos_por_importe(cur, hoy)
        if not rucs:
            cur.close()
            conn.close()
            return (json.dumps({"mensaje": "No hay RUC con emisión vencida"}), 200, HEADERS)
        cur.execute("""
            SELECT ruc, celular FROM sctr_ruc_celular
            WHERE ruc = ANY(%s)
        """, (rucs,))
        resultados = cur.fetchall()

        enviados = envia_wsp_sctr_vencidos(resultados)
        registrar_wsp_sctr_vencidos(cur, conn, resultados)

        cur.close()
        conn.close()

        return (json.dumps({"mensaje": f"Todo OK: {enviados} mensajes enviados"}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

# ────────────────────────────────────────────────────────
# Filtra los rucs de SCTR vencidos por importe (parametrizable)
# ────────────────────────────────────────────────────────
from datetime import datetime, timedelta

MONTO_MINIMO_ABONADO = 299

def obtener_rucs_vencidos_por_importe(cur, fecha):
    # Aceptar fecha como string "YYYY-MM-DD" o como objeto date
    if isinstance(fecha, str):
        fecha_dt = datetime.strptime(fecha, "%Y-%m-%d").date()
    else:
        fecha_dt = fecha

    fechas_a_buscar = [fecha_dt]

    # weekday(): lunes=0, martes=1, ..., domingo=6
    if fecha_dt.weekday() == 0:  # Lunes
        fechas_a_buscar.append(fecha_dt - timedelta(days=1))  # Domingo anterior

    cur.execute("""
        SELECT DISTINCT nro_doc
        FROM sctr_emisiones
        WHERE fin_vigencia = ANY(%s)
          AND fin_vigencia IS NOT NULL
          AND nro_doc IS NOT NULL
          AND monto_abonado > %s
    """, (fechas_a_buscar, MONTO_MINIMO_ABONADO))
    return [row[0] for row in cur.fetchall()]

# ───────────────────────────────────────────────────────
# Credenciales Chattigo (guardar en Secret Manager / variables de entorno)
# ───────────────────────────────────────────────────
CHATTIGO_USERNAME = os.environ.get("CHATTIGO_USERNAME")
CHATTIGO_PASSWORD = os.environ.get("CHATTIGO_PASSWORD")

CHATTIGO_LOGIN_URL   = "https://massive.chattigo.com/message/login"
CHATTIGO_INBOUND_URL = "https://massive.chattigo.com/message/inbound"

CHATTIGO_TEMPLATE_ID    = "12345"
CHATTIGO_NAMESPACE      = "04bfa3e8_1aef_4b29_a499_3946cd213c68"
CHATTIGO_TEMPLATE_NAME  = "sctr_renovacion"
CHATTIGO_HEADER_IMG_URL = "https://storage.googleapis.com/sanitas-recetas-archivos/static/sctr_vencimiento_hoy.jpg"
CHATTIGO_DID            = "51955973936"


def _obtener_token_chattigo():
    """Hace login en Chattigo y retorna el access_token."""
    resp = requests.post(
        CHATTIGO_LOGIN_URL,
        json={"username": CHATTIGO_USERNAME, "password": CHATTIGO_PASSWORD},
        timeout=10
    )
    resp.raise_for_status()
    data = resp.json()

    if not data.get("access") or not data.get("access_token"):
        raise Exception(f"Login Chattigo falló: {data}")

    return data["access_token"]

# ─────────────────────────────────────────────────────────────────
# Envia whatsApp masivos, previamente hace login para traer token
# ──────────────────────────────────────────────────────────────────
def envia_wsp_sctr_vencidos(resultados):
    """
    Envía un WhatsApp masivo (HSM) a todos los contactos de 'resultados'.
    resultados: lista de tuplas (ruc, celular)
    Retorna la cantidad de destinatarios enviados.
    """
    if not resultados:
        return 0

    token = _obtener_token_chattigo()

    destinations = [
        {
            "id": CHATTIGO_TEMPLATE_ID,
            "destination": celular,
            "parameters": [ruc]
        }
        for ruc, celular in resultados
    ]

    payload = {
        "id": "retencion_001",
        "did": CHATTIGO_DID,
        "type": "HSM",
        "channel": "WHATSAPP",
        "hsm": {
            "destinations": destinations,
            "component": {
                "header": {
                    "type": "image_url",
                    "value": CHATTIGO_HEADER_IMG_URL,
                    "isBinaryFile": False,
                    "filename": "header.jpg"
                }
            },
            "namespace": CHATTIGO_NAMESPACE,
            "template": CHATTIGO_TEMPLATE_NAME,
            "botAttention": True,
            "languageCode": "es"
        }
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    resp = requests.post(CHATTIGO_INBOUND_URL, json=payload, headers=headers, timeout=20)
    resp.raise_for_status()

    logging.info(f"[Chattigo] Envío masivo OK: {resp.text}")
    return len(destinations)

# ─────────────────────────────────────────────────────────────────
# Guarda en tabla el registro de envio de whatsapp por numero y Ruc
# ─────────────────────────────────────────────────────────────────
def registrar_wsp_sctr_vencidos(cur, conn, resultados):
    """
    Inserta un registro por cada contacto al que se le envió el WhatsApp.
    resultados: lista de tuplas (ruc, celular)
    """
    lima = pytz.timezone("America/Lima")
    fecha_envio = datetime.now(lima)

    cur.executemany("""
        INSERT INTO sofi_wsp_avisos (ruc, celular, fecha)
        VALUES (%s, %s, %s)
    """, [(ruc, celular, fecha_envio) for ruc, celular in resultados])

    conn.commit()

# ─────────────────────────────────────────────────────────────────
# RETENCION - Definicion para el control de errores
# ─────────────────────────────────────────────────────────────────
def con_log_errores(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Error en {func.__name__}: {e}")
            logging.error(traceback.format_exc())
            return (json.dumps({"error": f"Error interno en {func.__name__}: {str(e)}"}), 500, HEADERS)
    return wrapper

# ─────────────────────────────────────────────────────────────────
# RETENCION - Guarda excel con informacion de clientes y celulares
# ─────────────────────────────────────────────────────────────────
@con_log_errores
def handle_cargar_excel_campania(request):
    api_key = request.headers.get("X-API-Key")
    if api_key != os.environ.get("API_KEY"):
        return jsonify({"error": "No autorizado"}), 401

    if "archivo" not in request.files:
        return jsonify({"error": "archivo requerido (multipart/form-data)"}), 400

    archivo = request.files["archivo"]
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    ws = wb.active

    filas = []
    headers = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(h).strip().lower() for h in row]
            continue
        data = dict(zip(headers, row))
        telefono1 = str(data.get("telefono1") or "").strip()
        tipo_cliente = str(data.get("tipo_cliente") or "").strip().upper()
        if not telefono1 or not tipo_cliente:
            continue
        fecha_carga = data.get("fecha_carga")
        if isinstance(fecha_carga, datetime):
            fecha_carga = fecha_carga.date()
        filas.append((
            str(data.get("nombre") or "").strip(),
            telefono1,
            str(data.get("telefono2") or "").strip() or None,
            str(data.get("telefono3") or "").strip() or None,
            tipo_cliente,
            fecha_carga
        ))
    wb.close()

    if not filas:
        return jsonify({"error": "no se encontraron filas válidas"}), 400

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            execute_values(cur, """
                INSERT INTO campania_contactos
                    (nombre, telefono1, telefono2, telefono3, tipo_cliente, fecha_carga)
                VALUES %s
                ON CONFLICT (telefono1, tipo_cliente) DO UPDATE SET
                    nombre = EXCLUDED.nombre,
                    telefono2 = EXCLUDED.telefono2,
                    telefono3 = EXCLUDED.telefono3,
                    fecha_carga = EXCLUDED.fecha_carga,
                    actualizado_at = NOW()
            """, filas)
        conn.commit()
    finally:
        conn.close()

    return jsonify({"insertados_o_actualizados": len(filas)}), 200

def _generar_code_campania(celular: str) -> str:
    raw = f"{celular}{time.time()}".encode()
    h = hashlib.md5(raw).digest()
    return base64.urlsafe_b64encode(h)[:7].decode()

def _generar_links_para_campania(tipo_cliente, campana):
    """
    Genera un link corto por cada contacto de ese tipo_cliente,
    inserta en campania_links, y devuelve [(celular, nombre, mensaje), ...]
    """
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT mensaje, whatsapp_destino FROM campania_plantillas
                WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            plantilla = cur.fetchone()
            if not plantilla:
                raise ValueError(f"no existe plantilla para {tipo_cliente}")
            mensaje_tpl, whatsapp_destino = plantilla

            cur.execute("""
                SELECT nombre, telefono1 FROM campania_contactos
                WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            contactos = cur.fetchall()

            filas_out = []
            registros_insert = []
            for nombre, telefono1 in contactos:
                code = _generar_code_campania(telefono1)
                registros_insert.append((code, telefono1, nombre, tipo_cliente, campana, whatsapp_destino))
                link_corto = f"{base_url}/c/{code}"
                mensaje_final = mensaje_tpl.replace("{Nombre}", nombre or "").replace("{LINK}", link_corto)
                filas_out.append((telefono1, nombre, mensaje_final))

            if registros_insert:
                execute_values(cur, """
                    INSERT INTO campania_links
                        (code, celular, nombre, tipo_cliente, campana, whatsapp_destino)
                    VALUES %s
                """, registros_insert)
            conn.commit()
    finally:
        conn.close()

    return filas_out

@con_log_errores
def handle_generar_sms_campania(request):
    api_key = request.headers.get("X-API-Key")
    if api_key != os.environ.get("API_KEY"):
        return jsonify({"error": "No autorizado"}), 401

    data = request.get_json(force=True)
    tipo_cliente = (data.get("tipo_cliente") or "").strip().upper()
    campana = data.get("campana", "")
    if not tipo_cliente:
        return jsonify({"error": "tipo_cliente requerido"}), 400

    try:
        filas = _generar_links_para_campania(tipo_cliente, campana)
    except ValueError as e:
        return jsonify({"error": str(e)}), 404

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["celular", "mensaje"])
    writer.writerows([(c, m) for c, n, m in filas])

    return app.response_class(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=sms_{tipo_cliente}.csv"}
    )

@con_log_errores
def handle_generar_sms_web(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT tipo_cliente FROM campania_plantillas ORDER BY tipo_cliente")
            tipos = [r[0] for r in cur.fetchall()]
    finally:
        conn.close()

    opciones_html = "".join(f'<option value="{t}">{t}</option>' for t in tipos)
    resultado_html = ""
    campana_val = ""

    if request.method == "POST":
        tipo_val = (request.form.get("tipo_cliente") or "").strip().upper()
        campana_val = (request.form.get("campana") or "").strip()

        if not tipo_val or not campana_val:
            resultado_html = '<p style="color:red;">Selecciona tipo_cliente y escribe un nombre de campaña.</p>'
        else:
            try:
                resultados = _generar_y_enviar_campania(tipo_val, campana_val)
            except ValueError as e:
                resultados = None
                resultado_html = f'<p style="color:red;">{e}</p>'

            if resultados is not None:
                if not resultados:
                    resultado_html = '<p style="color:orange;">No hay contactos cargados para ese tipo_cliente.</p>'
                else:
                    exitosos = sum(1 for _, _, ok, _ in resultados if ok)
                    fallidos = len(resultados) - exitosos
                    filas_tabla = "".join(
                        f"""<tr>
                            <td>{cel}</td><td>{nom or ''}</td>
                            <td style="color:{'green' if ok else 'red'};">{'✅ Enviado' if ok else '❌ Error'}</td>
                            <td>{err or ''}</td>
                        </tr>"""
                        for cel, nom, ok, err in resultados
                    )
                    resultado_html = f"""
                    <p style="color:green;">✅ {exitosos} enviados correctamente{f', ❌ {fallidos} con error' if fallidos else ''} — tipo <b>{tipo_val}</b> / campaña <b>{campana_val}</b>.</p>
                    <table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:13px;">
                    <tr><th>Celular</th><th>Nombre</th><th>Estado</th><th>Detalle error</th></tr>
                    {filas_tabla}
                    </table>
                    <p>Los clics de estos links ya se pueden ver en el <a href="/campania/reporte-clicks?pw={pw}">reporte de clics</a>.</p>
                    """

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Generar y enviar SMS - Sanitas</title>
  <style>
    body {{ font-family: sans-serif; max-width: 900px; margin: 40px auto; padding: 0 20px; }}
    .box {{ border: 1px solid #ccc; border-radius: 8px; padding: 25px; margin-bottom: 20px; }}
    select, input[type="text"] {{ padding: 6px; margin: 8px 0; width: 100%; box-sizing: border-box; }}
    button {{ background: #0057A3; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }}
    button:hover {{ background: #003d73; }}
    table {{ margin-top: 15px; }}
  </style>
</head>
<body>
  <h2>Generar links y enviar SMS</h2>
  <div class="box">
    <form method="POST" action="/campania/generar-sms-web?pw={pw}"
          onsubmit="return confirm('Esto enviará SMS reales por Danaconnect a todos los contactos de este tipo_cliente. ¿Continuar?');">
      <label>Tipo de cliente:</label>
      <select name="tipo_cliente" required>
        <option value="">-- selecciona --</option>
        {opciones_html}
      </select>
      <label>Nombre de campaña (identifica este envío):</label>
      <input type="text" name="campana" placeholder="ej. retencion_jul2026" value="{campana_val}" required>
      <button type="submit">Generar y enviar SMS</button>
    </form>
    {resultado_html}
  </div>
  <p>
    <a href="/campania/mantenimiento?pw={pw}">← Volver a mantenimiento</a> |
    <a href="/campania/cargar-excel-web?pw={pw}">Cargar Excel</a> |
    <a href="/campania/reporte-clicks?pw={pw}">Ver reporte de clics</a>
  </p>
</body>
</html>"""
    return html, 200

@con_log_errores
def handle_generar_sms_csv(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    tipo_cliente = (request.args.get("tipo_cliente") or "").strip().upper()
    campana = (request.args.get("campana") or "").strip()
    if not tipo_cliente or not campana:
        return "Parámetros incompletos", 400

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT mensaje FROM campania_plantillas WHERE tipo_cliente = %s", (tipo_cliente,))
            row = cur.fetchone()
            if not row:
                return "Plantilla no encontrada", 404
            mensaje_tpl = row[0]

            cur.execute("""
                SELECT celular, nombre, code FROM campania_links
                WHERE tipo_cliente = %s AND campana = %s
                ORDER BY creado_at
            """, (tipo_cliente, campana))
            links = cur.fetchall()
    finally:
        conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["celular", "mensaje"])
    for celular, nombre, code in links:
        link_corto = f"{base_url}/c/{code}"
        mensaje_final = mensaje_tpl.replace("{Nombre}", nombre or "").replace("{LINK}", link_corto)
        writer.writerow([celular, mensaje_final])

    return app.response_class(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=sms_{tipo_cliente}_{campana}.csv"}
    )

IMG_VERSION = str(int(time.time()))

@con_log_errores
def handle_redirect_campania(code, request):
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT celular, tipo_cliente, whatsapp_destino FROM campania_links WHERE code = %s", (code,))
            row = cur.fetchone()
            if not row:
                return "Link no encontrado", 404
            celular, tipo_cliente, whatsapp_destino = row

            cur.execute("""
                UPDATE campania_links SET
                    clicked_at = COALESCE(clicked_at, NOW()),
                    click_count = click_count + 1,
                    ip_address = %s
                WHERE code = %s
            """, (request.remote_addr, code))
            conn.commit()

            cur.execute("""
                SELECT og_title, og_description, og_image
                FROM campania_plantillas WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            og = cur.fetchone()
    finally:
        conn.close()

    og_title, og_description, og_image = og if og else ("Sanitas Perú", "Contáctanos por WhatsApp", "sanitas-general.jpg")
    url_canonica = f"{base_url}/c/{code}"
    url_destino = f"https://wa.me/{whatsapp_destino}"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0;url={url_destino}">
  <meta property="og:type" content="website">
  <meta property="og:url" content="{url_canonica}">
  <meta property="og:site_name" content="Sanitas Perú">
  <meta property="og:title" content="{og_title}">
  <meta property="og:description" content="{og_description}">
  <meta property="og:image" content="{base_url}/static/{og_image}?v={IMG_VERSION}">
  <meta property="og:image:width" content="1200">
  <meta property="og:image:height" content="630">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{og_title}">
  <meta name="twitter:description" content="{og_description}">
  <meta name="twitter:image" content="{base_url}/static/{og_image}?v={IMG_VERSION}">
  <title>{og_title}</title>
</head>
<body>
  <script>window.location.href = "{url_destino}";</script>
  <p>Redirigiendo... <a href="{url_destino}">Haz clic aquí si no eres redirigido.</a></p>
</body>
</html>"""
    return html, 200

ADMIN_PW_CAMPANIA = os.environ.get("AGENTE_KB_CLAVE")

@con_log_errores
def handle_mantenimiento_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    conn = get_connection()
    try:
        if request.method == "POST":
            tipo_cliente = request.form.get("tipo_cliente")
            with conn.cursor() as cur:
                cur.execute("""
                    UPDATE campania_plantillas SET
                        mensaje = %s, og_title = %s, og_description = %s,
                        og_image = %s, whatsapp_destino = %s, actualizado_at = NOW()
                    WHERE tipo_cliente = %s
                """, (
                    request.form.get("mensaje"),
                    request.form.get("og_title"),
                    request.form.get("og_description"),
                    request.form.get("og_image"),
                    request.form.get("whatsapp_destino"),
                    tipo_cliente
                ))
            conn.commit()

        with conn.cursor() as cur:
            cur.execute("""
                SELECT tipo_cliente, mensaje, og_title, og_description, og_image, whatsapp_destino
                FROM campania_plantillas ORDER BY tipo_cliente
            """)
            plantillas = cur.fetchall()
    finally:
        conn.close()

    filas_html = ""
    for tc, mensaje, og_title, og_desc, og_img, wa in plantillas:
        filas_html += f"""
        <form method="POST" action="/campania/mantenimiento?pw={pw}" style="border:1px solid #ccc;padding:15px;margin-bottom:15px;">
            <input type="hidden" name="tipo_cliente" value="{tc}">
            <h3>{tc}</h3>
            <label>Mensaje (usa {{Nombre}} y {{LINK}}):</label><br>
            <textarea name="mensaje" rows="3" style="width:100%;">{mensaje}</textarea><br>
            <label>OG Title:</label><br>
            <input type="text" name="og_title" value="{og_title}" style="width:100%;"><br>
            <label>OG Description:</label><br>
            <input type="text" name="og_description" value="{og_desc}" style="width:100%;"><br>
            <label>OG Image (archivo en /static):</label><br>
            <input type="text" name="og_image" value="{og_img}" style="width:100%;"><br>
            <label>WhatsApp destino:</label><br>
            <input type="text" name="whatsapp_destino" value="{wa}" style="width:100%;"><br><br>
            <button type="submit">Guardar</button>
        </form>"""

    html = f"""<!DOCTYPE html>
        <html><head><meta charset="UTF-8"><title>Mantenimiento Campaña</title></head>
        <body style="font-family:sans-serif;max-width:700px;margin:30px auto;">
        <h2>Plantillas de campaña</h2>
        <p>
        <a href="/campania/generar-sms-web?pw={pw}">✉️ Generar SMS</a> |
        <a href="/campania/cargar-excel-web?pw={pw}">📤 Cargar Excel de contactos</a> |
        <a href="/campania/reporte-clicks?pw={pw}">📊 Ver reporte de clics</a>
        </p>
        {filas_html}
        </body></html>"""
    return html, 200


@con_log_errores
def handle_reporte_clicks_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT c.telefono1, c.nombre, c.tipo_cliente,
                       l.click_count, l.clicked_at, l.campana
                FROM campania_contactos c
                LEFT JOIN campania_links l
                    ON l.celular = c.telefono1 AND l.tipo_cliente = c.tipo_cliente
                ORDER BY l.clicked_at DESC NULLS LAST
            """)
            rows = cur.fetchall()
    finally:
        conn.close()

    filas_html = ""
    total_clicks = 0
    for telefono, nombre, tipo_cliente, click_count, clicked_at, campana in rows:
        clicked = click_count and click_count > 0
        if clicked:
            total_clicks += 1
        filas_html += f"""<tr>
            <td>{telefono}</td><td>{nombre or ''}</td><td>{tipo_cliente}</td>
            <td>{campana or ''}</td>
            <td>{'Sí' if clicked else 'No'}</td>
            <td>{click_count or 0}</td>
            <td>{clicked_at or ''}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Reporte de clics</title></head>
<body style="font-family:sans-serif;max-width:900px;margin:30px auto;">
<h2>Reporte de clics ({total_clicks} de {len(rows)} contactaron)</h2>
<p>
<a href="/campania/mantenimiento?pw={pw}">← Volver a mantenimiento</a> |
    <a href="/campania/cargar-excel-web?pw={pw}">Cargar Excel</a>
</p>
<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;">
<tr><th>Teléfono</th><th>Nombre</th><th>Tipo</th><th>Campaña</th><th>Clic</th><th># Clics</th><th>Fecha</th></tr>
{filas_html}
</table>
</body></html>"""
    return html, 200

def handle_static_campania(filename):
    return send_from_directory("static", filename)

def con_log_errores(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logging.error(f"Error en {func.__name__}: {e}")
            logging.error(traceback.format_exc())
            return (json.dumps({"error": f"Error interno en {func.__name__}: {str(e)}"}), 500, HEADERS)
    return wrapper

@con_log_errores
def handle_cargar_excel_web(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    mensaje = ""
    resumen = ""

    if request.method == "POST":
        if "archivo" not in request.files or request.files["archivo"].filename == "":
            mensaje = '<p style="color:red;">Selecciona un archivo Excel.</p>'
        else:
            archivo = request.files["archivo"]
            try:
                wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
                ws = wb.active

                filas = []
                headers = None
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i == 0:
                        headers = [str(h).strip().lower() for h in row]
                        continue
                    data = dict(zip(headers, row))
                    telefono1 = str(data.get("telefono1") or "").strip()
                    tipo_cliente = str(data.get("tipo_cliente") or "").strip().upper()
                    if not telefono1 or not tipo_cliente:
                        continue
                    fecha_carga = data.get("fecha_carga")
                    if isinstance(fecha_carga, datetime):
                        fecha_carga = fecha_carga.date()
                    filas.append((
                        str(data.get("nombre") or "").strip(),
                        telefono1,
                        str(data.get("telefono2") or "").strip() or None,
                        str(data.get("telefono3") or "").strip() or None,
                        tipo_cliente,
                        fecha_carga
                    ))
                wb.close()

                if not filas:
                    mensaje = '<p style="color:red;">No se encontraron filas válidas (revisa columnas: nombre, telefono1, telefono2, telefono3, tipo_cliente, fecha_carga).</p>'
                else:
                    conn = get_connection()
                    try:
                        with conn.cursor() as cur:
                            execute_values(cur, """
                                INSERT INTO campania_contactos
                                    (nombre, telefono1, telefono2, telefono3, tipo_cliente, fecha_carga)
                                VALUES %s
                                ON CONFLICT (telefono1, tipo_cliente) DO UPDATE SET
                                    nombre = EXCLUDED.nombre,
                                    telefono2 = EXCLUDED.telefono2,
                                    telefono3 = EXCLUDED.telefono3,
                                    fecha_carga = EXCLUDED.fecha_carga,
                                    actualizado_at = NOW()
                            """, filas)
                        conn.commit()
                    finally:
                        conn.close()

                    mensaje = f'<p style="color:green;">✅ {len(filas)} contactos insertados/actualizados correctamente.</p>'

                    with get_connection() as conn2:
                        with conn2.cursor() as cur2:
                            cur2.execute("SELECT tipo_cliente, COUNT(*) FROM campania_contactos GROUP BY tipo_cliente ORDER BY tipo_cliente")
                            resumen_rows = cur2.fetchall()
                    resumen = "<ul>" + "".join(f"<li>{tc}: {cnt} contactos</li>" for tc, cnt in resumen_rows) + "</ul>"

            except Exception as e:
                mensaje = f'<p style="color:red;">Error al procesar el archivo: {e}</p>'

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Cargar campaña - Sanitas</title>
  <style>
    body {{ font-family: sans-serif; max-width: 600px; margin: 40px auto; padding: 0 20px; }}
    .box {{ border: 1px solid #ccc; border-radius: 8px; padding: 25px; }}
    input[type="file"] {{ margin: 15px 0; }}
    button {{ background: #0057A3; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }}
    button:hover {{ background: #003d73; }}
  </style>
</head>
<body>
  <h2>Cargar Excel de campaña</h2>
  <div class="box">
    <p>Columnas esperadas: <b>nombre, telefono1, telefono2, telefono3, tipo_cliente, fecha_carga</b></p>
    <form method="POST" action="/campania/cargar-excel-web?pw={pw}" enctype="multipart/form-data">
      <input type="file" name="archivo" accept=".xlsx,.xls" required>
      <br>
      <button type="submit">Cargar archivo</button>
    </form>
    {mensaje}
    {resumen}
  </div>
  <p><a href="/campania/mantenimiento?pw={pw}">← Volver a mantenimiento</a> |
     <a href="/campania/reporte-clicks?pw={pw}">Ver reporte de clics</a></p>
</body>
</html>"""
    return html, 200

def _limpiar_nombre_sms(nombre):
    return re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre or '', flags=re.UNICODE
    )

def _enviar_sms_campania_uno(numero, nombre, mensaje_tpl, link_corto, code):
    nombre_limpio = _limpiar_nombre_sms(nombre)
    mensaje = mensaje_tpl.replace("{Nombre}", nombre_limpio)

    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  link_corto
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        return False, str(e)

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE campania_links SET enviado_at = NOW(), error_envio = NULL WHERE code = %s", (code,))
        conn.commit()
    finally:
        conn.close()

    return True, None

def _generar_y_enviar_campania(tipo_cliente, campana):
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT mensaje, whatsapp_destino FROM campania_plantillas
                WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            plantilla = cur.fetchone()
            if not plantilla:
                raise ValueError(f"no existe plantilla para {tipo_cliente}")
            mensaje_tpl, whatsapp_destino = plantilla

            cur.execute("""
                SELECT nombre, telefono1 FROM campania_contactos
                WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            contactos = cur.fetchall()

            registros_insert = []
            codes_por_contacto = []
            for nombre, telefono1 in contactos:
                code = _generar_code_campania(telefono1)
                registros_insert.append((code, telefono1, nombre, tipo_cliente, campana, whatsapp_destino))
                codes_por_contacto.append((telefono1, nombre, code))

            if registros_insert:
                execute_values(cur, """
                    INSERT INTO campania_links
                        (code, celular, nombre, tipo_cliente, campana, whatsapp_destino)
                    VALUES %s
                """, registros_insert)
            conn.commit()
    finally:
        conn.close()

    resultados = []
    for telefono1, nombre, code in codes_por_contacto:
        link_corto = f"{base_url}/c/{code}"
        ok, error = _enviar_sms_campania_uno(telefono1, nombre, mensaje_tpl, link_corto, code)
        if not ok:
            conn2 = get_connection()
            try:
                with conn2.cursor() as cur2:
                    cur2.execute("UPDATE campania_links SET error_envio = %s WHERE code = %s", (error, code))
                conn2.commit()
            finally:
                conn2.close()
        resultados.append((telefono1, nombre, ok, error))

    return resultados