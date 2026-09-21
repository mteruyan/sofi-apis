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
import functools
import secrets
from jinja2 import Environment, FileSystemLoader
jinja_env = Environment(loader=FileSystemLoader("templates"))
from google.cloud import storage
storage_client = storage.Client()
from weasyprint import HTML
import uuid
import random
import string
import zipfile
import hmac

from urllib.parse import quote
import google.auth
from google.auth.transport import requests as google_auth_requests
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO

from flask import send_from_directory
from sctr_forms import handle_sctr_form, handle_sctr_submit, handle_sctr_crear_grupo
from sctr_whatsapp import handle_sctr_resumen_whatsapp, handle_sctr_excel_whatsapp, handle_sctr_borrar_empleado, handle_sctr_excel_web
from admision_forms import handle_admision_sctr_nuevo, handle_admision_sctr_continuidad, handle_admision_eps_nuevo
from agente_consulta import (handle_consulta_agente, handle_agente_kb_upsert, handle_agente_kb_listar, handle_agente_kb_form_web, handle_agente_kb_perfil)
from crypto_utils import cifrar, descifrar, cifrar_json, descifrar_json

# ─────────────────────────────────────────────
# VARIABLES
# ─────────────────────────────────────────────
LIMA_TZ = pytz.timezone("America/Lima")
DANACONNECT_URL  = "https://appserv.danaconnect.com/api/1.0/rest/conversation/517078/start/data"
DANACONNECT_AUTH = os.environ["DANACONNECT_AUTH"]
WHATSAPP_URL     = "wa.me/51955973936"
TIMEOUT_OFERTA_SEGUNDOS   = 120    # 2 min para que el proveedor acepte/rechace
MAX_INTENTOS_ASIGNACION   = 5      # después de esto, se marca sin_cobertura
INCREMENTO_RADIO_METROS   = 2000   # cada reintento amplía el radio de búsqueda
RADIO_MAXIMO_METROS       = 15000

ESTADOS_VALIDOS = {"NUEVO", "ASIGNADO", "EN_PROCESO", "RESUELTO", "CERRADO"}

WHATSAPP_CLOUD_API_VERSION = "v21.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


CAMPANIA_CSS = """
:root {
  --sanitas-teal: #00A9B7;
  --sanitas-teal-dark: #007A87;
  --sanitas-navy: #003654;
  --sanitas-bg: #F4F7F9;
  --sanitas-white: #FFFFFF;
  --sanitas-gray: #6B7A82;
  --sanitas-border: #DCE4E8;
  --sanitas-success: #2E9E5B;
  --sanitas-danger: #D64545;
  --sanitas-warning: #E8A33D;
  --radius: 10px;
}

* { box-sizing: border-box; }

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  background: var(--sanitas-bg);
  color: var(--sanitas-navy);
  margin: 0;
  padding: 0;
  line-height: 1.5;
}

.contenedor {
  max-width: 1100px;
  margin: 0 auto;
  padding: 16px;
}

/* Header de marca */
.topbar {
  background: linear-gradient(135deg, var(--sanitas-navy), var(--sanitas-teal-dark));
  color: white;
  padding: 18px 16px;
}
.topbar h1 {
  margin: 0;
  font-size: 20px;
  font-weight: 600;
}
.topbar .subtitulo {
  font-size: 13px;
  opacity: 0.85;
  margin-top: 2px;
}

/* Navegación */
nav.tabs {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  padding: 12px 0 20px;
}
nav.tabs a {
  text-decoration: none;
  color: var(--sanitas-navy);
  background: var(--sanitas-white);
  border: 1px solid var(--sanitas-border);
  padding: 8px 14px;
  border-radius: 999px;
  font-size: 13px;
  font-weight: 500;
  white-space: nowrap;
  transition: all 0.15s ease;
}
nav.tabs a:hover, nav.tabs a.activo {
  background: var(--sanitas-teal);
  color: white;
  border-color: var(--sanitas-teal);
}

/* Tarjetas / cajas */
.card {
  background: var(--sanitas-white);
  border: 1px solid var(--sanitas-border);
  border-radius: var(--radius);
  padding: 20px;
  margin-bottom: 18px;
  box-shadow: 0 1px 3px rgba(0,54,84,0.06);
}
.card h2, .card h3 { margin-top: 0; color: var(--sanitas-navy); }

/* Formularios */
label {
  display: block;
  font-size: 12px;
  font-weight: 600;
  color: var(--sanitas-gray);
  margin-bottom: 4px;
  margin-top: 12px;
  text-transform: uppercase;
  letter-spacing: 0.3px;
}
input[type="text"], input[type="date"], input[type="file"],
select, textarea {
  width: 100%;
  padding: 10px 12px;
  border: 1px solid var(--sanitas-border);
  border-radius: 8px;
  font-size: 15px;
  font-family: inherit;
  background: var(--sanitas-white);
  color: var(--sanitas-navy);
}
input:focus, select:focus, textarea:focus {
  outline: none;
  border-color: var(--sanitas-teal);
  box-shadow: 0 0 0 3px rgba(0,169,183,0.15);
}
textarea { resize: vertical; min-height: 60px; }

button, .btn {
  background: var(--sanitas-teal);
  color: white;
  border: none;
  padding: 12px 22px;
  border-radius: 8px;
  font-size: 15px;
  font-weight: 600;
  cursor: pointer;
  margin-top: 16px;
  transition: background 0.15s ease;
  min-height: 44px; /* área táctil cómoda en celular */
}
button:hover, .btn:hover { background: var(--sanitas-teal-dark); }
button.danger { background: var(--sanitas-danger); }
button.danger:hover { background: #b83a3a; }

/* Filtros en fila (colapsan en columna en móvil) */
.filtros {
  display: flex;
  gap: 14px;
  flex-wrap: wrap;
  align-items: end;
}
.filtros > div { flex: 1; min-width: 140px; }
.filtros label { margin-top: 0; }
.filtros button { margin-top: 0; height: 44px; }

/* Tablas responsive: contenedor con scroll horizontal en pantallas chicas */
.table-wrap {
  overflow-x: auto;
  border: 1px solid var(--sanitas-border);
  border-radius: var(--radius);
  background: var(--sanitas-white);
  -webkit-overflow-scrolling: touch;
}
table {
  border-collapse: collapse;
  width: 100%;
  min-width: 600px; /* fuerza scroll en vez de aplastar columnas */
  font-size: 13px;
}
th {
  background: var(--sanitas-navy);
  color: white;
  text-align: left;
  padding: 10px 12px;
  font-weight: 600;
  position: sticky;
  top: 0;
}
td {
  padding: 9px 12px;
  border-bottom: 1px solid var(--sanitas-border);
}
tr:hover td { background: rgba(0,169,183,0.05); }

/* Mensajes de estado */
.msg-exito { color: var(--sanitas-success); font-weight: 600; }
.msg-error { color: var(--sanitas-danger); font-weight: 600; }
.msg-aviso { color: var(--sanitas-warning); font-weight: 600; }

.badge {
  display: inline-block;
  padding: 3px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 600;
}
.badge-ok { background: rgba(46,158,91,0.12); color: var(--sanitas-success); }
.badge-pendiente { background: rgba(232,163,61,0.12); color: var(--sanitas-warning); }
.badge-neutro { background: rgba(107,122,130,0.12); color: var(--sanitas-gray); }

/* Responsive: ajustes para celular */
@media (max-width: 600px) {
  .topbar h1 { font-size: 17px; }
  .card { padding: 14px; border-radius: 8px; }
  .filtros { flex-direction: column; }
  .filtros > div { width: 100%; }
  nav.tabs { gap: 6px; }
  nav.tabs a { font-size: 12px; padding: 7px 11px; }
  button { width: 100%; }
  th, td { padding: 8px; font-size: 12px; }
}
"""

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
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",  
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
        and not request.path.startswith("/s/") \
        and not request.path.startswith("/ps/") \
        and not request.path.startswith("/social/") \
        and not request.path.startswith("/assets/") \
        and not request.path.startswith("/chattigo/") \
        and not request.path.startswith("/track/") \
        and not request.path.startswith("/oferta/") \
        and not request.path.startswith("/proveedor") \
        and not request.path.startswith("/proveedores/") \
        and not request.path.startswith("/ofertas/") \
        and not request.path.startswith("/solicitudes/") \
        and not request.path.startswith("/backoffice/") \
        and not request.path.startswith("/programa-cronico/firmar-consentimiento") \
        and not request.path.startswith("/programa-cronico/mantenimiento") \
        and not request.path.startswith("/programa-cronico/reenviar-danaconnect") \
        and not request.path.startswith("/programa-cronico/reprocesar") \
        and not request.path.startswith("/programa-cronico/aprobar-manual") \
        and not request.path.startswith("/programa-cronico/forzar-recordatorio") \
        and not (request.path.startswith("/programa-cronico/") and request.method == "GET") \
        and not (request.path.startswith("/afiliacion/") and request.path != "/afiliacion/iniciar") \
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
    elif path == "/sync_rechazados_excel":
        return handle_sync_rechazados_excel(request)
    elif path == "/sctr/crear_grupo":
        return handle_sctr_crear_grupo(request)
    elif path == "/sctr/reporte-wsp" and request.method == "GET":
        return handle_reporte_wsp_sctr_vencidos(request)
    elif path == "/chattigo/dashboard":
        return handle_chattigo_dashboard(request)
    elif path == "/chattigo/agentes/snapshot" and request.method == "POST":
        return handle_chattigo_snapshot(request)
    elif path == "/chattigo/agentes/historico" and request.method == "POST":
        return handle_chattigo_historico(request)
    elif path == "/afiliacion/iniciar":
        return handle_iniciar_afiliacion(request)
    elif path == "/afiliacion/consulta-ruc":
        return handle_consulta_ruc_afiliacion(request)
    elif path == "/afiliacion/registrar-titular":
        return handle_registrar_titular_afiliacion(request)
    elif path == "/afiliacion/ocr-documento":
        return handle_ocr_documento_afiliacion(request)
    elif path == "/afiliacion/guardar-paso":
        return handle_guardar_paso_afiliacion(request)
    elif path == "/afiliacion/guardar-declaracion-salud":
        return handle_guardar_declaracion_salud(request)
    elif path == "/afiliacion/guardar-antecedentes-seguro":
        return handle_guardar_antecedentes_seguro(request)
    elif path == "/afiliacion/guardar-firma":
        return handle_guardar_firma_afiliacion(request)
    elif path == "/afiliacion/finalizar":
        return handle_finalizar_afiliacion(request)
    elif path == "/afiliacion/estado":
        return handle_consultar_estado_afiliacion(request)
    elif path == "/afiliacion/completar-datos":
        return handle_completar_datos_afiliacion(request)
    elif path == "/afiliacion/registrar-dependiente":
        return handle_registrar_dependiente_afiliacion(request)
    elif path == "/afiliacion/guardar-documento-civil":
        return handle_guardar_documento_civil_afiliacion(request)
    elif path == "/afiliacion/reenviar-sms" and request.method == "POST":
        return handle_reenviar_sms_afiliacion(request)
    elif path == "/afiliacion/reporte" and request.method == "GET":
        return handle_reporte_afiliacion(request)
    elif path == "/afiliacion/reporte-excel" and request.method == "GET":
        return handle_reporte_excel_afiliacion(request)
    elif path == "/programa-cronico/reenviar-danaconnect":
        return handle_reenviar_danaconnect(request)
    elif path == "/programa-cronico/reprocesar":
        return handle_reprocesar_cronico(request)
    elif path == "/programa-cronico/aprobar-manual":
        return handle_aprobar_manual_cronico(request)
    elif path == "/programa-cronico/recordatorio":
        return handle_recordatorio_cronico(request)
    elif path == "/programa-cronico/forzar-recordatorio":
        return handle_forzar_recordatorio_cronico(request)
    elif path.startswith("/oferta/") and path.endswith("/estado") and request.method == "GET":
        token = path.split("/oferta/")[1].rsplit("/estado", 1)[0]
        return handle_estado_oferta_token(request, token)
    elif path.startswith("/oferta/") and path.endswith("/aceptar") and request.method == "POST":
        token = path.split("/oferta/")[1].rsplit("/aceptar", 1)[0]
        return handle_aceptar_oferta_token(request, token)
    elif path.startswith("/oferta/") and path.endswith("/rechazar") and request.method == "POST":
        token = path.split("/oferta/")[1].rsplit("/rechazar", 1)[0]
        return handle_rechazar_oferta_token(request, token)
    elif path.startswith("/oferta/") and request.method == "GET":
        token = path.split("/oferta/")[1]
        return handle_servir_oferta(request, token)
    elif path.startswith("/proveedores/") and path.endswith("/disponibilidad") and request.method == "POST":
        proveedor_id = path.split("/proveedores/")[1].rsplit("/disponibilidad", 1)[0]
        return handle_toggle_disponibilidad(request, proveedor_id)
    elif path.startswith("/proveedores/") and path.endswith("/estado-actual") and request.method == "GET":
        proveedor_id = path.split("/proveedores/")[1].rsplit("/estado-actual", 1)[0]
        return handle_estado_actual_proveedor(request, proveedor_id)
    elif path.startswith("/ofertas/") and path.endswith("/rechazar") and request.method == "POST":
        oferta_id = path.split("/ofertas/")[1].rsplit("/rechazar", 1)[0]
        return handle_rechazar_oferta(request, oferta_id)
    elif path.startswith("/solicitudes/") and path.endswith("/estado") and request.method == "POST":
        solicitud_id = path.split("/solicitudes/")[1].rsplit("/estado", 1)[0]
        return handle_actualizar_estado_solicitud(request, solicitud_id)
    elif path == "/solicitudes/revisar-timeouts" and request.method == "GET":
        return handle_revisar_timeouts_ofertas(request)
    elif path == "/solicitudes/crear" and request.method == "POST":
        return handle_crear_solicitud(request)
    elif path.startswith("/solicitudes/") and path.endswith("/buscar-proveedor") and request.method == "GET":
        solicitud_id = path.split("/solicitudes/")[1].rsplit("/buscar-proveedor", 1)[0]
        return handle_buscar_proveedor(request, solicitud_id)
    elif path.startswith("/ofertas/") and path.endswith("/aceptar") and request.method == "POST":
        oferta_id = path.split("/ofertas/")[1].rsplit("/aceptar", 1)[0]
        return handle_aceptar_oferta(request, oferta_id)
    elif path.startswith("/proveedores/") and path.endswith("/ubicacion") and request.method == "POST":
        proveedor_id = path.split("/proveedores/")[1].rsplit("/ubicacion", 1)[0]
        return handle_actualizar_ubicacion_proveedor(request, proveedor_id)
    elif path.startswith("/track/") and path.endswith("/estado") and request.method == "GET":
        tracking_token = path.split("/track/")[1].rsplit("/estado", 1)[0]
        return obtener_estado_tracking(tracking_token)
    elif path.startswith("/track/") and request.method == "GET":
        tracking_token = path.split("/track/")[1]
        return handle_servir_tracking(request, tracking_token)
    elif path == "/proveedor" and request.method == "GET":
        return handle_servir_proveedor(request)
    elif path == "/sctr/whatsapp/resumen" and request.method == "POST":
        return handle_sctr_resumen_whatsapp(request)
    elif path == "/sctr/whatsapp/excel" and request.method == "POST":
        return handle_sctr_excel_whatsapp(request)
    elif path == "/sctr/whatsapp/borrar" and request.method == "POST":
        return handle_sctr_borrar_empleado(request)
    elif path == "/sctr/excel-web":
        return handle_sctr_excel_web(request)
    elif path == "/programa-cronico/reporte" and request.method == "GET":
        return handle_reporte_cronico(request)
    elif path == "/programa-cronico/estadistica" and request.method == "GET":
        return handle_estadistica_cronico(request)
    elif path == "/programa-cronico/actualizar-contacto" and request.method == "POST":
        return handle_actualizar_contacto_cronico(request)
    elif path == "/programa-cronico/mantenimiento" and request.method in ("GET", "POST"):
        return handle_mantenimiento_contacto_cronico(request) 
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
    elif path == "/campania/top-mensajes" and request.method == "GET":
        return handle_top_mensajes_campania(request)
    elif path == "/campania/reporte-envios" and request.method == "GET":
        return handle_reporte_envios_campania(request)
    elif path == "/campania/ver-contactos" and request.method == "GET":
        return handle_ver_contactos_campania(request)
    elif path == "/campania/imagenes" and request.method in ("GET", "POST"):
        return handle_imagenes_campania(request)
    elif path == "/campania/mensajes-por-dia" and request.method == "GET":
        return handle_mensajes_por_dia_campania(request)
    elif path == "/campania/envio-especial":
        return handle_envio_especial_web(request)
    elif path == "/programa-cronico/recibir-dni":
        return handle_recibir_dni_cronico(request)
    elif path == "/programa-cronico/firmar-consentimiento":
        return handle_firmar_consentimiento_cronico(request)
    elif request.path.startswith("/programa-cronico/") and request.method == "GET":
        token = request.path.split("/programa-cronico/")[1]
        return handle_ver_formulario_cronico(request, token)
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
    elif path.startswith("/c/confirm/"):
        code = path.split("/c/confirm/")[1]
        return handle_confirmar_click_campania(code, request)
    elif path.startswith("/c/"):
        code = path.split("/c/")[1]
        return handle_redirect_campania(code, request)
    elif path.startswith("/ps/confirm/"):
        code = path.split("/ps/confirm/")[1]
        return handle_confirmar_click_programa_salud(code, request)
    elif path.startswith("/ps/"):
        code = path.split("/ps/")[1]
        return handle_redirect_programa_salud(code, request)
    elif path.startswith("/static/"):
        filename = path.split("/static/")[1]
        return handle_static_campania(filename)
    elif path.startswith("/assets/"):
        filename = path.split("/assets/")[1]
        return handle_static_assets(filename)
    elif request.path.startswith("/afiliacion/") and request.method == "GET":
        token = request.path.rstrip("/").split("/afiliacion/")[1]
        return handle_ver_wizard_afiliacion(request, token)
    elif path == "/social/mantenimiento" and request.method in ("GET", "POST"):
        return handle_mantenimiento_social(request)
    elif path == "/social/reporte" and request.method == "GET":
        return handle_reporte_social(request)
    elif path.startswith("/s/confirm/"):
        code = path.split("/s/confirm/")[1]
        return handle_confirmar_click_social(code, request)
    elif path.startswith("/s/"):
        code = path.split("/s/")[1]
        return handle_redirect_social(code, request)
    elif path == '/backoffice/login' and request.method == 'POST':
        return bo_login(request)
    elif path == '/backoffice/logout' and request.method == 'POST':
        return bo_logout(request)
    elif path == '/backoffice/solicitar-reset' and request.method == 'POST':
        return bo_solicitar_reset(request)
    elif path == '/backoffice/confirmar-reset' and request.method == 'POST':
        return bo_confirmar_reset(request)
    elif path == '/backoffice/tareas' and request.method == 'POST':
        return bo_crear_tarea(request)
    elif path == '/backoffice/tareas' and request.method == 'GET':
        return bo_listar_tareas(request)
    elif path.startswith('/backoffice/tareas/') and path.endswith('/comentarios') and request.method == 'POST':
        tarea_id = int(path.split('/')[-2])
        return bo_agregar_comentario(request, tarea_id)
    elif path.startswith('/backoffice/tareas/') and request.method == 'GET':
        tarea_id = int(path.split('/')[-1])
        return bo_detalle_tarea(request, tarea_id)
    elif path.startswith('/backoffice/tareas/') and request.method == 'PUT':
        tarea_id = int(path.split('/')[-1])
        return bo_actualizar_tarea(request, tarea_id)
    elif path == '/backoffice/asesores' and request.method == 'POST':
        return bo_crear_asesor(request)
    elif path == '/backoffice/asesores' and request.method == 'GET':
        return bo_listar_asesores(request)
    elif path.startswith('/backoffice/asesores/') and path.endswith('/productividad') and request.method == 'GET':
        asesor_id = int(path.split('/')[-2])
        return bo_productividad_asesor(request, asesor_id)
    elif path.startswith('/backoffice/asesores/') and request.method == 'PUT':
        asesor_id = int(path.split('/')[-1])
        return bo_actualizar_asesor(request, asesor_id)
    elif path == '/backoffice/tablero' and request.method == 'GET':
        with open(os.path.join(BASE_DIR, 'static/backoffice/tablero.html')) as f:
            return f.read(), 200, {'Content-Type': 'text/html'}
    elif path == '/backoffice/admin-asesores' and request.method == 'GET':
        with open(os.path.join(BASE_DIR, 'static/backoffice/admin_asesores.html')) as f:
            return f.read(), 200, {'Content-Type': 'text/html'}

    
    return (json.dumps({"error": f"Ruta no encontrada: {path}"}), 404, HEADERS)

# ─────────────────────────────────────────────
# UTILIDAD: oculta texto dejando solo primera letra
# ─────────────────────────────────────────────
def _celulares_coinciden(celular_a, celular_b):
    """
    Compara dos celulares peruanos ignorando diferencias de prefijo de país
    (+51 / 51) u otros caracteres. Se comparan los últimos 9 dígitos, que
    son el número de celular real (ej. 997855645 == 51997855645).
    """
    a = re.sub(r"[^\d]", "", celular_a or "")
    b = re.sub(r"[^\d]", "", celular_b or "")
    if not a or not b:
        return False
    return a[-9:] == b[-9:]

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
        if not _celulares_coinciden(celular_consulta, nro_contacto1) and not _celulares_coinciden(celular_consulta, nro_contacto2):
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
            lima  = pytz.timezone("America/Lima")
            fecha_creacion = datetime.now(lima)

            cur.execute("""
                INSERT INTO sofi_usuarios (
                    doi, nombres, apellidos, direccion1, distrito1,
                    nro_contacto1, nro_contacto2, email, escoge, fecha_creacion
                ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,'1',%s)
            """, (
                doi, nombres, apellidos, direccion1, distrito1,
                nro_contacto1, nro_contacto2, email, fecha_creacion
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
        if not _celulares_coinciden(celular_consulta, nro_contacto1) and not _celulares_coinciden(celular_consulta, nro_contacto2):
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
        cur.execute("SELECT DISTINCT ON (celular) id, nombre, celular FROM sofi_leads_medicamentos WHERE enviado = 'N' ORDER BY celular, id")
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
        cur.execute("SELECT DISTINCT ON (celular) id, nombre, celular FROM sofi_leads_consultas WHERE enviado = 'N' ORDER BY celular, id")
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
def safe_read_csv(url, tries=3, delay=2):
    import time
    import requests

    for intento in range(1, tries + 1):
        try:
            print(f"[CSV] Intento {intento}/{tries}")
            print(f"[CSV] URL: {url}")

            response = requests.get(
                url,
                timeout=(10, 60),  # 10s conexión / 60s lectura
                allow_redirects=True,
                headers={
                    "User-Agent": "Mozilla/5.0"
                }
            )

            print(f"[CSV] HTTP Status: {response.status_code}")
            print(f"[CSV] Content-Type: {response.headers.get('content-type')}")
            print(f"[CSV] Tamaño: {len(response.content)} bytes")
            print(f"[CSV] URL final: {response.url}")

            if response.status_code == 200:
                response.encoding = "utf-8"

                lines = response.text.splitlines()

                print(f"[CSV] Líneas recibidas: {len(lines)}")

                return lines

            print(
                f"[CSV] Google respondió HTTP {response.status_code}"
            )

        except requests.exceptions.ConnectTimeout as e:
            print(f"[CSV] Timeout de conexión: {e}")

        except requests.exceptions.ReadTimeout as e:
            print(f"[CSV] Timeout esperando respuesta: {e}")

        except requests.exceptions.ConnectionError as e:
            print(f"[CSV] Error de conexión: {e}")

        except requests.exceptions.RequestException as e:
            print(f"[CSV] Error de requests: {e}")

        except Exception as e:
            print(f"[CSV] Error inesperado: {type(e).__name__}: {e}")

        if intento < tries:
            print(f"[CSV] Esperando {delay} segundos...")
            time.sleep(delay)

    print("[CSV] No se pudo obtener el CSV después de todos los intentos.")

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
    cur.execute(f"SELECT DISTINCT ON (celular) id, nombre, celular FROM {tabla} WHERE enviado='N'")
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
    lima    = pytz.timezone("America/Lima")
    ahora   = datetime.now(lima)
    hoy_str = ahora.strftime("%Y-%m-%d %H:%M:%S")

    # Si es lunes, tomar anteayer; si no, tomar ayer (misma regla que antes)
    dias_atras      = 2 if ahora.isoweekday() == 1 else 1
    fecha_objetivo  = (ahora - timedelta(days=dias_atras)).date()

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("""
        SELECT sol, nombre, celular, fecha_rechazo
        FROM sofi_rechazados_raw
        WHERE fecha_rechazo = %s
    """, (fecha_objetivo,))
    filas = cur.fetchall()

    contador             = 0
    contador_con_celular = 0
    rows_to_insert       = []

    for sol, nombre, celular, fecha_rechazo in filas:
        rows_to_insert.append((
            sol,
            fecha_rechazo.strftime("%Y-%m-%d"),
            celular,
            nombre,
            hoy_str,
            "N",
            None,
            "N"
        ))

        if celular and len(celular) == 9 and celular[0] == "9" and celular != "999999999":
            nombre_corto = nombre[:64] if nombre and len(nombre) > 64 else nombre
            envia_sms_cartas_rechazadas(celular, nombre_corto)
            contador_con_celular += 1

        contador += 1

    if rows_to_insert:
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
# RETENCION - Header + navegación reutilizable
# ─────────────────────────────────────────────────────────────────
def _layout_campania(pw, titulo, subtitulo, contenido_html, activo=""):
    def link(href, texto, key):
        clase = "activo" if key == activo else ""
        return f'<a href="{href}?pw={pw}" class="{clase}">{texto}</a>'

    nav_html = f"""
    <nav class="tabs">
      {link("/campania/mantenimiento", "🏠 Mantenimiento", "mantenimiento")}
      {link("/campania/cargar-excel-web", "📤 Cargar Excel", "cargar")}
      {link("/campania/ver-contactos", "🔍 Contactos", "contactos")}
      {link("/campania/generar-sms-web", "✉️ Enviar SMS", "enviar")}
      {link("/campania/imagenes", "🖼️ Imágenes", "imagenes")}
      {link("/campania/reporte-clicks", "📊 Clicks", "retencion")}
      {link("/campania/reporte-envios", "📅 Envíos", "envios")}
      {link("/campania/mensajes-por-dia", "📆 Por día", "por-dia")}
      {link("/campania/top-mensajes", "🏆 Top mensajes", "top")}
      {link("/campania/envio-especial", "✈️ Envío Especial", "especial")}
    </nav>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="noindex, nofollow">
  <title>{titulo} - Sanitas</title>
  <link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
  <style>{CAMPANIA_CSS}</style>
</head>
<body>
  <div class="topbar">
    <h1>{titulo}</h1>
    <div class="subtitulo">{subtitulo}</div>
  </div>
  <div class="contenedor">
    {nav_html}
    {contenido_html}
  </div>
</body>
</html>"""   

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
            filas.append((
                str(data.get("nombre") or "").strip(),
                telefono1,
                tipo_cliente
            ))
            
    wb.close()

    if not filas:
        return jsonify({"error": "no se encontraron filas válidas"}), 400

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            execute_values(cur, """
                INSERT INTO campania_contactos
                    (nombre, telefono1, tipo_cliente)
                VALUES %s
                ON CONFLICT (telefono1, tipo_cliente) DO UPDATE SET
                    nombre = EXCLUDED.nombre,
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
    max_intentos_val = 3

    if request.method == "POST":
        tipo_val = (request.form.get("tipo_cliente") or "").strip().upper()
        campana_val = (request.form.get("campana") or "").strip()
        excluir_convertidos = request.form.get("excluir_convertidos") == "on"
        try:
            max_intentos_val = int(request.form.get("max_intentos") or 3)
        except ValueError:
            max_intentos_val = 3

        if not tipo_val or not campana_val:
            resultado_html = '<p class="msg-error">Selecciona tipo_cliente y escribe un nombre de campaña.</p>'
        else:
            try:
                resultados, omitidos = _generar_y_enviar_campania(
                    tipo_val, campana_val, excluir_convertidos, max_intentos_val
                )
            except ValueError as e:
                resultados = None
                omitidos = 0
                resultado_html = f'<p class="msg-error">{e}</p>'

            if resultados is not None:
                aviso_omitidos = f'<p class="msg-aviso">⏭️ {omitidos} contactos omitidos por haber alcanzado el tope de {max_intentos_val} intentos.</p>' if omitidos else ""
                if not resultados:
                    resultado_html = f'<p class="msg-aviso">No hay contactos pendientes de envío para ese tipo_cliente.</p>{aviso_omitidos}'
                else:
                    exitosos = sum(1 for _, _, ok, _ in resultados if ok)
                    fallidos = len(resultados) - exitosos

                    solo_errores = [r for r in resultados if not r[2]]
                    if solo_errores:
                        filas_tabla = "".join(
                            f"""<tr>
                                <td>{cel}</td><td>{nom or ''}</td>
                                <td><span class="badge badge-pendiente" style="background:rgba(214,69,69,0.12);color:var(--sanitas-danger);">❌ Error</span></td>
                                <td>{err or ''}</td>
                            </tr>"""
                            for cel, nom, ok, err in solo_errores
                        )
                        tabla_html = f"""
                        <p style="margin-top:16px;font-weight:600;">Detalle de {len(solo_errores)} errores:</p>
                        <div class="table-wrap">
                          <table>
                          <tr><th>Celular</th><th>Nombre</th><th>Estado</th><th>Detalle error</th></tr>
                          {filas_tabla}
                          </table>
                        </div>"""
                    else:
                        tabla_html = ""

                    resultado_html = f"""
                    <p class="msg-exito">✅ {exitosos} enviados correctamente{f', ❌ {fallidos} con error' if fallidos else ''} — tipo <b>{tipo_val}</b> / campaña <b>{campana_val}</b>.</p>
                    {aviso_omitidos}
                    {tabla_html}
                    <p style="margin-top:14px;"><a href="/campania/reporte-clicks?pw={pw}">Ver reporte de retención →</a> |
                       <a href="/campania/mensajes-por-dia?pw={pw}">Ver mensajes por día →</a></p>
                    """

    contenido = f"""
    <div class="card">
      <form method="POST" action="/campania/generar-sms-web?pw={pw}"
            onsubmit="return confirm('Esto enviará SMS reales por Danaconnect a todos los contactos de este tipo_cliente. ¿Continuar?');">
        <label>Tipo de cliente</label>
        <select name="tipo_cliente" required>
          <option value="">-- selecciona --</option>
          {opciones_html}
        </select>
        <label>Nombre de campaña (identifica este envío)</label>
        <input type="text" name="campana" placeholder="ej. retencion_jul2026" value="{campana_val}" required>
        <label style="display:flex;align-items:center;gap:8px;text-transform:none;font-weight:normal;margin-top:14px;">
          <input type="checkbox" name="excluir_convertidos" checked style="width:auto;">
          No reenviar a quienes ya hicieron clic (recomendado)
        </label>
        <label>Máximo de intentos por cliente</label>
        <input type="text" name="max_intentos" value="{max_intentos_val}" style="max-width:100px;">
        <button type="submit">Generar y enviar SMS</button>
      </form>
    </div>
    <div class="card">
      {resultado_html if resultado_html else '<p style="color:var(--sanitas-gray);margin:0;">Los resultados del envío aparecerán aquí.</p>'}
    </div>
    """

    return _layout_campania(pw, "Generar y enviar SMS", "Envío manual vía Danaconnect", contenido, activo="enviar")

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
    IMAGEN_OG = os.environ.get("IMAGEN_OG")
    
    
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
                    raw_hits = raw_hits + 1,
                    ip_address = %s,
                    user_agent = %s,
                    entregado = TRUE,
                    entregado_at = COALESCE(entregado_at, NOW())
                WHERE code = %s
            """, (request.remote_addr, request.headers.get("User-Agent", ""), code))
            conn.commit()

            cur.execute("""
                SELECT og_title, og_description, og_image
                FROM campania_plantillas WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            og = cur.fetchone()
    finally:
        conn.close()

    og_title, og_description, og_image = og if og else ("Sanitas Perú", "Contáctanos por WhatsApp", "sanitas-general.jpg")
    
    if og_image and og_image.startswith("http"):
        img_url = og_image
    else:
        img_url = f"{IMAGEN_OG}{og_image}?v={IMG_VERSION}"

    url_canonica = f"{base_url}/c/{code}"
    url_destino = f"https://wa.me/{whatsapp_destino}"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta property="og:type" content="website">
  <meta property="og:url" content="{url_canonica}">
  <meta property="og:site_name" content="Sanitas Perú">
  <meta property="og:title" content="{og_title}">
  <meta property="og:description" content="{og_description}">
  <meta property="og:image" content="{IMAGEN_OG}{og_image}?v={IMG_VERSION}">
  <meta property="og:image:width" content="1200">
  <meta property="og:image:height" content="630">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{og_title}">
  <meta name="twitter:description" content="{og_description}">
  <meta name="twitter:image" content="{IMAGEN_OG}{og_image}?v={IMG_VERSION}">
  <title>{og_title}</title>
</head>
<body>
  <script>
    if (navigator.sendBeacon) {{
      navigator.sendBeacon("/c/confirm/{code}");
    }} else {{
      fetch("/c/confirm/{code}", {{ method: "POST", keepalive: true }});
    }}
    window.location.href = "{url_destino}";
  </script>
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
                for n in range(2, 6):
                    texto = (request.form.get(f"mensaje_intento_{n}") or "").strip()
                    if texto:
                        cur.execute("""
                            INSERT INTO campania_plantillas_intentos (tipo_cliente, intento, mensaje, actualizado_at)
                            VALUES (%s, %s, %s, NOW())
                            ON CONFLICT (tipo_cliente, intento) DO UPDATE SET
                                mensaje = EXCLUDED.mensaje, actualizado_at = NOW()
                        """, (tipo_cliente, n, texto))
                    else:
                        cur.execute("""
                            DELETE FROM campania_plantillas_intentos
                            WHERE tipo_cliente = %s AND intento = %s
                        """, (tipo_cliente, n))
            conn.commit()

        with conn.cursor() as cur:
            cur.execute("""
                SELECT tipo_cliente, mensaje, og_title, og_description, og_image, whatsapp_destino
                FROM campania_plantillas ORDER BY tipo_cliente
            """)
            plantillas = cur.fetchall()

        tarjetas_html = ""
        for tc, mensaje, og_title, og_desc, og_img, wa in plantillas:
            with conn.cursor() as cur_v:
                cur_v.execute("""
                    SELECT intento, mensaje FROM campania_plantillas_intentos
                    WHERE tipo_cliente = %s ORDER BY intento
                """, (tc,))
                variantes_dict = dict(cur_v.fetchall())

            variantes_html = ""
            for n in range(2, 6):
                texto_variante = variantes_dict.get(n, "")
                variantes_html += f"""
                <label>Mensaje intento #{n} (vacío = usa el mensaje base)</label>
                <textarea name="mensaje_intento_{n}" rows="2">{texto_variante}</textarea>"""

            tarjetas_html += f"""
            <div class="card">
              <form method="POST" action="/campania/mantenimiento?pw={pw}">
                <input type="hidden" name="tipo_cliente" value="{tc}">
                <h3>{tc}</h3>
                <label>Mensaje base (intento #1)</label>
                <textarea name="mensaje" rows="3">{mensaje}</textarea>
                {variantes_html}
                <label>OG Title</label>
                <input type="text" name="og_title" value="{og_title}">
                <label>OG Description</label>
                <input type="text" name="og_description" value="{og_desc}">
                <label>OG Image (URL del bucket)</label>
                <input type="text" name="og_image" value="{og_img}">
                <label>WhatsApp destino</label>
                <input type="text" name="whatsapp_destino" value="{wa}">
                <button type="submit">Guardar {tc}</button>
              </form>
            </div>"""
    finally:
        conn.close()

    contenido = f"""
    <div class="card">
      <p style="color:var(--sanitas-gray);margin:0;">Edita el mensaje base y las variantes por intento de cada tipo de cliente.</p>
    </div>
    {tarjetas_html}
    """

    return _layout_campania(pw, "Plantillas de campaña", "Panel de mantenimiento RETENCIONES", contenido, activo="mantenimiento")

@con_log_errores
def handle_reporte_clicks_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT
                    c.telefono1, c.nombre, c.tipo_cliente,
                    c.convertido, c.fecha_conversion,
                    COUNT(l.id) AS total_intentos,
                    MIN(l.creado_at) AS primer_envio,
                    MAX(l.creado_at) AS ultimo_envio,
                    BOOL_OR(l.entregado) AS entregado,
                    MIN(l.entregado_at) AS primera_entrega
                FROM campania_contactos c
                LEFT JOIN campania_links l
                    ON l.celular = c.telefono1 AND l.tipo_cliente = c.tipo_cliente
                GROUP BY c.telefono1, c.nombre, c.tipo_cliente, c.convertido, c.fecha_conversion
                ORDER BY c.convertido ASC, entregado ASC, total_intentos DESC
            """)
            rows = cur.fetchall()
    finally:
        conn.close()

    filas_html = ""
    total_convertidos = 0
    total_entregados = 0
    for telefono, nombre, tipo_cliente, convertido, fecha_conv, intentos, primer_envio, ultimo_envio, entregado, primera_entrega in rows:
        if convertido:
            total_convertidos += 1
        if entregado:
            total_entregados += 1

        if convertido:
            badge = '<span class="badge badge-ok">✅ Convertido</span>'
        elif entregado:
            badge = '<span class="badge" style="background:rgba(0,169,183,0.12);color:var(--sanitas-teal-dark);">📩 Entregado</span>'
        elif intentos:
            badge = '<span class="badge badge-pendiente">⏳ Sin confirmar</span>'
        else:
            badge = '<span class="badge badge-neutro">— Sin enviar</span>'

        filas_html += f"""<tr>
            <td>{telefono}</td><td>{nombre or ''}</td><td>{tipo_cliente}</td>
            <td>{intentos or 0}</td>
            <td>{primer_envio.strftime('%Y-%m-%d %H:%M') if primer_envio else ''}</td>
            <td>{ultimo_envio.strftime('%Y-%m-%d %H:%M') if ultimo_envio else ''}</td>
            <td>{primera_entrega.strftime('%Y-%m-%d %H:%M') if primera_entrega else ''}</td>
            <td>{badge}</td>
            <td>{fecha_conv.strftime('%Y-%m-%d %H:%M') if fecha_conv else ''}</td>
        </tr>"""

    contenido = f"""
    <div class="card" style="display:flex;gap:30px;flex-wrap:wrap;">
      <div><div style="font-size:12px;color:var(--sanitas-gray);">✅ Convertidos</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-success);">{total_convertidos} / {len(rows)}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">📩 Entregados sin abrir</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-teal-dark);">{total_entregados - total_convertidos}</div></div>
    </div>
    <div class="table-wrap">
      <table>
      <tr><th>Teléfono</th><th>Nombre</th><th>Tipo</th><th>Intentos</th><th>1er envío</th><th>Último envío</th><th>1ra entrega</th><th>Estado</th><th>Fecha conversión</th></tr>
      {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, "Reporte de retención", "Estado de conversión por contacto", contenido, activo="retencion")

def handle_static_campania(filename):
    return send_from_directory("static", filename)

def handle_static_assets(filename):
    return send_from_directory("assets", filename)

@con_log_errores
def handle_cargar_excel_web(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    mensaje = ""
    resumen = ""

    if request.method == "POST":
        if "archivo" not in request.files or request.files["archivo"].filename == "":
            mensaje = '<p class="msg-error">Selecciona un archivo Excel.</p>'
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
                    filas.append((
                        str(data.get("nombre") or "").strip(),
                        telefono1,
                        tipo_cliente
                    ))
                wb.close()

                if not filas:
                    mensaje = '<p class="msg-error">No se encontraron filas válidas (revisa columnas: nombre, telefono1, tipo_cliente).</p>'
                else:
                    conn = get_connection()
                    try:
                        with conn.cursor() as cur:
                            execute_values(cur, """
                                INSERT INTO campania_contactos
                                    (nombre, telefono1, tipo_cliente)
                                VALUES %s
                                ON CONFLICT (telefono1, tipo_cliente) DO UPDATE SET
                                    nombre = EXCLUDED.nombre,
                                    actualizado_at = NOW()
                            """, filas)
                        conn.commit()
                    finally:
                        conn.close()

                    mensaje = f'<p class="msg-exito">✅ {len(filas)} contactos insertados/actualizados correctamente.</p>'

                    with get_connection() as conn2:
                        with conn2.cursor() as cur2:
                            cur2.execute("SELECT tipo_cliente, COUNT(*) FROM campania_contactos GROUP BY tipo_cliente ORDER BY tipo_cliente")
                            resumen_rows = cur2.fetchall()
                    filas_resumen = "".join(f"<li><b>{tc}</b>: {cnt} contactos</li>" for tc, cnt in resumen_rows)
                    resumen = f"<ul style='margin-top:10px;'>{filas_resumen}</ul>"

            except Exception as e:
                mensaje = f'<p class="msg-error">Error al procesar el archivo: {e}</p>'

    contenido = f"""
    <div class="card">
      <p style="color:var(--sanitas-gray);margin-top:0;">Columnas esperadas en el Excel: <b>nombre, telefono1, tipo_cliente</b></p>
      <form method="POST" action="/campania/cargar-excel-web?pw={pw}" enctype="multipart/form-data">
        <label>Archivo Excel (.xlsx)</label>
        <input type="file" name="archivo" accept=".xlsx,.xls" required>
        <button type="submit">Cargar archivo</button>
      </form>
      {mensaje}
      {resumen}
    </div>
    """

    return _layout_campania(pw, "Cargar Excel", "Sube tu base de contactos de campaña", contenido, activo="cargar")

def _limpiar_nombre_sms(nombre):
    return re.sub(
        r'[\U0001F300-\U0001F9FF\U00002600-\U000026FF'
        r'\U00002700-\U000027BF\U0001F1E6-\U0001F1FF]',
        '', nombre or '', flags=re.UNICODE
    )

def _enviar_sms_campania_uno(numero, nombre, mensaje_tpl, link_corto, code):
    nombre_limpio = _limpiar_nombre_sms(nombre)
    mensaje = mensaje_tpl.replace("{Nombre}", nombre_limpio)

    # ── Validación: SMS no debe exceder 160 caracteres ──
    if len(mensaje) > 160:
        return False, f"Mensaje excede 160 caracteres ({len(mensaje)}): {mensaje[:50]}..."

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

from concurrent.futures import ThreadPoolExecutor, as_completed
MAX_HILOS_ENVIO_SMS = 15  # cuántos SMS se mandan en paralelo; ajusta según límite de Danaconnect

def _generar_y_enviar_campania(tipo_cliente, campana, excluir_convertidos=True, max_intentos=3):
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
            mensaje_base, whatsapp_destino = plantilla

            cur.execute("""
                SELECT intento, mensaje FROM campania_plantillas_intentos
                WHERE tipo_cliente = %s
            """, (tipo_cliente,))
            variantes = {intento: mensaje for intento, mensaje in cur.fetchall()}

            query_contactos = """
                SELECT nombre, telefono1 FROM campania_contactos
                WHERE tipo_cliente = %s
            """
            params = [tipo_cliente]
            if excluir_convertidos:
                query_contactos += " AND COALESCE(convertido, FALSE) = FALSE"
            cur.execute(query_contactos, params)
            contactos = cur.fetchall()

            registros_insert = []
            codes_por_contacto = []
            omitidos_por_tope = 0

            for nombre, telefono1 in contactos:
                cur.execute("""
                    SELECT COUNT(*) FROM campania_links
                    WHERE celular = %s AND tipo_cliente = %s
                """, (telefono1, tipo_cliente))
                intentos_previos = cur.fetchone()[0]

                if intentos_previos >= max_intentos:
                    omitidos_por_tope += 1
                    continue

                intento = intentos_previos + 1
                mensaje_para_este_envio = variantes.get(intento, mensaje_base)

                code = _generar_code_campania(telefono1)
                registros_insert.append((code, telefono1, nombre, tipo_cliente, campana, whatsapp_destino, intento, mensaje_para_este_envio))
                codes_por_contacto.append((telefono1, nombre, code, mensaje_para_este_envio))

            if registros_insert:
                execute_values(cur, """
                    INSERT INTO campania_links
                        (code, celular, nombre, tipo_cliente, campana, whatsapp_destino, intento, mensaje_enviado)
                    VALUES %s
                """, registros_insert)
            conn.commit()
    finally:
        conn.close()

    # NUEVO: envío en paralelo con hilos, en vez de uno por uno
    resultados = []
    with ThreadPoolExecutor(max_workers=MAX_HILOS_ENVIO_SMS) as executor:
        futuros = {
            executor.submit(_enviar_sms_campania_uno, telefono1, nombre, mensaje_tpl, f"{base_url}/c/{code}", code): (telefono1, nombre, code)
            for telefono1, nombre, code, mensaje_tpl in codes_por_contacto
        }

        for futuro in as_completed(futuros):
            telefono1, nombre, code = futuros[futuro]
            try:
                ok, error = futuro.result()
            except Exception as e:
                ok, error = False, str(e)

            if ok:
                aviso_descripcion = "retencion"
                sms_aviso(telefono1, aviso_descripcion)

            else:
                conn2 = get_connection()
                try:
                    with conn2.cursor() as cur2:
                        cur2.execute(
                            "UPDATE campania_links SET error_envio = %s WHERE code = %s",
                            (error, code)
                        )
                    conn2.commit()
                finally:
                    conn2.close()

            resultados.append((telefono1, nombre, ok, error))

    return resultados, omitidos_por_tope

@con_log_errores
def handle_top_mensajes_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    try:
        minimo_envios = int(request.args.get("min", 1))
    except ValueError:
        minimo_envios = 1

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT tipo_cliente, intento, mensaje_enviado,
                       COUNT(*) AS enviados,
                       SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END) AS abiertos
                FROM campania_links
                WHERE mensaje_enviado IS NOT NULL
                GROUP BY tipo_cliente, intento, mensaje_enviado
                HAVING COUNT(*) >= %s
                ORDER BY (SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END)::float / COUNT(*)) DESC
                LIMIT 10
            """, (minimo_envios,))
            rows = cur.fetchall()
    finally:
        conn.close()

    filas_html = ""
    for tipo_cliente, intento, mensaje, enviados, abiertos in rows:
        tasa = round(100.0 * abiertos / enviados, 1) if enviados else 0
        color_tasa = "var(--sanitas-success)" if tasa >= 30 else ("var(--sanitas-warning)" if tasa >= 10 else "var(--sanitas-gray)")
        filas_html += f"""<tr>
            <td>{tipo_cliente}</td><td>#{intento}</td>
            <td style="max-width:380px;">{mensaje}</td>
            <td>{enviados}</td><td>{abiertos}</td>
            <td style="color:{color_tasa};font-weight:700;">{tasa}%</td>
        </tr>"""

    if not rows:
        filas_html = '<tr><td colspan="6" style="text-align:center;color:var(--sanitas-gray);padding:20px;">Sin datos suficientes todavía (o baja el mínimo de envíos con el filtro de arriba).</td></tr>'

    contenido = f"""
    <div class="card">
      <form method="GET" action="/campania/top-mensajes" class="filtros">
        <input type="hidden" name="pw" value="{pw}">
        <div><label>Mínimo de envíos para calificar</label><input type="text" name="min" value="{minimo_envios}"></div>
        <button type="submit">Aplicar</button>
      </form>
    </div>
    <div class="table-wrap">
      <table>
      <tr><th>Tipo</th><th>Intento</th><th>Mensaje</th><th>Enviados</th><th>Abiertos</th><th>Tasa</th></tr>
      {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, "Top mensajes", "Ranking por tasa de apertura", contenido, activo="top")

@con_log_errores
def handle_confirmar_click_campania(code, request):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT celular, tipo_cliente FROM campania_links WHERE code = %s", (code,))
            row = cur.fetchone()
            if not row:
                return ("", 204)
            celular, tipo_cliente = row

            cur.execute("""
                UPDATE campania_links SET
                    clicked_at = COALESCE(clicked_at, NOW()),
                    click_count = click_count + 1
                WHERE code = %s
            """, (code,))

            cur.execute("""
                UPDATE campania_contactos SET
                    convertido = TRUE,
                    fecha_conversion = COALESCE(fecha_conversion, NOW())
                WHERE telefono1 = %s AND tipo_cliente = %s
            """, (celular, tipo_cliente))

            conn.commit()
    finally:
        conn.close()

    return ("", 204)

@con_log_errores
def handle_reporte_envios_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    fecha_inicio = request.args.get("fecha_inicio", "")
    fecha_fin = request.args.get("fecha_fin", "")
    tipo_filtro = request.args.get("tipo_cliente", "")
    campana_filtro = request.args.get("campana", "")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT tipo_cliente FROM campania_plantillas ORDER BY tipo_cliente")
            tipos = [r[0] for r in cur.fetchall()]

            cur.execute("SELECT DISTINCT campana FROM campania_links WHERE campana IS NOT NULL AND campana != '' ORDER BY campana")
            campanas = [r[0] for r in cur.fetchall()]

            where = ["1=1"]
            params = []
            if fecha_inicio:
                where.append("creado_at >= %s")
                params.append(fecha_inicio)
            if fecha_fin:
                where.append("creado_at < (%s::date + INTERVAL '1 day')")
                params.append(fecha_fin)
            if tipo_filtro:
                where.append("tipo_cliente = %s")
                params.append(tipo_filtro)
            if campana_filtro:
                where.append("campana = %s")
                params.append(campana_filtro)
            where_sql = " AND ".join(where)

            cur.execute(f"""
                SELECT
                    COUNT(*) AS total_enviados,
                    SUM(CASE WHEN entregado THEN 1 ELSE 0 END) AS total_entregados,
                    SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END) AS total_abiertos,
                    SUM(CASE WHEN error_envio IS NOT NULL THEN 1 ELSE 0 END) AS total_errores
                FROM campania_links
                WHERE {where_sql}
            """, params)
            total_env, total_ent, total_abi, total_err = cur.fetchone()

            cur.execute(f"""
                SELECT
                    creado_at::date AS fecha,
                    tipo_cliente,
                    campana,
                    COUNT(*) AS enviados,
                    SUM(CASE WHEN entregado THEN 1 ELSE 0 END) AS entregados,
                    SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END) AS abiertos,
                    SUM(CASE WHEN error_envio IS NOT NULL THEN 1 ELSE 0 END) AS errores
                FROM campania_links
                WHERE {where_sql}
                GROUP BY creado_at::date, tipo_cliente, campana
                ORDER BY fecha DESC, tipo_cliente
            """, params)
            filas = cur.fetchall()
    finally:
        conn.close()

    opciones_tipo = "".join(
        f'<option value="{t}" {"selected" if t == tipo_filtro else ""}>{t}</option>' for t in tipos
    )
    opciones_campana = "".join(
        f'<option value="{c}" {"selected" if c == campana_filtro else ""}>{c}</option>' for c in campanas
    )

    filas_html = ""
    for fecha, tipo_cliente, campana, enviados, entregados, abiertos, errores in filas:
        tasa_apertura = round(100.0 * (abiertos or 0) / enviados, 1) if enviados else 0
        color_err = 'style="color:var(--sanitas-danger);font-weight:600;"' if errores else ""
        filas_html += f"""<tr>
            <td>{fecha}</td><td>{tipo_cliente}</td><td>{campana or ''}</td>
            <td>{enviados}</td><td>{entregados or 0}</td>
            <td>{abiertos or 0} ({tasa_apertura}%)</td>
            <td {color_err}>{errores or 0}</td>
        </tr>"""

    contenido = f"""
    <div class="card">
      <form method="GET" action="/campania/reporte-envios" class="filtros">
        <input type="hidden" name="pw" value="{pw}">
        <div><label>Desde</label><input type="date" name="fecha_inicio" value="{fecha_inicio}"></div>
        <div><label>Hasta</label><input type="date" name="fecha_fin" value="{fecha_fin}"></div>
        <div><label>Tipo cliente</label>
          <select name="tipo_cliente"><option value="">Todos</option>{opciones_tipo}</select>
        </div>
        <div><label>Campaña</label>
          <select name="campana"><option value="">Todas</option>{opciones_campana}</select>
        </div>
        <button type="submit">Filtrar</button>
      </form>
    </div>

    <div class="card" style="display:flex;gap:30px;flex-wrap:wrap;">
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Total enviados</div><div style="font-size:24px;font-weight:700;">{total_env or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Entregados</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-teal-dark);">{total_ent or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Abiertos</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-success);">{total_abi or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Errores</div><div style="font-size:24px;font-weight:700;color:{'var(--sanitas-danger)' if total_err else 'inherit'};">{total_err or 0}</div></div>
    </div>

    <div class="table-wrap">
      <table>
      <tr><th>Fecha</th><th>Tipo</th><th>Campaña</th><th>Enviados</th><th>Entregados</th><th>Abiertos</th><th>Errores</th></tr>
      {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, "Reporte de envíos", "Filtra por fecha, tipo y campaña", contenido, activo="envios")

@con_log_errores
def handle_ver_contactos_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    fecha_inicio = request.args.get("fecha_inicio", "")
    fecha_fin = request.args.get("fecha_fin", "")
    celular_busqueda = request.args.get("celular", "").strip()
    tipo_filtro = request.args.get("tipo_cliente", "")
    convertido_filtro = request.args.get("convertido", "")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT tipo_cliente FROM campania_contactos ORDER BY tipo_cliente")
            tipos = [r[0] for r in cur.fetchall()]

            where = ["1=1"]
            params = []
            if fecha_inicio:
                where.append("creado_at::date >= %s")
                params.append(fecha_inicio)
            if fecha_fin:
                where.append("creado_at::date <= %s")
                params.append(fecha_fin)
            if celular_busqueda:
                where.append("telefono1 ILIKE %s")
                params.append(f"%{celular_busqueda}%")
            if tipo_filtro:
                where.append("tipo_cliente = %s")
                params.append(tipo_filtro)
            if convertido_filtro == "si":
                where.append("convertido = TRUE")
            elif convertido_filtro == "no":
                where.append("(convertido = FALSE OR convertido IS NULL)")
            where_sql = " AND ".join(where)

            cur.execute(f"""
                SELECT nombre, telefono1, tipo_cliente, creado_at, convertido
                FROM campania_contactos
                WHERE {where_sql}
                ORDER BY creado_at DESC, nombre
                LIMIT 300
            """, params)
            filas = cur.fetchall()

            cur.execute(f"SELECT COUNT(*) FROM campania_contactos WHERE {where_sql}", params)
            total = cur.fetchone()[0]
    finally:
        conn.close()

    opciones_tipo = "".join(
        f'<option value="{t}" {"selected" if t == tipo_filtro else ""}>{t}</option>' for t in tipos
    )

    opciones_convertido = "".join(
        f'<option value="{v}" {"selected" if v == convertido_filtro else ""}>{label}</option>'
        for v, label in [("si", "Sí"), ("no", "No")]
    )

    filas_html = ""
    for nombre, tel1, tipo_cliente, creado_at, convertido in filas:
        fecha_mostrada = creado_at.strftime('%Y-%m-%d') if creado_at else ''
        badge = '<span class="badge badge-ok">✅ Sí</span>' if convertido else '<span class="badge badge-neutro">—</span>'
        filas_html += f"""<tr>
            <td>{nombre or ''}</td><td>{tel1}</td>
            <td>{tipo_cliente}</td><td>{fecha_mostrada}</td>
            <td>{badge}</td>
        </tr>"""

    aviso_limite = f'<p class="msg-aviso">Mostrando primeros 300 de {total} resultados. Afina los filtros para ver menos.</p>' if total > 300 else ""

    contenido = f"""
    <div class="card">
      <form method="GET" action="/campania/ver-contactos" class="filtros">
        <input type="hidden" name="pw" value="{pw}">
        <div><label>Fecha carga desde</label><input type="date" name="fecha_inicio" value="{fecha_inicio}"></div>
        <div><label>Fecha carga hasta</label><input type="date" name="fecha_fin" value="{fecha_fin}"></div>
        <div><label>Celular contiene</label><input type="text" name="celular" value="{celular_busqueda}" placeholder="ej. 9878"></div>
        <div><label>Tipo cliente</label>
          <select name="tipo_cliente"><option value="">Todos</option>{opciones_tipo}</select>
        </div>
        <div><label>Convertido</label>
          <select name="convertido"><option value="">Todos</option>{opciones_convertido}</select>
        </div>
        <button type="submit">Buscar</button>
      </form>
    </div>
    {aviso_limite}
    <div class="table-wrap">
      <table>
        <tr><th>Nombre</th><th>Tel 1</th><th>Tipo</th><th>Fecha carga</th><th>Convertido</th></tr>
        {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, f"Contactos ({total})", "Búsqueda y filtros de la base cargada", contenido, activo="contactos")

from google.cloud import storage

BUCKET_CAMPANIA_IMAGENES = os.environ.get("BUCKET_ARCHIVOS", "sanitas-campania-imagenes")
PREFIX_CAMPANIA_IMAGENES = "campania-static/"

def _storage_client_campania():
    return storage.Client()

def _listar_imagenes_campania():
    client = _storage_client_campania()
    bucket = client.bucket(BUCKET_CAMPANIA_IMAGENES)
    blobs = bucket.list_blobs(prefix=PREFIX_CAMPANIA_IMAGENES)
    resultado = []
    for blob in blobs:
        if blob.name == PREFIX_CAMPANIA_IMAGENES:
            continue
        resultado.append({
            "nombre": blob.name.replace(PREFIX_CAMPANIA_IMAGENES, ""),
            "url": blob.public_url,
            "tamano_kb": round(blob.size / 1024, 1) if blob.size else 0,
        })
    return resultado

def _subir_imagen_campania(archivo_flask):
    client = _storage_client_campania()
    bucket = client.bucket(BUCKET_CAMPANIA_IMAGENES)
    blob = bucket.blob(f"{PREFIX_CAMPANIA_IMAGENES}{archivo_flask.filename}")
    blob.upload_from_file(archivo_flask.stream, content_type=archivo_flask.content_type)
    return blob.public_url

def _eliminar_imagen_campania(nombre_archivo):
    client = _storage_client_campania()
    bucket = client.bucket(BUCKET_CAMPANIA_IMAGENES)
    bucket.blob(f"{PREFIX_CAMPANIA_IMAGENES}{nombre_archivo}").delete()

@con_log_errores
def handle_imagenes_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    mensaje = ""

    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "subir":
            if "archivo" not in request.files or request.files["archivo"].filename == "":
                mensaje = '<p class="msg-error">Selecciona una imagen.</p>'
            else:
                try:
                    url = _subir_imagen_campania(request.files["archivo"])
                    mensaje = f'<p class="msg-exito">✅ Imagen subida. URL pública: <code>{url}</code></p>'
                except Exception as e:
                    mensaje = f'<p class="msg-error">Error al subir: {e}</p>'
        elif accion == "eliminar":
            nombre = request.form.get("nombre")
            try:
                _eliminar_imagen_campania(nombre)
                mensaje = f'<p class="msg-exito">🗑️ Imagen eliminada: {nombre}</p>'
            except Exception as e:
                mensaje = f'<p class="msg-error">Error al eliminar: {e}</p>'

    imagenes = _listar_imagenes_campania()

    filas_html = ""
    for img in imagenes:
        filas_html += f"""<tr>
            <td><img src="{img['url']}" style="max-width:70px;max-height:70px;border-radius:6px;"></td>
            <td>{img['nombre']}</td>
            <td>{img['tamano_kb']} KB</td>
            <td><input type="text" value="{img['url']}" readonly style="width:100%;font-size:11px;" onclick="this.select();"></td>
            <td>
                <form method="POST" action="/campania/imagenes?pw={pw}" onsubmit="return confirm('¿Eliminar {img['nombre']}?');" style="margin:0;">
                    <input type="hidden" name="accion" value="eliminar">
                    <input type="hidden" name="nombre" value="{img['nombre']}">
                    <button type="submit" class="danger" style="margin:0;padding:8px 14px;font-size:13px;">Eliminar</button>
                </form>
            </td>
        </tr>"""

    contenido = f"""
    <div class="card">
      <form method="POST" action="/campania/imagenes?pw={pw}" enctype="multipart/form-data">
        <input type="hidden" name="accion" value="subir">
        <label>Nueva imagen (usada en previews de WhatsApp / Open Graph)</label>
        <input type="file" name="archivo" accept="image/*" required>
        <button type="submit">Subir imagen</button>
      </form>
      {mensaje}
    </div>
    <div class="table-wrap">
      <table>
        <tr><th>Vista previa</th><th>Nombre</th><th>Tamaño</th><th>URL pública</th><th></th></tr>
        {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, "Gestor de imágenes", "Imágenes públicas para Open Graph (Cloud Storage)", contenido, activo="imagenes")

from datetime import date

@con_log_errores
def handle_mensajes_por_dia_campania(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    hoy = date.today()
    primer_dia_mes = hoy.replace(day=1).isoformat()

    fecha_inicio = request.args.get("fecha_inicio") or primer_dia_mes
    fecha_fin = request.args.get("fecha_fin") or hoy.isoformat()
    tipo_filtro = request.args.get("tipo_cliente", "")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT tipo_cliente FROM campania_plantillas ORDER BY tipo_cliente")
            tipos = [r[0] for r in cur.fetchall()]

            where = ["creado_at >= %s", "creado_at < (%s::date + INTERVAL '1 day')"]
            params = [fecha_inicio, fecha_fin]
            if tipo_filtro:
                where.append("tipo_cliente = %s")
                params.append(tipo_filtro)
            where_sql = " AND ".join(where)

            cur.execute(f"""
                SELECT
                    creado_at::date AS fecha,
                    COUNT(*) AS enviados,
                    SUM(CASE WHEN entregado THEN 1 ELSE 0 END) AS entregados,
                    SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END) AS abiertos,
                    SUM(CASE WHEN error_envio IS NOT NULL THEN 1 ELSE 0 END) AS errores
                FROM campania_links
                WHERE {where_sql}
                GROUP BY creado_at::date
                ORDER BY fecha DESC
            """, params)
            filas = cur.fetchall()

            cur.execute(f"""
                SELECT COUNT(*),
                       SUM(CASE WHEN entregado THEN 1 ELSE 0 END),
                       SUM(CASE WHEN clicked_at IS NOT NULL THEN 1 ELSE 0 END),
                       SUM(CASE WHEN error_envio IS NOT NULL THEN 1 ELSE 0 END)
                FROM campania_links
                WHERE {where_sql}
            """, params)
            total_env, total_ent, total_abi, total_err = cur.fetchone()
    finally:
        conn.close()

    opciones_tipo = "".join(
        f'<option value="{t}" {"selected" if t == tipo_filtro else ""}>{t}</option>' for t in tipos
    )

    filas_html = ""
    for fecha, enviados, entregados, abiertos, errores in filas:
        tasa_apertura = round(100.0 * (abiertos or 0) / enviados, 1) if enviados else 0
        tasa_entrega = round(100.0 * (entregados or 0) / enviados, 1) if enviados else 0
        color_err = 'style="color:var(--sanitas-danger);font-weight:600;"' if errores else ""
        filas_html += f"""<tr>
            <td>{fecha.strftime('%d/%m/%Y')} <span style="color:var(--sanitas-gray);font-size:11px;">({fecha.strftime('%A')})</span></td>
            <td><b>{enviados}</b></td>
            <td>{entregados or 0} <span style="color:var(--sanitas-gray);font-size:11px;">({tasa_entrega}%)</span></td>
            <td>{abiertos or 0} <span style="color:var(--sanitas-gray);font-size:11px;">({tasa_apertura}%)</span></td>
            <td {color_err}>{errores or 0}</td>
        </tr>"""

    if not filas:
        filas_html = '<tr><td colspan="5" style="text-align:center;color:var(--sanitas-gray);padding:20px;">No hay envíos en el rango seleccionado.</td></tr>'

    contenido = f"""
    <div class="card">
      <form method="GET" action="/campania/mensajes-por-dia" class="filtros">
        <input type="hidden" name="pw" value="{pw}">
        <div><label>Desde</label><input type="date" name="fecha_inicio" value="{fecha_inicio}"></div>
        <div><label>Hasta</label><input type="date" name="fecha_fin" value="{fecha_fin}"></div>
        <div><label>Tipo cliente</label>
          <select name="tipo_cliente"><option value="">Todos</option>{opciones_tipo}</select>
        </div>
        <button type="submit">Filtrar</button>
      </form>
    </div>

    <div class="card" style="display:flex;gap:30px;flex-wrap:wrap;">
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Total enviados</div><div style="font-size:24px;font-weight:700;">{total_env or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Entregados</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-teal-dark);">{total_ent or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Abiertos</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-success);">{total_abi or 0}</div></div>
      <div><div style="font-size:12px;color:var(--sanitas-gray);">Errores</div><div style="font-size:24px;font-weight:700;color:{'var(--sanitas-danger)' if total_err else 'inherit'};">{total_err or 0}</div></div>
    </div>

    <div class="table-wrap">
      <table>
      <tr><th>Día</th><th>Enviados</th><th>Entregados</th><th>Abiertos</th><th>Errores</th></tr>
      {filas_html}
      </table>
    </div>
    """

    return _layout_campania(pw, "Mensajes por día", "Vista diaria de envíos (mes actual por defecto)", contenido, activo="por-dia")

@con_log_errores
def handle_envio_especial_web(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CAMPANIA:
        return "No autorizado", 401

    mensaje_resultado = ""

    # Lista de imágenes ya subidas al bucket, para elegir en un <select>
    imagenes = _listar_imagenes_campania()
    opciones_imagenes = "".join(
        f'<option value="{img["nombre"]}">{img["nombre"]} ({img["tamano_kb"]} KB)</option>'
        for img in imagenes
    )

    if request.method == "POST":
        tipo_cliente = (request.form.get("tipo_cliente") or "").strip().upper()
        campana = (request.form.get("campana") or "").strip()
        mensaje_tpl = (request.form.get("mensaje") or "").strip()
        whatsapp_destino = (request.form.get("whatsapp_destino") or "").strip()
        og_title = (request.form.get("og_title") or "Sanitas Perú").strip()
        og_description = (request.form.get("og_description") or "Contáctanos por WhatsApp").strip()
        og_image = (request.form.get("og_image") or "").strip()
        archivo = request.files.get("archivo")

        errores_validacion = []
        if not tipo_cliente:
            errores_validacion.append("nombre identificador requerido")
        if not campana:
            errores_validacion.append("nombre de campaña requerido")
        if not mensaje_tpl:
            errores_validacion.append("mensaje requerido")
        if not whatsapp_destino:
            errores_validacion.append("whatsapp destino requerido")
        if not og_image:
            errores_validacion.append("selecciona una imagen")
        if not archivo or archivo.filename == "":
            errores_validacion.append("archivo Excel requerido")

        filas = None

        if errores_validacion:
            mensaje_resultado = f'<p class="msg-error">{"; ".join(errores_validacion)}</p>'
        else:
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
                    if not telefono1:
                        continue
                    filas.append((
                        str(data.get("nombre") or "").strip(),
                        telefono1,
                        tipo_cliente
                    ))
                wb.close()
            except Exception as e:
                filas = None
                mensaje_resultado = f'<p class="msg-error">Error al leer el Excel: {e}</p>'

            if filas is not None:
                if not filas:
                    mensaje_resultado = '<p class="msg-error">No se encontraron filas válidas (columna telefono1 requerida).</p>'
                else:
                    conn = get_connection()
                    try:
                        with conn.cursor() as cur:
                            # Crea o actualiza la plantilla de este envío especial
                            cur.execute("""
                                INSERT INTO campania_plantillas
                                    (tipo_cliente, mensaje, og_title, og_description, og_image, whatsapp_destino, actualizado_at)
                                VALUES (%s, %s, %s, %s, %s, %s, NOW())
                                ON CONFLICT (tipo_cliente) DO UPDATE SET
                                    mensaje = EXCLUDED.mensaje,
                                    og_title = EXCLUDED.og_title,
                                    og_description = EXCLUDED.og_description,
                                    og_image = EXCLUDED.og_image,
                                    whatsapp_destino = EXCLUDED.whatsapp_destino,
                                    actualizado_at = NOW()
                            """, (tipo_cliente, mensaje_tpl, og_title, og_description, og_image, whatsapp_destino))

                            execute_values(cur, """
                                INSERT INTO campania_contactos
                                    (nombre, telefono1, tipo_cliente)
                                VALUES %s
                                ON CONFLICT (telefono1, tipo_cliente) DO UPDATE SET
                                    nombre = EXCLUDED.nombre,
                                    actualizado_at = NOW()
                            """, filas)
                        conn.commit()
                    finally:
                        conn.close()

                    try:
                        # excluir_convertidos=False: es una lista nueva y puntual,
                        # queremos enviar a todos los que están en el Excel.
                        # max_intentos=1: un envío especial normalmente no necesita
                        # reintentos automáticos, pero puedes volver a correr este
                        # mismo formulario con el mismo tipo_cliente si hace falta.
                        resultados, omitidos = _generar_y_enviar_campania(
                            tipo_cliente, campana, excluir_convertidos=False, max_intentos=1
                        )
                    except ValueError as e:
                        mensaje_resultado = f'<p class="msg-error">{e}</p>'
                        resultados = None

                    if resultados is not None:
                        exitosos = sum(1 for _, _, ok, _ in resultados if ok)
                        fallidos = len(resultados) - exitosos
                        mensaje_resultado = f"""
                        <p class="msg-exito">✅ {len(filas)} contactos cargados en <b>{tipo_cliente}</b>.
                        Envío: {exitosos} exitosos{f', ❌ {fallidos} con error' if fallidos else ''}.</p>
                        <p style="margin-top:14px;">
                          <a href="/campania/reporte-clicks?pw={pw}">Ver seguimiento de aperturas (clics) →</a>
                        </p>
                        """

    contenido = f"""
    <div class="card">
      <p style="color:var(--sanitas-gray);margin-top:0;">
        Envío puntual con lista de números, mensaje e imagen propios (no se mezcla
        con las campañas de retención regulares). El Excel debe tener columnas
        <b>nombre, telefono1</b>.
      </p>
      <form method="POST" action="/campania/envio-especial?pw={pw}" enctype="multipart/form-data"
            onsubmit="return confirm('Esto enviará SMS reales por Danaconnect a los números del Excel. ¿Continuar?');">
        <label>Nombre identificador (ej. ESPECIAL_DIA_MADRE)</label>
        <input type="text" name="tipo_cliente" required>
        <label>Nombre de campaña (para reportes)</label>
        <input type="text" name="campana" placeholder="ej. especial_ago2026" required>
        <label>Archivo Excel (columnas: nombre, telefono1)</label>
        <input type="file" name="archivo" accept=".xlsx,.xls" required>
        <label>Mensaje (usa {{Nombre}} y {{LINK}})</label>
        <textarea name="mensaje" rows="3" required></textarea>
        <label>WhatsApp destino (código país + número, sin +)</label>
        <input type="text" name="whatsapp_destino" required>
        <label>Imagen (preview de WhatsApp)</label>
        <select name="og_image" required>
          <option value="">-- selecciona --</option>
          {opciones_imagenes}
        </select>
        <label>OG Title</label>
        <input type="text" name="og_title" value="Sanitas Perú">
        <label>OG Description</label>
        <input type="text" name="og_description" value="Contáctanos por WhatsApp">
        <button type="submit">Cargar y enviar envío especial</button>
      </form>
      {mensaje_resultado}
    </div>
    """

    return _layout_campania(pw, "Envío especial", "Carga puntual: Excel + mensaje + imagen propios", contenido, activo="especial")

# ───────────────────────────────────────────────────────────────
# Enlaces con contador para redes sociales
# ───────────────────────────────────────────────────────────────
ADMIN_PW_SOCIAL = os.environ.get("ADMIN_PW_SOCIAL")


def _generar_code_social(nombre: str) -> str:
    raw = f"{nombre}{time.time()}".encode()
    h = hashlib.md5(raw).digest()
    return base64.urlsafe_b64encode(h)[:7].decode()


def _layout_social(pw, titulo, subtitulo, contenido_html, activo=""):
    def link(href, texto, key):
        clase = "activo" if key == activo else ""
        return f'<a href="{href}?pw={pw}" class="{clase}">{texto}</a>'

    nav_html = f"""
    <nav class="tabs">
      {link("/social/mantenimiento", "🔗 Crear / Gestionar", "mantenimiento")}
      {link("/social/reporte", "📊 Reporte de clics", "reporte")}
    </nav>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="noindex, nofollow">
  <title>{titulo} - Sanitas</title>
  <link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
  <style>{CAMPANIA_CSS}</style>
</head>
<body>
  <div class="topbar">
    <h1>{titulo}</h1>
    <div class="subtitulo">{subtitulo}</div>
  </div>
  <div class="contenedor">
    {nav_html}
    {contenido_html}
  </div>
</body>
</html>"""


# ─────────────────────────────────────────────
# WEB: crear y listar enlaces sociales
# GET  /social/mantenimiento?pw=...          -> lista + formulario de creación
# POST /social/mantenimiento?pw=...          -> crea un nuevo enlace
# ─────────────────────────────────────────────
@con_log_errores
def handle_mantenimiento_social(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_SOCIAL:
        return "No autorizado", 401

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    mensaje_resultado = ""

    if request.method == "POST":
        nombre = (request.form.get("nombre") or "").strip()
        plataforma = (request.form.get("plataforma") or "").strip().upper()
        whatsapp_destino = (request.form.get("whatsapp_destino") or "").strip()
        mensaje_prellenado = (request.form.get("mensaje_prellenado") or "").strip()
        og_title = (request.form.get("og_title") or "Sanitas Perú").strip()
        og_description = (request.form.get("og_description") or "Contáctanos por WhatsApp").strip()
        og_image = (request.form.get("og_image") or "").strip()

        if not nombre or not whatsapp_destino:
            mensaje_resultado = '<p class="msg-error">Nombre y WhatsApp destino son obligatorios.</p>'
        else:
            code = _generar_code_social(nombre)
            conn = get_connection()
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO enlaces_sociales
                            (code, nombre, plataforma, whatsapp_destino, mensaje_prellenado,
                             og_title, og_description, og_image)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    """, (code, nombre, plataforma, whatsapp_destino, mensaje_prellenado,
                          og_title, og_description, og_image))
                conn.commit()
            finally:
                conn.close()

            link_generado = f"{base_url}/s/{code}"
            mensaje_resultado = f"""
            <p class="msg-exito">✅ Enlace creado: <code>{link_generado}</code></p>
            <p style="margin-top:8px;">
              <input type="text" value="{link_generado}" readonly style="width:100%;font-size:13px;" onclick="this.select();">
            </p>
            """

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT code, nombre, plataforma, whatsapp_destino, activo, raw_hits, creado_at, og_image,
                       (SELECT COUNT(*) FROM enlaces_sociales_clicks c WHERE c.code = enlaces_sociales.code) AS clics_confirmados
                FROM enlaces_sociales
                ORDER BY creado_at DESC
            """)
            enlaces = cur.fetchall()
    finally:
        conn.close()

    filas_html = ""
    for code, nombre, plataforma, whatsapp_destino, activo, raw_hits, creado_at, og_image, clics_confirmados in enlaces:
        link_completo = f"{base_url}/s/{code}"
        badge_activo = '<span class="badge badge-ok">Activo</span>' if activo else '<span class="badge badge-neutro">Inactivo</span>'
        miniatura = (
            f'<img src="{og_image}" style="max-width:50px;max-height:50px;border-radius:6px;object-fit:cover;">'
            if og_image else '<span class="badge badge-neutro">Sin imagen</span>'
        )
        filas_html += f"""<tr>
            <td>{miniatura}</td>
            <td>{nombre}<br><span style="font-size:11px;color:var(--sanitas-gray);">{plataforma or ''}</span></td>
            <td><input type="text" value="{link_completo}" readonly style="width:180px;font-size:12px;" onclick="this.select();"></td>
            <td>{whatsapp_destino}</td>
            <td>{raw_hits or 0}</td>
            <td><b>{clics_confirmados or 0}</b></td>
            <td>{badge_activo}</td>
            <td>{creado_at.strftime('%d/%m/%Y') if creado_at else ''}</td>
        </tr>"""

    if not enlaces:
        filas_html = '<tr><td colspan="7" style="text-align:center;color:var(--sanitas-gray);padding:20px;">Aún no hay enlaces creados.</td></tr>'

    contenido = f"""
    <div class="card">
      <form method="POST" action="/social/mantenimiento?pw={pw}">
        <label>Nombre identificador (ej. post_fb_agosto2026)</label>
        <input type="text" name="nombre" required>
        <label>Plataforma</label>
        <select name="plataforma">
          <option value="FACEBOOK">Facebook</option>
          <option value="INSTAGRAM">Instagram</option>
          <option value="TIKTOK">TikTok</option>
          <option value="OTRO">Otro</option>
        </select>
        <label>WhatsApp destino (código país + número, sin +)</label>
        <input type="text" name="whatsapp_destino" placeholder="51955973936" required>
        <label>Mensaje prellenado (opcional, se abrirá ya escrito en WhatsApp)</label>
        <textarea name="mensaje_prellenado" rows="2" placeholder="Hola, vengo de Instagram y quiero información sobre..."></textarea>
        <label>OG Title (título de la vista previa al compartir)</label>
        <input type="text" name="og_title" value="Sanitas Perú">
        <label>OG Description</label>
        <input type="text" name="og_description" value="Contáctanos por WhatsApp">
        <label>OG Image (URL pública de imagen, opcional)</label>
        <input type="text" name="og_image" placeholder="https://storage.googleapis.com/.../imagen.jpg">
        <button type="submit">Crear enlace</button>
      </form>
      {mensaje_resultado}
    </div>
    <div class="table-wrap">
      <table>
        <tr><th>Imagen</th><th>Nombre</th><th>Link</th><th>WhatsApp</th><th>Hits crudos</th><th>Clics confirmados</th><th>Estado</th><th>Creado</th></tr>
        {filas_html}
      </table>
    </div>
    """

    return _layout_social(pw, "Enlaces sociales", "Links con seguimiento de clics para redes sociales", contenido, activo="mantenimiento")


# ─────────────────────────────────────────────
# WEB: reporte de clics
# ─────────────────────────────────────────────
@con_log_errores
@con_log_errores
def handle_reporte_social(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_SOCIAL:
        return "No autorizado", 401

    code_filtro = request.args.get("code", "")
    fecha_inicio = request.args.get("fecha_inicio", "")
    fecha_fin = request.args.get("fecha_fin", "")

    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT code, nombre FROM enlaces_sociales ORDER BY nombre")
            enlaces_todos = cur.fetchall()
            nombre_por_code = {code: nombre for code, nombre in enlaces_todos}

            where = ["1=1"]
            params = []
            if code_filtro:
                where.append("c.code = %s")
                params.append(code_filtro)
            if fecha_inicio:
                where.append("c.fecha >= %s")
                params.append(fecha_inicio)
            if fecha_fin:
                where.append("c.fecha < (%s::date + INTERVAL '1 day')")
                params.append(fecha_fin)
            where_sql = " AND ".join(where)

            # ── Detalle día x enlace (para el pivot y el gráfico) ──
            cur.execute(f"""
                SELECT c.fecha::date AS fecha, c.code, COUNT(*) AS clics
                FROM enlaces_sociales_clicks c
                WHERE {where_sql}
                GROUP BY c.fecha::date, c.code
                ORDER BY fecha DESC
            """, params)
            filas_pivot = cur.fetchall()

            cur.execute(f"SELECT COUNT(*) FROM enlaces_sociales_clicks c WHERE {where_sql}", params)
            total_clics = cur.fetchone()[0]

            # ── Detalle de clics individuales, solo si hay un enlace filtrado ──
            detalle_clicks = []
            if code_filtro:
                cur.execute("""
                    SELECT fecha, ip_address, user_agent
                    FROM enlaces_sociales_clicks
                    WHERE code = %s
                    ORDER BY fecha DESC
                    LIMIT 300
                """, (code_filtro,))
                detalle_clicks = cur.fetchall()
    finally:
        conn.close()

    # ── Armar el pivot en memoria: {fecha: {code: clics}} ──
    fechas_ordenadas = sorted({f for f, c, n in filas_pivot}, reverse=True)
    codes_presentes = sorted({c for f, c, n in filas_pivot}, key=lambda cod: nombre_por_code.get(cod, cod))

    pivot = {}
    for fecha, code, clics in filas_pivot:
        pivot.setdefault(fecha, {})[code] = clics

    totales_por_code = {
        code: sum(pivot.get(f, {}).get(code, 0) for f in fechas_ordenadas)
        for code in codes_presentes
    }

    # ── Tabla pivot ──
    columnas_html = "".join(f"<th>{nombre_por_code.get(code, code)}</th>" for code in codes_presentes)

    filas_pivot_html = ""
    for fecha in fechas_ordenadas:
        celdas = "".join(
            f"<td>{pivot.get(fecha, {}).get(code, 0) or '—'}</td>" for code in codes_presentes
        )
        total_dia = sum(pivot.get(fecha, {}).values())
        filas_pivot_html += f"<tr><td>{fecha.strftime('%d/%m/%Y')}</td>{celdas}<td><b>{total_dia}</b></td></tr>"

    if not fechas_ordenadas:
        colspan = len(codes_presentes) + 2
        filas_pivot_html = f'<tr><td colspan="{colspan}" style="text-align:center;color:var(--sanitas-gray);padding:20px;">Sin clics en el rango seleccionado.</td></tr>'

    fila_totales_html = "".join(f"<td><b>{totales_por_code.get(code, 0)}</b></td>" for code in codes_presentes)

    # ── Gráfico de tendencia (Chart.js), una línea por enlace ──
    grafico_html = ""
    if fechas_ordenadas:
        labels_chart = [f.strftime('%d/%m') for f in reversed(fechas_ordenadas)]
        colores = ["#00A9B7", "#003654", "#2E9E5B", "#E8A33D", "#D64545", "#6B7A82", "#8E44AD", "#16A085"]
        datasets_js = []
        for i, code in enumerate(codes_presentes):
            data_serie = [pivot.get(f, {}).get(code, 0) for f in reversed(fechas_ordenadas)]
            datasets_js.append({
                "label": nombre_por_code.get(code, code),
                "data": data_serie,
                "borderColor": colores[i % len(colores)],
                "backgroundColor": colores[i % len(colores)],
                "tension": 0.3,
            })

        grafico_html = f"""
        <div class="card">
          <canvas id="chartClicksSocial" height="90"></canvas>
        </div>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
        <script>
          new Chart(document.getElementById('chartClicksSocial').getContext('2d'), {{
            type: 'line',
            data: {{
              labels: {json.dumps(labels_chart)},
              datasets: {json.dumps(datasets_js)}
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom' }} }} }}
          }});
        </script>
        """

    opciones_enlace = "".join(
        f'<option value="{code}" {"selected" if code == code_filtro else ""}>{nombre}</option>'
        for code, nombre in enlaces_todos
    )

    # ── Detalle de clics individuales (solo cuando hay un enlace seleccionado) ──
    detalle_html = ""
    if code_filtro:
        filas_detalle_html = ""
        for fecha, ip, ua in detalle_clicks:
            ua_corto = (ua or "")[:70]
            filas_detalle_html += f"""<tr>
                <td>{fecha.strftime('%d/%m/%Y %H:%M:%S')}</td>
                <td>{ip or ''}</td>
                <td style="font-size:11px;color:var(--sanitas-gray);">{ua_corto}</td>
            </tr>"""
        if not detalle_clicks:
            filas_detalle_html = '<tr><td colspan="3" style="text-align:center;color:var(--sanitas-gray);padding:14px;">Sin registros.</td></tr>'

        detalle_html = f"""
        <div class="card">
          <h3 style="margin:0 0 4px;">Detalle de clics individuales — {nombre_por_code.get(code_filtro, code_filtro)}</h3>
          <p style="color:var(--sanitas-gray);font-size:12px;margin:0;">Últimos 300 registros</p>
        </div>
        <div class="table-wrap">
          <table>
            <tr><th>Fecha y hora</th><th>IP</th><th>Navegador</th></tr>
            {filas_detalle_html}
          </table>
        </div>
        """

    contenido = f"""
    <div class="card">
      <form method="GET" action="/social/reporte" class="filtros">
        <input type="hidden" name="pw" value="{pw}">
        <div><label>Enlace</label>
          <select name="code"><option value="">Todos</option>{opciones_enlace}</select>
        </div>
        <div><label>Desde</label><input type="date" name="fecha_inicio" value="{fecha_inicio}"></div>
        <div><label>Hasta</label><input type="date" name="fecha_fin" value="{fecha_fin}"></div>
        <button type="submit">Filtrar</button>
      </form>
    </div>
    <div class="card">
      <div style="font-size:12px;color:var(--sanitas-gray);">Total de clics confirmados</div>
      <div style="font-size:28px;font-weight:700;color:var(--sanitas-success);">{total_clics or 0}</div>
    </div>
    {grafico_html}
    <div class="table-wrap">
      <table>
        <tr><th>Fecha</th>{columnas_html}<th>Total día</th></tr>
        {filas_pivot_html}
        <tr style="background:rgba(0,169,183,0.08);"><td><b>Total</b></td>{fila_totales_html}<td><b>{total_clics or 0}</b></td></tr>
      </table>
    </div>
    {detalle_html}
    """

    return _layout_social(pw, "Reporte de clics", "Detalle por día y por enlace", contenido, activo="reporte")



# ─────────────────────────────────────────────
# PÚBLICO: redirección + registro del clic
# GET /s/<code>
# ─────────────────────────────────────────────
@con_log_errores
def handle_redirect_social(code, request):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT whatsapp_destino, mensaje_prellenado, og_title, og_description, og_image, activo
                FROM enlaces_sociales WHERE code = %s
            """, (code,))
            row = cur.fetchone()
            if not row:
                return "Enlace no encontrado", 404

            whatsapp_destino, mensaje_prellenado, og_title, og_description, og_image, activo = row

            if not activo:
                return "Este enlace ya no está disponible", 410

            cur.execute("""
                UPDATE enlaces_sociales SET raw_hits = raw_hits + 1 WHERE code = %s
            """, (code,))
            conn.commit()
    finally:
        conn.close()

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    url_canonica = f"{base_url}/s/{code}"

    texto_wa = quote(mensaje_prellenado) if mensaje_prellenado else ""
    url_destino = f"https://wa.me/{whatsapp_destino}"
    if texto_wa:
        url_destino += f"?text={texto_wa}"

    # ── Meta tags de imagen: cache-busting + dimensiones + twitter:image ──
    # Sin twitter:image algunos clientes arman la tarjeta sin imagen aunque
    # og:image esté correcto. El parámetro ?v= evita que, al cambiar la
    # imagen de un link ya compartido, siga sirviendo la versión vieja
    # cacheada por WhatsApp/Facebook la próxima vez que se re-escanee.
    if og_image:
        separador = "&" if "?" in og_image else "?"
        img_url_meta = f"{og_image}{separador}v={int(time.time())}"
        imagen_meta = f"""<meta property="og:image" content="{img_url_meta}">
  <meta property="og:image:width" content="1200">
  <meta property="og:image:height" content="630">
  <meta name="twitter:image" content="{img_url_meta}">"""
    else:
        imagen_meta = ""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta property="og:type" content="website">
  <meta property="og:url" content="{url_canonica}">
  <meta property="og:site_name" content="Sanitas Perú">
  <meta property="og:title" content="{og_title or 'Sanitas Perú'}">
  <meta property="og:description" content="{og_description or 'Contáctanos por WhatsApp'}">
  {imagen_meta}
  <meta name="twitter:card" content="summary_large_image">
  <title>{og_title or 'Sanitas Perú'}</title>
</head>
<body>
  <script>
    if (navigator.sendBeacon) {{
      navigator.sendBeacon("/s/confirm/{code}");
    }} else {{
      fetch("/s/confirm/{code}", {{ method: "POST", keepalive: true }});
    }}
    window.location.href = "{url_destino}";
  </script>
  <p>Redirigiendo a WhatsApp... <a href="{url_destino}">Haz clic aquí si no eres redirigido.</a></p>
</body>
</html>"""
    return html, 200


# ─────────────────────────────────────────────
# PÚBLICO: confirmación real del clic (via sendBeacon)
# POST /s/confirm/<code>
# Solo se llama si el JS del navegador se ejecutó de verdad,
# lo que filtra la mayoría de crawlers/bots de preview de Facebook/Instagram.
# ─────────────────────────────────────────────
@con_log_errores
def handle_confirmar_click_social(code, request):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM enlaces_sociales WHERE code = %s", (code,))
            if not cur.fetchone():
                return "", 204

            cur.execute("""
                INSERT INTO enlaces_sociales_clicks (code, ip_address, user_agent)
                VALUES (%s, %s, %s)
            """, (code, request.remote_addr, request.headers.get("User-Agent", "")))
            conn.commit()
    finally:
        conn.close()

    return "", 204

# ───────────────────────────────────────────────────────────────
# Prioridad de programas (más complejo = número más alto)
# ───────────────────────────────────────────────────────────────
import xml.etree.ElementTree as ET

DANACONNECT_USER = os.environ.get("DANACONNECT_USER", "canalesAPI@sanitasperu")
DANACONNECT_PASSWORD = os.environ.get("DANACONNECT_PASSWORD")
DANACONNECT_CONVERSATION_ID = "601417"
DANACONNECT_CONVERSATION_ID_CRONICO = "597709"

DANACONNECT_FILE2_POR_PROGRAMA = {
    "vive_sano_completo": "s3://WS/2026/7/1a20c4f50f4cc80ecbf06c9ad0ac3923.jpg",
    "krystal": "s3://WS/2026/7/291e2a5cbbd575dce140de36ad935c2f.jpg",
}

DANACONNECT_NOMBRE_PROGRAMA = {
    "vive_sano_completo": "VIVE SANO",
    "krystal": "KRYSTAL",
}

def _content_type_por_url(url):
    return "application/pdf" if url.lower().endswith(".pdf") else "image/jpeg"

def descargar_blob_desde_url(url):
    """
    Descarga un archivo privado de Cloud Storage usando la librería oficial,
    evitando requests.get() y errores 403.
    """

    # Ejemplo:
    # https://storage.googleapis.com/sanitas-programa-cronico/programa-cronico/token/dni.jpg

    partes = url.replace("https://storage.googleapis.com/", "").split("/", 1)

    if len(partes) != 2:
        raise ValueError(f"URL de Cloud Storage inválida: {url}")

    bucket_name, blob_name = partes

    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_name)

    return blob.download_as_bytes()


def subir_archivo_danaconnect(archivo_bytes, nombre_archivo, content_type):
    url = "https://appserv.danaconnect.com/dana/conversation/http/rest/file/upload"
    resp = requests.post(
        url,
        auth=(DANACONNECT_USER, DANACONNECT_PASSWORD),
        files={"file": (nombre_archivo, archivo_bytes, content_type)},
        timeout=30,
    )
    resp.raise_for_status()

    root = ET.fromstring(resp.text)
    file_id = root.findtext("fileID")
    if not file_id:
        raise ValueError(f"Danaconnect no devolvió fileID: {resp.text}")
    return file_id


def iniciar_conversacion_danaconnect(n_documento, nombre_completo, email, programa, file_id_pdf,
                                       file_id_dni_a, file_id_dni_r, file_id_receta):
    url = f"https://appserv.danaconnect.com/api/1.0/rest/conversation/{DANACONNECT_CONVERSATION_ID}/start/data"
    nombre_programa_dana = DANACONNECT_NOMBRE_PROGRAMA.get(programa, programa.upper())
    file2 = DANACONNECT_FILE2_POR_PROGRAMA.get(programa)

    payload = {
        "N_DOCUMENTO": n_documento,
        "NOMBRE": nombre_completo,
        "EMAIL": email,
        "PROGRAMA": nombre_programa_dana,
        "FILE": file_id_pdf,
        "FILE2": file2,
        "FILE_DNI_A": file_id_dni_a,
        "FILE_DNI_R": file_id_dni_r,
        "FILE_RECETA": file_id_receta,
    }

    resp = requests.post(
        url,
        auth=(DANACONNECT_USER, DANACONNECT_PASSWORD),
        json=payload,
        timeout=30,
    )
    resp.raise_for_status()
    return resp.text

def iniciar_conversacion_cronico_complejo(nombre, email_destino, programa, celular, n_documento, correo_paciente, file_id_dni_a, file_id_dni_r,
    file_id_receta):
    
    url = (
        f"https://appserv.danaconnect.com/api/1.0/rest/conversation/"
        f"{DANACONNECT_CONVERSATION_ID_CRONICO}/start/data"
    )

    payload = {
        "NOMBRE": nombre,
        "EMAIL": email_destino,
        "PROGRAMA": programa,
        "CELULAR": celular,
        "N_DOCUMENTO": n_documento,
        "CORREO_PACIENTE": correo_paciente,
        "FILE_DNI_A": file_id_dni_a,
        "FILE_DNI_R": file_id_dni_r,
        "FILE_RECETA": file_id_receta,
    }

    resp = requests.post(
        url,
        auth=(DANACONNECT_USER, DANACONNECT_PASSWORD),
        json=payload,
        timeout=30,
    )

    print(f"[DEBUG Danaconnect Cronico] status={resp.status_code} respuesta={resp.text}")

    resp.raise_for_status()
    return resp.text

def notificar_danaconnect(n_documento, nombres, apellidos, correo, programa, pdf_bytes,
                            dni_anverso_url, dni_reverso_url, url_receta_bucket):
    nombre_completo = f"{nombres} {apellidos}".strip()

    file_id_pdf = subir_archivo_danaconnect(pdf_bytes, "consentimiento.pdf", "application/pdf")

    # ── DNI anverso ──
    archivo_dni_a = descargar_blob_desde_url(dni_anverso_url)

    ext_a = "pdf" if _content_type_por_url(dni_anverso_url) == "application/pdf" else "jpg"

    file_id_dni_a = subir_archivo_danaconnect(
        archivo_dni_a,
        f"dni_anverso.{ext_a}",
        _content_type_por_url(dni_anverso_url)
    )

    # ── DNI reverso (si no existe, reusamos el anverso para no romper el payload) ──
    if dni_reverso_url:
        archivo_dni_r = descargar_blob_desde_url(dni_reverso_url)

        ext_r = "pdf" if _content_type_por_url(dni_reverso_url) == "application/pdf" else "jpg"

        file_id_dni_r = subir_archivo_danaconnect(
            archivo_dni_r,
            f"dni_reverso.{ext_r}",
            _content_type_por_url(dni_reverso_url)      
        )
    else:
        file_id_dni_r = file_id_dni_a

    # ── Receta ──
    archivo_receta = descargar_blob_desde_url(url_receta_bucket)

    ext_receta = ("pdf" if _content_type_por_url(url_receta_bucket) == "application/pdf" else "jpg")

    file_id_receta = subir_archivo_danaconnect(
            archivo_receta,
            f"receta.{ext_receta}",
            _content_type_por_url(url_receta_bucket)
    )

    iniciar_conversacion_danaconnect(
        n_documento, nombre_completo, correo, programa, file_id_pdf,
        file_id_dni_a, file_id_dni_r, file_id_receta
    )

PRIORIDAD_PROGRAMAS = {
    "cronico_complejo": 4,
    "krystal": 3,
    "vive_sano_completo": 2,
    "vive_sano_peas": 1,
}

STOPWORDS_NOMBRES = {"DE", "DEL", "LA", "LAS", "LOS", "Y"}

def _normalizar_nombre(nombre):
     nombre = (nombre or "").upper()
     nombre = re.sub(r"[.,;:]", " ", nombre)  # quita comas, puntos, etc.
     nombre = re.sub(r"\s+", " ", nombre).strip()
     return nombre

def _dni_similar(dni_a, dni_b, max_diferencias=1):
    """
    Compara dos números de documento tolerando hasta max_diferencias dígitos
    distintos, para absorber errores comunes de OCR en un solo carácter
    (ej. '4' leído como '9', '0' leído como '8'). Solo aplica si ambos
    tienen la misma longitud; si difieren en longitud, se considera que
    no coinciden (evita falsos positivos por documentos truncados).
    """
    a = re.sub(r"[^\d]", "", dni_a or "")
    b = re.sub(r"[^\d]", "", dni_b or "")
    if not a or not b or len(a) != len(b):
        return False
    diferencias = sum(1 for c1, c2 in zip(a, b) if c1 != c2)
    return diferencias <= max_diferencias

def _nombres_coinciden(nombre1, nombre2, umbral=0.85):
    import jellyfish
    n1 = _normalizar_nombre(nombre1)
    n2 = _normalizar_nombre(nombre2)
    if not n1 or not n2:
        return False

    # Comparación 1: string completo
    similitud_completa = jellyfish.jaro_winkler_similarity(n1, n2)
    if similitud_completa >= umbral:
        return True

    # Comparación 2: por conjunto de palabras (tolera orden distinto, ej. apellidos/nombres invertidos)
    palabras1 = set(n1.split()) - STOPWORDS_NOMBRES
    palabras2 = set(n2.split()) - STOPWORDS_NOMBRES
    if not palabras1 or not palabras2:
        return False

    coincidencias = 0
    for p1 in palabras1:
        mejor = max((jellyfish.jaro_winkler_similarity(p1, p2) for p2 in palabras2), default=0)
        if mejor >= 0.85:
            coincidencias += 1

    proporcion = coincidencias / max(len(palabras1), len(palabras2))
    return proporcion >= 0.75

def extraer_datos_receta(archivo_bytes, mime_type="image/jpeg"):
    nombre_procesador = docai_client.processor_path(DOCAI_PROJECT_ID, DOCAI_LOCATION, DOCAI_PROCESSOR_ID)
    raw_document = documentai.RawDocument(content=archivo_bytes, mime_type=mime_type)
    request_docai = documentai.ProcessRequest(name=nombre_procesador, raw_document=raw_document)
    resultado = docai_client.process_document(request=request_docai)
    texto_crudo = resultado.document.text

    if not texto_crudo or len(texto_crudo.strip()) < 10:
        return {"nombre_paciente": None, "fecha_emision": None, "diagnosticos": []}

    prompt = f"""
Extrae los siguientes datos de esta receta médica peruana. Responde SOLO con JSON,
sin explicación, sin markdown, sin backticks.

Formato exacto:
{{
  "nombre_paciente": "nombre completo del paciente tal como aparece, en mayúsculas o null",
  "n_documento_paciente": "número de DNI del paciente si aparece impreso en la receta, o null",
  "fecha_emision": "YYYY-MM-DD o null si no se encuentra",
  "diagnosticos": [
    {{"cie10": "código CIE-10 EXACTO tal como está impreso, incluyendo cualquier sufijo después del punto (ej. I10.X, E11.9), sin recortar ni simplificar nada", "descripcion": "descripción del diagnóstico"}}
  ]
}}

IMPORTANTE sobre el código CIE-10: el PRIMER carácter de un código CIE-10 SIEMPRE es una
letra (A-Z), nunca un dígito. Si el texto original parece mostrar un número en esa primera
posición (por ejemplo, un "1" que en realidad es la letra "I", o un "0" que en realidad es
la letra "O"), corrígelo a la letra correspondiente. El resto del código sí puede tener dígitos.

Si hay más de un diagnóstico, inclúyelos todos en la lista. Si no encuentras código CIE-10
explícito para un diagnóstico, deja cie10 como null.

Texto extraído del documento:
{texto_crudo}
"""

    respuesta = gemini_model.generate_content(prompt)
    texto_respuesta = respuesta.text.strip().replace("```json", "").replace("```", "").strip()
    try:
        datos = json.loads(texto_respuesta)
    except json.JSONDecodeError:
        datos = {"nombre_paciente": None, "fecha_emision": None, "diagnosticos": []}
    return datos

_CONFUSIONES_OCR_GENERAL = {
    "1": "I", "0": "O", "5": "S", "8": "B", "6": "G", "2": "Z",
    "I": "1", "O": "0", "S": "5", "B": "8", "G": "6", "Z": "2",
}

def _generar_candidatos_cie10(cie10_limpio, max_candidatos=8):
    """
    Genera variantes del código probando corregir, de a una posición a la
    vez, caracteres que suelen confundirse por OCR (letra/dígito). La
    primera posición SIEMPRE debe ser letra (regla CIE-10), así que ahí se
    fuerza dígito->letra; en el resto de posiciones se prueban ambos
    sentidos porque el error puede ir en cualquier dirección.
    """
    if not cie10_limpio:
        return []

    candidatos = {cie10_limpio}

    # Posición 0: siempre debe terminar siendo letra
    if cie10_limpio[0] in _CONFUSIONES_OCR_GENERAL:
        candidatos.add(_CONFUSIONES_OCR_GENERAL[cie10_limpio[0]] + cie10_limpio[1:])

    # Resto de posiciones: prueba sustitución individual (una por variante)
    for i in range(1, len(cie10_limpio)):
        char = cie10_limpio[i]
        if char in _CONFUSIONES_OCR_GENERAL:
            variante = cie10_limpio[:i] + _CONFUSIONES_OCR_GENERAL[char] + cie10_limpio[i + 1:]
            candidatos.add(variante)
            if len(candidatos) >= max_candidatos:
                break

    return list(candidatos)

def determinar_programa_por_diagnosticos(diagnosticos_receta, cur):
    matches = []  # lista de (programa, cie10, descripcion) por cada diagnóstico encontrado
    
    # Catálogo completo cargado una sola vez, usado tanto para el match por
    # código como para el fallback por similitud de texto.
    cur.execute("SELECT programa, cie_10, descripcion FROM sofi_programa_diagnosticos")
    catalogo = cur.fetchall()

    UMBRAL_SIMILITUD_DESCRIPCION = 75  # calibrar según falsos positivos/negativos observados

    for diag in diagnosticos_receta:
        cie10 = _limpiar_cie10(diag.get("cie10"))
        descripcion_diag = (diag.get("descripcion") or "").strip()

        match_encontrado = None
        origen_match = None

        # 1. Intento por código CIE-10 (exacto o por prefijo, con variantes OCR)
        if cie10:
            for candidato in _generar_candidatos_cie10(cie10):
                for programa_bd, cie_bd, desc_bd in catalogo:
                    cie_bd_norm = (cie_bd or "").strip().upper()
                    if cie_bd_norm == candidato or cie_bd_norm.startswith(candidato):
                        match_encontrado = (programa_bd, cie_bd, desc_bd)
                        origen_match = f"cie10:{candidato}"
                        break
                if match_encontrado:
                    break

        # 2. Fallback: si no hubo match por código (o no vino código),
        # compara la descripción del diagnóstico contra el catálogo con la
        # misma función de similitud que ya usa /busca_actividad.
        if not match_encontrado and descripcion_diag:
            mejor_score = 0
            mejor_row = None
            for programa_bd, cie_bd, desc_bd in catalogo:
                score = similitud_avanzada(descripcion_diag, desc_bd or "")
                if score > mejor_score:
                    mejor_score, mejor_row = score, (programa_bd, cie_bd, desc_bd)
            if mejor_row and mejor_score >= UMBRAL_SIMILITUD_DESCRIPCION:
                match_encontrado = mejor_row
                origen_match = f"descripcion:{mejor_score}"

        if match_encontrado:
            programa_encontrado, cie10_encontrado, descripcion_bd = match_encontrado
            matches.append((programa_encontrado, cie10_encontrado, descripcion_bd))
            logging.info(f"[programa-cronico] Match diagnóstico '{descripcion_diag}' -> {programa_encontrado} (origen: {origen_match})")

    if not matches:
        return None, None, None

    # ── Regla especial: 2 o más diagnósticos de Vive Sano → se sube a Krystal ──
    #conteo_vive_sano = sum(1 for m in matches if m[0] == "vive_sano_completo")
    #if conteo_vive_sano >= 2:
    #    matches = [
    #        ("krystal", cie10, desc) if programa == "vive_sano_completo" else (programa, cie10, desc)
    #        for (programa, cie10, desc) in matches
    #    ]

    # ── Elegir el de mayor prioridad entre todos los matches (ya con la regla aplicada) ──
    mejor_programa, mejor_cie10, mejor_descripcion = max(
        matches, key=lambda m: PRIORIDAD_PROGRAMAS.get(m[0], 0)
    )

    return mejor_programa, mejor_cie10, mejor_descripcion

def enviar_correo_cronico_complejo(nombres, apellidos, n_documento, celular, correo_paciente, programa, dni_anverso_url, dni_reverso_url,
    receta_url):
    
    email_destino = "programasdesalud@sanitasperu.com"
    
    nombre_completo = f"{nombres} {apellidos}".strip()

    # -------------------------
    # DNI Anverso
    # -------------------------
    archivo_dni_a = descargar_blob_desde_url(dni_anverso_url)

    ext_a = (
        "pdf"
        if _content_type_por_url(dni_anverso_url) == "application/pdf"
        else "jpg"
    )

    file_id_dni_a = subir_archivo_danaconnect(
        archivo_dni_a,
        f"dni_anverso.{ext_a}",
        _content_type_por_url(dni_anverso_url)
    )

    # -------------------------
    # DNI Reverso
    # -------------------------
    if dni_reverso_url:

        archivo_dni_r = descargar_blob_desde_url(dni_reverso_url)

        ext_r = (
            "pdf"
            if _content_type_por_url(dni_reverso_url) == "application/pdf"
            else "jpg"
        )

        file_id_dni_r = subir_archivo_danaconnect(
            archivo_dni_r,
            f"dni_reverso.{ext_r}",
            _content_type_por_url(dni_reverso_url)
        )

    else:

        file_id_dni_r = file_id_dni_a

    # -------------------------
    # Receta
    # -------------------------
    archivo_receta = descargar_blob_desde_url(receta_url)

    ext_receta = (
        "pdf"
        if _content_type_por_url(receta_url) == "application/pdf"
        else "jpg"
    )

    file_id_receta = subir_archivo_danaconnect(
        archivo_receta,
        f"receta.{ext_receta}",
        _content_type_por_url(receta_url)
    )

    iniciar_conversacion_cronico_complejo(
        nombre=nombre_completo,
        email_destino=email_destino,
        programa="CRÓNICO COMPLEJO",
        celular=celular,
        n_documento=n_documento,
        correo_paciente=correo_paciente,
        file_id_dni_a=file_id_dni_a,
        file_id_dni_r=file_id_dni_r,
        file_id_receta=file_id_receta,
    )

def _limpiar_cie10(cie10):
    cie10 = (cie10 or "").upper()
    cie10 = re.sub(r"[^A-Z0-9.]", "", cie10)  # quita corchetes, espacios, guiones raros, etc.
    return cie10.strip()

def _extraer_head_styles(html_completo):
    match = re.search(r"<style[^>]*>(.*?)</style>", html_completo, re.DOTALL)
    return match.group(1) if match else ""

def _extraer_body(html_completo):
    match = re.search(r"<body[^>]*>(.*)</body>", html_completo, re.DOTALL)
    return match.group(1) if match else html_completo
    
# ──────────────────────────────────────────
# API: Programas SUBIR DNI en imagen o PDF
# ──────────────────────────────────────────
from google.cloud import documentai_v1 as documentai
import vertexai
from vertexai.generative_models import GenerativeModel

DOCAI_PROJECT_ID = "sanitas-recetas-prod"
DOCAI_LOCATION = "us"
DOCAI_PROCESSOR_ID = "6837fb55f199826a"  # sanitas-ocr-recetas (Document OCR genérico)

docai_client = documentai.DocumentProcessorServiceClient(
    client_options={"api_endpoint": f"{DOCAI_LOCATION}-documentai.googleapis.com"}
)

def _ocr_texto_documento(archivo_bytes, mime_type):
    nombre_procesador = docai_client.processor_path(DOCAI_PROJECT_ID, DOCAI_LOCATION, DOCAI_PROCESSOR_ID)
    raw_document = documentai.RawDocument(content=archivo_bytes, mime_type=mime_type)
    request_docai = documentai.ProcessRequest(name=nombre_procesador, raw_document=raw_document)
    resultado = docai_client.process_document(request=request_docai)
    return resultado.document.text

vertexai.init(project=DOCAI_PROJECT_ID, location="us-central1")
gemini_model = GenerativeModel("gemini-2.5-flash-lite")

def extraer_datos_dni(bytes_anverso, mime_anverso, bytes_reverso=None, mime_reverso=None):
    texto_anverso = _ocr_texto_documento(bytes_anverso, mime_anverso)
    texto_reverso = _ocr_texto_documento(bytes_reverso, mime_reverso) if bytes_reverso else ""

    texto_crudo = f"{texto_anverso}\n\n--- REVERSO DEL DOCUMENTO ---\n\n{texto_reverso}"

    if not texto_crudo.strip() or len((texto_anverso + texto_reverso).strip()) < 10:
        return {
            "tipo_documento": None, "n_documento": None, "nombres": None, "apellidos": None,
            "fecha_nacimiento": None, "domicilio": None, "distrito": None,
            "provincia": None, "departamento": None, "sexo": None, "estado_civil": None,
            "nacionalidad": None,
            "error": "No se pudo leer texto del documento"
        }

    prompt = f"""
Extrae los siguientes datos de este texto de un documento de identidad peruano (puede ser
DNI, Carné de Extranjería o Pasaporte; incluye anverso y reverso si aplica).
Responde SOLO con JSON, sin explicación, sin markdown, sin backticks.

Formato exacto:
{{
  "tipo_documento": "uno de: DNI, CE, Pasaporte — según el tipo de documento identificado",
  "n_documento": "número del documento o null si no se encuentra",
  "nombres": "nombres del titular en mayúsculas o null",
  "apellidos": "apellidos del titular (paterno y materno juntos) en mayúsculas o null",
  "fecha_nacimiento": "YYYY-MM-DD o null si no se encuentra",
  "sexo": "M o F, según lo impreso en el documento (campo 'SEXO'), o null si no se encuentra",
  "estado_civil": "uno de: SOLTERO, CASADO, VIUDO, DIVORCIADO, CONVIVIENTE -- SOLO si aparece
    impreso explícitamente en el documento (algunos DNI antiguos lo traen, el DNI electrónico
    actual normalmente NO). Si no aparece impreso, responde null; NO lo infieras ni lo asumas.",
  "domicilio": "dirección completa impresa en el reverso del documento (calle, número, urbanización) o null si no aparece",
  "distrito": "nombre del distrito de residencia, en formato Título (ej. Santiago De Surco), o null",
  "provincia": "nombre de la provincia de residencia, en formato Título, o null",
  "departamento": "nombre del departamento de residencia, en formato Título, o null",
  "nacionalidad": "código de 3 letras impreso en el documento tal como aparece (ej. PER, VEN, COL, ECU), o null si no se encuentra"
}}

Pistas para identificar el tipo de documento: si el texto menciona "DOCUMENTO NACIONAL DE
IDENTIDAD" o "RENIEC", es DNI. Si menciona "CARNÉ DE EXTRANJERÍA" o "CARNET DE EXTRANJERIA",
es CE. Si menciona "PASSPORT" o "PASAPORTE", es Pasaporte. Si no hay indicios claros, asume DNI.

El documento suele traer una línea "Departamento/Provincia/Distrito" seguida de los tres
valores en ese orden separados por "/". Úsala para completar distrito, provincia y departamento
por separado.

El campo de nacionalidad suele aparecer como un código de 3 letras (ej. "PER" para peruano).
Transcríbelo tal cual está impreso, sin traducirlo ni expandirlo.

Texto extraído del documento:
{texto_crudo}
"""

    respuesta = gemini_model.generate_content(prompt)
    texto_respuesta = respuesta.text.strip().replace("```json", "").replace("```", "").strip()

    try:
        datos = json.loads(texto_respuesta)
    except json.JSONDecodeError:
        datos = {"tipo_documento": None, "n_documento": None, "nombres": None, "apellidos": None,
                 "fecha_nacimiento": None, "domicilio": None, "distrito": None,
                 "provincia": None, "departamento": None, "sexo": None, "estado_civil": None,
                 "nacionalidad": None}

    # Asegura que las claves nuevas siempre existan aunque Gemini no las devuelva
    # (retrocompatibilidad: programa-cronico usa esta misma función y no las usa)
    datos.setdefault("sexo", None)
    datos.setdefault("estado_civil", None)
    datos.setdefault("nacionalidad", None)

    return datos

# ──────────────────────────────────────────
# API: ENDPOINT PRINCIPAL DE PROGRAMAS
# ──────────────────────────────────────────
BUCKET_NAME = "sanitas-programa-cronico"
ESTADOS_BLOQUEADOS_REINTENTO = ("firmado",)

@con_log_errores
def handle_recibir_dni_cronico(request):
    data = request.get_json(force=True, silent=True) or {}
    if not data:
        return (json.dumps({"error": "Body JSON inválido o vacío"}), 400, HEADERS)

    celular = data.get("celular")
    correo = data.get("correo")
    n_documento = data.get("n_documento")
    url_anverso = data.get("url_anverso")
    url_reverso = data.get("url_reverso")
    url_documento = data.get("url_documento")
    url_receta = data.get("url_receta")

    if not celular or not n_documento or not url_receta:
        return (json.dumps({"error": "celular, n_documento y url_receta son requeridos"}), 400, HEADERS)

    if not url_documento and not (url_anverso and url_reverso):
        return (json.dumps({"error": "Debes enviar url_documento, o url_anverso y url_reverso"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()

    # ── Buscar o crear inscripción ──
    cur.execute(
        "SELECT id, estado, token FROM sofi_programa_cronico_inscripciones WHERE celular = %s AND n_documento = %s",
        (celular, n_documento)
    )
    row = cur.fetchone()
    if row:
        inscripcion_id, estado_actual, token = row

        if estado_actual in ESTADOS_BLOQUEADOS_REINTENTO:
            cur.close(); conn.close()
            return (json.dumps({
                "ok": False,
                "token": token,
                "estado": estado_actual,
                "error": "Esta inscripción ya fue firmada y no puede modificarse. Contacta al equipo de Programa de Salud si necesitas hacer un cambio.",
            }), 409, HEADERS)

        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET n_documento = %s, correo = %s, url_receta = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (n_documento, correo, url_receta, inscripcion_id))
    else:
        token = secrets.token_urlsafe(24)
        cur.execute("""
            INSERT INTO sofi_programa_cronico_inscripciones (token, n_documento, celular, correo, url_receta, estado)
            VALUES (%s, %s, %s, %s, %s, 'esperando_dni')
            RETURNING id
        """, (token, n_documento, celular, correo, url_receta))
        inscripcion_id = cur.fetchone()[0]
    conn.commit()

    bucket = storage_client.bucket(BUCKET_NAME)

    # ── Procesar DNI ──
    if url_documento:
        resp = requests.get(url_documento, timeout=20)
        resp.raise_for_status()
        archivo_bytes = resp.content
        content_type = "application/pdf"
        blob = bucket.blob(f"programa-cronico/{token}/dni_documento.pdf")
        blob.upload_from_string(archivo_bytes, content_type=content_type)
        cur.execute("UPDATE sofi_programa_cronico_inscripciones SET dni_foto_anverso_url = %s, actualizado_at = NOW() WHERE id = %s",
                    (blob.public_url, inscripcion_id))
        datos_dni = extraer_datos_dni(archivo_bytes, content_type)
    else:
        resp_a = requests.get(url_anverso, timeout=20); resp_a.raise_for_status()
        bytes_a = resp_a.content
        resp_r = requests.get(url_reverso, timeout=20); resp_r.raise_for_status()
        bytes_r = resp_r.content
        blob_a = bucket.blob(f"programa-cronico/{token}/dni_anverso.jpg")
        blob_a.upload_from_string(bytes_a, content_type="image/jpeg")
        blob_r = bucket.blob(f"programa-cronico/{token}/dni_reverso.jpg")
        blob_r.upload_from_string(bytes_r, content_type="image/jpeg")
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET dni_foto_anverso_url = %s, dni_foto_reverso_url = %s, actualizado_at = NOW() WHERE id = %s
        """, (blob_a.public_url, blob_r.public_url, inscripcion_id))
        datos_dni = extraer_datos_dni(bytes_a, "image/jpeg", bytes_r, "image/jpeg")

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET dni_validado = TRUE, nombres = %s, apellidos = %s, fecha_nacimiento = %s,
            domicilio = %s, distrito = %s, provincia = %s, departamento = %s,
            tipo_documento = %s, actualizado_at = NOW()
        WHERE id = %s
    """, (datos_dni.get("nombres"), datos_dni.get("apellidos"), datos_dni.get("fecha_nacimiento"),
          datos_dni.get("domicilio"), datos_dni.get("distrito"),
          datos_dni.get("provincia"), datos_dni.get("departamento"),
          datos_dni.get("tipo_documento"), inscripcion_id))
    conn.commit()

    # ── Procesar receta ──
    resp_receta = requests.get(url_receta, timeout=20)
    resp_receta.raise_for_status()
    receta_bytes = resp_receta.content
    receta_mime = "application/pdf" if url_receta.lower().endswith(".pdf") else "image/jpeg"

    blob_receta = bucket.blob(f"programa-cronico/{token}/receta.{'pdf' if receta_mime == 'application/pdf' else 'jpg'}")
    blob_receta.upload_from_string(receta_bytes, content_type=receta_mime)

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones SET url_receta_bucket = %s, actualizado_at = NOW() WHERE id = %s
    """, (blob_receta.public_url, inscripcion_id))
    conn.commit()

    datos_receta = extraer_datos_receta(receta_bytes, mime_type=receta_mime)

    # ── Regla 1 (receta): mismo paciente que el DNI ──
    nombre_dni_completo = f"{datos_dni.get('nombres', '')} {datos_dni.get('apellidos', '')}".strip()
    nombre_receta = datos_receta.get("nombre_paciente") or ""
    dni_en_receta = datos_receta.get("n_documento_paciente")

    coincide_por_dni = bool(dni_en_receta) and _dni_similar(dni_en_receta.strip(), n_documento.strip())
    coincide_por_nombre = _nombres_coinciden(nombre_dni_completo, nombre_receta)
    paciente_valido = coincide_por_dni or coincide_por_nombre

    #if not paciente_valido:
    #    motivo = "El paciente de la receta no coincide con el DNI / CE"
    #    cur.execute("""
    #        UPDATE sofi_programa_cronico_inscripciones
    #        SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s,
    #            datos_ocr_raw = %s, actualizado_at = NOW()
    #        WHERE id = %s
    #    """, (motivo, nombre_receta, json.dumps({
    #            "datos_dni": datos_dni,
    #            "datos_receta": datos_receta,
    #            "coincide_por_dni": coincide_por_dni,
    #            "coincide_por_nombre": coincide_por_nombre,
    #        }, ensure_ascii=False), inscripcion_id))
    #    conn.commit(); cur.close(); conn.close()
    #    return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 2: receta no mayor a 3 meses ──
    fecha_receta_str = datos_receta.get("fecha_emision")
    fecha_receta_parsed = None
    receta_vigente = False
    if fecha_receta_str:
        try:
            fecha_receta_parsed = datetime.strptime(fecha_receta_str, "%Y-%m-%d").date()
            receta_vigente = fecha_receta_parsed >= (datetime.now() - timedelta(days=92)).date()
        except ValueError:
            receta_vigente = False

    if not receta_vigente:
        motivo = "La receta tiene una antigüedad mayor a tres meses o no se pudo leer la fecha"
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s, fecha_receta = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (motivo, nombre_receta, fecha_receta_parsed, inscripcion_id))
        conn.commit(); cur.close(); conn.close()
        return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 3: clasificar programa según diagnósticos ──
    diagnosticos = datos_receta.get("diagnosticos", [])
    programa_detectado, cie10_match, descripcion_match = determinar_programa_por_diagnosticos(diagnosticos, cur)

    if not programa_detectado:
        motivo = "Ninguno de los diagnósticos de la receta corresponde a un programa disponible"
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s, fecha_receta = %s,
                datos_ocr_raw = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (motivo, nombre_receta, fecha_receta_parsed, json.dumps({
                "datos_dni": datos_dni,
                "datos_receta": datos_receta,
            }, ensure_ascii=False), inscripcion_id))
        conn.commit(); cur.close(); conn.close()
        return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 4: crónico complejo → manejo manual por correo ──
    if programa_detectado == "cronico_complejo":
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'derivado_manual', programa = %s, cie10_detectado = %s, diagnostico_detectado = %s,
                nombre_paciente_receta = %s, fecha_receta = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (programa_detectado, cie10_match, descripcion_match, nombre_receta, fecha_receta_parsed, inscripcion_id))
        conn.commit()

        dni_anverso_bucket = blob.public_url if url_documento else blob_a.public_url
        dni_reverso_bucket = None if url_documento else blob_r.public_url
        receta_bucket = blob_receta.public_url

        enviar_correo_cronico_complejo(
        nombres=datos_dni.get("nombres"),
        apellidos=datos_dni.get("apellidos"),
        n_documento=n_documento,
        celular=celular,
        correo_paciente=correo,
        programa=programa_detectado,
        dni_anverso_url=dni_anverso_bucket,
        dni_reverso_url=dni_reverso_bucket,
        receta_url=receta_bucket,
        )   

        cur.close(); conn.close()
        return (json.dumps({
            "ok": True, "token": token, "estado": "derivado_manual",
            "motivo_rechazo": "Tu caso requiere evaluación del equipo de Programa de Salud. Te contactarán pronto."
        }), 200, HEADERS)

    # ── Todo OK: listo para firmar ──
    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET estado = 'dni_recibido', programa = %s, cie10_detectado = %s, diagnostico_detectado = %s,
            nombre_paciente_receta = %s, fecha_receta = %s, fecha_dni_recibido = NOW(), actualizado_at = NOW()
        WHERE id = %s
    """, (programa_detectado, cie10_match, descripcion_match, nombre_receta, fecha_receta_parsed, inscripcion_id))
    conn.commit()
    cur.close()
    conn.close()

    return (json.dumps({
        "ok": True, "token": token, "estado": "dni_recibido", "programa": programa_detectado,
        "link_firma": f"https://sntspe.com/programa-cronico/{token}"
    }), 200, HEADERS)

def _pagina_mensaje(titulo, mensaje):
    return f"""<!DOCTYPE html>
    <html lang="es">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="robots" content="noindex, nofollow">
    <title>{titulo}</title>
    <link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
    <style>
        body {{
            font-family: 'Helvetica Neue', Arial, sans-serif;
            background: #f7f9fa;
            margin: 0;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
        }}
        .card {{
            background: #fff;
            border-radius: 12px;
            padding: 40px 28px;
            max-width: 420px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }}
        img.logo {{ width: 120px; margin-bottom: 20px; }}
        h2 {{ color: #0090A6; margin: 0 0 10px 0; }}
        p {{ color: #444; font-size: 15px; line-height: 1.5; margin: 0; }}
    </style>
    </head>
    <body>
        <div class="card">
            <img class="logo" src="/assets/logo_sanitasperu.png" alt="Sanitas">
            <h2>{titulo}</h2>
            <p>{mensaje}</p>
        </div>
    </body>
    </html>"""

# ──────────────────────────────────────────
# API: Programas Firmar consentimiento 
# ──────────────────────────────────────────
@con_log_errores
def handle_firmar_consentimiento_cronico(request):
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token")
    firma_carta_base64 = data.get("firma_carta_base64")
    firma_ficha_base64 = data.get("firma_ficha_base64")
    modalidad_atencion = data.get("modalidad_atencion")
    ip_firma = request.headers.get("X-Forwarded-For", request.remote_addr)

    # datos de la ficha que el paciente completa si no vinieron del DNI
    domicilio_form = data.get("domicilio")
    distrito_form = data.get("distrito")
    provincia_form = data.get("provincia")
    departamento_form = data.get("departamento")
    nombre_titular = data.get("nombre_titular")
    empleador = data.get("empleador")

    # datos opcionales de cuidador (solo carta, si requiere_cuidador = TRUE)
    firma_cuidador_base64 = data.get("firma_cuidador_base64")
    nombre_cuidador = data.get("nombre_cuidador")
    dni_cuidador = data.get("dni_cuidador")
    parentesco_cuidador = data.get("parentesco_cuidador")

    if not token or not firma_carta_base64 or not firma_ficha_base64:
        return (json.dumps({"error": "token, firma_carta_base64 y firma_ficha_base64 son requeridos"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT n_documento, nombres, apellidos, fecha_nacimiento, programa,
               requiere_cuidador, estado, celular, correo, domicilio, distrito,
               provincia, departamento, tipo_documento, diagnostico_detectado,
               dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Token no encontrado"}), 404, HEADERS)

    (n_documento, nombres, apellidos, fecha_nacimiento, programa,
     requiere_cuidador, estado, celular, correo, domicilio_bd, distrito_bd,
     provincia_bd, departamento_bd, tipo_documento, diagnostico_detectado,
     dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket) = row

    if estado != "dni_recibido":
        cur.close(); conn.close()
        return (json.dumps({"error": "El DNI aún no ha sido validado"}), 400, HEADERS)

    if requiere_cuidador and (not firma_cuidador_base64 or not dni_cuidador):
        cur.close(); conn.close()
        return (json.dumps({"error": "Este caso requiere datos y firma del cuidador"}), 400, HEADERS)

    # ── Priorizar lo extraído del DNI; si vino vacío, usar lo que completó el paciente ──
    domicilio_final = domicilio_bd or domicilio_form
    distrito_final = distrito_bd or distrito_form
    provincia_final = provincia_bd or provincia_form
    departamento_final = departamento_bd or departamento_form

    if not domicilio_final or not distrito_final or not provincia_final or not departamento_final:
        cur.close(); conn.close()
        return (json.dumps({"error": "Falta domicilio o distrito o provincia o departamento"}), 400, HEADERS)

    if programa == "krystal" and (not nombre_titular or not empleador):
        cur.close(); conn.close()
        return (json.dumps({"error": "Falta nombre_titular o empleador"}), 400, HEADERS)

    # ── Subir firmas a Cloud Storage ──
    bucket = storage_client.bucket(BUCKET_NAME)

    firma_carta_bytes = base64.b64decode(firma_carta_base64)
    bucket.blob(f"programa-cronico/{token}/firma_carta.png").upload_from_string(firma_carta_bytes, content_type="image/png")

    firma_ficha_bytes = base64.b64decode(firma_ficha_base64)
    bucket.blob(f"programa-cronico/{token}/firma_ficha.png").upload_from_string(firma_ficha_bytes, content_type="image/png")

    firma_cuidador_url = None
    if firma_cuidador_base64:
        firma_cuidador_bytes = base64.b64decode(firma_cuidador_base64)
        blob_c = bucket.blob(f"programa-cronico/{token}/firma_cuidador.png")
        blob_c.upload_from_string(firma_cuidador_bytes, content_type="image/png")
        firma_cuidador_url = blob_c.public_url

    fecha_hoy = datetime.now(LIMA_TZ).strftime("%d/%m/%Y")
    fecha_firma_completa = datetime.now(LIMA_TZ).strftime("%d/%m/%Y %H:%M")

     # ── Renderizar carta de compromiso ──
    template_carta = jinja_env.get_template(f"consentimiento_{programa}.html")
    html_carta = template_carta.render(
        nombres=nombres, apellidos=apellidos, n_documento=n_documento,
        tipo_documento=tipo_documento or "DNI",
        fecha_nacimiento=fecha_nacimiento, modalidad_atencion=modalidad_atencion,
        firma_afiliado_b64=firma_carta_base64,
        requiere_cuidador=requiere_cuidador,
        nombre_cuidador=nombre_cuidador, dni_cuidador=dni_cuidador,
        parentesco_cuidador=parentesco_cuidador,
        firma_cuidador_b64=firma_cuidador_base64,
        fecha_firma=fecha_firma_completa,
        ip_firma=ip_firma,
        token=token,
    )

    # ── Renderizar carta de compromiso ──
    html_carta = template_carta.render(
        nombres=nombres, apellidos=apellidos, n_documento=n_documento,
        tipo_documento=tipo_documento or "DNI",
        fecha_nacimiento=fecha_nacimiento, modalidad_atencion=modalidad_atencion,
        firma_afiliado_b64=firma_carta_base64,
        requiere_cuidador=requiere_cuidador,
        nombre_cuidador=nombre_cuidador, dni_cuidador=dni_cuidador,
        parentesco_cuidador=parentesco_cuidador,
        firma_cuidador_b64=firma_cuidador_base64,
        fecha_firma=fecha_firma_completa,
        ip_firma=ip_firma,
        token=token,
    )

    # ── Renderizar ficha de inscripción ──
    template_ficha = jinja_env.get_template(f"ficha_{programa}.html")
    html_ficha = template_ficha.render(
        nombres=nombres, apellidos=apellidos, n_documento=n_documento,
        fecha_nacimiento=fecha_nacimiento, telefono=celular, correo=correo,
        domicilio=domicilio_final, distrito=distrito_final, provincia=provincia_final,
        departamento=departamento_final,
        requiere_cuidador=requiere_cuidador,
        nombre_titular=nombre_titular, empleador=empleador,
        patologia_1=diagnostico_detectado, patologia_2="", patologia_3="",
        patologia_4="", patologia_5="", patologia_6="",
        firma_ficha_b64=firma_ficha_base64,
        fecha_inscripcion=fecha_hoy, fecha_firma=fecha_firma_completa,
        ip_firma=ip_firma,
        token=token,
    )

    # ── Combinar ambos documentos en un solo PDF ──
    estilos_combinados = _extraer_head_styles(html_carta) + "\n" + _extraer_head_styles(html_ficha)
    body_carta = _extraer_body(html_carta)
    body_ficha = _extraer_body(html_ficha)

    html_combinado = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<style>
{estilos_combinados}
.salto-documento {{ page-break-before: always; }}
</style>
</head>
<body>
{body_carta}
<div class="salto-documento"></div>
{body_ficha}
</body>
</html>"""

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    pdf_bytes = HTML(string=html_combinado, base_url=f"file://{BASE_DIR}/").write_pdf()
    hash_documento = hashlib.sha256(pdf_bytes).hexdigest()

    blob_pdf = bucket.blob(f"programa-cronico/{token}/consentimiento_firmado.pdf")
    blob_pdf.upload_from_string(pdf_bytes, content_type="application/pdf")
    pdf_url = blob_pdf.public_url

    # ── Notificar a Danaconnect (no debe bloquear el flujo si falla) ──
    danaconnect_enviado = False
    danaconnect_error_msg = None
    try:
        notificar_danaconnect(n_documento, nombres, apellidos, correo, programa, pdf_bytes,
                                dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket)
        danaconnect_enviado = True
    except Exception as e:
        danaconnect_error_msg = str(e)

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET estado = 'firmado', pdf_consentimiento_url = %s, hash_documento = %s,
            ip_firma = %s, firmado_at = NOW(), modalidad_atencion = %s,
            nombre_cuidador = %s, dni_cuidador = %s, parentesco_cuidador = %s,
            firma_cuidador_url = %s, firma_ficha_url = %s,
            domicilio = %s, distrito = %s, provincia = %s, departamento = %s,
            nombre_titular = %s, empleador = %s,
            danaconnect_enviado = %s, danaconnect_error = %s,
            actualizado_at = NOW()
        WHERE token = %s
    """, (pdf_url, hash_documento, ip_firma, modalidad_atencion,
          nombre_cuidador, dni_cuidador, parentesco_cuidador,
          firma_cuidador_url, f"programa-cronico/{token}/firma_ficha.png",
          domicilio_final, distrito_final, provincia_final, departamento_final,
          nombre_titular, empleador,
          danaconnect_enviado, danaconnect_error_msg, token))
    conn.commit()
    cur.close()
    conn.close()

    return (json.dumps({"ok": True, "pdf_url": pdf_url, "hash_documento": hash_documento}), 200, HEADERS)

# ─────────────────────
# API: /programa-cronico/{token} (GET - sirve el formulario web)
# ─────────────────────
def handle_ver_formulario_cronico(request, token):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT nombres, apellidos, n_documento, programa, estado, requiere_cuidador,
               fecha_nacimiento, celular, correo, domicilio, distrito, provincia, departamento,
               diagnostico_detectado
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    cur.close()
    conn.close()

    if not row:
        return (_pagina_mensaje("Enlace no válido", "Este enlace expiró o no es correcto. Contáctanos por WhatsApp si necesitas ayuda."), 404, {"Content-Type": "text/html; charset=utf-8"})

    (nombres, apellidos, n_documento, programa, estado, requiere_cuidador,
     fecha_nacimiento, celular, correo, domicilio, distrito, provincia, departamento,
     diagnostico_detectado) = row

    if estado == "firmado":
        return (_pagina_mensaje("¡Ya completaste tu inscripción!", "Revisa tu correo para ver tus constancias."), 200, {"Content-Type": "text/html; charset=utf-8"})

    if estado != "dni_recibido":
        return (_pagina_mensaje("Aún estamos validando tu documento", "Te escribiremos por WhatsApp en cuanto puedas firmar."), 200, {"Content-Type": "text/html; charset=utf-8"})

    if programa not in ("vive_sano_completo", "krystal"):
        return (_pagina_mensaje("Programa no reconocido", "Contáctanos por WhatsApp para ayudarte."), 400, {"Content-Type": "text/html; charset=utf-8"})

    template = jinja_env.get_template(f"formulario_firma_{programa}.html")
    html_renderizado = template.render(
        token=token, nombres=nombres, apellidos=apellidos, n_documento=n_documento,
        requiere_cuidador=requiere_cuidador, fecha_nacimiento=fecha_nacimiento,
        celular=celular, correo=correo or "", domicilio=domicilio or "",
        distrito=distrito or "", provincia=provincia or "", departamento=departamento or "",
        diagnostico_detectado=diagnostico_detectado or "",
    )
    return (html_renderizado, 200, {"Content-Type": "text/html; charset=utf-8"})


PROGRAMAS_VALIDOS_APROBACION_MANUAL = {"vive_sano_completo", "krystal"}

# ─────────────────────────────────────────────
# API: /programa-cronico/aprobar-manual
# Fuerza el paso a 'dni_recibido' sin re-evaluar OCR/reglas,
# para casos donde el rechazo automático fue un error.
# ─────────────────────────────────────────────
@con_log_errores
def handle_aprobar_manual_cronico(request):
    data = request.get_json(force=True, silent=True) or {}
    if data.get("pw") != ADMIN_PW_CRONICO:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    token = data.get("token")
    programa = (data.get("programa") or "").strip()
    diagnostico = (data.get("diagnostico") or "").strip() or "Aprobado manualmente por equipo de Programa de Salud"

    if not token:
        return (json.dumps({"error": "token es requerido"}), 400, HEADERS)
    if programa not in PROGRAMAS_VALIDOS_APROBACION_MANUAL:
        return (json.dumps({
            "error": f"programa debe ser uno de: {', '.join(PROGRAMAS_VALIDOS_APROBACION_MANUAL)}"
        }), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT estado, datos_ocr_raw, nombres, celular, correo
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Token no encontrado"}), 404, HEADERS)

    estado, datos_ocr_raw, nombres, celular, correo = row

    if estado not in ("rechazado", "derivado_manual"):
        cur.close(); conn.close()
        return (json.dumps({
            "error": f"Solo se pueden aprobar manualmente inscripciones rechazadas o derivadas (estado actual: {estado})"
        }), 409, HEADERS)

    # ── Auditoría: dejar registro de que fue una aprobación manual, sin pisar el OCR original ──
    try:
        auditoria = json.loads(datos_ocr_raw) if isinstance(datos_ocr_raw, str) else (datos_ocr_raw or {})
    except (TypeError, json.JSONDecodeError):
        auditoria = {}
    if not isinstance(auditoria, dict):
        auditoria = {}

    auditoria["aprobacion_manual"] = {
        "programa_asignado": programa,
        "diagnostico_asignado": diagnostico,
        "fecha": datetime.now(LIMA_TZ).isoformat(),
    }

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET estado = 'dni_recibido', programa = %s, diagnostico_detectado = %s,
            motivo_rechazo = NULL, datos_ocr_raw = %s,
            fecha_dni_recibido = NOW(), intentos_recordatorio = 0,
            fecha_ultimo_recordatorio = NULL, actualizado_at = NOW()
        WHERE token = %s
    """, (programa, diagnostico, json.dumps(auditoria, ensure_ascii=False), token))
    conn.commit()
    cur.close(); conn.close()

    # ── Avisar al paciente que ya puede firmar (best-effort, no bloquea la respuesta) ──
    try:
        _enviar_recordatorio_cronico(token, celular, nombres, correo)
    except Exception as e:
        logging.warning(f"[aprobar-manual-cronico] No se pudo notificar al paciente: {e}")

    return (json.dumps({
        "ok": True, "token": token, "estado": "dni_recibido", "programa": programa,
        "link_firma": f"https://sntspe.com/programa-cronico/{token}"
    }), 200, HEADERS)

# ─────────────────────────────────────────────
# API: /programa-cronico/reprocesar
# Vuelve a correr OCR + reglas de validación sobre los archivos
# ya almacenados, sin que el paciente reenvíe nada por WhatsApp.
# ─────────────────────────────────────────────
@con_log_errores
def handle_reprocesar_cronico(request):
    data = request.get_json(force=True, silent=True) or {}
    if data.get("pw") != ADMIN_PW_CRONICO:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    token = data.get("token")
    if not token:
        return (json.dumps({"error": "token es requerido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, estado, celular, n_documento, correo,
               dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Token no encontrado"}), 404, HEADERS)

    (inscripcion_id, estado, celular, n_documento, correo,
     dni_anverso_url, dni_reverso_url, url_receta_bucket) = row

    if estado not in ("rechazado", "derivado_manual"):
        cur.close(); conn.close()
        return (json.dumps({
            "error": f"Solo se pueden reprocesar inscripciones rechazadas o derivadas (estado actual: {estado})"
        }), 409, HEADERS)

    if not dni_anverso_url or not url_receta_bucket:
        cur.close(); conn.close()
        return (json.dumps({"error": "Faltan los archivos originales (DNI o receta) para reprocesar"}), 409, HEADERS)

    # ── Redescargar los archivos originales desde el bucket ──
    try:
        bytes_anverso = descargar_blob_desde_url(dni_anverso_url)
        mime_anverso = _content_type_por_url(dni_anverso_url)
        bytes_reverso = descargar_blob_desde_url(dni_reverso_url) if dni_reverso_url else None
        mime_reverso = _content_type_por_url(dni_reverso_url) if dni_reverso_url else None
        bytes_receta = descargar_blob_desde_url(url_receta_bucket)
        mime_receta = _content_type_por_url(url_receta_bucket)
    except Exception as e:
        cur.close(); conn.close()
        return (json.dumps({"error": f"No se pudieron descargar los archivos originales: {str(e)}"}), 500, HEADERS)

    # ── Volver a correr OCR (puede corregir un error de lectura anterior) ──
    datos_dni = extraer_datos_dni(bytes_anverso, mime_anverso, bytes_reverso, mime_reverso)
    datos_receta = extraer_datos_receta(bytes_receta, mime_type=mime_receta)

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET nombres = %s, apellidos = %s, fecha_nacimiento = %s,
            domicilio = %s, distrito = %s, provincia = %s, departamento = %s,
            tipo_documento = %s, actualizado_at = NOW()
        WHERE id = %s
    """, (datos_dni.get("nombres"), datos_dni.get("apellidos"), datos_dni.get("fecha_nacimiento"),
          datos_dni.get("domicilio"), datos_dni.get("distrito"), datos_dni.get("provincia"),
          datos_dni.get("departamento"), datos_dni.get("tipo_documento"), inscripcion_id))
    conn.commit()

    # ── Regla 1: paciente de la receta coincide con el DNI ──
    nombre_dni_completo = f"{datos_dni.get('nombres', '')} {datos_dni.get('apellidos', '')}".strip()
    nombre_receta = datos_receta.get("nombre_paciente") or ""
    dni_en_receta = datos_receta.get("n_documento_paciente")

    coincide_por_dni = bool(dni_en_receta) and _dni_similar(dni_en_receta.strip(), (n_documento or "").strip())
    coincide_por_nombre = _nombres_coinciden(nombre_dni_completo, nombre_receta)
    paciente_valido = coincide_por_dni or coincide_por_nombre

    if not paciente_valido:
        motivo = "El paciente de la receta no coincide con el DNI / CE"
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s,
                datos_ocr_raw = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (motivo, nombre_receta, json.dumps({
                "datos_dni": datos_dni, "datos_receta": datos_receta,
                "coincide_por_dni": coincide_por_dni, "coincide_por_nombre": coincide_por_nombre,
                "reprocesado": True,
            }, ensure_ascii=False), inscripcion_id))
        conn.commit(); cur.close(); conn.close()
        return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 2: receta no mayor a 3 meses ──
    fecha_receta_str = datos_receta.get("fecha_emision")
    fecha_receta_parsed = None
    receta_vigente = False
    if fecha_receta_str:
        try:
            fecha_receta_parsed = datetime.strptime(fecha_receta_str, "%Y-%m-%d").date()
            receta_vigente = fecha_receta_parsed >= (datetime.now() - timedelta(days=92)).date()
        except ValueError:
            receta_vigente = False

    if not receta_vigente:
        motivo = "La receta tiene una antigüedad mayor a tres meses o no se pudo leer la fecha"
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s, fecha_receta = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (motivo, nombre_receta, fecha_receta_parsed, inscripcion_id))
        conn.commit(); cur.close(); conn.close()
        return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 3: clasificar programa según diagnósticos ──
    diagnosticos = datos_receta.get("diagnosticos", [])
    programa_detectado, cie10_match, descripcion_match = determinar_programa_por_diagnosticos(diagnosticos, cur)

    if not programa_detectado:
        motivo = "Ninguno de los diagnósticos de la receta corresponde a un programa disponible"
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'rechazado', motivo_rechazo = %s, nombre_paciente_receta = %s, fecha_receta = %s,
                datos_ocr_raw = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (motivo, nombre_receta, fecha_receta_parsed, json.dumps({
                "datos_dni": datos_dni, "datos_receta": datos_receta, "reprocesado": True,
            }, ensure_ascii=False), inscripcion_id))
        conn.commit(); cur.close(); conn.close()
        return (json.dumps({"ok": True, "token": token, "estado": "rechazado", "motivo_rechazo": motivo}), 200, HEADERS)

    # ── Regla 4: crónico complejo → manejo manual por correo ──
    if programa_detectado == "cronico_complejo":
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET estado = 'derivado_manual', programa = %s, cie10_detectado = %s, diagnostico_detectado = %s,
                nombre_paciente_receta = %s, fecha_receta = %s, actualizado_at = NOW()
            WHERE id = %s
        """, (programa_detectado, cie10_match, descripcion_match, nombre_receta, fecha_receta_parsed, inscripcion_id))
        conn.commit()

        try:
            enviar_correo_cronico_complejo(
                nombres=datos_dni.get("nombres"), apellidos=datos_dni.get("apellidos"),
                n_documento=n_documento, celular=celular, correo_paciente=correo,
                programa=programa_detectado, dni_anverso_url=dni_anverso_url,
                dni_reverso_url=dni_reverso_url, receta_url=url_receta_bucket,
            )
        except Exception as e:
            logging.warning(f"[reprocesar-cronico] No se pudo enviar correo derivado_manual: {e}")

        cur.close(); conn.close()
        return (json.dumps({
            "ok": True, "token": token, "estado": "derivado_manual",
            "motivo_rechazo": "Tu caso requiere evaluación del equipo de Programa de Salud. Te contactarán pronto."
        }), 200, HEADERS)

    # ── Todo OK: listo para firmar ──
    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET estado = 'dni_recibido', programa = %s, cie10_detectado = %s, diagnostico_detectado = %s,
            nombre_paciente_receta = %s, fecha_receta = %s, fecha_dni_recibido = NOW(),
            motivo_rechazo = NULL, intentos_recordatorio = 0, fecha_ultimo_recordatorio = NULL,
            actualizado_at = NOW()
        WHERE id = %s
    """, (programa_detectado, cie10_match, descripcion_match, nombre_receta, fecha_receta_parsed, inscripcion_id))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({
        "ok": True, "token": token, "estado": "dni_recibido", "programa": programa_detectado,
        "link_firma": f"https://sntspe.com/programa-cronico/{token}"
    }), 200, HEADERS)

# ──────────────────────────────────
# API: /Reenvio manual de correos
# ──────────────────────────────────
@con_log_errores
def handle_reenviar_danaconnect(request):
    data = request.get_json(force=True, silent=True) or {}

    if data.get("pw") != ADMIN_PW_CRONICO:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    token = data.get("token")

    if not token:
        return (json.dumps({"error": "token es requerido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT n_documento, nombres, apellidos, correo, programa, estado,
               pdf_consentimiento_url, danaconnect_enviado,
               dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Token no encontrado"}), 404, HEADERS)

    (n_documento, nombres, apellidos, correo, programa, estado,
     pdf_consentimiento_url, danaconnect_enviado,
     dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket) = row

    if estado != "firmado":
        cur.close(); conn.close()
        return (json.dumps({"error": f"Esta inscripción no está firmada (estado actual: {estado})"}), 400, HEADERS)

    if not pdf_consentimiento_url:
        cur.close(); conn.close()
        return (json.dumps({"error": "No hay PDF de consentimiento generado para esta inscripción"}), 400, HEADERS)

    # ── Descargar el PDF ya generado desde el bucket ──
    try:
        #resp_pdf = requests.get(pdf_consentimiento_url, timeout=20)
        #resp_pdf.raise_for_status()
        #pdf_bytes = resp_pdf.content
        pdf_bytes = descargar_blob_desde_url(pdf_consentimiento_url)
    except Exception as e:
        cur.close(); conn.close()
        return (json.dumps({"error": f"No se pudo descargar el PDF: {str(e)}"}), 500, HEADERS)

    # ── Reintentar el envío a Danaconnect ──
    danaconnect_enviado_nuevo = False
    danaconnect_error_msg = None
    try:
        notificar_danaconnect(n_documento, nombres, apellidos, correo, programa, pdf_bytes,
                                dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket)
        danaconnect_enviado_nuevo = True
    except Exception as e:
        danaconnect_error_msg = str(e)

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET danaconnect_enviado = %s, danaconnect_error = %s, actualizado_at = NOW()
        WHERE token = %s
    """, (danaconnect_enviado_nuevo, danaconnect_error_msg, token))
    conn.commit()
    cur.close()
    conn.close()

    if danaconnect_enviado_nuevo:
        return (json.dumps({"ok": True, "token": token, "danaconnect_enviado": True}), 200, HEADERS)
    else:
        return (json.dumps({"ok": False, "token": token, "danaconnect_enviado": False, "error": danaconnect_error_msg}), 502, HEADERS)


# ─────────────────────────────────────────────
# TAREA: /programa-cronico/recordatorio
# SMS a inscripciones con DNI recibido que aún no firman
# Cloud Scheduler: 1 vez al día, ej. 10:00 lun-sab
# ─────────────────────────────────────────────
INTERVALOS_RECORDATORIO_CRONICO = {
    1: timedelta(hours=2),   # 1er recordatorio
    2: timedelta(hours=6),   # 2do: recordatorio
    3: timedelta(hours=24),     # 3ro (último): 
}
MAX_INTENTOS_RECORDATORIO_CRONICO = 3
DANACONNECT_CONVERSATION_ID_RECORDATORIO_CRONICO = "605087"

def enviar_correo_recordatorio_cronico(token, correo, nombres):
    """
    Envía el recordatorio por correo vía Danaconnect (conversación 605086).
    Usa Basic Auth (DANACONNECT_USER/DANACONNECT_PASSWORD), igual que
    iniciar_conversacion_danaconnect / iniciar_conversacion_cronico_complejo,
    NO el header DANACONNECT_AUTH que usa enviar_danaconnect() para SMS.
    """
    if not correo:
        return False

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link = f"{base_url}/programa-cronico/{token}"

    # Solo el primer nombre (nombres puede traer más de uno separado por espacio)
    primer_nombre = (nombres or "").strip().split(" ")[0] if nombres else ""

    url = f"https://appserv.danaconnect.com/api/1.0/rest/conversation/{DANACONNECT_CONVERSATION_ID_RECORDATORIO_CRONICO}/start/data"
    payload = {
        "EMAIL": correo,
        "NOMBRE": primer_nombre,
        "URL": link,
    }

    try:
        resp = requests.post(
            url,
            auth=(DANACONNECT_USER, DANACONNECT_PASSWORD),
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        return True
    except Exception as e:
        print(f"Error al enviar correo de recordatorio a {correo}: {e}")
        return False

@con_log_errores
def handle_forzar_recordatorio_cronico(request):
    data = request.get_json(force=True, silent=True) or {}
    if data.get("pw") != ADMIN_PW_CRONICO:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    token = data.get("token")
    if not token:
        return (json.dumps({"error": "token es requerido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT estado, celular, nombres, correo, intentos_recordatorio
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Token no encontrado"}), 404, HEADERS)

    estado, celular, nombres, correo, intentos = row

    if estado != "dni_recibido":
        cur.close(); conn.close()
        return (json.dumps({
            "error": f"Solo se puede forzar el envío en estado dni_recibido (actual: {estado})"
        }), 409, HEADERS)

    ahora = datetime.now(LIMA_TZ).replace(tzinfo=None)
    siguiente_intento = (intentos or 0) + 1

    enviado = _enviar_recordatorio_cronico(token, celular, nombres, correo)

    if enviado:
        cur.execute("""
            UPDATE sofi_programa_cronico_inscripciones
            SET intentos_recordatorio = %s, fecha_ultimo_recordatorio = %s
            WHERE token = %s
        """, (siguiente_intento, ahora, token))
        conn.commit()

    cur.close(); conn.close()

    return (json.dumps({"ok": enviado, "token": token, "intento_numero": siguiente_intento}), 200, HEADERS)

def _enviar_recordatorio_cronico(token, celular, nombres, correo=None):
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link = f"{base_url}/programa-cronico/{token}"
    nombre_limpio = _limpiar_nombre_sms(nombres or "")
    saludo = f"Hola {nombre_limpio}, " if nombre_limpio else "Hola, "
    mensaje = f"Sanitas: {saludo}completa y firma tu inscripcion al Programa de Salud: {link}"

    numero = _normalizar_celular_pe(celular)
    sms_enviado = False
    if numero:
        enviar_danaconnect(numero, mensaje)
        sms_aviso(numero, "recordatorio programa cronico")
        sms_enviado = True

    correo_enviado = enviar_correo_recordatorio_cronico(token, correo, nombres)

    # Se considera "enviado" si al menos uno de los dos canales funcionó
    return sms_enviado or correo_enviado


def handle_recordatorio_cronico(request):
    conn = None
    cur = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT token, celular, nombres, correo, intentos_recordatorio,
                   COALESCE(fecha_ultimo_recordatorio, fecha_dni_recibido) AS fecha_referencia
            FROM sofi_programa_cronico_inscripciones
            WHERE estado = 'dni_recibido'
              AND COALESCE(intentos_recordatorio, 0) < %s
              AND fecha_dni_recibido IS NOT NULL
        """, (MAX_INTENTOS_RECORDATORIO_CRONICO,))
        candidatos = cur.fetchall()

        ahora = datetime.now(LIMA_TZ).replace(tzinfo=None)
        enviados = 0
        omitidos_sin_celular = 0
        errores = []

        for token, celular, nombres, correo, intentos, fecha_referencia in candidatos:
            try:
                intentos = intentos or 0
                siguiente_intento = intentos + 1
                espera = INTERVALOS_RECORDATORIO_CRONICO.get(siguiente_intento)
                if not espera:
                    continue
                if not celular and not correo:
                    omitidos_sin_celular += 1
                    continue

                if fecha_referencia and getattr(fecha_referencia, "tzinfo", None):
                    fecha_referencia = fecha_referencia.replace(tzinfo=None)

                if not fecha_referencia or (ahora - fecha_referencia) < espera:
                    continue

                if _enviar_recordatorio_cronico(token, celular, nombres, correo):
                    cur.execute("""
                        UPDATE sofi_programa_cronico_inscripciones
                        SET intentos_recordatorio = %s, fecha_ultimo_recordatorio = %s
                        WHERE token = %s
                    """, (siguiente_intento, ahora, token))
                    conn.commit()
                    enviados += 1

            except Exception as e_candidato:
                conn.rollback()
                errores.append({"token": token, "error": str(e_candidato)})
                continue

        return (json.dumps({
            "mensaje": f"Recordatorios enviados: {enviados} de {len(candidatos)} candidatos evaluados",
            "omitidos_sin_celular": omitidos_sin_celular,
            "errores": errores
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)

    finally:
        if cur:
            cur.close()
        if conn:
            conn.close()

ADMIN_PW_CRONICO = os.environ.get("ADMIN_PW_CRONICO")
# ─────────────────────
# API: /programa-cronico/reporte
def handle_reporte_cronico(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_CRONICO:
        return "No autorizado", 401

    filtro_estado = request.args.get("estado", "")
    filtro_programa = request.args.get("programa", "")
    filtro_busqueda = request.args.get("q", "").strip()

    conn = get_connection()
    with conn.cursor() as cur:
        query = """
            SELECT token, n_documento, nombres, apellidos, celular, correo,
                   programa, estado, motivo_rechazo, diagnostico_detectado, cie10_detectado,
                   dni_foto_anverso_url, dni_foto_reverso_url, url_receta,
                   pdf_consentimiento_url, danaconnect_enviado, danaconnect_error,
                   creado_at, firmado_at,
                   intentos_recordatorio, fecha_ultimo_recordatorio, datos_ocr_raw
            FROM sofi_programa_cronico_inscripciones
            WHERE 1=1
        """
        params = []

        if filtro_estado:
            query += " AND estado = %s"
            params.append(filtro_estado)

        if filtro_programa:
            query += " AND programa = %s"
            params.append(filtro_programa)

        if filtro_busqueda:
            query += " AND (n_documento ILIKE %s OR nombres ILIKE %s OR apellidos ILIKE %s OR celular ILIKE %s)"
            comodin = f"%{filtro_busqueda}%"
            params.extend([comodin, comodin, comodin, comodin])

        query += " ORDER BY creado_at DESC LIMIT 500"
        cur.execute(query, params)
        registros = cur.fetchall()
    conn.close()

    COLOR_ESTADO = {
        "firmado": "#2e7d32",
        "rechazado": "#c62828",
        "derivado_manual": "#ef6c00",
        "dni_recibido": "#0090A6",
        "esperando_dni": "#888",
    }

    filas_html = ""
    for row in registros:
        (token, n_documento, nombres, apellidos, celular, correo,
         programa, estado, motivo_rechazo, diagnostico_detectado, cie10_detectado,
         dni_anverso_url, dni_reverso_url, url_receta,
         pdf_url, dana_enviado, dana_error,
         creado_at, firmado_at,
         intentos_recordatorio, fecha_ultimo_recordatorio, datos_ocr_raw) = row

        color_estado = COLOR_ESTADO.get(estado, "#555")
        nombre_completo = f"{nombres or ''} {apellidos or ''}".strip() or "—"

        links = []
        if dni_anverso_url:
            links.append(f'<a href="{dni_anverso_url}" target="_blank">DNI</a>')
        if dni_reverso_url:
            links.append(f'<a href="{dni_reverso_url}" target="_blank">DNI reverso</a>')
        if url_receta:
            links.append(f'<a href="{url_receta}" target="_blank">Receta orig.</a>')
        if pdf_url:
            links.append(f'<a href="{pdf_url}" target="_blank"><strong>PDF firmado</strong></a>')
        links_html = " · ".join(links) if links else "—"

        dana_badge = ""
        if estado == "firmado":
            if dana_enviado:
                dana_badge = '<span style="color:#2e7d32;">✓ Dana</span>'
            else:
                error_txt = (dana_error or "sin enviar")[:60]
                dana_badge = (
                    f'<span style="color:#c62828;" title="{dana_error or ""}">✗ Dana ({error_txt})</span>'
                    f'<br><button class="btn-reenviar" data-token="{token}" onclick="reenviarDana(this)">🔁 Reenviar</button>'
                )

        # ── Badge de recordatorios ──
        recordatorio_html = "—"
        if estado == "dni_recibido":
            intentos_val = intentos_recordatorio or 0
            if intentos_val > 0:
                fecha_ult = fecha_ultimo_recordatorio.strftime('%d/%m/%Y %H:%M') if fecha_ultimo_recordatorio else "—"
                color_rec = "#ef6c00" if intentos_val < MAX_INTENTOS_RECORDATORIO_CRONICO else "#c62828"
                recordatorio_html = f'<span style="color:{color_rec};">{intentos_val}/{MAX_INTENTOS_RECORDATORIO_CRONICO} enviados</span><br><span style="font-size:11px;color:#888;">{fecha_ult}</span>'
            else:
                recordatorio_html = '<span style="color:#888;">Sin enviar</span>'

        motivo_html = f'<div style="font-size:12px;color:#c62828;">{motivo_rechazo}</div>' if motivo_rechazo else ""
        if estado == "rechazado" and datos_ocr_raw:
            import html as html_lib
            ocr_pretty = html_lib.escape(json.dumps(datos_ocr_raw, indent=2, ensure_ascii=False))
            motivo_html += f'<details style="margin-top:4px;"><summary style="cursor:pointer;font-size:11px;color:#888;">Ver OCR crudo</summary><pre style="font-size:10px;white-space:pre-wrap;max-width:280px;overflow-x:auto;">{ocr_pretty}</pre></details>'

        if estado in ("rechazado", "derivado_manual"):
            motivo_html += f'<br><button class="btn-reprocesar" data-token="{token}" onclick="reprocesarCronico(this)">🔄 Reprocesar</button>'
            motivo_html += f"""
            <div style="margin-top:6px;">
                <select class="select-aprobar" id="sel_{token}" style="font-size:11px;padding:3px;border-radius:4px;">
                    <option value="vive_sano_completo">Vive Sano</option>
                    <option value="krystal">Krystal</option>
                </select>
                <button class="btn-aprobar" data-token="{token}" onclick="aprobarManualCronico(this)">✅ Aprobar manual</button>
            </div>"""

        if estado == "dni_recibido":
            motivo_html += f'<br><button class="btn-forzar" data-token="{token}" onclick="forzarRecordatorioCronico(this)">📨 Enviar recordatorio ahora</button>'
        
        diagnostico_html = f"{diagnostico_detectado or ''} ({cie10_detectado or ''})" if diagnostico_detectado else "—"

        filas_html += f"""
        <tr>
            <td>{creado_at.strftime('%d/%m/%Y %H:%M') if creado_at else '—'}</td>
            <td>{nombre_completo}<br><span style="font-size:12px;color:#888;">{n_documento or ''}</span></td>
            <td>{celular or '—'}<br><span style="font-size:12px;color:#888;">{correo or ''}</span></td>
            <td>{programa or '—'}</td>
            <td>{diagnostico_html}</td>
            <td><span style="color:{color_estado};font-weight:600;">{estado}</span>{motivo_html}</td>
            <td>{recordatorio_html}</td>
            <td>{links_html}</td>
            <td>{dana_badge}</td>
            <td><span style="font-size:11px;color:#aaa;">{token}</span></td>
        </tr>"""

    opciones_estado = ["", "esperando_dni", "dni_recibido", "rechazado", "derivado_manual", "firmado"]
    opciones_estado_html = "".join(
        f'<option value="{e}" {"selected" if e == filtro_estado else ""}>{e or "Todos"}</option>'
        for e in opciones_estado
    )

    opciones_programa = ["", "vive_sano_completo", "krystal", "cronico_complejo"]
    opciones_programa_html = "".join(
        f'<option value="{p}" {"selected" if p == filtro_programa else ""}>{p or "Todos"}</option>'
        for p in opciones_programa
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Reporte Programa Crónico</title>
<link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
<style>
    body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background: #f7f9fa; margin: 0; padding: 20px; color: #222; }}
    h1 {{ color: #0090A6; font-size: 20px; margin-bottom: 4px; }}
    .subtitulo {{ color: #888; font-size: 13px; margin-bottom: 20px; }}
    .filtros {{ background: #fff; border-radius: 10px; padding: 14px 18px; margin-bottom: 16px; display: flex; gap: 12px; align-items: end; flex-wrap: wrap; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
    .filtros label {{ font-size: 12px; font-weight: 600; display: block; margin-bottom: 4px; }}
    .filtros input, .filtros select {{ padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }}
    .filtros button {{ padding: 8px 16px; background: #0090A6; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }}
    table {{ width: 100%; border-collapse: collapse; background: #fff; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
    th {{ background: #0090A6; color: #fff; text-align: left; padding: 10px 12px; font-size: 12px; }}
    td {{ padding: 10px 12px; font-size: 13px; border-bottom: 1px solid #eee; vertical-align: top; }}
    tr:hover {{ background: #f0f7f8; }}
    .contador {{ font-size: 12px; color: #888; margin-bottom: 10px; }}
    a {{ color: #0090A6; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .btn-reenviar {{ margin-top:4px; padding:4px 10px; font-size:11px; background:#0090A6; color:#fff; border:none; border-radius:5px; cursor:pointer; }}
    .btn-reenviar:disabled {{ background:#9bb6d6; cursor:not-allowed; }}
    .btn-reprocesar {{ margin-top:4px; padding:4px 10px; font-size:11px; background:#ef6c00; color:#fff; border:none; border-radius:5px; cursor:pointer; }}
    .btn-reprocesar:disabled {{ background:#f0b783; cursor:not-allowed; }}
    .btn-aprobar {{ margin-top:2px; padding:4px 10px; font-size:11px; background:#2e7d32; color:#fff; border:none; border-radius:5px; cursor:pointer; }}
    .btn-aprobar:disabled {{ background:#a5c9a8; cursor:not-allowed; }}
    .btn-forzar {{ margin-top:4px; padding:4px 10px; font-size:11px; background:#0090A6; color:#fff; border:none; border-radius:5px; cursor:pointer; }}
    .btn-forzar:disabled {{ background:#8fc7d0; cursor:not-allowed; }}
</style>
</head>
<body>

<h1>Reporte — Programa de Pacientes Crónicos</h1>
<div class="subtitulo">Vive Sano · Krystal · Crónico Complejo</div>

<form class="filtros" method="GET" action="/programa-cronico/reporte">
    <input type="hidden" name="pw" value="{pw}">
    <div>
        <label>Buscar (DNI, nombre, celular)</label>
        <input type="text" name="q" value="{filtro_busqueda}" placeholder="Ej. 47313070">
    </div>
    <div>
        <label>Estado</label>
        <select name="estado">{opciones_estado_html}</select>
    </div>
    <div>
        <label>Programa</label>
        <select name="programa">{opciones_programa_html}</select>
    </div>
    <button type="submit">Filtrar</button>
    <a href="/programa-cronico/estadistica?pw={pw}" style="padding:8px 16px;background:#0090A6;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;text-decoration:none;display:inline-block;">📊 Ver estadísticas</a>
</form>

<div class="contador">{len(registros)} registro(s) (máx. 500)</div>

<table>
    <thead>
        <tr>
            <th>Fecha</th>
            <th>Paciente</th>
            <th>Contacto</th>
            <th>Programa</th>
            <th>Diagnóstico (CIE-10)</th>
            <th>Estado</th>
            <th>Recordatorios</th>
            <th>Archivos</th>
            <th>Danaconnect</th>
            <th>Token</th>
        </tr>
    </thead>
    <tbody>
        {filas_html}
    </tbody>
    </table>

    <script>
    async function reenviarDana(btn) {{
        const token = btn.dataset.token;
        if (!confirm("¿Reenviar el consentimiento a Danaconnect?")) return;
        btn.disabled = true;
        btn.textContent = "Enviando...";
        try {{
            const resp = await fetch("/programa-cronico/reenviar-danaconnect", {{
                method: "POST",
                headers: {{ "Content-Type": "application/json" }},
                body: JSON.stringify({{ pw: "{pw}", token: token }})
            }});
            const data = await resp.json();
            if (resp.ok && data.ok) {{
                btn.closest("td").innerHTML = '<span style="color:#2e7d32;">✓ Dana</span>';
            }} else {{
                alert("Error: " + (data.error || "No se pudo reenviar"));
                btn.disabled = false;
                btn.textContent = "🔁 Reenviar";
            }}
        }} catch (e) {{
            alert("Error de conexión");
            btn.disabled = false;
            btn.textContent = "🔁 Reenviar";
        }}
    }}

    async function reprocesarCronico(btn) {{
        const token = btn.dataset.token;
        if (!confirm("¿Volver a leer los documentos y re-evaluar esta inscripción?")) return;
        btn.disabled = true;
        btn.textContent = "Procesando...";
        try {{
            const resp = await fetch("/programa-cronico/reprocesar", {{
                method: "POST",
                headers: {{ "Content-Type": "application/json" }},
                body: JSON.stringify({{ pw: "{pw}", token: token }})
            }});
            const data = await resp.json();
            if (resp.ok && data.ok) {{
                alert("Nuevo estado: " + data.estado + (data.motivo_rechazo ? "\\nMotivo: " + data.motivo_rechazo : ""));
                location.reload();
            }} else {{
                alert("Error: " + (data.error || "No se pudo reprocesar"));
                btn.disabled = false;
                btn.textContent = "🔄 Reprocesar";
            }}
        }} catch (e) {{
            alert("Error de conexión");
            btn.disabled = false;
            btn.textContent = "🔄 Reprocesar";
        }}
    }}

    async function aprobarManualCronico(btn) {{
        const token = btn.dataset.token;
        const select = document.getElementById("sel_" + token);
        const programa = select.value;
        const nombrePrograma = select.options[select.selectedIndex].text;

        if (!confirm(`¿Aprobar manualmente esta inscripción como "${{nombrePrograma}}"? Se le avisará al paciente que ya puede firmar.`)) return;

        btn.disabled = true;
        btn.textContent = "Aprobando...";
        try {{
            const resp = await fetch("/programa-cronico/aprobar-manual", {{
                method: "POST",
                headers: {{ "Content-Type": "application/json" }},
                body: JSON.stringify({{ pw: "{pw}", token: token, programa: programa }})
            }});
            const data = await resp.json();
            if (resp.ok && data.ok) {{
                alert("Aprobado. El paciente ya puede firmar en: " + data.link_firma);
                location.reload();
            }} else {{
                alert("Error: " + (data.error || "No se pudo aprobar"));
                btn.disabled = false;
                btn.textContent = "✅ Aprobar manual";
            }}
        }} catch (e) {{
            alert("Error de conexión");
            btn.disabled = false;
            btn.textContent = "✅ Aprobar manual";
        }}
    }}

    async function forzarRecordatorioCronico(btn) {{
        const token = btn.dataset.token;
        if (!confirm("¿Enviar el recordatorio (SMS + correo) ahora mismo, sin esperar el intervalo programado?")) return;

        btn.disabled = true;
        btn.textContent = "Enviando...";
        try {{
            const resp = await fetch("/programa-cronico/forzar-recordatorio", {{
                method: "POST",
                headers: {{ "Content-Type": "application/json" }},
                body: JSON.stringify({{ pw: "{pw}", token: token }})
            }});
            const data = await resp.json();
            if (resp.ok && data.ok) {{
                alert("Recordatorio #" + data.intento_numero + " enviado correctamente.");
                location.reload();
            }} else {{
                alert("Error: " + (data.error || "No se pudo enviar"));
                btn.disabled = false;
                btn.textContent = "📨 Enviar recordatorio ahora";
            }}
        }} catch (e) {{
            alert("Error de conexión");
            btn.disabled = false;
            btn.textContent = "📨 Enviar recordatorio ahora";
        }}
    }}
    </script>
    </body>
    </html>"""

    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})

# ─────────────────────────────────────────────────────────────
# API: /programa-cronico/estadistica
# ─────────────────────────────────────────────────────────────
def handle_estadistica_cronico(request):
    import html as html_lib
    from collections import Counter

    pw = request.args.get("pw")
    if pw != ADMIN_PW_CRONICO:
        return "No autorizado", 401

    fecha_inicio = request.args.get("fecha_inicio", "").strip()
    fecha_fin = request.args.get("fecha_fin", "").strip()
    filtro_programa = request.args.get("programa", "").strip()

    # Sin fechas -> por defecto, el mes actual completo hasta hoy
    if not fecha_inicio or not fecha_fin:
        hoy = datetime.now(LIMA_TZ).date()
        fecha_inicio = hoy.replace(day=1).strftime("%Y-%m-%d")
        fecha_fin = hoy.strftime("%Y-%m-%d")

    try:
        fecha_inicio_obj = datetime.strptime(fecha_inicio, "%Y-%m-%d").date()
        fecha_fin_obj = datetime.strptime(fecha_fin, "%Y-%m-%d").date()
    except ValueError:
        return ("""<html><body style="font-family:Arial;padding:40px;">
            <h2>Error en las fechas</h2><p>Usa el formato YYYY-MM-DD.</p>
            <a href="javascript:history.back()">Volver</a></body></html>""",
            400, {"Content-Type": "text/html; charset=utf-8"})

    if fecha_inicio_obj > fecha_fin_obj:
        return ("""<html><body style="font-family:Arial;padding:40px;">
            <h2>Rango de fechas inválido</h2><p>La fecha desde no puede ser mayor que la fecha hasta.</p>
            <a href="javascript:history.back()">Volver</a></body></html>""",
            400, {"Content-Type": "text/html; charset=utf-8"})

    where = "WHERE creado_at >= %s::date AND creado_at < (%s::date + INTERVAL '1 day')"
    params_base = [fecha_inicio, fecha_fin]
    if filtro_programa:
        where += " AND programa = %s"
        params_base.append(filtro_programa)

    conn = None
    try:
        conn = get_connection()
        with conn.cursor() as cur:
            cur.execute(f"""
                SELECT
                    COUNT(*) AS total,
                    COUNT(*) FILTER (WHERE estado = 'firmado') AS firmados,
                    COUNT(*) FILTER (WHERE estado = 'rechazado') AS rechazados,
                    COUNT(*) FILTER (WHERE estado = 'dni_recibido') AS dni_recibido,
                    COUNT(*) FILTER (WHERE estado = 'esperando_dni') AS esperando_dni,
                    COUNT(*) FILTER (WHERE estado = 'derivado_manual') AS derivados_manual
                FROM sofi_programa_cronico_inscripciones {where}
            """, params_base)
            total, firmados, rechazados, dni_recibido, esperando_dni, derivados_manual = cur.fetchone()

            total = total or 0
            firmados = firmados or 0
            rechazados = rechazados or 0
            dni_recibido = dni_recibido or 0
            esperando_dni = esperando_dni or 0
            derivados_manual = derivados_manual or 0

            porcentaje_rechazo = (rechazados / total * 100) if total else 0
            porcentaje_firmados = (firmados / total * 100) if total else 0

            cur.execute(f"""
                SELECT DATE(creado_at) AS fecha, COUNT(*) AS total,
                       COUNT(*) FILTER (WHERE estado = 'firmado') AS firmados,
                       COUNT(*) FILTER (WHERE estado = 'rechazado') AS rechazados
                FROM sofi_programa_cronico_inscripciones {where}
                GROUP BY DATE(creado_at) ORDER BY DATE(creado_at)
            """, params_base)
            registros_por_dia = cur.fetchall()

            cur.execute(f"""
                SELECT COALESCE(NULLIF(TRIM(motivo_rechazo), ''), 'Sin motivo registrado') AS motivo, COUNT(*) AS cantidad
                FROM sofi_programa_cronico_inscripciones {where} AND estado = 'rechazado'
                GROUP BY COALESCE(NULLIF(TRIM(motivo_rechazo), ''), 'Sin motivo registrado')
                ORDER BY cantidad DESC
            """, params_base)
            rechazos_motivos = cur.fetchall()

            cur.execute(f"""
                SELECT COALESCE(programa, 'Sin programa') AS programa, COUNT(*) AS cantidad
                FROM sofi_programa_cronico_inscripciones {where} AND estado = 'rechazado'
                GROUP BY programa ORDER BY cantidad DESC
            """, params_base)
            rechazos_programa = cur.fetchall()

            motivo_diagnostico = "Ninguno de los diagnósticos de la receta corresponde a un programa disponible"
            cur.execute(f"""
                SELECT token, n_documento, nombres, apellidos, datos_ocr_raw
                FROM sofi_programa_cronico_inscripciones {where}
                  AND estado = 'rechazado' AND motivo_rechazo = %s AND datos_ocr_raw IS NOT NULL
            """, params_base + [motivo_diagnostico])
            rechazados_diagnostico = cur.fetchall()

            contador_diagnosticos = Counter()
            detalles_diagnosticos = {}

            for token, n_documento, nombres, apellidos, datos_ocr_raw in rechazados_diagnostico:
                try:
                    if isinstance(datos_ocr_raw, str):
                        datos_ocr_raw = json.loads(datos_ocr_raw)
                    if not isinstance(datos_ocr_raw, dict):
                        continue

                    datos_receta = datos_ocr_raw.get("datos_receta", {})
                    if not isinstance(datos_receta, dict):
                        continue

                    diagnosticos = datos_receta.get("diagnosticos", [])
                    if not isinstance(diagnosticos, list):
                        continue

                    for diagnostico in diagnosticos:
                        if not isinstance(diagnostico, dict):
                            continue

                        cie10 = str(diagnostico.get("cie10") or diagnostico.get("CIE10") or "").strip()
                        descripcion = str(diagnostico.get("descripcion") or diagnostico.get("diagnostico") or "").strip()
                        if not cie10 and not descripcion:
                            continue

                        clave = (cie10.upper(), descripcion.upper())
                        contador_diagnosticos[clave] += 1
                        detalles_diagnosticos.setdefault(clave, {
                            "cie10": cie10, "descripcion": descripcion, "casos": []
                        })
                        detalles_diagnosticos[clave]["casos"].append({
                            "token": token, "dni": n_documento,
                            "nombre": f"{nombres or ''} {apellidos or ''}".strip()
                        })

                except Exception as e_diag:
                    logging.warning(f"[estadistica-cronico] Error procesando OCR token={token}: {e_diag}")

            diagnosticos_estadistica = [
                {"cie10": detalles_diagnosticos[clave]["cie10"],
                 "descripcion": detalles_diagnosticos[clave]["descripcion"],
                 "cantidad": cantidad,
                 "casos": detalles_diagnosticos[clave]["casos"]}
                for clave, cantidad in contador_diagnosticos.most_common()
            ]

    except Exception as e:
        logging.exception("[estadistica-cronico] Error generando estadísticas")
        return (json.dumps({"error": str(e)}), 500, HEADERS)

    finally:
        if conn:
            conn.close()

    # ═════════════════════════════════════════════
    # HTML
    # ═════════════════════════════════════════════
    def esc(value):
        return html_lib.escape(str(value if value is not None else ""))

    nombres_programa = {
        "vive_sano_completo": "Vive Sano", "krystal": "Krystal",
        "cronico_complejo": "Crónico Complejo", "Sin programa": "Sin programa"
    }
    def nombre_programa(valor):
        return nombres_programa.get(valor, valor or "Todos")

    filas_dias = "".join(f"""
        <tr>
            <td>{fecha.strftime('%d/%m/%Y') if fecha else '—'}</td>
            <td><strong>{cantidad}</strong></td>
            <td style="color:#2e7d32;">{firmados_dia}</td>
            <td style="color:#c62828;">{rechazados_dia}</td>
            <td>{(rechazados_dia / cantidad * 100) if cantidad else 0:.1f}%</td>
        </tr>
    """ for fecha, cantidad, firmados_dia, rechazados_dia in registros_por_dia) or """
        <tr><td colspan="5" style="text-align:center;color:#888;">No hay registros para el rango seleccionado.</td></tr>
    """

    filas_motivos = "".join(f"""
        <tr>
            <td>{esc(motivo)}</td>
            <td><strong>{cantidad}</strong></td>
            <td>{(cantidad / rechazados * 100) if rechazados else 0:.1f}%</td>
        </tr>
    """ for motivo, cantidad in rechazos_motivos) or """
        <tr><td colspan="3" style="text-align:center;color:#888;">No hay rechazos en el rango seleccionado.</td></tr>
    """

    filas_programas_rechazo = "".join(f"""
        <tr>
            <td>{esc(nombre_programa(programa))}</td>
            <td><strong>{cantidad}</strong></td>
            <td>{(cantidad / rechazados * 100) if rechazados else 0:.1f}%</td>
        </tr>
    """ for programa, cantidad in rechazos_programa) or """
        <tr><td colspan="3" style="text-align:center;color:#888;">No hay rechazos por programa.</td></tr>
    """

    filas_diagnosticos = "".join(f"""
        <tr>
            <td><strong>{esc(item["cie10"]) or "—"}</strong></td>
            <td>{esc(item["descripcion"]) or "Sin descripción"}</td>
            <td><strong>{item["cantidad"]}</strong></td>
            <td>{(item["cantidad"] / rechazados * 100) if rechazados else 0:.1f}%</td>
        </tr>
    """ for item in diagnosticos_estadistica) or """
        <tr><td colspan="4" style="text-align:center;color:#888;padding:20px;">
            No se encontraron diagnósticos asociados a rechazos por diagnóstico en este período.
        </td></tr>
    """

    opciones_programa = ["", "vive_sano_completo", "krystal", "cronico_complejo"]
    opciones_programa_html = "".join(
        f'<option value="{esc(p)}" {"selected" if p == filtro_programa else ""}>{esc(nombre_programa(p))}</option>'
        for p in opciones_programa
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Estadística Programa Crónico</title>
<link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
<style>
body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background: #f7f9fa; margin: 0; padding: 20px; color: #222; }}
h1 {{ color: #0090A6; font-size: 22px; margin: 0 0 4px 0; }}
.subtitulo {{ color: #888; font-size: 13px; margin-bottom: 20px; }}
.filtros {{ background: #fff; border-radius: 10px; padding: 16px 18px; margin-bottom: 20px; display: flex; gap: 14px; align-items: end; flex-wrap: wrap; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
.filtros label {{ font-size: 12px; font-weight: 600; display: block; margin-bottom: 5px; }}
.filtros input, .filtros select {{ padding: 8px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }}
.filtros button {{ padding: 9px 18px; background: #0090A6; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; }}
.filtros button:hover {{ opacity: .9; }}
.rango {{ font-size: 12px; color: #777; margin-bottom: 16px; }}
.cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 20px; }}
.card {{ background: #fff; border-radius: 10px; padding: 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-top: 4px solid #0090A6; }}
.card.rojo {{ border-top-color: #c62828; }}
.card.verde {{ border-top-color: #2e7d32; }}
.card.naranja {{ border-top-color: #ef6c00; }}
.card.gris {{ border-top-color: #888; }}
.card-titulo {{ color: #777; font-size: 12px; margin-bottom: 8px; }}
.card-numero {{ font-size: 28px; font-weight: 700; color: #222; }}
.card-extra {{ color: #888; font-size: 12px; margin-top: 5px; }}
.seccion {{ background: #fff; border-radius: 10px; padding: 18px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
.seccion h2 {{ margin: 0 0 14px 0; font-size: 16px; color: #0090A6; }}
.grid-dos {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 20px; margin-bottom: 20px; }}
table {{ width: 100%; border-collapse: collapse; background: #fff; }}
th {{ background: #0090A6; color: #fff; text-align: left; padding: 10px 12px; font-size: 12px; }}
td {{ padding: 10px 12px; font-size: 13px; border-bottom: 1px solid #eee; vertical-align: top; }}
tr:hover {{ background: #f0f7f8; }}
.contador {{ font-size: 12px; color: #888; margin-bottom: 10px; }}
.info {{ background: #f0f7f8; border-left: 4px solid #0090A6; padding: 12px 14px; margin-bottom: 15px; border-radius: 5px; font-size: 13px; }}
@media (max-width: 700px) {{
    body {{ padding: 10px; }}
    .grid-dos {{ grid-template-columns: 1fr; }}
    .seccion {{ padding: 12px; overflow-x: auto; }}
}}
</style>
</head>
<body>
<h1>Estadística — Programa de Pacientes Crónicos</h1>
<div class="subtitulo">Vive Sano · Krystal · Crónico Complejo</div>

<form class="filtros" method="GET" action="/programa-cronico/estadistica">
    <input type="hidden" name="pw" value="{esc(pw)}">
    <div><label>Fecha desde</label><input type="date" name="fecha_inicio" value="{esc(fecha_inicio)}"></div>
    <div><label>Fecha hasta</label><input type="date" name="fecha_fin" value="{esc(fecha_fin)}"></div>
    <div><label>Programa</label><select name="programa">{opciones_programa_html}</select></div>
    <button type="submit">Actualizar estadísticas</button>
    <a href="/programa-cronico/reporte?pw={esc(pw)}" style="padding:9px 18px;background:#888;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;text-decoration:none;display:inline-block;">← Volver al reporte</a>
</form>

<div class="rango">
    Período: <strong>{fecha_inicio_obj.strftime('%d/%m/%Y')}</strong> al <strong>{fecha_fin_obj.strftime('%d/%m/%Y')}</strong>
    · Programa: <strong>{esc(nombre_programa(filtro_programa))}</strong>
</div>

<div class="cards">
    <div class="card">
        <div class="card-titulo">TOTAL INSCRIPCIONES</div>
        <div class="card-numero">{total}</div>
        <div class="card-extra">En el período seleccionado</div>
    </div>
    <div class="card verde">
        <div class="card-titulo">FIRMADOS</div>
        <div class="card-numero" style="color:#2e7d32;">{firmados}</div>
        <div class="card-extra">{porcentaje_firmados:.1f}% del total</div>
    </div>
    <div class="card rojo">
        <div class="card-titulo">RECHAZADOS</div>
        <div class="card-numero" style="color:#c62828;">{rechazados}</div>
        <div class="card-extra">{porcentaje_rechazo:.1f}% del total</div>
    </div>
    <div class="card">
        <div class="card-titulo">DNI RECIBIDO</div>
        <div class="card-numero">{dni_recibido}</div>
        <div class="card-extra">Pendientes de firma</div>
    </div>
    <div class="card gris">
        <div class="card-titulo">ESPERANDO DNI</div>
        <div class="card-numero">{esperando_dni}</div>
        <div class="card-extra">Inscripciones iniciadas</div>
    </div>
    <div class="card naranja">
        <div class="card-titulo">DERIVADOS MANUALMENTE</div>
        <div class="card-numero" style="color:#ef6c00;">{derivados_manual}</div>
        <div class="card-extra">Revisión del equipo</div>
    </div>
</div>

<div class="seccion">
    <h2>Rechazos por motivo</h2>
    <div class="contador">Total rechazados: <strong>{rechazados}</strong></div>
    <table>
        <thead><tr><th>Motivo de rechazo</th><th>Cantidad</th><th>% de rechazados</th></tr></thead>
        <tbody>{filas_motivos}</tbody>
    </table>
</div>

<div class="grid-dos">
    <div class="seccion">
        <h2>Rechazos por programa</h2>
        <table>
            <thead><tr><th>Programa</th><th>Cantidad</th><th>%</th></tr></thead>
            <tbody>{filas_programas_rechazo}</tbody>
        </table>
    </div>
    <div class="seccion">
        <h2>Resumen del período</h2>
        <table>
            <tbody>
                <tr><td>Total</td><td><strong>{total}</strong></td></tr>
                <tr><td>Firmados</td><td><strong>{firmados}</strong> ({porcentaje_firmados:.1f}%)</td></tr>
                <tr><td>Rechazados</td><td><strong>{rechazados}</strong> ({porcentaje_rechazo:.1f}%)</td></tr>
                <tr><td>DNI recibido</td><td><strong>{dni_recibido}</strong></td></tr>
                <tr><td>Esperando DNI</td><td><strong>{esperando_dni}</strong></td></tr>
                <tr><td>Derivados manualmente</td><td><strong>{derivados_manual}</strong></td></tr>
            </tbody>
        </table>
    </div>
</div>

<div class="seccion">
    <h2>Diagnósticos que generan rechazo</h2>
    <div class="info">
        Esta sección analiza los diagnósticos encontrados en el OCR de las recetas rechazadas porque ninguno
        de sus diagnósticos corresponde a un programa disponible. Sirve para identificar diagnósticos frecuentes
        que podrían estar quedando fuera de los programas actuales.
    </div>
    <table>
        <thead><tr><th>CIE-10</th><th>Diagnóstico</th><th>Casos</th><th>% de rechazados</th></tr></thead>
        <tbody>{filas_diagnosticos}</tbody>
    </table>
</div>

<div class="seccion">
    <h2>Inscripciones por día</h2>
    <table>
        <thead><tr><th>Fecha</th><th>Total</th><th>Firmados</th><th>Rechazados</th><th>% rechazo</th></tr></thead>
        <tbody>{filas_dias}</tbody>
    </table>
</div>
</body>
</html>
"""

    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})


# ──────────────────────────────────────────
# Lógica compartida: actualizar celular/correo
# (usada tanto por el endpoint API como por el formulario web)
# ──────────────────────────────────────────
def _actualizar_contacto(token, celular_nuevo, correo_nuevo, reenviar_si_firmado=True):
    celular_nuevo = (celular_nuevo or "").strip() or None
    correo_nuevo = (correo_nuevo or "").strip() or None

    if not token:
        return {"ok": False, "error": "token es requerido"}
    if not celular_nuevo and not correo_nuevo:
        return {"ok": False, "error": "Debes enviar celular_nuevo y/o correo_nuevo"}

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, celular, correo, estado
        FROM sofi_programa_cronico_inscripciones
        WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return {"ok": False, "error": "Token no encontrado"}

    inscripcion_id, celular_actual, correo_actual, estado = row

    if celular_nuevo and celular_nuevo != celular_actual:
        cur.execute("""
            SELECT token FROM sofi_programa_cronico_inscripciones
            WHERE celular = %s AND id != %s
        """, (celular_nuevo, inscripcion_id))
        conflicto = cur.fetchone()
        if conflicto:
            cur.close(); conn.close()
            return {"ok": False, "error": f"Ese celular ya está en uso por otra inscripción (token: {conflicto[0]})"}

    campos = []
    valores = []

    if celular_nuevo and celular_nuevo != celular_actual:
        campos.append("celular = %s")
        valores.append(celular_nuevo)

    if correo_nuevo and correo_nuevo != correo_actual:
        campos.append("correo = %s")
        valores.append(correo_nuevo)

    if not campos:
        cur.close(); conn.close()
        return {"ok": True, "mensaje": "No hubo cambios (valores idénticos a los actuales)"}

    campos.append("actualizado_at = NOW()")
    valores.append(inscripcion_id)

    cur.execute(f"""
        UPDATE sofi_programa_cronico_inscripciones
        SET {", ".join(campos)}
        WHERE id = %s
    """, valores)
    conn.commit()
    cur.close()
    conn.close()

    resultado = {
        "ok": True,
        "token": token,
        "celular_anterior": celular_actual,
        "celular_actual": celular_nuevo or celular_actual,
        "correo_anterior": correo_actual,
        "correo_actual": correo_nuevo or correo_actual,
    }

    if reenviar_si_firmado and correo_nuevo and correo_nuevo != correo_actual and estado == "firmado":
        try:
            resultado["reenvio_danaconnect"] = _reenviar_danaconnect_por_token(token)
        except Exception as e:
            resultado["reenvio_danaconnect"] = {"ok": False, "error": str(e)}

    return resultado


def _reenviar_danaconnect_por_token(token):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT n_documento, nombres, apellidos, correo, programa, estado,
               pdf_consentimiento_url, danaconnect_enviado,
               dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket
        FROM sofi_programa_cronico_inscripciones WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return {"ok": False, "error": "Token no encontrado"}

    (n_documento, nombres, apellidos, correo, programa, estado,
     pdf_consentimiento_url, danaconnect_enviado,
     dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket) = row

    if estado != "firmado" or not pdf_consentimiento_url:
        cur.close(); conn.close()
        return {"ok": False, "error": f"Inscripción no está firmada o sin PDF (estado: {estado})"}

    try:
        pdf_bytes = descargar_blob_desde_url(pdf_consentimiento_url)
    except Exception as e:
        cur.close(); conn.close()
        return {"ok": False, "error": f"No se pudo descargar el PDF: {str(e)}"}

    danaconnect_enviado_nuevo = False
    danaconnect_error_msg = None
    try:
        notificar_danaconnect(n_documento, nombres, apellidos, correo, programa, pdf_bytes,
                                dni_foto_anverso_url, dni_foto_reverso_url, url_receta_bucket)
        danaconnect_enviado_nuevo = True
    except Exception as e:
        danaconnect_error_msg = str(e)

    cur.execute("""
        UPDATE sofi_programa_cronico_inscripciones
        SET danaconnect_enviado = %s, danaconnect_error = %s, actualizado_at = NOW()
        WHERE token = %s
    """, (danaconnect_enviado_nuevo, danaconnect_error_msg, token))
    conn.commit()
    cur.close()
    conn.close()

    return {"ok": danaconnect_enviado_nuevo, "error": danaconnect_error_msg}


# ──────────────────────────────────────────
# API: mantiene compatibilidad si alguna vez quieres
# llamarlo por API/integración además del formulario web
# ──────────────────────────────────────────
@con_log_errores
def handle_actualizar_contacto_cronico(request):
    data = request.get_json(force=True, silent=True) or {}

    if data.get("pw") != ADMIN_PW_CRONICO:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    resultado = _actualizar_contacto(
        token=data.get("token"),
        celular_nuevo=data.get("celular_nuevo"),
        correo_nuevo=data.get("correo_nuevo"),
        reenviar_si_firmado=data.get("reenviar_si_firmado", True),
    )
    status = 200 if resultado.get("ok") else 400
    return (json.dumps(resultado), status, HEADERS)


# ──────────────────────────────────────────
# WEB: página de mantenimiento amigable
# GET  /programa-cronico/mantenimiento?pw=...            -> buscador
# GET  /programa-cronico/mantenimiento?pw=...&q=...       -> resultados de búsqueda
# GET  /programa-cronico/mantenimiento?pw=...&token=...   -> formulario de edición de 1 registro
# POST /programa-cronico/mantenimiento (form-urlencoded)  -> guarda cambios
# ──────────────────────────────────────────
def _layout_mantenimiento(pw, contenido, mensaje_html=""):
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Mantenimiento de contacto — Programa Crónico</title>
<link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
<style>
    body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background: #f7f9fa; margin: 0; padding: 24px; color: #222; }}
    .contenedor {{ max-width: 640px; margin: 0 auto; }}
    h1 {{ color: #0090A6; font-size: 20px; margin-bottom: 4px; }}
    .subtitulo {{ color: #888; font-size: 13px; margin-bottom: 20px; }}
    .card {{ background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); margin-bottom: 16px; }}
    label {{ font-size: 12px; font-weight: 600; display: block; margin: 12px 0 4px; }}
    input[type=text], input[type=email], input[type=tel] {{
        width: 100%; box-sizing: border-box; padding: 9px 12px; border: 1px solid #ccc;
        border-radius: 6px; font-size: 14px;
    }}
    button {{
        margin-top: 16px; padding: 10px 18px; background: #0090A6; color: #fff; border: none;
        border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600;
    }}
    button:hover {{ background: #007686; }}
    a.volver {{ display: inline-block; margin-bottom: 14px; color: #0090A6; font-size: 13px; text-decoration: none; }}
    a.volver:hover {{ text-decoration: underline; }}
    table {{ width: 100%; border-collapse: collapse; background: #fff; border-radius: 10px; overflow: hidden; }}
    th {{ background: #0090A6; color: #fff; text-align: left; padding: 9px 10px; font-size: 12px; }}
    td {{ padding: 9px 10px; font-size: 13px; border-bottom: 1px solid #eee; }}
    tr:hover {{ background: #f0f7f8; }}
    .msg-ok {{ background: #e6f5ea; color: #2e7d32; padding: 12px 16px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; }}
    .msg-error {{ background: #fdecea; color: #c62828; padding: 12px 16px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; }}
    .campo-actual {{ font-size: 12px; color: #888; margin-top: 2px; }}
</style>
</head>
<body>
<div class="contenedor">
    <h1>Mantenimiento de contacto</h1>
    <div class="subtitulo">Programa de Pacientes Crónicos — corregir celular / correo</div>
    {mensaje_html}
    {contenido}
</div>
</body>
</html>"""


def _form_busqueda(pw, q=""):
    return f"""
    <div class="card">
        <form method="GET" action="/programa-cronico/mantenimiento">
            <input type="hidden" name="pw" value="{pw}">
            <label>Buscar por DNI, nombre, celular o token</label>
            <input type="text" name="q" value="{q}" placeholder="Ej. 47313070 o Juan Pérez" autofocus>
            <button type="submit">Buscar</button>
        </form>
    </div>
    """


def _tabla_resultados(pw, registros):
    if not registros:
        return '<div class="card">No se encontraron inscripciones con ese criterio.</div>'

    filas = ""
    for token, n_documento, nombres, apellidos, celular, correo, programa, estado in registros:
        nombre_completo = f"{nombres or ''} {apellidos or ''}".strip() or "—"
        filas += f"""
        <tr>
            <td>{nombre_completo}<br><span class="campo-actual">{n_documento or ''}</span></td>
            <td>{celular or '—'}<br><span class="campo-actual">{correo or ''}</span></td>
            <td>{programa or '—'}</td>
            <td>{estado or '—'}</td>
            <td><a href="/programa-cronico/mantenimiento?pw={pw}&token={token}">Editar</a></td>
        </tr>"""

    return f"""
    <div class="card" style="padding:0; overflow:hidden;">
        <table>
            <thead>
                <tr><th>Paciente</th><th>Contacto</th><th>Programa</th><th>Estado</th><th></th></tr>
            </thead>
            <tbody>{filas}</tbody>
        </table>
    </div>
    """


def _form_edicion(pw, token, row):
    (n_documento, nombres, apellidos, celular, correo, programa, estado) = row
    nombre_completo = f"{nombres or ''} {apellidos or ''}".strip() or "—"

    aviso_reenvio = ""
    if estado == "firmado":
        aviso_reenvio = '<div class="campo-actual" style="margin-top:10px;">Si corriges el correo, se reenviará automáticamente el consentimiento firmado.</div>'

    return f"""
    <a class="volver" href="/programa-cronico/mantenimiento?pw={pw}">&larr; Volver a la búsqueda</a>
    <div class="card">
        <div style="font-size:14px; font-weight:600;">{nombre_completo}</div>
        <div class="campo-actual">DNI: {n_documento or '—'} · Programa: {programa or '—'} · Estado: {estado or '—'}</div>

        <form method="POST" action="/programa-cronico/mantenimiento">
            <input type="hidden" name="pw" value="{pw}">
            <input type="hidden" name="token" value="{token}">

            <label>Celular</label>
            <input type="tel" name="celular_nuevo" value="{celular or ''}">

            <label>Correo</label>
            <input type="email" name="correo_nuevo" value="{correo or ''}">

            {aviso_reenvio}
            <button type="submit">Guardar cambios</button>
        </form>
    </div>
    """


def handle_mantenimiento_contacto_cronico(request):
    pw = request.args.get("pw") or request.form.get("pw")
    if pw != ADMIN_PW_CRONICO:
        return ("No autorizado", 401, {"Content-Type": "text/html; charset=utf-8"})

    HTML_HEADERS = {"Content-Type": "text/html; charset=utf-8"}

    # ── POST: guardar cambios ──
    if request.method == "POST":
        token = request.form.get("token")
        celular_nuevo = request.form.get("celular_nuevo")
        correo_nuevo = request.form.get("correo_nuevo")

        resultado = _actualizar_contacto(token, celular_nuevo, correo_nuevo)

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT n_documento, nombres, apellidos, celular, correo, programa, estado
            FROM sofi_programa_cronico_inscripciones WHERE token = %s
        """, (token,))
        row = cur.fetchone()
        cur.close(); conn.close()

        if not row:
            return (_layout_mantenimiento(pw, _form_busqueda(pw), '<div class="msg-error">Token no encontrado.</div>'), 404, HTML_HEADERS)

        if resultado.get("ok"):
            texto_reenvio = ""
            if resultado.get("reenvio_danaconnect"):
                texto_reenvio = " Se reenvió a Danaconnect." if resultado["reenvio_danaconnect"].get("ok") else " (No se pudo reenviar a Danaconnect, revisa el reporte)."
            mensaje = f'<div class="msg-ok">Datos actualizados correctamente.{texto_reenvio}</div>'
        else:
            mensaje = f'<div class="msg-error">{resultado.get("error")}</div>'

        return (_layout_mantenimiento(pw, _form_edicion(pw, token, row), mensaje), 200, HTML_HEADERS)

    # ── GET con token: mostrar formulario de edición de un registro puntual ──
    token = request.args.get("token")
    if token:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT n_documento, nombres, apellidos, celular, correo, programa, estado
            FROM sofi_programa_cronico_inscripciones WHERE token = %s
        """, (token,))
        row = cur.fetchone()
        cur.close(); conn.close()

        if not row:
            return (_layout_mantenimiento(pw, _form_busqueda(pw), '<div class="msg-error">Token no encontrado.</div>'), 404, HTML_HEADERS)

        return (_layout_mantenimiento(pw, _form_edicion(pw, token, row)), 200, HTML_HEADERS)

    # ── GET con q: mostrar resultados de búsqueda ──
    q = request.args.get("q", "").strip()
    if q:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT token, n_documento, nombres, apellidos, celular, correo, programa, estado
            FROM sofi_programa_cronico_inscripciones
            WHERE n_documento ILIKE %s OR nombres ILIKE %s OR apellidos ILIKE %s
               OR celular ILIKE %s OR token ILIKE %s
            ORDER BY creado_at DESC LIMIT 50
        """, tuple([f"%{q}%"] * 4 + [f"{q}%"]))
        registros = cur.fetchall()
        cur.close(); conn.close()

        contenido = _form_busqueda(pw, q) + _tabla_resultados(pw, registros)
        return (_layout_mantenimiento(pw, contenido), 200, HTML_HEADERS)

    # ── GET sin nada: solo el buscador ──
    return (_layout_mantenimiento(pw, _form_busqueda(pw)), 200, HTML_HEADERS)

# ────────────────────────────────────────────────────────────────────────────────────
# API: Afiliacion a titulares y dependientes con imagen de su documento de identidad
# ────────────────────────────────────────────────────────────────────────────────────
MAPA_NACIONALIDADES = {
    "AFG": "Afganistán", "ALB": "Albania", "DEU": "Alemania", "AND": "Andorra",
    "AGO": "Angola", "AIA": "Anguila", "ATA": "Antártida", "ATG": "Antigua y Barbuda",
    "ANT": "Antillas Holandesas", "SAU": "Arabia Saudita", "DZA": "Argelia", "ARG": "Argentina",
    "ARM": "Armenia", "ABW": "Aruba", "AUS": "Australia", "AUT": "Austria",
    "AZE": "Azerbaiyán", "BHS": "Bahamas", "BHR": "Bahrein", "BGD": "Bangladesh",
    "BRB": "Barbados", "BLR": "Belarús", "BEL": "Bélgica", "BLZ": "Belice",
    "BEN": "Benin", "BMU": "Bermuda", "BOL": "Bolivia", "BES": "Bonaire, San Eustaquio y Saba",
    "BIH": "Bosnia y Herzegovina", "BWA": "Botswana", "BRA": "Brasil", "BRN": "Brunei Darussalam",
    "BGR": "Bulgaria", "BFA": "Burkina Faso", "BDI": "Burundi", "BTN": "Bután",
    "CPV": "Cabo Verde", "KHM": "Camboya", "CMR": "Camerún", "CAN": "Canadá",
    "TCD": "Chad", "CHL": "Chile", "CHN": "China", "CYP": "Chipre",
    "VAT": "Ciudad del Vaticano", "COL": "Colombia", "COG": "Congo (Brazzaville)", "COD": "Congo RD",
    "PRK": "Corea del Norte", "KOR": "Corea del Sur", "CIV": "Costa de Marfil", "CRI": "Costa Rica",
    "HRV": "Croacia", "CUB": "Cuba", "CUW": "Curaçao", "DNK": "Dinamarca",
    "DMA": "Dominica", "ECU": "Ecuador", "EGY": "Egipto", "SLV": "El Salvador",
    "ARE": "Emiratos Árabes Unidos", "ERI": "Eritrea", "SVK": "Eslovaquia", "SVN": "Eslovenia",
    "ESP": "España", "USA": "Estados Unidos", "UMI": "Estados Unidos (Islas Menores Alejadas)", "EST": "Estonia",
    "ETH": "Etiopía", "PHL": "Filipinas", "FIN": "Finlandia", "FJI": "Fiyi",
    "FRA": "Francia", "GAB": "Gabón", "GMB": "Gambia", "GEO": "Georgia",
    "SGS": "Georgia del Sur e Islas Sandwich", "GHA": "Ghana", "GIB": "Gibraltar", "GRD": "Granada",
    "GRC": "Grecia", "GRL": "Groenlandia", "GLP": "Guadalupe", "GUM": "Guam",
    "GTM": "Guatemala", "GUY": "Guayana", "GUF": "Guayana Francesa", "GGY": "Guernesey",
    "GIN": "Guinea", "GNQ": "Guinea Ecuatorial", "GNB": "Guinea-Bissau", "HTI": "Haití",
    "HND": "Honduras", "HKG": "Hong Kong", "HUN": "Hungría", "IND": "India",
    "IDN": "Indonesia", "IRQ": "Irak", "IRN": "Irán", "IRL": "Irlanda",
    "BVT": "Isla Bouvet", "IMN": "Isla de Man", "CXR": "Isla de Navidad", "NFK": "Isla Norfolk",
    "ISL": "Islandia", "ALA": "Islas Åland", "CYM": "Islas Caimán", "CCK": "Islas Cocos (Keeling)",
    "COM": "Islas Comoros", "COK": "Islas Cook", "FRO": "Islas Feroe", "HMD": "Islas Heard y McDonald",
    "FLK": "Islas Malvinas (Falkland)", "MNP": "Islas Marianas del Norte", "MHL": "Islas Marshall", "PCN": "Islas Pitcairn",
    "SLB": "Islas Salomón", "TCA": "Islas Turcas y Caicos", "VGB": "Islas Vírgenes Británicas", "VIR": "Islas Vírgenes de EE.UU.",
    "ISR": "Israel", "ITA": "Italia", "JAM": "Jamaica", "JPN": "Japón",
    "JEY": "Jersey", "JOR": "Jordania", "KAZ": "Kazajstán", "KEN": "Kenia",
    "KGZ": "Kirguistán", "KIR": "Kiribati", "XKX": "Kosovo", "KWT": "Kuwait",
    "LAO": "Laos", "LSO": "Lesoto", "LVA": "Letonia", "LBN": "Líbano",
    "LBR": "Liberia", "LBY": "Libia", "LIE": "Liechtenstein", "LTU": "Lituania",
    "LUX": "Luxemburgo", "MAC": "Macao", "MKD": "Macedonia", "MDG": "Madagascar",
    "MYS": "Malasia", "MWI": "Malawi", "MDV": "Maldivas", "MLI": "Malí",
    "MLT": "Malta", "MAR": "Marruecos", "MTQ": "Martinica", "MUS": "Mauricio",
    "MRT": "Mauritania", "MYT": "Mayotte", "MEX": "México", "FSM": "Micronesia",
    "MDA": "Moldova", "MCO": "Mónaco", "MNG": "Mongolia", "MNE": "Montenegro",
    "MSR": "Montserrat", "MOZ": "Mozambique", "MMR": "Myanmar", "NAM": "Namibia",
    "NRU": "Nauru", "NPL": "Nepal", "NIC": "Nicaragua", "NER": "Níger",
    "NGA": "Nigeria", "NIU": "Niue", "NOR": "Noruega", "NCL": "Nueva Caledonia",
    "NZL": "Nueva Zelanda", "OMN": "Omán", "NLD": "Países Bajos", "PAK": "Pakistán",
    "PLW": "Palau", "PSE": "Palestina (Gaza)", "PAN": "Panamá", "PNG": "Papua Nueva Guinea",
    "PRY": "Paraguay", "PER": "Perú", "PYF": "Polinesia Francés", "POL": "Polonia",
    "PRT": "Portugal", "PRI": "Puerto Rico", "QAT": "Qatar", "GBR": "Reino Unido",
    "CAF": "República Centroafricana", "CZE": "República Checa", "DOM": "República Dominicana", "REU": "Reunión",
    "RWA": "Ruanda", "ROU": "Rumania", "RUS": "Rusia", "ESH": "Sáhara Occidental",
    "MAF": "Saint Martin", "BLM": "Saint-Barthélemy", "WSM": "Samoa", "ASM": "Samoa Americana",
    "KNA": "San Cristóbal y Nieves", "SMR": "San Marino", "SPM": "San Pedro y Miquelón", "VCT": "San Vicente y las Granadinas",
    "SHN": "Santa Helena - Ascensión y Tristán da Cunha", "LCA": "Santa Lucía", "STP": "Santo Tomé y Príncipe", "SEN": "Senegal",
    "SRB": "Serbia", "SYC": "Seychelles", "SLE": "Sierra Leona", "SGP": "Singapur",
    "SYR": "Siria", "SOM": "Somalia", "LKA": "Sri Lanka", "SWZ": "Suazilandia",
    "ZAF": "Sudáfrica", "SDN": "Sudán", "SSD": "Sudan Sur", "SWE": "Suecia",
    "CHE": "Suiza", "SUR": "Suriname", "SJM": "Svalbard y Jan Mayen", "THA": "Tailandia",
    "TWN": "Taiwan", "TZA": "Tanzania", "TJK": "Tayikistán", "IOT": "Territorio Brtiánico del Océano lndico",
    "ATF": "Territorios Franceses del Sur", "TLS": "Timor Oriental", "TGO": "Togo", "TKL": "Tokelau",
    "TON": "Tonga", "TTO": "Trinidad y Tobago", "TUN": "Túnez", "TKM": "Turkmenistán",
    "TUR": "Turquía", "TUV": "Tuvalu", "UKR": "Ucrania", "UGA": "Uganda",
    "URY": "Uruguay", "UZB": "Uzbekistán", "VUT": "Vanuatu", "VEN": "Venezuela",
    "VNM": "Viet Nam", "WLF": "Wallis y Futuna", "YEM": "Yemen", "DJI": "Yibuti",
    "ZMB": "Zambia", "ZWE": "Zimbabue",
}

def _mapear_nacionalidad(codigo):
    if not codigo:
        return None
    codigo = codigo.strip().upper()
    return MAPA_NACIONALIDADES.get(codigo, codigo)  # si no está en el mapa, deja el código tal cual como fallback

def _formatear_ubigeo_texto(distrito, provincia, departamento):
    """
    Arma el texto de ubicación tal como debe quedar guardado, en formato
    DEPARTAMENTO-PROVINCIA-DISTRITO (ej. 'LIMA-LIMA-LINCE'), a partir de lo
    que la OCR extrajo directamente de la imagen del DNI/CE. Ya no consulta
    la tabla sofi_distritos ni depende de que el ubigeo numérico exista ahí.
    """
    partes = [
        (departamento or "").strip().upper(),
        (provincia or "").strip().upper(),
        (distrito or "").strip().upper(),
    ]
    partes = [p for p in partes if p]
    return "-".join(partes) if partes else None

def _buscar_ubigeo(cur, distrito, provincia, departamento):
    """
    Busca el ubigeo en sofi_distritos a partir de distrito/provincia/departamento
    extraídos por OCR (Document AI + Gemini). Intenta primero un match exacto
    por los tres campos; si no hay coincidencia (variaciones de redacción/tildes),
    cae a un match solo por distrito (toma el primero si hay más de uno).
    """
    if not distrito:
        return None

    distrito_norm     = distrito.strip().upper()
    provincia_norm     = (provincia or "").strip().upper()
    departamento_norm  = (departamento or "").strip().upper()

    if provincia_norm and departamento_norm:
        cur.execute("""
            SELECT ubigeo FROM sofi_distritos
            WHERE distrito = %s AND provincia = %s AND departamento = %s
        """, (distrito_norm, provincia_norm, departamento_norm))
        row = cur.fetchone()
        if row:
            return row[0]

    cur.execute("SELECT ubigeo FROM sofi_distritos WHERE distrito = %s LIMIT 1", (distrito_norm,))
    row = cur.fetchone()
    return row[0] if row else None

def _fmt_fecha(valor):
    if valor is None:
        return "—"
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return valor or "—"



APISPERU_TOKEN = os.environ.get("APISPERU_TOKEN")
CODIGOS_PERSONA_VALIDOS = {"T", "D1", "D2", "D3", "D4"}
# =====================================================================
# 1. INICIAR AFILIACION
# =====================================================================
@con_log_errores
def handle_iniciar_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    dni_titular = (data.get("dni_titular") or "").strip()
    celular = (data.get("celular") or "").strip()
    tipo_solicitud = (data.get("tipo_solicitud") or "AFILIACION").strip().upper()

    if not dni_titular:
        return (json.dumps({"error": "El campo dni_titular es obligatorio"}), 400, HEADERS)
    if not celular:
        return (json.dumps({"error": "El campo celular es obligatorio"}), 400, HEADERS)
    if tipo_solicitud not in ("AFILIACION", "INCLUSION", "TRASLADO"):
        return (json.dumps({"error": "tipo_solicitud inválido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()

    # ─: Bloquear si el DNI no está en la lista autorizada (Congreso) ──
    cur.execute("SELECT 1 FROM afiliacion_dnis_congreso WHERE dni = %s", (dni_titular,))
    if not cur.fetchone():
        cur.close(); conn.close()
        return (json.dumps({
            "es_reanudada": "None",
            "error": "Este DNI no se encuentra habilitado para iniciar el proceso de afiliación."
        }), 403, HEADERS)

    # ── Bloquear si este DNI ya tiene una afiliación COMPLETADA ──
    cur.execute("""
        SELECT id, token, fecha_completada FROM afiliacion_solicitudes
        WHERE titular_num_documento = %s AND estado = 'COMPLETADA'
        ORDER BY fecha_completada DESC LIMIT 1
    """, (dni_titular,))
    row_completada = cur.fetchone()
    if row_completada:
        cur.close(); conn.close()
        solicitud_previa_id, token_previo, fecha_completada = row_completada
        return (json.dumps({
            "error": "Este DNI ya cuenta con una solicitud de afiliación completada. En caso quiera agregar dependientes deberá acercarse al departamento de afiliación del Congreso.",
            "solicitud_previa_id": solicitud_previa_id,
            "es_reanudada": "None",
            "fecha_completada": str(fecha_completada) if fecha_completada else None,
        }), 409, HEADERS)

    # ── Bloquear si este DNI ya está registrado como dependiente en una solicitud COMPLETADA ──
    cur.execute("""
        SELECT ad.solicitud_id, ad.codigo, s.token
        FROM afiliacion_dependientes ad
        JOIN afiliacion_solicitudes s ON s.id = ad.solicitud_id
        WHERE ad.num_documento = %s
          AND s.estado = 'COMPLETADA'
        LIMIT 1
    """, (dni_titular,))
    row_dependiente = cur.fetchone()
    if row_dependiente:
        cur.close(); conn.close()
        solicitud_dep_id, codigo_dep, token_dep = row_dependiente
        return (json.dumps({
            "error": "Este DNI ya está registrado como dependiente en una afiliación completada y no puede afiliarse como titular",
            "solicitud_id": solicitud_dep_id,
            "persona_codigo": codigo_dep,
            "es_reanudada": "None",
        }), 409, HEADERS)

    # ── Find-or-create por DNI del titular (solo entre las EN_PROCESO) ──
    cur.execute("""
        SELECT id, token FROM afiliacion_solicitudes
        WHERE titular_num_documento = %s AND estado = 'EN_PROCESO'
        ORDER BY id DESC LIMIT 1
    """, (dni_titular,))
    row = cur.fetchone()

    if row:
        solicitud_id, token = row

        # La solicitud ya existe y está EN_PROCESO.
        # Actualizamos ambos campos para garantizar que el celular
        # quede disponible desde el inicio del proceso.
        cur.execute("""
            UPDATE afiliacion_solicitudes
            SET celular_origen_wsp = %s,
                titular_celular = %s
            WHERE id = %s
        """, (celular, celular, solicitud_id))

        conn.commit()
        es_reanudada = True

    else:
        token = str(uuid.uuid4())

        # Guardamos el celular desde el inicio tanto como
        # celular de origen WhatsApp como celular del titular.
        cur.execute("""
            INSERT INTO afiliacion_solicitudes
                (
                    token,
                    tipo_solicitud,
                    celular_origen_wsp,
                    titular_celular,
                    titular_num_documento
                )
            VALUES (%s, %s, %s, %s, %s)
            RETURNING id
        """, (
            token,
            tipo_solicitud,
            celular,
            celular,
            dni_titular
        ))

        solicitud_id = cur.fetchone()[0]
        conn.commit()
        es_reanudada = False

    cur.close(); conn.close()

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link = f"{base_url}/afiliacion/{token}"

    return (json.dumps({
        "solicitud_id": solicitud_id,
        "token": token,
        "error": f"Ingresar al link generado para terminar el flujo {link}",
        "link": link,
        "es_reanudada": es_reanudada,
    }, ensure_ascii=False), 200, HEADERS)

 
 
# =====================================================================
# 2. CONSULTA RUC (reusa sctr_empresa, sin telefono)
# =====================================================================
@con_log_errores
def handle_consulta_ruc_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    ruc = data.get("ruc", "").strip()
    token = data.get("token", "").strip()
 
    if not ruc:
        return (json.dumps({"error": "El campo ruc es obligatorio"}), 400, HEADERS)
    if len(ruc) != 11 or not ruc.isdigit():
        return (json.dumps({"error": "RUC inválido, debe tener 11 dígitos"}), 400, HEADERS)
 
    conn = get_connection()
    cur = conn.cursor()
 
    cur.execute("""
        SELECT razon_social, direccion, departamento, provincia, distrito
        FROM sctr_empresa WHERE ruc = %s
    """, (ruc,))
    row = cur.fetchone()
 
    if row:
        razon_social, direccion, departamento, provincia, distrito = row
    else:
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
 
        if not razon_social:
            cur.close(); conn.close()
            return (json.dumps({"error": "RUC no encontrado"}), 404, HEADERS)
 
        cur.execute("""
            INSERT INTO sctr_empresa
                (ruc, razon_social, direccion, departamento, provincia, distrito, fecha_creacion, fecha_ult_acceso)
            VALUES (%s, %s, %s, %s, %s, %s, now(), now())
            ON CONFLICT (ruc) DO NOTHING
        """, (ruc, razon_social, direccion, departamento, provincia, distrito))
        conn.commit()
 
    # Si viene un token de solicitud, asociamos el RUC de una vez
    if token:
        cur.execute("""
            UPDATE afiliacion_solicitudes SET ruc = %s
            WHERE token = %s AND estado = 'EN_PROCESO'
        """, (ruc, token))
        conn.commit()
 
    cur.close(); conn.close()
 
    return (json.dumps({
        "ruc": ruc,
        "razon_social": razon_social,
        "direccion": direccion,
        "departamento": departamento,
        "provincia": provincia,
        "distrito": distrito,
    }, ensure_ascii=False), 200, HEADERS)
 
 
# =====================================================================
# 3. OCR DE DOCUMENTO (DNI/CE) - titular o dependiente
# =====================================================================
@con_log_errores
def handle_ocr_documento_afiliacion(request):
    """
    Recibe una o dos imágenes (anverso/reverso, base64) de DNI/CE de una
    persona (titular o dependiente) y reusa extraer_datos_dni() -- la misma
    función de Document AI + Gemini que ya usa programa-cronico -- para
    devolver los campos estructurados SIN guardarlos aún (el front los
    muestra para confirmación/edición antes de llamar a guardar-paso).

    Body esperado:
    {
      "token": "...",
      "persona_codigo": "T",
      "imagen_anverso_base64": "data:image/jpeg;base64,....",
      "imagen_reverso_base64": "data:image/jpeg;base64,...."   # opcional
    }
    """
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    persona_codigo = data.get("persona_codigo", "").strip().upper()
    imagen_anverso_base64 = data.get("imagen_anverso_base64", "")
    imagen_reverso_base64 = data.get("imagen_reverso_base64", "")

    if not token or not imagen_anverso_base64:
        return (json.dumps({"error": "token e imagen_anverso_base64 son obligatorios"}), 400, HEADERS)
    if persona_codigo not in CODIGOS_PERSONA_VALIDOS:
        return (json.dumps({"error": "persona_codigo inválido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, estado, fecha_expira FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    solicitud_id, estado, fecha_expira = row
    ahora = datetime.now(LIMA_TZ)
    if estado != "EN_PROCESO" or (fecha_expira and fecha_expira.replace(tzinfo=LIMA_TZ) < ahora):
        cur.close(); conn.close()
        return (json.dumps({"error": "La solicitud ya no está vigente"}), 409, HEADERS)

    # Subir imágenes a Cloud Storage (para trazabilidad, igual que en cronico)
    ruta_anverso = _subir_imagen_storage(token, persona_codigo, "DNI_ANVERSO", imagen_anverso_base64)
    cur.execute("""
        INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
        VALUES (%s, %s, 'DNI_ANVERSO', %s)
    """, (solicitud_id, persona_codigo, ruta_anverso))

    ruta_reverso = None
    if imagen_reverso_base64:
        ruta_reverso = _subir_imagen_storage(token, persona_codigo, "DNI_REVERSO", imagen_reverso_base64)
        cur.execute("""
            INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
            VALUES (%s, %s, 'DNI_REVERSO', %s)
        """, (solicitud_id, persona_codigo, ruta_reverso))

    conn.commit()
    cur.close(); conn.close()

    # Reusa la función real de OCR + Gemini que ya usa programa-cronico
    bytes_anverso = base64.b64decode(imagen_anverso_base64.split(",")[-1])
    bytes_reverso = base64.b64decode(imagen_reverso_base64.split(",")[-1]) if imagen_reverso_base64 else None

    datos_dni = extraer_datos_dni(
        bytes_anverso, "image/jpeg",
        bytes_reverso, "image/jpeg" if bytes_reverso else None
    )

    # extraer_datos_dni() devuelve nombres y apellidos juntos; el formulario
    # de afiliación los pide separados (paterno/materno), así que se separa
    # de forma simple. Si el resultado no calza, el front deja editar a mano.
    apellidos_completos = (datos_dni.get("apellidos") or "").strip()
    partes_apellido = apellidos_completos.split(" ", 1)
    apellido_paterno = partes_apellido[0] if partes_apellido else ""
    apellido_materno = partes_apellido[1] if len(partes_apellido) > 1 else ""

    campos_extraidos = {
        "apellido_paterno": apellido_paterno,
        "apellido_materno": apellido_materno,
        "nombres": datos_dni.get("nombres"),
        "tipo_documento": datos_dni.get("tipo_documento"),
        "num_documento": datos_dni.get("n_documento"),
        "fecha_nacimiento": datos_dni.get("fecha_nacimiento"),
        "direccion": datos_dni.get("domicilio"),
        "distrito": datos_dni.get("distrito"),
        "provincia": datos_dni.get("provincia"),
        "departamento": datos_dni.get("departamento"),
    }

    return (json.dumps({
        "persona_codigo": persona_codigo,
        "ruta_anverso": ruta_anverso,
        "ruta_reverso": ruta_reverso,
        "campos_extraidos": campos_extraidos,
    }, ensure_ascii=False), 200, HEADERS)
 
def _subir_archivo_generico_storage(token, persona_codigo, tipo_documento, archivo_base64):
    bucket_name = os.environ.get("BUCKET_AFILIACION", "sanitas-afiliacion-docs")

    if "," in archivo_base64:
        header, contenido_b64 = archivo_base64.split(",", 1)
    else:
        header, contenido_b64 = "", archivo_base64

    if "application/pdf" in header:
        content_type, ext = "application/pdf", "pdf"
    else:
        content_type, ext = "image/jpeg", "jpg"

    blob_path = f"afiliacion-regular/{token}/{persona_codigo}_{tipo_documento}.{ext}"
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_path)
    blob.upload_from_string(base64.b64decode(contenido_b64), content_type=content_type)

    return f"gs://{bucket_name}/{blob_path}" 

def _subir_imagen_storage(token, persona_codigo, tipo_documento, imagen_base64):
    """
    Sube la imagen decodificada al bucket PRIVADO de afiliación.
    Devuelve la ruta interna gs://... (no una URL pública, porque el
    bucket tiene public_access_prevention activado). Para ver el archivo
    después, usar generar_url_firmada_afiliacion().
    """
    bucket_name = os.environ.get("BUCKET_AFILIACION", "sanitas-afiliacion-docs")
    blob_path = f"afiliacion-regular/{token}/{persona_codigo}_{tipo_documento}.jpg"

    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_path)

    img_bytes = base64.b64decode(imagen_base64.split(",")[-1])
    blob.upload_from_string(img_bytes, content_type="image/jpeg")

    return f"gs://{bucket_name}/{blob_path}"
 
_credenciales_firma = None

def _obtener_credenciales_firma():
    """
    Obtiene (y cachea) credenciales con access_token + service_account_email,
    necesarias para firmar URLs vía la API de IAM (SignBlob) en vez de con
    una llave privada local -- que la cuenta de servicio default de Compute
    no tiene. Requiere que esa cuenta tenga el rol
    roles/iam.serviceAccountTokenCreator sobre sí misma.
    """
    global _credenciales_firma
    if _credenciales_firma is None:
        credenciales, _ = google.auth.default()
        credenciales.refresh(google_auth_requests.Request())
        _credenciales_firma = credenciales
    else:
        if not _credenciales_firma.valid:
            _credenciales_firma.refresh(google_auth_requests.Request())
    return _credenciales_firma


def generar_url_firmada_afiliacion(ruta_gs, minutos_expiracion=30):
    """
    Genera una URL temporal firmada para ver un documento del bucket privado
    de afiliación. Úsala solo desde endpoints internos protegidos (ej.
    handle_consultar_estado_afiliacion, handle_reporte_afiliacion), nunca la
    devuelvas en respuestas públicas del wizard.
    """
    if not ruta_gs or not ruta_gs.startswith("gs://"):
        return None

    bucket_name, blob_path = ruta_gs.replace("gs://", "", 1).split("/", 1)
    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_path)

    credenciales = _obtener_credenciales_firma()

    return blob.generate_signed_url(
        version="v4",
        expiration=timedelta(minutes=minutos_expiracion),
        method="GET",
        service_account_email=credenciales.service_account_email,
        access_token=credenciales.token,
    )

def _document_ai_extraer_texto(imagen_base64):
    """
    Placeholder: reusar el cliente de Document AI ya configurado en
    procesar-receta / programa-cronico (mismo processor_id de OCR genérico
    o uno dedicado a DNI/CE si ya lo tienes).
    Debe devolver el texto plano extraído del documento.
    """
    raise NotImplementedError("Conectar con el cliente Document AI existente en sofi-apis")
 
 
def _gemini_extraer_campos_dni(texto_ocr):
    """
    Placeholder: prompt a Gemini (gemini-2.5-flash-lite, con backoff exponencial
    en ResourceExhausted, igual que en agente_consulta.py) para estructurar el
    texto OCR en JSON:
    {
      "apellido_paterno": "...",
      "apellido_materno": "...",
      "nombres": "...",
      "tipo_documento": "DNI",
      "num_documento": "...",
      "fecha_nacimiento": "YYYY-MM-DD",
      "sexo": "M" | "F"
    }
    """
    raise NotImplementedError("Conectar con el cliente Gemini existente en sofi-apis")
 
 
# =====================================================================
# 4. GUARDAR PASO (persistencia incremental del wizard)
# =====================================================================
@con_log_errores
def handle_guardar_paso_afiliacion(request):
    """
    Guarda el progreso de un paso del wizard. El 'paso' determina qué tabla(s)
    se actualizan. Diseñado para poder llamarse repetidamente sin duplicar datos
    (upsert por token + persona_codigo donde aplique).
 
    Body esperado:
    {
      "token": "...",
      "paso": 2,                 # 1=empleadora/plan, 2=titular, 3=dependiente, ...
      "persona_codigo": "T",     # requerido en pasos 2 y 3
      "datos": { ... }           # campos propios de ese paso
    }
    """
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    paso = data.get("paso")
    persona_codigo = (data.get("persona_codigo") or "").strip().upper()
    datos = data.get("datos") or {}
 
    if not token or paso is None:
        return (json.dumps({"error": "token y paso son obligatorios"}), 400, HEADERS)
 
    conn = get_connection()
    cur = conn.cursor()
 
    cur.execute("SELECT id, estado FROM afiliacion_solicitudes WHERE token = %s", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)
    solicitud_id, estado = row
    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": "La solicitud ya no está en proceso"}), 409, HEADERS)
 
    if paso == 1:
        # Plan (y RUC ya fue asociado en consulta_ruc_afiliacion)
        cur.execute("""
            UPDATE afiliacion_solicitudes SET plan = %s, paso_actual = GREATEST(paso_actual, 2)
            WHERE id = %s
        """, (datos.get("plan"), solicitud_id))
 
    elif paso == 2:
        # Datos del titular
        cur.execute("""
            UPDATE afiliacion_solicitudes SET
                titular_apellido_paterno = %s, titular_apellido_materno = %s, titular_nombres = %s,
                titular_fecha_nacimiento = %s, titular_sexo = %s, titular_tipo_documento = %s,
                titular_num_documento = %s, titular_nacionalidad = %s, titular_estado_civil = %s,
                titular_direccion = %s, titular_distrito = %s, titular_provincia = %s,
                titular_departamento = %s, titular_telefono = %s, titular_celular = %s,
                titular_email = %s, paso_actual = GREATEST(paso_actual, 3)
            WHERE id = %s
        """, (
            datos.get("apellido_paterno"), datos.get("apellido_materno"), datos.get("nombres"),
            datos.get("fecha_nacimiento"), datos.get("sexo"), datos.get("tipo_documento"),
            datos.get("num_documento"), datos.get("nacionalidad"), datos.get("estado_civil"),
            datos.get("direccion"), datos.get("distrito"), datos.get("provincia"),
            datos.get("departamento"), datos.get("telefono"), datos.get("celular"),
            datos.get("email"), solicitud_id
        ))
 
    elif paso == 3:
        # Dependiente (upsert por solicitud_id + codigo)
        if persona_codigo not in {"D1", "D2", "D3", "D4"}:
            cur.close(); conn.close()
            return (json.dumps({"error": "persona_codigo inválido para dependiente"}), 400, HEADERS)
        cur.execute("""
            INSERT INTO afiliacion_dependientes
                (solicitud_id, codigo, apellido_paterno, apellido_materno, nombres,
                 sexo, fecha_nacimiento, tipo_documento, num_documento, parentesco)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (solicitud_id, codigo) DO UPDATE SET
                apellido_paterno = EXCLUDED.apellido_paterno,
                apellido_materno = EXCLUDED.apellido_materno,
                nombres = EXCLUDED.nombres,
                sexo = EXCLUDED.sexo,
                fecha_nacimiento = EXCLUDED.fecha_nacimiento,
                tipo_documento = EXCLUDED.tipo_documento,
                num_documento = EXCLUDED.num_documento,
                parentesco = EXCLUDED.parentesco
        """, (
            solicitud_id, persona_codigo, datos.get("apellido_paterno"), datos.get("apellido_materno"),
            datos.get("nombres"), datos.get("sexo"), datos.get("fecha_nacimiento"),
            datos.get("tipo_documento"), datos.get("num_documento"), datos.get("parentesco")
        ))
        cur.execute("""
            UPDATE afiliacion_solicitudes SET paso_actual = GREATEST(paso_actual, 4) WHERE id = %s
        """, (solicitud_id,))
 
    else:
        cur.close(); conn.close()
        return (json.dumps({"error": f"Paso {paso} no manejado por este endpoint"}), 400, HEADERS)
 
    conn.commit()
    cur.close(); conn.close()
 
    return (json.dumps({"ok": True, "solicitud_id": solicitud_id, "paso_guardado": paso}), 200, HEADERS)
 
 
# =====================================================================
# 5. GUARDAR DECLARACION DE SALUD (cifrada)
# =====================================================================
@con_log_errores
def handle_guardar_declaracion_salud(request):
    """
    Body esperado:
    {
      "token": "...",
      "persona_codigo": "T",
      "peso_kg": "75",
      "talla_cm": "170",
      "respuestas": {"1.2": "N", "1.3": "S", ..., "3.20": "N"},
      "embarazo_meses": null,
      "diagnosticos": [
        {
          "pregunta_codigo": "3.9",
          "diagnostico": "Hernia discal L4-L5",
          "extremidad": "N/A",
          "medico_tratante": "Dr. Perez",
          "fecha_diagnostico": "2024-03-10",
          "situacion_actual": "En tratamiento conservador"
        }
      ]
    }
    """
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    persona_codigo = (data.get("persona_codigo") or "").strip().upper()
 
    if not token or persona_codigo not in CODIGOS_PERSONA_VALIDOS:
        return (json.dumps({"error": "token y persona_codigo (T/D1-D4) son obligatorios"}), 400, HEADERS)

    peso_kg_raw = str(data.get("peso_kg") or "").strip()
    talla_cm_raw = str(data.get("talla_cm") or "").strip()

    if not peso_kg_raw or not talla_cm_raw:
        return (json.dumps({
            "error": "peso_kg y talla_cm son obligatorios para continuar",
            "persona_codigo": persona_codigo,
        }), 400, HEADERS)

    try:
        peso_valor = float(peso_kg_raw)
        talla_valor = float(talla_cm_raw)
    except (TypeError, ValueError):
        return (json.dumps({
            "error": "peso_kg y talla_cm deben ser valores numéricos",
            "persona_codigo": persona_codigo,
        }), 400, HEADERS)

    if not (1 <= peso_valor <= 300):
        return (json.dumps({
            "error": "peso_kg debe estar entre 1 y 300 kg",
            "persona_codigo": persona_codigo,
        }), 400, HEADERS)

    if not (30 <= talla_valor <= 250):
        return (json.dumps({
            "error": "talla_cm debe estar entre 30 y 250 cm",
            "persona_codigo": persona_codigo,
        }), 400, HEADERS)
 
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id FROM afiliacion_solicitudes WHERE token = %s AND estado = 'EN_PROCESO'", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada o no vigente"}), 404, HEADERS)
    solicitud_id = row[0]
 
    peso_enc = cifrar(peso_kg_raw)
    talla_enc = cifrar(talla_cm_raw)

    respuestas_enc = cifrar_json(data.get("respuestas") or {})
    embarazo_enc = cifrar(data.get("embarazo_meses")) if data.get("embarazo_meses") else None
 
    cur.execute("""
        INSERT INTO afiliacion_declaracion_salud
            (solicitud_id, persona_codigo, peso_kg_enc, talla_cm_enc, respuestas_enc, embarazo_meses_enc)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (solicitud_id, persona_codigo) DO UPDATE SET
            peso_kg_enc = EXCLUDED.peso_kg_enc,
            talla_cm_enc = EXCLUDED.talla_cm_enc,
            respuestas_enc = EXCLUDED.respuestas_enc,
            embarazo_meses_enc = EXCLUDED.embarazo_meses_enc,
            fecha_registro = now()
    """, (solicitud_id, persona_codigo, peso_enc, talla_enc, respuestas_enc, embarazo_enc))
 
    # Reemplazar detalle de diagnósticos de esta persona (borra e inserta de nuevo, más simple que diffing)
    cur.execute("""
        DELETE FROM afiliacion_diagnosticos_detalle WHERE solicitud_id = %s AND persona_codigo = %s
    """, (solicitud_id, persona_codigo))
 
    for diag in data.get("diagnosticos") or []:
        cur.execute("""
            INSERT INTO afiliacion_diagnosticos_detalle
                (solicitud_id, persona_codigo, pregunta_codigo, diagnostico_enc,
                 extremidad, medico_tratante_enc, fecha_diagnostico, situacion_actual_enc)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            solicitud_id, persona_codigo, diag.get("pregunta_codigo"),
            cifrar(diag.get("diagnostico")), diag.get("extremidad"),
            cifrar(diag.get("medico_tratante")), diag.get("fecha_diagnostico"),
            cifrar(diag.get("situacion_actual"))
        ))
 
    conn.commit()
    cur.close(); conn.close()
 
    return (json.dumps({"ok": True, "persona_codigo": persona_codigo}), 200, HEADERS)
 
 
# =====================================================================
# 6. GUARDAR ANTECEDENTES EN SISTEMA DE SALUD (seguro anterior)
# =====================================================================
@con_log_errores
def handle_guardar_antecedentes_seguro(request):
    """
    Body esperado:
    {
      "token": "...",
      "persona_codigo": "T",
      "tuvo_seguro": true,
      "tipo_seguro": "EPS",
      "modalidad": "COLECTIVO",
      "asegurador_anterior": "Rimac EPS",
      "fecha_desde": "2019-01-01",
      "fecha_hasta": "2023-12-31"
    }
    No es dato de salud propiamente, pero se guarda sin cifrar salvo el
    nombre del asegurador si prefieres tratarlo como sensible; aquí se deja
    en claro porque no revela condición de salud, solo historial comercial.
    """
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    persona_codigo = (data.get("persona_codigo") or "").strip().upper()
 
    if not token or persona_codigo not in CODIGOS_PERSONA_VALIDOS:
        return (json.dumps({"error": "token y persona_codigo (T/D1-D4) son obligatorios"}), 400, HEADERS)
 
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id FROM afiliacion_solicitudes WHERE token = %s AND estado = 'EN_PROCESO'", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada o no vigente"}), 404, HEADERS)
    solicitud_id = row[0]
 
    # Upsert manual (no hay UNIQUE(solicitud_id, persona_codigo) aquí porque
    # una persona podría declarar más de un seguro anterior; se borra e
    # inserta de nuevo cada vez que el front envía este paso, igual que
    # con diagnosticos_detalle).
    cur.execute("""
        DELETE FROM afiliacion_antecedentes_seguro WHERE solicitud_id = %s AND persona_codigo = %s
    """, (solicitud_id, persona_codigo))
 
    registros = data.get("registros")
    if registros is None:
        # compatibilidad si el front manda un solo registro plano
        registros = [data] if data.get("tuvo_seguro") is not None else []
 
    for reg in registros:
        cur.execute("""
            INSERT INTO afiliacion_antecedentes_seguro
                (solicitud_id, persona_codigo, tuvo_seguro, tipo_seguro, modalidad,
                 asegurador_anterior, fecha_desde, fecha_hasta)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            solicitud_id, persona_codigo, bool(reg.get("tuvo_seguro")),
            reg.get("tipo_seguro"), reg.get("modalidad"), reg.get("asegurador_anterior"),
            reg.get("fecha_desde"), reg.get("fecha_hasta")
        ))
 
    conn.commit()
    cur.close(); conn.close()
 
    return (json.dumps({"ok": True, "persona_codigo": persona_codigo}), 200, HEADERS)
 
 
# =====================================================================
# 7. GUARDAR FIRMA
# =====================================================================
TIPOS_DOCUMENTO_CIVIL_VALIDOS = {"PARTIDA_MATRIMONIO", "UNION_DE_HECHO"}
TIPOS_DOCUMENTO_ADICIONAL_VALIDOS = {"CARNET_CONADIS"}
TIPOS_DOCUMENTO_GENERICO_VALIDOS = TIPOS_DOCUMENTO_CIVIL_VALIDOS | TIPOS_DOCUMENTO_ADICIONAL_VALIDOS

@con_log_errores
def handle_guardar_documento_civil_afiliacion(request):
    """
    Sube el documento que sustenta el estado civil declarado:
    - CASADO      -> PARTIDA_MATRIMONIO
    - CONVIVIENTE -> UNION_DE_HECHO
    Aplica tanto al titular como a cualquier dependiente.
    """
    data = request.get_json(force=True, silent=True) or {}
    token = (data.get("token") or "").strip()
    persona_codigo = (data.get("persona_codigo") or "").strip().upper()
    tipo_documento = (data.get("tipo_documento") or "").strip().upper()
    archivo_base64 = data.get("archivo_base64", "")

    if not token or not archivo_base64:
        return (json.dumps({"error": "token y archivo_base64 son obligatorios"}), 400, HEADERS)
    if persona_codigo not in CODIGOS_PERSONA_VALIDOS:
        return (json.dumps({"error": "persona_codigo inválido"}), 400, HEADERS)
    if tipo_documento not in TIPOS_DOCUMENTO_GENERICO_VALIDOS:
        return (json.dumps({"error": "tipo_documento inválido"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, estado FROM afiliacion_solicitudes WHERE token = %s", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)
    solicitud_id, estado = row
    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": f"La solicitud ya no está en proceso (estado: {estado})"}), 409, HEADERS)

    # Si ya había uno del mismo tipo para esa persona, lo reemplaza (permite corregir)
    cur.execute("""
        DELETE FROM afiliacion_documentos
        WHERE solicitud_id = %s AND persona_codigo = %s AND tipo_documento = %s
    """, (solicitud_id, persona_codigo, tipo_documento))

    ruta = _subir_archivo_generico_storage(token, persona_codigo, tipo_documento, archivo_base64)

    cur.execute("""
        INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
        VALUES (%s, %s, %s, %s)
    """, (solicitud_id, persona_codigo, tipo_documento, ruta))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({"ok": True, "ruta_storage": ruta}), 200, HEADERS)

@con_log_errores
def handle_guardar_firma_afiliacion(request):
    """
    Solo el titular firma (a diferencia del programa crónico que usa firma dual).
    Body esperado: { "token": "...", "firma_base64": "data:image/png;base64,...." }
    """
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    firma_base64 = data.get("firma_base64", "")

    if not token or not firma_base64:
        return (json.dumps({"error": "token y firma_base64 son obligatorios"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id FROM afiliacion_solicitudes WHERE token = %s AND estado = 'EN_PROCESO'", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada o no vigente"}), 404, HEADERS)
    solicitud_id = row[0]

    ruta_firma = _subir_imagen_storage(token, "T", "FIRMA", firma_base64)

    cur.execute("""
        INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
        VALUES (%s, 'T', 'FIRMA', %s)
    """, (solicitud_id, ruta_firma))
    cur.execute("""
        UPDATE afiliacion_solicitudes SET ruta_firma = %s, paso_actual = GREATEST(paso_actual, 7)
        WHERE id = %s
    """, (ruta_firma, solicitud_id))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({"ok": True, "ruta_firma": ruta_firma}), 200, HEADERS)
 
 
# =====================================================================
# 8. FINALIZAR AFILIACION (genera PDF, envía correo y WhatsApp)
# =====================================================================
def _armar_zip_dni_afiliacion(cur, solicitud_id):
    """
    Descarga las fotos de DNI (anverso/reverso), el documento de estado civil
    (Partida de Matrimonio o Unión de Hecho) y el carné de CONADIS de cada
    dependiente incapacitado, y los empaqueta en un ZIP en memoria. No incluye
    la firma ni el PDF final.
    """
    cur.execute("""
        SELECT persona_codigo, tipo_documento, ruta_storage
        FROM afiliacion_documentos
        WHERE solicitud_id = %s AND tipo_documento IN
            ('DNI_ANVERSO', 'DNI_REVERSO', 'PARTIDA_MATRIMONIO', 'UNION_DE_HECHO', 'CARNET_CONADIS')
        ORDER BY persona_codigo, tipo_documento
    """, (solicitud_id,))
    documentos = cur.fetchall()

    if not documentos:
        return None

    buffer_zip = io.BytesIO()
    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for persona_codigo, tipo_documento, ruta_storage in documentos:
            if not ruta_storage or not ruta_storage.startswith("gs://"):
                continue
            bucket_name, blob_path = ruta_storage.replace("gs://", "", 1).split("/", 1)
            blob = storage_client.bucket(bucket_name).blob(blob_path)
            try:
                contenido = blob.download_as_bytes()
            except Exception as e:
                logging.warning(f"No se pudo descargar {ruta_storage} para el ZIP: {e}")
                continue

            extension = blob_path.rsplit(".", 1)[-1] if "." in blob_path else "jpg"
            nombre_en_zip = f"{persona_codigo}_{tipo_documento}.{extension}"
            zf.writestr(nombre_en_zip, contenido)

    buffer_zip.seek(0)
    return buffer_zip.read()

@con_log_errores
def handle_finalizar_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    token = data.get("token", "").strip()
    if not token:
        return (json.dumps({"error": "token es obligatorio"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, estado, titular_email, titular_nombres, titular_apellido_paterno,
               titular_apellido_materno, titular_celular, titular_num_documento,
               celular_origen_wsp
        FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    (solicitud_id, estado, email, nombres, apellido_paterno,
     apellido_materno, celular, num_documento, celular_origen_wsp) = row

    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": f"La solicitud ya está en estado {estado}"}), 409, HEADERS)

    contexto = _armar_contexto_pdf(cur, solicitud_id)
    zip_bytes = _armar_zip_dni_afiliacion(cur, solicitud_id)
    cur.close(); conn.close()

    ruta_pdf_local = _generar_pdf_afiliacion(token, contexto)
    pdf_bytes = open(ruta_pdf_local, "rb").read()
    ruta_storage_pdf = _subir_pdf_storage(token, ruta_pdf_local)

    _enviar_correo_afiliacion(
        email, nombres, apellido_paterno, apellido_materno,
        celular, num_documento, pdf_bytes, zip_bytes
    )
    _notificar_whatsapp_afiliacion(celular_origen_wsp or celular, nombres)  # ← ahora usa celular_origen_wsp, con fallback

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        UPDATE afiliacion_solicitudes
        SET estado = 'COMPLETADA', ruta_pdf_final = %s, paso_actual = 8, fecha_completada = now()
        WHERE id = %s
    """, (ruta_storage_pdf, solicitud_id))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({
        "ok": True,
        "solicitud_id": solicitud_id,
        "ruta_pdf": ruta_storage_pdf,
    }, ensure_ascii=False), 200, HEADERS)

 
 
def _armar_contexto_pdf(cur, solicitud_id):
    """
    Arma el diccionario de contexto para el template Jinja2, descifrando
    únicamente lo necesario y solo en memoria (nunca se vuelve a persistir
    en claro).
    """
    cur.execute("""
        SELECT * FROM afiliacion_solicitudes WHERE id = %s
    """, (solicitud_id,))
    col_names = [d[0] for d in cur.description]
    solicitud = dict(zip(col_names, cur.fetchone()))

    # ── Datos de la empresa (viven en sctr_empresa, no en afiliacion_solicitudes) ──
    if solicitud.get("ruc"):
        cur.execute("""
            SELECT razon_social, direccion, departamento, provincia, distrito
            FROM sctr_empresa WHERE ruc = %s
        """, (solicitud["ruc"],))
        row_empresa = cur.fetchone()
        if row_empresa:
            solicitud["empresa_razon_social"] = row_empresa[0]
            solicitud["empresa_direccion"] = row_empresa[1]
            solicitud["empresa_departamento"] = row_empresa[2]
            solicitud["empresa_provincia"] = row_empresa[3]
            solicitud["empresa_distrito"] = row_empresa[4]

    cur.execute("""
        SELECT codigo, apellido_paterno, apellido_materno, nombres, sexo,
               fecha_nacimiento, tipo_documento, num_documento, parentesco, nacionalidad, ubigeo
        FROM afiliacion_dependientes WHERE solicitud_id = %s ORDER BY codigo
    """, (solicitud_id,))

    dep_cols = [d[0] for d in cur.description]
    dependientes = [dict(zip(dep_cols, r)) for r in cur.fetchall()]

    cur.execute("""
        SELECT persona_codigo, peso_kg_enc, talla_cm_enc, respuestas_enc, embarazo_meses_enc
        FROM afiliacion_declaracion_salud WHERE solicitud_id = %s
    """, (solicitud_id,))
    declaracion_salud = {}
    for persona_codigo, peso_enc, talla_enc, respuestas_enc, embarazo_enc in cur.fetchall():
        declaracion_salud[persona_codigo] = {
            "peso_kg": descifrar(peso_enc),
            "talla_cm": descifrar(talla_enc),
            "respuestas": descifrar_json(respuestas_enc) or {},
            "embarazo_meses": descifrar(embarazo_enc),
        }

    cur.execute("""
    SELECT persona_codigo, pregunta_codigo, diagnostico_enc, extremidad,
           medico_tratante_enc, fecha_diagnostico, situacion_actual_enc
        FROM afiliacion_diagnosticos_detalle WHERE solicitud_id = %s
    """, (solicitud_id,))
        
    diagnosticos = []
    for pc, pregunta_codigo, diag_enc, extremidad, medico_enc, fecha_diag, situ_enc in cur.fetchall():
            diagnosticos.append({
                "persona_codigo": pc,
                "pregunta_codigo": pregunta_codigo,
                "diagnostico": descifrar(diag_enc),
                "extremidad": extremidad,
                "medico_tratante": descifrar(medico_enc),
                "fecha_diagnostico": fecha_diag,
                "situacion_actual": descifrar(situ_enc),
            })

    cur.execute("""
        SELECT persona_codigo, tuvo_seguro, tipo_seguro, modalidad,
               asegurador_anterior, fecha_desde, fecha_hasta
        FROM afiliacion_antecedentes_seguro WHERE solicitud_id = %s
    """, (solicitud_id,))
    ant_cols = [d[0] for d in cur.description]
    antecedentes_seguro = [dict(zip(ant_cols, r)) for r in cur.fetchall()]

    return {
        "solicitud": solicitud,
        "dependientes": dependientes,
        "declaracion_salud": declaracion_salud,
        "diagnosticos": diagnosticos,
        "antecedentes_seguro": antecedentes_seguro,
        "fecha_emision": datetime.now(LIMA_TZ).strftime("%d/%m/%Y"),
    }
 
PREGUNTAS_SALUD = [
    {"codigo": "1.2", "texto": "¿Pérdida de peso no explicada de más de 10 kg en los últimos 6 meses?", "solo_mujeres": False},
    {"codigo": "1.3", "texto": "¿Presenta alguna discapacidad física o mental?", "solo_mujeres": False},
    {"codigo": "1.4", "texto": "¿Hospitalización o atención de emergencia en los últimos 6 meses?", "solo_mujeres": False},
    {"codigo": "1.5", "texto": "¿Consume o ha consumido en forma repetida fármacos?", "solo_mujeres": False},
    {"codigo": "2.1", "texto": "¿Está actualmente embarazada?", "solo_mujeres": True},
    {"codigo": "3.1", "texto": "Neurológicas (migraña, epilepsia, convulsiones, vértigo, Parkinson, Alzheimer, ACV, etc.)", "solo_mujeres": False},
    {"codigo": "3.2", "texto": "Psicológicas o psiquiátricas (ansiedad, depresión, demencia, psicosis, etc.)", "solo_mujeres": False},
    {"codigo": "3.3", "texto": "Cardiovasculares (hipertensión, enfermedad coronaria, arritmias, várices, trombosis, etc.)", "solo_mujeres": False},
    {"codigo": "3.4", "texto": "Vías respiratorias (asma, bronquitis crónica, EPOC, sinusitis crónica, etc.)", "solo_mujeres": False},
    {"codigo": "3.5", "texto": "Gastrointestinales (gastritis crónica, úlcera péptica, colitis, hepatitis crónica, etc.)", "solo_mujeres": False},
    {"codigo": "3.6", "texto": "Pared abdominal (hernias, eventración, diastasis de rectos)", "solo_mujeres": False},
    {"codigo": "3.7", "texto": "Endócrinas (diabetes, hipo/hipertiroidismo, bocio, SOP, etc.)", "solo_mujeres": False},
    {"codigo": "3.8", "texto": "Aparatos urinario y genital (enf. renal crónica, cálculos, próstata, etc.)", "solo_mujeres": False},
    {"codigo": "3.9", "texto": "Huesos, articulaciones y tejido conjuntivo (artritis, artrosis, hernias de columna, etc.)", "solo_mujeres": False},
    {"codigo": "3.10", "texto": "Hematológicas (anemia, linfoma, leucemia, trastornos de coagulación, etc.)", "solo_mujeres": False},
    {"codigo": "3.11", "texto": "Oftalmológicas (estrabismo, catarata, glaucoma, ceguera, etc.)", "solo_mujeres": False},
    {"codigo": "3.12", "texto": "Otorrinolaringológicas (hipoacusia, sordera, otitis crónica, etc.)", "solo_mujeres": False},
    {"codigo": "3.13", "texto": "Dermatológicas (psoriasis, acné, carcinomas de piel, melanomas, etc.)", "solo_mujeres": False},
    {"codigo": "3.14", "texto": "Enfermedades congénitas o genéticas", "solo_mujeres": False},
    {"codigo": "3.15", "texto": "Enfermedades infecciosas (hepatitis B/C, VIH/SIDA, tuberculosis, etc.)", "solo_mujeres": False},
    {"codigo": "3.16", "texto": "Oncológicas (tumor benigno, maligno o cáncer en cualquier parte del cuerpo)", "solo_mujeres": False},
    {"codigo": "3.17", "texto": "¿Ha sido o tiene pendiente alguna intervención quirúrgica?", "solo_mujeres": False},
    {"codigo": "3.18", "texto": "Trasplante de órganos o tejidos", "solo_mujeres": False},
    {"codigo": "3.19", "texto": "Otras enfermedades o antecedentes médicos no consignados en la lista", "solo_mujeres": False},
    {"codigo": "3.20", "texto": "¿Tiene que ser hospitalizado o recibir tratamiento médico próximamente?", "solo_mujeres": False},
]

def _generar_pdf_afiliacion(token, contexto):
    jinja_env.filters["fmt_fecha"] = _fmt_fecha

    # Descargar la firma del bucket privado y convertirla a data URI para el PDF
    ruta_firma = contexto["solicitud"].get("ruta_firma")
    if ruta_firma:
        bucket_name, blob_path = ruta_firma.replace("gs://", "", 1).split("/", 1)
        firma_bytes = storage_client.bucket(bucket_name).blob(blob_path).download_as_bytes()
        contexto["firma_base64_para_pdf"] = f"data:image/png;base64,{base64.b64encode(firma_bytes).decode()}"
    else:
        contexto["firma_base64_para_pdf"] = ""

    contexto["preguntas_salud"] = PREGUNTAS_SALUD

    template = jinja_env.get_template("afiliacion_regular.html")
    html_renderizado = template.render(**contexto)

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    ruta_local = f"/tmp/afiliacion_{token}.pdf"
    HTML(string=html_renderizado, base_url=f"file://{BASE_DIR}/").write_pdf(ruta_local)
    
    return ruta_local
 
 
def _subir_pdf_storage(token, ruta_local):
    """Sube el PDF final al bucket privado. Devuelve la ruta interna gs://..."""
    bucket_name = os.environ.get("BUCKET_AFILIACION", "sanitas-afiliacion-docs")
    blob_path = f"afiliacion-regular/{token}/solicitud_final.pdf"

    bucket = storage_client.bucket(bucket_name)
    blob = bucket.blob(blob_path)
    blob.upload_from_filename(ruta_local, content_type="application/pdf")

    return f"gs://{bucket_name}/{blob_path}"
 
 
DANACONNECT_CONVERSATION_ID_AFILIACION = "600073"

# Correos que reciben copia de TODAS las solicitudes de afiliación completadas
CORREOS_ADICIONALES_AFILIACION = [
    "balcazar@sanitasperu.com",
    "mesadecontrol_eps@sanitasperu.com",
    "mteruya@sanitasperu.com",
    # agrega hasta 5 en total
]

def _enviar_correo_afiliacion(email, nombres, apellido_paterno, apellido_materno, celular, num_documento, pdf_bytes, zip_bytes=None):
    """
    Sube el PDF final (y opcionalmente el ZIP de documentos de identidad) a
    Danaconnect UNA sola vez, y dispara la conversación de afiliación una vez
    por cada destinatario: el correo del titular + los correos fijos definidos
    en CORREOS_ADICIONALES_AFILIACION.
    """
    if not DANACONNECT_CONVERSATION_ID_AFILIACION:
        logging.error("DANACONNECT_CONVERSATION_ID_AFILIACION no está configurada; no se envió el correo de afiliación")
        return

    file_id_pdf = subir_archivo_danaconnect(pdf_bytes, "solicitud_afiliacion.pdf", "application/pdf")
    logging.info(f"[DEBUG Danaconnect Afiliacion] file_id_pdf obtenido: {file_id_pdf}")

    file_id_zip = None
    if zip_bytes:
        file_id_zip = subir_archivo_danaconnect(zip_bytes, "documentos_dni.zip", "application/zip")
        logging.info(f"[DEBUG Danaconnect Afiliacion] file_id_zip obtenido: {file_id_zip}")

    url = f"https://appserv.danaconnect.com/api/1.0/rest/conversation/{DANACONNECT_CONVERSATION_ID_AFILIACION}/start/data"

    nombre_completo = " ".join(
        parte for parte in [nombres, apellido_paterno, apellido_materno] if parte
    ).strip()

    # ── Titular + correos fijos, sin duplicados ──
    destinatarios = []
    if email:
        destinatarios.append(email.strip())
    for correo_extra in CORREOS_ADICIONALES_AFILIACION[:5]:
        if correo_extra.lower() not in [d.lower() for d in destinatarios]:
            destinatarios.append(correo_extra)

    if not destinatarios:
        logging.warning("[Danaconnect Afiliacion] No hay destinatarios de correo para enviar")
        return

    for destinatario in destinatarios:
        payload = {
            "NOMBRE": nombre_completo,
            "EMAIL": destinatario,
            "CELULAR": celular,
            "DOI": num_documento,
            "FILE": file_id_pdf,
        }
        if file_id_zip:
            payload["FILE2"] = file_id_zip

        resp = requests.post(
            url,
            auth=(DANACONNECT_USER, DANACONNECT_PASSWORD),
            json=payload,
            timeout=30,
        )

        logging.info(f"[DEBUG Danaconnect Afiliacion] destinatario={destinatario} status={resp.status_code} body={resp.text}")
        resp.raise_for_status()

def _normalizar_celular_pe(celular):
    """
    Asegura que el número quede en formato +51XXXXXXXXX.
    Acepta: '987654321' (9 dígitos), '51987654321', '+51987654321'.
    Si no calza ningún patrón conocido, lo deja tal cual con un '+' antepuesto
    (mejor que nada, pero se loguea para revisión manual).
    """
    if not celular:
        return None

    # Detecta wa_id ofuscado de WhatsApp (número oculto), ej: "PE.1081517251058054"
    if re.match(r"^[A-Z]{2}\.\d+$", celular):
        logging.warning(f"[Afiliacion] Celular es un wa_id ofuscado (número oculto por WhatsApp), no se puede enviar SMS: {celular}")
        return None
        
    limpio = re.sub(r"[^\d]", "", celular)  # solo dígitos

    if len(limpio) == 9 and limpio.startswith("9"):
        return f"+51{limpio}"
    if len(limpio) == 11 and limpio.startswith("51"):
        return f"+{limpio}"
    if celular.startswith("+"):
        return celular

    logging.warning(f"[Afiliacion] Número de celular con formato inesperado, se envía tal cual: {celular}")
    return f"+{limpio}"

def _notificar_whatsapp_afiliacion(celular, nombres):
    """Reusa enviar_danaconnect(), ya definida en main.py, para confirmar por WhatsApp."""
    numero = _normalizar_celular_pe(celular)
    if not numero:
        return
    nombre_limpio = _limpiar_nombre_sms(nombres or "")
    mensaje = f"Sanitas: Hola {nombre_limpio}, hemos recibido tu solicitud de preafiliación correctamente. Te enviamos una copia a tu correo."
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "afilicion congreso")
 
def _obtener_documentos_con_urls_firmadas(cur, solicitud_id):
    """
    Trae todos los documentos (DNI, firma, PDF) de una solicitud con URLs
    firmadas temporales (30 min). Solo se usa desde endpoints internos
    protegidos -- nunca exponer esto al wizard público.
    """
    cur.execute("""
        SELECT persona_codigo, tipo_documento, ruta_storage
        FROM afiliacion_documentos WHERE solicitud_id = %s
        ORDER BY persona_codigo, tipo_documento
    """, (solicitud_id,))
    documentos = []
    for persona_codigo, tipo_documento, ruta_storage in cur.fetchall():
        documentos.append({
            "persona_codigo": persona_codigo,
            "tipo_documento": tipo_documento,
            "url_temporal": generar_url_firmada_afiliacion(ruta_storage),
        })
    return documentos

# =====================================================================
# 9. CONSULTAR ESTADO (uso interno - área médica / suscripción)
# =====================================================================
@con_log_errores
def handle_consultar_estado_afiliacion(request):
    """
    Endpoint interno (protegerlo con el mismo mecanismo de auth que uses para
    endpoints administrativos, ej. API key o IAM). Es el ÚNICO punto del
    sistema que descifra y expone la declaración de salud en claro, y el
    único que entrega URLs firmadas temporales para ver los documentos
    (DNI, firma, PDF final) del bucket privado sanitas-afiliacion-docs.
    """
    token = request.args.get("token", "").strip()
    if not token:
        return (json.dumps({"error": "token es obligatorio"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, estado FROM afiliacion_solicitudes WHERE token = %s", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    solicitud_id, estado = row
    contexto = _armar_contexto_pdf(cur, solicitud_id)
    contexto["documentos"] = _obtener_documentos_con_urls_firmadas(cur, solicitud_id)
    cur.close(); conn.close()

    contexto["solicitud"]["estado"] = estado
    return (json.dumps(contexto, ensure_ascii=False, default=str), 200, HEADERS)

@con_log_errores
def handle_registrar_titular_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    token = (data.get("token") or "").strip()
    correo = (data.get("correo") or "").strip()
    celular_body = (data.get("celular") or "").strip()
    url_dni_anverso = data.get("url_dni_anverso")
    url_dni_reverso = data.get("url_dni_reverso")

    if not token or not correo or not url_dni_anverso or not url_dni_reverso:
        return (json.dumps({
            "error": "token, correo, url_dni_anverso y url_dni_reverso son obligatorios"
        }), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, estado, titular_num_documento, celular_origen_wsp
        FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    solicitud_id, estado, dni, celular_origen_wsp = row
    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": f"La solicitud ya no está en proceso (estado: {estado})"}), 409, HEADERS)
    if not dni:
        cur.close(); conn.close()
        return (json.dumps({"error": "Esta solicitud no tiene dni_titular registrado (falta llamar /afiliacion/iniciar primero)"}), 409, HEADERS)

    # ── Respaldo: si el bot no vuelve a mandar 'celular' en este paso,
    # se usa el que ya se guardó en /afiliacion/iniciar (celular_origen_wsp),
    # para que titular_celular NUNCA quede vacío si ya se tiene un número.
    celular = celular_body or celular_origen_wsp
    if not celular:
        cur.close(); conn.close()
        return (json.dumps({"error": "No hay celular disponible (ni en el body ni en la solicitud original)"}), 400, HEADERS)

    # ── Descargar ambas imágenes ──
    try:
        resp_a = requests.get(url_dni_anverso, timeout=20)
        resp_a.raise_for_status()
        bytes_a = resp_a.content

        resp_r = requests.get(url_dni_reverso, timeout=20)
        resp_r.raise_for_status()
        bytes_r = resp_r.content
    except Exception as e:
        cur.close(); conn.close()
        return (json.dumps({"error": f"No se pudo descargar alguna de las imágenes: {str(e)}"}), 500, HEADERS)

    # ── Subir a Cloud Storage (bucket privado) para trazabilidad ──
    bucket_name = os.environ.get("BUCKET_AFILIACION", "sanitas-afiliacion-docs")
    bucket = storage_client.bucket(bucket_name)

    blob_a = bucket.blob(f"afiliacion-regular/{token}/T_DNI_ANVERSO.jpg")
    blob_a.upload_from_string(bytes_a, content_type="image/jpeg")
    ruta_a = f"gs://{bucket_name}/afiliacion-regular/{token}/T_DNI_ANVERSO.jpg"

    blob_r = bucket.blob(f"afiliacion-regular/{token}/T_DNI_REVERSO.jpg")
    blob_r.upload_from_string(bytes_r, content_type="image/jpeg")
    ruta_r = f"gs://{bucket_name}/afiliacion-regular/{token}/T_DNI_REVERSO.jpg"

    cur.execute("""
        INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
        VALUES (%s, 'T', 'DNI_ANVERSO', %s), (%s, 'T', 'DNI_REVERSO', %s)
    """, (solicitud_id, ruta_a, solicitud_id, ruta_r))
    conn.commit()

    # ── OCR + Gemini ──
    datos_dni = extraer_datos_dni(bytes_a, "image/jpeg", bytes_r, "image/jpeg")

    #ubigeo_titular = _buscar_ubigeo(cur, datos_dni.get("distrito"), datos_dni.get("provincia"), datos_dni.get("departamento"))
    ubigeo_titular = _formatear_ubigeo_texto(datos_dni.get("distrito"), datos_dni.get("provincia"), datos_dni.get("departamento"))

    apellidos_completos = (datos_dni.get("apellidos") or "").strip()
    partes_apellido = apellidos_completos.split(" ", 1)
    apellido_paterno = partes_apellido[0] if partes_apellido else ""
    apellido_materno = partes_apellido[1] if len(partes_apellido) > 1 else ""

    dni_ocr = (datos_dni.get("n_documento") or "").strip()
    coincide_dni = (not dni_ocr) or (dni_ocr == dni)

    sexo_ocr = datos_dni.get("sexo")
    estado_civil_ocr = datos_dni.get("estado_civil")
    nacionalidad_ocr = _mapear_nacionalidad(datos_dni.get("nacionalidad"))

    cur.execute("""
        UPDATE afiliacion_solicitudes SET
            titular_apellido_paterno = %s, titular_apellido_materno = %s, titular_nombres = %s,
            titular_fecha_nacimiento = %s, titular_tipo_documento = %s, titular_num_documento = %s,
            titular_direccion = %s, titular_distrito = %s, titular_provincia = %s,
            titular_departamento = %s, titular_ubigeo = %s, titular_celular = %s, titular_email = %s,
            titular_sexo = COALESCE(%s, titular_sexo),
            titular_estado_civil = COALESCE(%s, titular_estado_civil),
            titular_nacionalidad = COALESCE(%s, titular_nacionalidad),
            paso_actual = GREATEST(paso_actual, 3)
        WHERE id = %s
    """, (
        apellido_paterno, apellido_materno, datos_dni.get("nombres"),
        datos_dni.get("fecha_nacimiento"), datos_dni.get("tipo_documento") or "DNI", dni,
        datos_dni.get("domicilio"), datos_dni.get("distrito"), datos_dni.get("provincia"),
        datos_dni.get("departamento"), ubigeo_titular, celular, correo,
        sexo_ocr, estado_civil_ocr, nacionalidad_ocr,
        solicitud_id
    ))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({
        "ok": True,
        "solicitud_id": solicitud_id,
        "token": token,
        "coincide_dni_ocr": coincide_dni,
        "celular_guardado": celular,
        "campos_extraidos": {
            "nombres": datos_dni.get("nombres"),
            "apellido_paterno": apellido_paterno,
            "apellido_materno": apellido_materno,
            "fecha_nacimiento": datos_dni.get("fecha_nacimiento"),
            "distrito": datos_dni.get("distrito"),
            "sexo": sexo_ocr,
            "estado_civil": estado_civil_ocr,
            "nacionalidad": nacionalidad_ocr,
        },
        "faltan_por_completar_manualmente": {
            "sexo": sexo_ocr is None,
            "estado_civil": estado_civil_ocr is None,
            "nacionalidad": nacionalidad_ocr is None,
        },
    }, ensure_ascii=False), 200, HEADERS)

PARENTESCOS_VALIDOS = {"CONYUGE", "CONVIVIENTE", "HIJO", "HIJO_INCAPACITADO"}

@con_log_errores
def handle_registrar_dependiente_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    token = (data.get("token") or "").strip()
    parentesco = (data.get("parentesco") or "").strip().upper()
    url_dni_anverso = data.get("url_dni_anverso")
    url_dni_reverso = data.get("url_dni_reverso")

    if not token or not parentesco or not url_dni_anverso:
        return (json.dumps({
            "error": "token, parentesco y url_dni_anverso son obligatorios"
        }), 400, HEADERS)
    if parentesco not in PARENTESCOS_VALIDOS:
        return (json.dumps({"error": f"parentesco debe ser uno de: {', '.join(PARENTESCOS_VALIDOS)}"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, estado FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    solicitud_id, estado = row
    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": f"La solicitud ya no está en proceso (estado: {estado})"}), 409, HEADERS)

    # ── Asignar el siguiente código disponible (D1-D4) ──
    cur.execute("""
        SELECT codigo FROM afiliacion_dependientes WHERE solicitud_id = %s
    """, (solicitud_id,))
    codigos_usados = {r[0] for r in cur.fetchall()}
    codigo_asignado = None
    for candidato in ("D1", "D2", "D3", "D4"):
        if candidato not in codigos_usados:
            codigo_asignado = candidato
            break

    if not codigo_asignado:
        cur.close(); conn.close()
        return (json.dumps({"error": "Ya se registraron los 4 dependientes máximos permitidos"}), 409, HEADERS)

    # ── Descargar imágenes ──
    try:
        resp_a = requests.get(url_dni_anverso, timeout=20)
        resp_a.raise_for_status()
        bytes_a = resp_a.content

        bytes_r = None
        if url_dni_reverso:
            resp_r = requests.get(url_dni_reverso, timeout=20)
            resp_r.raise_for_status()
            bytes_r = resp_r.content
    except Exception as e:
        cur.close(); conn.close()
        return (json.dumps({"error": f"No se pudo descargar alguna imagen: {str(e)}"}), 500, HEADERS)

    # ── OCR + Gemini PRIMERO (aquí sale el DNI, ya no viene del body) ──
    datos_dni = extraer_datos_dni(bytes_a, "image/jpeg", bytes_r, "image/jpeg" if bytes_r else None)

    dni_dependiente = (datos_dni.get("n_documento") or "").strip()
    if not dni_dependiente:
        cur.close(); conn.close()
        return (json.dumps({
            "error": "No se pudo leer el número de documento en la foto. Pide una foto más clara del DNI (anverso) e intenta de nuevo."
        }), 422, HEADERS)

    # ── Evitar registrar el mismo DNI dos veces como dependiente ──
    cur.execute("""
        SELECT codigo FROM afiliacion_dependientes
        WHERE solicitud_id = %s AND num_documento = %s
    """, (solicitud_id, dni_dependiente))
    ya_existe = cur.fetchone()
    if ya_existe:
        cur.close(); conn.close()
        return (json.dumps({
            "error": f"Este DNI ya fue registrado como dependiente {ya_existe[0]} en esta solicitud"
        }), 409, HEADERS)

    # ── Subir a Cloud Storage (bucket privado) ──
    bucket_name = os.environ.get("BUCKET_AFILIACION", "sanitas-afiliacion-docs")
    bucket = storage_client.bucket(bucket_name)

    blob_a = bucket.blob(f"afiliacion-regular/{token}/{codigo_asignado}_DNI_ANVERSO.jpg")
    blob_a.upload_from_string(bytes_a, content_type="image/jpeg")
    ruta_a = f"gs://{bucket_name}/afiliacion-regular/{token}/{codigo_asignado}_DNI_ANVERSO.jpg"

    ruta_r = None
    if bytes_r:
        blob_r = bucket.blob(f"afiliacion-regular/{token}/{codigo_asignado}_DNI_REVERSO.jpg")
        blob_r.upload_from_string(bytes_r, content_type="image/jpeg")
        ruta_r = f"gs://{bucket_name}/afiliacion-regular/{token}/{codigo_asignado}_DNI_REVERSO.jpg"

    cur.execute("""
        INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
        VALUES (%s, %s, 'DNI_ANVERSO', %s)
    """, (solicitud_id, codigo_asignado, ruta_a))
    if ruta_r:
        cur.execute("""
            INSERT INTO afiliacion_documentos (solicitud_id, persona_codigo, tipo_documento, ruta_storage)
            VALUES (%s, %s, 'DNI_REVERSO', %s)
        """, (solicitud_id, codigo_asignado, ruta_r))
    conn.commit()

    apellidos_completos = (datos_dni.get("apellidos") or "").strip()
    partes_apellido = apellidos_completos.split(" ", 1)
    apellido_paterno = partes_apellido[0] if partes_apellido else ""
    apellido_materno = partes_apellido[1] if len(partes_apellido) > 1 else ""
    sexo_ocr = datos_dni.get("sexo")
    estado_civil_ocr = datos_dni.get("estado_civil")
    nacionalidad_ocr = _mapear_nacionalidad(datos_dni.get("nacionalidad"))

    # El dependiente SIEMPRE hereda el mismo texto de ubicación que el titular,
    # sin importar lo que su propio DNI diga (ya no se calcula por OCR propio).
    cur.execute("SELECT titular_ubigeo FROM afiliacion_solicitudes WHERE id = %s", (solicitud_id,))
    row_ubigeo_titular = cur.fetchone()
    ubigeo_dependiente = row_ubigeo_titular[0] if row_ubigeo_titular else None

    cur.execute("""
        INSERT INTO afiliacion_dependientes
            (solicitud_id, codigo, apellido_paterno, apellido_materno, nombres,
             sexo, fecha_nacimiento, tipo_documento, num_documento, parentesco,
             nacionalidad, ubigeo, estado_civil)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        solicitud_id, codigo_asignado, apellido_paterno, apellido_materno,
        datos_dni.get("nombres"), sexo_ocr, datos_dni.get("fecha_nacimiento"),
        datos_dni.get("tipo_documento") or "DNI", dni_dependiente, parentesco,
        nacionalidad_ocr, ubigeo_dependiente, estado_civil_ocr
    ))

    cur.execute("""
        UPDATE afiliacion_solicitudes SET paso_actual = GREATEST(paso_actual, 4) WHERE id = %s
    """, (solicitud_id,))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({
        "ok": True,
        "solicitud_id": solicitud_id,
        "persona_codigo": codigo_asignado,
        "campos_extraidos": {
            "nombres": datos_dni.get("nombres"),
            "apellido_paterno": apellido_paterno,
            "apellido_materno": apellido_materno,
            "num_documento": dni_dependiente,
            "fecha_nacimiento": datos_dni.get("fecha_nacimiento"),
            "parentesco": parentesco,
            "sexo": sexo_ocr,
            "nacionalidad": nacionalidad_ocr,
            "estado_civil": estado_civil_ocr,
        },
        "falta_completar_sexo_manualmente": sexo_ocr is None,
        "falta_completar_nacionalidad_manualmente": nacionalidad_ocr is None,
    }, ensure_ascii=False), 200, HEADERS)

def handle_ver_wizard_afiliacion(request, token):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, estado, plan, titular_apellido_paterno, titular_apellido_materno,
               titular_nombres, titular_sexo, titular_estado_civil, titular_num_documento,
               ruc, titular_ubigeo
        FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()

    if not row:
        cur.close(); conn.close()
        return (_pagina_mensaje("Enlace no válido", "Este enlace expiró o no es correcto. Contáctanos por WhatsApp si necesitas ayuda."), 404, {"Content-Type": "text/html; charset=utf-8"})

    (solicitud_id, estado, plan, apellido_paterno, apellido_materno, nombres, sexo, estado_civil,
     num_documento, ruc, titular_ubigeo) = row

    if estado == "COMPLETADA":
        cur.close(); conn.close()
        return (_pagina_mensaje("¡Ya completaste tu solicitud!", "Revisa tu correo para ver la constancia de tu afiliación."), 200, {"Content-Type": "text/html; charset=utf-8"})

    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (_pagina_mensaje("Solicitud no disponible", f"Esta solicitud está en estado: {estado}. Contáctanos por WhatsApp si necesitas ayuda."), 200, {"Content-Type": "text/html; charset=utf-8"})

    cur.execute("""
        SELECT codigo, apellido_paterno, apellido_materno, nombres, sexo, parentesco, ubigeo, estado_civil
        FROM afiliacion_dependientes WHERE solicitud_id = %s ORDER BY codigo
    """, (solicitud_id,))
    dep_rows = cur.fetchall()

    dependientes = [
        {
            "codigo": d[0], "apellido_paterno": d[1], "apellido_materno": d[2],
            "nombres": d[3], "sexo": d[4], "parentesco": d[5], "ubigeo": d[6],
            "estado_civil": d[7],
        }
        for d in dep_rows
    ]

    # sexo se conserva en el diccionario (lo usa la declaración de salud para
    # filtrar la pregunta de embarazo), pero YA NO se pide/edita en el wizard.
    personas = [{
        "codigo": "T",
        "nombre_completo": f"{nombres or ''} {apellido_paterno or ''} {apellido_materno or ''}".strip(),
        "sexo": sexo,
        "estado_civil": estado_civil,
        "es_titular": True,
    }] + [
        {
            "codigo": d["codigo"],
            "nombre_completo": f"{d['nombres'] or ''} {d['apellido_paterno'] or ''} {d['apellido_materno'] or ''}".strip(),
            "sexo": d["sexo"],
            "es_titular": False,
            "parentesco": d["parentesco"],
            "estado_civil": d["estado_civil"],
        }
        for d in dependientes
    ]

    # ── Documento civil: se solicita UNA sola vez por solicitud, no por persona ──
    MAPA_DOCUMENTO_CIVIL = {"CASADO": "PARTIDA_MATRIMONIO", "CONVIVIENTE": "UNION_DE_HECHO"}

    tipos_civiles_requeridos = sorted({
        MAPA_DOCUMENTO_CIVIL[p["estado_civil"]]
        for p in personas
        if p.get("estado_civil") in MAPA_DOCUMENTO_CIVIL
    })

    cur.execute("""
        SELECT DISTINCT tipo_documento FROM afiliacion_documentos
        WHERE solicitud_id = %s AND tipo_documento IN ('PARTIDA_MATRIMONIO', 'UNION_DE_HECHO')
    """, (solicitud_id,))
    tipos_civiles_subidos = {r[0] for r in cur.fetchall()}

    documentos_civiles = [
        {"tipo": tipo, "subido": tipo in tipos_civiles_subidos}
        for tipo in tipos_civiles_requeridos
    ]
    faltan_documentos_civiles = any(not d["subido"] for d in documentos_civiles)

    # ── NUEVO: Carné de CONADIS -- uno por cada dependiente HIJO_INCAPACITADO ──
    mapa_dependientes = {d["codigo"]: d for d in dependientes}
    codigos_incapacitados = [
        d["codigo"] for d in dependientes if d["parentesco"] == "HIJO_INCAPACITADO"
    ]

    conadis_subidos = set()
    if codigos_incapacitados:
        cur.execute("""
            SELECT DISTINCT persona_codigo FROM afiliacion_documentos
            WHERE solicitud_id = %s AND tipo_documento = 'CARNET_CONADIS'
              AND persona_codigo = ANY(%s)
        """, (solicitud_id, codigos_incapacitados))
        conadis_subidos = {r[0] for r in cur.fetchall()}

    documentos_conadis = [
        {
            "persona_codigo": cod,
            "nombre_completo": f"{mapa_dependientes[cod].get('nombres') or ''} "
                                f"{mapa_dependientes[cod].get('apellido_paterno') or ''} "
                                f"{mapa_dependientes[cod].get('apellido_materno') or ''}".strip(),
            "subido": cod in conadis_subidos,
        }
        for cod in codigos_incapacitados
    ]
    faltan_documentos_conadis = any(not d["subido"] for d in documentos_conadis)

    cur.close(); conn.close()

    # Ya no se exige sexo ni ubigeo (se obtienen de la foto del documento
    # en el registro de titular/dependiente, sin intervención manual).
    faltan_datos_basicos = (
        (not plan) or (not estado_civil)
        or any(not p.get("estado_civil") for p in personas)
        or faltan_documentos_civiles
        or faltan_documentos_conadis  # ← NUEVO
    )

    template = jinja_env.get_template("wizard_afiliacion.html")
    html_renderizado = template.render(
        token=token,
        personas=personas,
        faltan_datos_basicos=faltan_datos_basicos,
        plan_actual=plan or "",
        titular_estado_civil=estado_civil or "",
        documentos_civiles=documentos_civiles,
        documentos_conadis=documentos_conadis,  # ← NUEVO
        preguntas_salud=PREGUNTAS_SALUD,
    )

    return (html_renderizado, 200, {"Content-Type": "text/html; charset=utf-8"})

@con_log_errores
def handle_completar_datos_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    token = (data.get("token") or "").strip()
    plan = data.get("plan")
    titular_sexo = data.get("titular_sexo")
    titular_estado_civil = data.get("titular_estado_civil")
    titular_ubigeo = data.get("titular_ubigeo")
    dependientes = data.get("dependientes") or []

    if not token:
        return (json.dumps({"error": "token es obligatorio"}), 400, HEADERS)

    ESTADOS_CIVILES_VALIDOS = {"SOLTERO", "CASADO", "VIUDO", "DIVORCIADO", "CONVIVIENTE"}

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT id, estado FROM afiliacion_solicitudes WHERE token = %s", (token,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    solicitud_id, estado = row
    if estado != "EN_PROCESO":
        cur.close(); conn.close()
        return (json.dumps({"error": f"La solicitud ya no está en proceso (estado: {estado})"}), 409, HEADERS)

    cur.execute("""
        UPDATE afiliacion_solicitudes SET
            plan = COALESCE(%s, plan),
            titular_sexo = COALESCE(%s, titular_sexo),
            titular_estado_civil = COALESCE(%s, titular_estado_civil),
            titular_ubigeo = COALESCE(%s, titular_ubigeo),
            paso_actual = GREATEST(paso_actual, 5)
        WHERE id = %s
    """, (plan, titular_sexo, titular_estado_civil, titular_ubigeo, solicitud_id))

    for dep in dependientes:
        codigo = (dep.get("codigo") or "").strip().upper()
        sexo = dep.get("sexo")
        ubigeo_dep = dep.get("ubigeo")
        estado_civil_dep = dep.get("estado_civil")

        if estado_civil_dep and estado_civil_dep not in ESTADOS_CIVILES_VALIDOS:
            cur.close(); conn.close()
            return (json.dumps({"error": f"estado_civil de {codigo} no es válido"}), 400, HEADERS)

        if codigo in {"D1", "D2", "D3", "D4"} and (sexo or ubigeo_dep or estado_civil_dep):
            cur.execute("""
                UPDATE afiliacion_dependientes SET
                    sexo = COALESCE(%s, sexo),
                    ubigeo = COALESCE(%s, ubigeo),
                    estado_civil = COALESCE(%s, estado_civil)
                WHERE solicitud_id = %s AND codigo = %s
            """, (sexo, ubigeo_dep, estado_civil_dep, solicitud_id, codigo))

    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({"ok": True}), 200, HEADERS)

@con_log_errores
def handle_reenviar_sms_afiliacion(request):
    data = request.get_json(force=True, silent=True) or {}
    if data.get("pw") != ADMIN_PW_AFILIACION:
        return (json.dumps({"error": "No autorizado"}), 401, HEADERS)

    token = (data.get("token") or "").strip()
    if not token:
        return (json.dumps({"error": "token es obligatorio"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT estado, celular_origen_wsp, titular_nombres
        FROM afiliacion_solicitudes WHERE token = %s
    """, (token,))
    row = cur.fetchone()
    cur.close(); conn.close()

    if not row:
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    estado, celular_origen_wsp, titular_nombres = row

    if estado != "EN_PROCESO":
        return (json.dumps({
            "error": f"Solo se puede reenviar el link para solicitudes EN_PROCESO (estado actual: {estado})"
        }), 409, HEADERS)

    if not celular_origen_wsp:
        return (json.dumps({"error": "Esta solicitud no tiene celular registrado"}), 409, HEADERS)

    numero = _normalizar_celular_pe(celular_origen_wsp)
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link = f"{base_url}/afiliacion/{token}"
    nombre_limpio = _limpiar_nombre_sms(titular_nombres or "")
    saludo = f"Hola {nombre_limpio}, " if nombre_limpio else "Hola, "
    mensaje = f"Sanitas: {saludo}completa tu solicitud de afiliación aquí: {link}"

    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "afiliacion reenvio link")

    return (json.dumps({"ok": True, "celular": numero}), 200, HEADERS)

"""
handle_reporte_afiliacion
-----------------------------
Panel de reporte de solicitudes de afiliación. Mismo patrón de seguridad
que /programa-cronico/reporte y /campania/*: password por query string
(?pw=...), comparado contra una variable de entorno.

Ruta sugerida: GET /afiliacion/reporte?pw=...

IMPORTANTE: agregar la ruta ANTES del catch-all
    elif request.path.startswith("/afiliacion/") and request.method == "GET":
en el router, o como un elif exacto de path == "/afiliacion/reporte" en
cualquier posición previa a ese catch-all (el orden del if/elif ya lo
resuelve solo, ya que Python evalúa de arriba hacia abajo).

Requiere la variable de entorno ADMIN_PW_AFILIACION (Secret Manager o
env var directa, igual que ADMIN_PW_CRONICO / AGENTE_KB_CLAVE).
"""

ADMIN_PW_AFILIACION = os.environ.get("ADMIN_PW_AFILIACION")

@con_log_errores
def handle_reporte_afiliacion(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_AFILIACION:
        return ("No autorizado", 401, {"Content-Type": "text/html; charset=utf-8"})

    filtro_estado = request.args.get("estado", "")
    filtro_busqueda = request.args.get("q", "").strip()
    fecha_inicio = request.args.get("fecha_inicio", "").strip()
    fecha_fin = request.args.get("fecha_fin", "").strip()

    conn = get_connection()
    with conn.cursor() as cur:
        query = """
            SELECT
                s.token, s.tipo_solicitud, s.plan, s.estado,
                s.titular_num_documento, s.titular_nombres,
                s.titular_apellido_paterno, s.titular_apellido_materno,
                s.celular_origen_wsp, s.titular_email,
                s.ruc, e.razon_social,
                s.fecha_creacion, s.fecha_completada, s.ruta_pdf_final,
                (SELECT COUNT(*) FROM afiliacion_dependientes d WHERE d.solicitud_id = s.id) AS num_dependientes
            FROM afiliacion_solicitudes s
            LEFT JOIN sctr_empresa e ON e.ruc = s.ruc
            WHERE 1=1
        """
        params = []

        if filtro_estado:
            query += " AND s.estado = %s"
            params.append(filtro_estado)

        if filtro_busqueda:
            query += """ AND (
                s.titular_num_documento ILIKE %s OR
                s.titular_nombres ILIKE %s OR
                s.titular_apellido_paterno ILIKE %s OR
                s.titular_apellido_materno ILIKE %s OR
                s.celular_origen_wsp ILIKE %s OR
                e.razon_social ILIKE %s
            )"""
            comodin = f"%{filtro_busqueda}%"
            params.extend([comodin] * 6)

        if fecha_inicio:
            query += " AND s.fecha_creacion >= %s"
            params.append(fecha_inicio)

        if fecha_fin:
            query += " AND s.fecha_creacion < (%s::date + INTERVAL '1 day')"
            params.append(fecha_fin)

        query += " ORDER BY s.fecha_creacion DESC LIMIT 500"
        cur.execute(query, params)
        registros = cur.fetchall()
    conn.close()

    COLOR_ESTADO = {
        "COMPLETADA": "#2e7d32",
        "EN_PROCESO": "#0090A6",
        "VENCIDA": "#c62828",
    }

    filas_html = ""
    for row in registros:
        (token, tipo_solicitud, plan, estado, num_documento, nombres,
         apellido_paterno, apellido_materno, celular, email,
         ruc, razon_social, fecha_creacion, fecha_completada, ruta_pdf_final,
         num_dependientes) = row
        color_estado = COLOR_ESTADO.get(estado, "#555")
        nombre_completo = f"{nombres or ''} {apellido_paterno or ''} {apellido_materno or ''}".strip() or "—"

        links = [f'<a href="/afiliacion/{token}" target="_blank">Ver wizard</a>']
        if ruta_pdf_final:
            url_firmada = generar_url_firmada_afiliacion(ruta_pdf_final, minutos_expiracion=30)
            if url_firmada:
                links.append(f'<a href="{url_firmada}" target="_blank"><strong>PDF</strong></a>')
        links_html = " · ".join(links)

        btn_sms_html = "—"
        if estado == "EN_PROCESO" and celular:
            btn_sms_html = f'<button class="btn-sms" data-token="{token}" onclick="reenviarSMS(this)">📲 Enviar SMS</button>'

        filas_html += f"""
        <tr>
            <td>{fecha_creacion.strftime('%d/%m/%Y %H:%M') if fecha_creacion else '—'}</td>
            <td>{nombre_completo}<br><span style="font-size:12px;color:#888;">{num_documento or ''}</span></td>
            <td>{celular or '—'}<br><span style="font-size:12px;color:#888;">{email or ''}</span></td>
            <td>{razon_social or '—'}<br><span style="font-size:12px;color:#888;">{ruc or ''}</span></td>
            <td>{plan or '—'}</td>
            <td>{tipo_solicitud or '—'}</td>
            <td style="text-align:center;">{num_dependientes}</td>
            <td><span style="color:{color_estado};font-weight:600;">{estado}</span></td>
            <td>{links_html}</td>
            <td>{btn_sms_html}</td>
            <td><span style="font-size:11px;color:#aaa;">{token}</span></td>
        </tr>"""

    opciones_estado = ["", "EN_PROCESO", "COMPLETADA", "VENCIDA"]
    opciones_estado_html = "".join(
        f'<option value="{e}" {"selected" if e == filtro_estado else ""}>{e or "Todos"}</option>'
        for e in opciones_estado
    )

    # ── Construye el link de descarga de Excel arrastrando los mismos filtros activos ──
    from urllib.parse import urlencode
    params_excel = {"pw": pw}
    if filtro_busqueda:
        params_excel["q"] = filtro_busqueda
    if fecha_inicio:
        params_excel["fecha_inicio"] = fecha_inicio
    if fecha_fin:
        params_excel["fecha_fin"] = fecha_fin
    url_excel = f"/afiliacion/reporte-excel?{urlencode(params_excel)}"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Reporte de Afiliaciones</title>
<link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
<style>
    body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background: #f7f9fa; margin: 0; padding: 20px; color: #222; }}
    h1 {{ color: #0090A6; font-size: 20px; margin-bottom: 4px; }}
    .subtitulo {{ color: #888; font-size: 13px; margin-bottom: 20px; }}
    .filtros {{ background: #fff; border-radius: 10px; padding: 14px 18px; margin-bottom: 16px; display: flex; gap: 12px; align-items: end; flex-wrap: wrap; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
    .filtros label {{ font-size: 12px; font-weight: 600; display: block; margin-bottom: 4px; }}
    .filtros input, .filtros select {{ padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }}
    .filtros button {{ padding: 8px 16px; background: #0090A6; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }}
    .btn-excel {{ padding: 8px 16px; background: #2e7d32; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; text-decoration: none; display: inline-block; }}
    .btn-excel:hover {{ background: #256428; text-decoration: none; }}
    table {{ width: 100%; border-collapse: collapse; background: #fff; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
    th {{ background: #0090A6; color: #fff; text-align: left; padding: 10px 12px; font-size: 12px; }}
    td {{ padding: 10px 12px; font-size: 13px; border-bottom: 1px solid #eee; vertical-align: top; }}
    tr:hover {{ background: #f0f7f8; }}
    .contador {{ font-size: 12px; color: #888; margin-bottom: 10px; }}
    a {{ color: #0090A6; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .table-wrap {{ overflow-x: auto; }}
    .btn-sms {{ padding: 6px 12px; background: #0090A6; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 12px; }}
    .btn-sms:hover {{ background: #007686; }}
    .btn-sms:disabled {{ background: #9bb6d6; cursor: not-allowed; }}
</style>
</head>
<body>

<h1>Reporte — Afiliación Regular</h1>
<div class="subtitulo">Solicitudes de titulares y sus dependientes</div>

<form class="filtros" method="GET" action="/afiliacion/reporte">
    <input type="hidden" name="pw" value="{pw}">
    <div>
        <label>Buscar (DNI, nombre, celular, empresa)</label>
        <input type="text" name="q" value="{filtro_busqueda}" placeholder="Ej. 47313070">
    </div>
    <div>
        <label>Estado</label>
        <select name="estado">{opciones_estado_html}</select>
    </div>
    <div>
        <label>Desde</label>
        <input type="date" name="fecha_inicio" value="{fecha_inicio}">
    </div>
    <div>
        <label>Hasta</label>
        <input type="date" name="fecha_fin" value="{fecha_fin}">
    </div>
    <button type="submit">Filtrar</button>
    <a class="btn-excel" href="{url_excel}">📥 Descargar Excel</a>
</form>

<div class="contador">{len(registros)} registro(s) (máx. 500)</div>

<div class="table-wrap">
<table>
    <thead>
        <tr>
            <th>Fecha</th>
            <th>Titular</th>
            <th>Contacto</th>
            <th>Empresa</th>
            <th>Plan</th>
            <th>Tipo</th>
            <th>Dependientes</th>
            <th>Estado</th>
            <th>Enlaces</th>
            <th>SMS</th>
            <th>Token</th>
        </tr>
    </thead>
    <tbody>
        {filas_html}
    </tbody>
</table>
</div>
<script>
async function reenviarSMS(btn) {{
    const token = btn.dataset.token;
    if (!confirm("¿Enviar SMS con el link de afiliación a este celular?")) return;

    btn.disabled = true;
    btn.textContent = "Enviando...";

    try {{
        const resp = await fetch("/afiliacion/reenviar-sms", {{
            method: "POST",
            headers: {{ "Content-Type": "application/json" }},
            body: JSON.stringify({{ pw: "{pw}", token: token }})
        }});
        const data = await resp.json();

        if (resp.ok && data.ok) {{
            btn.textContent = "✓ Enviado";
        }} else {{
            alert("Error: " + (data.error || "No se pudo enviar"));
            btn.disabled = false;
            btn.textContent = "📲 Enviar SMS";
        }}
    }} catch (e) {{
        alert("Error de conexión");
        btn.disabled = false;
        btn.textContent = "📲 Enviar SMS";
    }}
}}
</script>
</body>
</html>"""

    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})

# Lista de campos en la cabecera para generar la trama (excel)
REPORTE_EXCEL_HEADERS = [
    "PLAN*",
    "VIGENCIA*",
    "TIPO DOCUMENTO TITULAR*",
    "NÚMERO DOCUMENTO DEL TITULAR*",
    "TIPO DOCUMENTO USUARIO*",
    "NÚMERO DOCUMENTO DE IDENTIDAD DEL USUARIO*",
    "1ER NOMBRE*",
    "2DO NOMBRE O MAS NOMBRES*",
    "1ER APELLIDO*",
    "2DO APELLIDO O MAS APELLIDOS*",
    "SEXO*",
    "FECHA DE NACIMIENTO*",
    "PAIS DE ORIGEN (NACIONALIDAD)*",
    "ESTADO CIVIL*",
    "PARENTESCO*",
    "TIPO VIA*",
    "DIRECCION*",
    "NUMERO*",
    "UBIGEO*",
    "TELEFONO FIJO*",
    "CELULAR*",
    "E-MAIL*",
    "TIPO*",
    "ANTERIOR ASEGURADOR SALUD*",
    "ANTERIOR EPS*",
    "FECHA INICIO EPS*",
    "FECHA FIN EPS*",
    "CONTACTO REFERENTE*",
]

# Prefijos de tipo de vía reconocidos al inicio de titular_direccion
_TIPOS_VIA_REPORTE = [
    "AV.", "AV", "JR.", "JR", "CALLE", "CA.", "CA", "PSJE.", "PSJE",
    "PASAJE", "PROLONG.", "PROLONGACION", "PROLONGACIÓN", "MZ.", "MZ",
    "URB.", "URB", "ALAMEDA", "MALECON", "MALECÓN",
]
_TIPO_VIA_REGEX_REPORTE = re.compile(
    r"^\s*(" + "|".join(re.escape(t) for t in _TIPOS_VIA_REPORTE) + r")\.?\s+",
    re.IGNORECASE,
)
_NUMERO_REGEX_REPORTE = re.compile(r"\b(\d+)\b")


def _parsear_direccion_reporte(direccion_raw):
    """
    'JR. ALMIRANTE GUISSE 1904' -> ('JR', 'ALMIRANTE GUISSE', '1904')
    'AV. PASEO DE LA REPUBLICA 6339 DPTO. 1204'
        -> ('AV', 'PASEO DE LA REPUBLICA DPTO. 1204', '6339')
    'LOS PRECURSORES 453 DPT.101' -> ('', 'LOS PRECURSORES DPT.101', '453')
    """
    if not direccion_raw:
        return "", "", ""

    texto = direccion_raw.strip()

    tipo_via = ""
    m = _TIPO_VIA_REGEX_REPORTE.match(texto)
    if m:
        tipo_via = m.group(1).rstrip(".").upper()
        texto = texto[m.end():].strip()

    numero = ""
    m_num = _NUMERO_REGEX_REPORTE.search(texto)
    if m_num:
        numero = m_num.group(1)
        antes = texto[: m_num.start()].strip()
        despues = texto[m_num.end():].strip()
        direccion = " ".join(p for p in [antes, despues] if p).strip()
    else:
        direccion = texto

    return tipo_via, direccion, numero


def _separar_nombres_reporte(nombres_raw):
    """1er nombre = primera palabra; 2do nombre o más = el resto."""
    if not nombres_raw:
        return "", ""
    partes = nombres_raw.strip().split()
    if not partes:
        return "", ""
    return partes[0], " ".join(partes[1:])


def _calcular_vigencia_reporte(fecha_completada):
    """
    La vigencia siempre es el primer día del mismo mes en que se completó
    la solicitud (firma), sin importar si fue el día 1 o el día 31.
    Ej: completada el 08/08/2026 -> vigencia 01/08/2026.
        completada el 31/08/2026 -> vigencia 01/08/2026.
    """
    if not fecha_completada:
        return None

    f = fecha_completada.date() if hasattr(fecha_completada, "date") else fecha_completada
    return date(f.year, f.month, 1)


def _contacto_referente_reporte(celular_origen_wsp):
    return "SI" if celular_origen_wsp else "NO"


def _obtener_datos_reporte_afiliacion(cur, fecha_inicio=None, fecha_fin=None, busqueda=None):
    query = """
        SELECT s.id, s.plan, s.titular_tipo_documento, s.titular_num_documento,
               s.titular_nombres, s.titular_apellido_paterno, s.titular_apellido_materno,
               s.titular_sexo, s.titular_fecha_nacimiento, s.titular_nacionalidad,
               s.titular_estado_civil, s.titular_direccion, s.titular_ubigeo, s.titular_telefono,
               s.titular_celular, s.celular_origen_wsp, s.titular_email, s.fecha_completada
        FROM afiliacion_solicitudes s
        LEFT JOIN sctr_empresa e ON e.ruc = s.ruc
        WHERE s.estado = 'COMPLETADA'
    """
    params = []

    if fecha_inicio:
        query += " AND s.fecha_creacion >= %s"
        params.append(fecha_inicio)

    if fecha_fin:
        query += " AND s.fecha_creacion < (%s::date + INTERVAL '1 day')"
        params.append(fecha_fin)

    if busqueda:
        query += """ AND (
            s.titular_num_documento ILIKE %s OR
            s.titular_nombres ILIKE %s OR
            s.titular_apellido_paterno ILIKE %s OR
            s.titular_apellido_materno ILIKE %s OR
            s.celular_origen_wsp ILIKE %s OR
            e.razon_social ILIKE %s
        )"""
        comodin = f"%{busqueda}%"
        params.extend([comodin] * 6)

    query += " ORDER BY s.id"
    cur.execute(query, params)

    col_sol = [d[0] for d in cur.description]
    solicitudes = [dict(zip(col_sol, r)) for r in cur.fetchall()]

    ids = [s["id"] for s in solicitudes]
    if not ids:
        return solicitudes, {}, {}

    cur.execute("""
        SELECT solicitud_id, codigo, apellido_paterno, apellido_materno, nombres,
               sexo, fecha_nacimiento, tipo_documento, num_documento, parentesco,
               nacionalidad, ubigeo, estado_civil
        FROM afiliacion_dependientes
        WHERE solicitud_id = ANY(%s)
        ORDER BY solicitud_id, id
    """, (ids,))

    col_dep = [d[0] for d in cur.description]
    dependientes_por_solicitud = {}
    for r in cur.fetchall():
        dep = dict(zip(col_dep, r))
        dependientes_por_solicitud.setdefault(dep["solicitud_id"], []).append(dep)

    cur.execute("""
        SELECT DISTINCT ON (solicitud_id, persona_codigo)
               solicitud_id, persona_codigo, tipo_seguro, modalidad,
               asegurador_anterior, fecha_desde, fecha_hasta
        FROM afiliacion_antecedentes_seguro
        WHERE solicitud_id = ANY(%s)
        ORDER BY solicitud_id, persona_codigo, id DESC
    """, (ids,))
    col_ant = [d[0] for d in cur.description]
    antecedentes_por_key = {}
    for r in cur.fetchall():
        ant = dict(zip(col_ant, r))
        antecedentes_por_key[(ant["solicitud_id"], ant["persona_codigo"])] = ant

    return solicitudes, dependientes_por_solicitud, antecedentes_por_key


def _fila_antecedente_reporte(ant):
    if not ant:
        return "", "", "", "", ""
    return (
        ant.get("tipo_seguro") or "",
        ant.get("modalidad") or "",
        ant.get("asegurador_anterior") or "",
        ant.get("fecha_desde"),
        ant.get("fecha_hasta"),
    )


def _construir_filas_reporte_afiliacion(solicitudes, dependientes_por_solicitud, antecedentes_por_key):
    filas = []

    for sol in solicitudes:
        vigencia = _calcular_vigencia_reporte(sol["fecha_completada"])
        tipo_via, direccion, numero = _parsear_direccion_reporte(sol["titular_direccion"])
        primer_nombre, resto_nombres = _separar_nombres_reporte(sol["titular_nombres"])
        referente = _contacto_referente_reporte(sol["celular_origen_wsp"])

        ant_titular = antecedentes_por_key.get((sol["id"], "T"))
        tipo, ant_asegurador_salud, ant_eps, f_inicio, f_fin = _fila_antecedente_reporte(ant_titular)

        # ── Fila del titular ──
        filas.append([
            sol["plan"] or "",
            vigencia,
            sol["titular_tipo_documento"] or "",
            sol["titular_num_documento"] or "",
            sol["titular_tipo_documento"] or "",
            sol["titular_num_documento"] or "",
            primer_nombre,
            resto_nombres,
            sol["titular_apellido_paterno"] or "",
            sol["titular_apellido_materno"] or "",
            sol["titular_sexo"] or "",
            sol["titular_fecha_nacimiento"],
            sol["titular_nacionalidad"] or "",
            sol["titular_estado_civil"] or "",
            "TITULAR",
            tipo_via,
            direccion,
            numero,
            sol["titular_ubigeo"] or "",
            sol["titular_telefono"] or "",
            sol["celular_origen_wsp"] or "",  
            sol["titular_email"] or "",
            tipo,
            ant_asegurador_salud,
            ant_eps,
            f_inicio,
            f_fin,
            referente,
        ])

        # ── Filas de dependientes ──
        for dep in dependientes_por_solicitud.get(sol["id"], []):
            primer_nombre_dep, resto_nombres_dep = _separar_nombres_reporte(dep["nombres"])
            ant_dep = antecedentes_por_key.get((sol["id"], dep["codigo"]))
            tipo_d, ant_asegurador_salud_d, ant_eps_d, f_inicio_d, f_fin_d = _fila_antecedente_reporte(ant_dep)

            filas.append([
                sol["plan"] or "",
                vigencia,
                sol["titular_tipo_documento"] or "",
                sol["titular_num_documento"] or "",
                dep["tipo_documento"] or "",
                dep["num_documento"] or "",
                primer_nombre_dep,
                resto_nombres_dep,
                dep["apellido_paterno"] or "",
                dep["apellido_materno"] or "",
                dep["sexo"] or "",
                dep["fecha_nacimiento"],
                dep["nacionalidad"] or "",
                dep["estado_civil"] or "",
                dep["parentesco"] or "",
                tipo_via,   # misma dirección del titular
                direccion,  # misma dirección del titular
                numero,     # misma dirección del titular
                dep["ubigeo"] or sol["titular_ubigeo"] or "",
                sol["titular_telefono"] or "",  # mismo contacto del titular
                sol["celular_origen_wsp"] or "",
                sol["titular_email"] or "",
                tipo_d,
                ant_asegurador_salud_d,
                ant_eps_d,
                f_inicio_d,
                f_fin_d,
                referente,
            ])

    return filas


def _generar_excel_reporte_afiliacion(filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Afiliacion"

    header_fill = PatternFill(start_color="00A9B7", end_color="00A9B7", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, header in enumerate(REPORTE_EXCEL_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    date_cols = {2, 12, 26, 27}  # VIGENCIA, FECHA DE NACIMIENTO, FECHA INICIO EPS, FECHA FIN EPS

    for row_idx, fila in enumerate(filas, start=2):
        for col_idx, valor in enumerate(fila, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            if col_idx in date_cols and valor:
                cell.number_format = "DD/MM/YYYY"

    for col_idx, header in enumerate(REPORTE_EXCEL_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(18, len(header) * 0.9)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(REPORTE_EXCEL_HEADERS))}1"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ─────────────────────────────────────────────
# API: /afiliacion/reporte-excel (GET) — descarga el .xlsx de 28 campos
# Protegido con ?pw= igual que /afiliacion/reporte
# ─────────────────────────────────────────────
@con_log_errores
def handle_reporte_excel_afiliacion(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_AFILIACION:
        return ("No autorizado", 401, {"Content-Type": "text/html; charset=utf-8"})

    fecha_inicio = request.args.get("fecha_inicio", "").strip()
    fecha_fin = request.args.get("fecha_fin", "").strip()
    busqueda = request.args.get("q", "").strip()

    conn = get_connection()
    cur = conn.cursor()
    try:
        solicitudes, dependientes_por_solicitud, antecedentes_por_key = _obtener_datos_reporte_afiliacion(
            cur,
            fecha_inicio=fecha_inicio or None,
            fecha_fin=fecha_fin or None,
            busqueda=busqueda or None,
        )
    finally:
        cur.close()
        conn.close()

    filas = _construir_filas_reporte_afiliacion(solicitudes, dependientes_por_solicitud, antecedentes_por_key)
    excel_bytes = _generar_excel_reporte_afiliacion(filas)

    nombre_archivo = f"reporte_afiliacion_{datetime.now(LIMA_TZ).strftime('%Y%m%d_%H%M')}.xlsx"

    return (
        excel_bytes,
        200,
        {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": f"attachment; filename={nombre_archivo}",
        },
    )

# ─────────────────────────────────────────────
# SYNC: Excel Rechazados -> tabla local (cada 15 min)
# ─────────────────────────────────────────────
def sincronizar_rechazados_excel():
    #url = (
    #    "https://docs.google.com/spreadsheets/d/"
    #    "1dJdYBQqzjilP7SXLsb6OSgi_wOqn8DrxfiQ6-OtNcM8"
    #    "/gviz/tq?tqx=out:csv"
    #    "&gid=577147525"
    #    f"&t={int(time.time())}"
        
    #)
    url = ("https://docs.google.com/spreadsheets/d/e/2PACX-1vTzkoWIfO26i83Vfzejemnxjq40HRpkJH9NxrO3eGwqcNAk6iWcNoC0r712wKDwIxGL16EpyEC68V1Z/pub?gid=577147525&single=true&output=csv&t={int(time.time())}")

    lines = safe_read_csv(url)
    if not lines:
        print("[sync_rechazados] No se pudo leer el archivo, se reintentará en el próximo ciclo.")
        return "Error lectura CSV"

    last2000 = lines[max(0, len(lines) - 2000):]
    reader   = csv.reader(last2000)

    valores = []
    for row in reader:
        if not row:
            continue

        row_2  = row[2].strip()  if len(row) > 2  else ""
        row_4  = row[4].strip()  if len(row) > 4  else ""
        row_21 = row[21].strip() if len(row) > 21 else ""
        row_23 = row[23].strip() if len(row) > 23 else ""

        if not row_23:
            continue

        try:
            fecha_parsed = datetime.strptime(row_23, "%d/%m/%Y").date()
        except ValueError:
            continue

        valores.append((row_2, row_4, row_21, row_23, fecha_parsed))

    if not valores:
        return "Sin filas válidas para sincronizar"

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("TRUNCATE TABLE sofi_rechazados_raw")
    execute_values(cur, """
        INSERT INTO sofi_rechazados_raw
            (sol, nombre, celular, fecha_rechazo_raw, fecha_rechazo)
        VALUES %s
    """, valores, page_size=500)
    conn.commit()
    cur.close()
    conn.close()

    return f"Sincronizados: {len(valores)} registros"


def handle_sync_rechazados_excel(request):
    try:
        resultado = sincronizar_rechazados_excel()
        return (json.dumps({"mensaje": resultado}), 200, HEADERS)
    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)


# ───────────────────────────────────────────────────
# Reporte de la venta por whatsApp
# ───────────────────────────────────────────────────
ADMIN_PW_SCTR_WSP = os.environ.get("ADMIN_PW_SCTR_WSP")

BROKERS_PERMITIDOS_WSP_SCTR = ["CANAL DIGITAL VENTA DIRECTA", "VENTA DIRECTA"]

@con_log_errores
def handle_reporte_wsp_sctr_vencidos(request):
    pw = request.args.get("pw")
    if pw != ADMIN_PW_SCTR_WSP:
        return ("No autorizado", 401, {"Content-Type": "text/html; charset=utf-8"})

    hoy = date.today()
    primer_dia_mes = hoy.replace(day=1).isoformat()

    fecha_inicio = request.args.get("fecha_inicio") or primer_dia_mes
    fecha_fin = request.args.get("fecha_fin") or hoy.isoformat()
    solo_ventas = request.args.get("solo_ventas", "") == "si"

    conn = get_connection()
    with conn.cursor() as cur:
        filtro_venta_sql = "AND sub.monto_abonado > 0" if solo_ventas else ""
        cur.execute(f"""
            SELECT * FROM (
                SELECT
                    a.ruc, a.celular, a.fecha AS fecha_envio,
                    e.razon_social,
                    EXISTS (
                        SELECT 1 FROM sofi_bitacora b
                        WHERE RIGHT(b.user_identifier, 9) = RIGHT(a.celular, 9)
                          AND b.fecha > a.fecha
                          AND b.fecha <= a.fecha + INTERVAL '5 days'
                    ) AS leido,
                    em.broker,
                    em.monto_abonado
                FROM sofi_wsp_avisos a
                LEFT JOIN sctr_empresa e ON e.ruc = a.ruc
                LEFT JOIN LATERAL (
                    SELECT broker, monto_abonado
                    FROM sctr_emisiones
                    WHERE nro_doc = a.ruc
                      AND inicio_vigencia >= a.fecha::date
                      AND broker = ANY(%s)
                    ORDER BY inicio_vigencia DESC
                    LIMIT 1
                ) em ON TRUE
                WHERE a.fecha::date BETWEEN %s AND %s
            ) sub
            WHERE 1=1 {filtro_venta_sql}
            ORDER BY fecha_envio DESC
        """, (BROKERS_PERMITIDOS_WSP_SCTR, fecha_inicio, fecha_fin))
        filas = cur.fetchall()
    conn.close()

    total_enviados = len(filas)
    total_leidos = sum(1 for f in filas if f[4])

    # Deduplicar ventas por RUC (un RUC puede tener 2+ celulares y por tanto
    # 2+ filas de envío, pero la venta y el importe son un único registro)
    ventas_por_ruc = {}
    for ruc, celular, fecha_envio, razon_social, leido, broker, monto_abonado in filas:
        if monto_abonado is not None and ruc not in ventas_por_ruc:
            ventas_por_ruc[ruc] = float(monto_abonado)

    total_ventas = len(ventas_por_ruc)
    suma_montos = sum(ventas_por_ruc.values())

    rucs_unicos_enviados = len({f[0] for f in filas})
    tasa_lectura = round(100.0 * total_leidos / total_enviados, 1) if total_enviados else 0
    tasa_conversion = round(100.0 * total_ventas / rucs_unicos_enviados, 1) if rucs_unicos_enviados else 0

    filas_html = ""
    for ruc, celular, fecha_envio, razon_social, leido, broker, monto_abonado in filas[:500]:
        badge_leido = '<span class="badge badge-ok">✅ Leído</span>' if leido else '<span class="badge badge-neutro">— Sin confirmar</span>'
        if monto_abonado is not None:
            badge_venta = f'<span class="badge badge-ok">💰 S/ {monto_abonado:,.2f}</span>'
        else:
            badge_venta = '<span class="badge badge-neutro">—</span>'
        filas_html += f"""<tr>
            <td>{fecha_envio.strftime('%d/%m/%Y %H:%M') if fecha_envio else ''}</td>
            <td>{ruc}</td>
            <td>{razon_social or '—'}</td>
            <td>{celular}</td>
            <td>{badge_leido}</td>
            <td>{broker or '—'}</td>
            <td>{badge_venta}</td>
        </tr>"""

    aviso_limite = f'<p class="msg-aviso">Mostrando primeros 500 de {total_enviados} resultados. Afina el rango de fechas para ver menos.</p>' if total_enviados > 500 else ""

    opcion_todos = 'selected' if not solo_ventas else ''
    opcion_solo_ventas = 'selected' if solo_ventas else ''

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>Venta de SCTR por whatsApp</title>
<link rel="icon" type="image/x-icon" href="/assets/favicon_sanitas.ico">
<style>{CAMPANIA_CSS}</style>
</head>
<body>
<div class="topbar">
    <h1>Venta de SCTR por whatsApp</h1>
    <div class="subtitulo">Envíos, lecturas (accesos en los 5 días posteriores) y ventas de Venta Directa / Canal Digital</div>
</div>
<div class="contenedor">
    <div class="card">
        <form method="GET" action="/sctr/reporte-wsp" class="filtros">
            <input type="hidden" name="pw" value="{pw}">
            <div><label>Desde</label><input type="date" name="fecha_inicio" value="{fecha_inicio}"></div>
            <div><label>Hasta</label><input type="date" name="fecha_fin" value="{fecha_fin}"></div>
            <div>
                <label>Mostrar</label>
                <select name="solo_ventas">
                    <option value="no" {opcion_todos}>Todos los envíos</option>
                    <option value="si" {opcion_solo_ventas}>Solo ventas (importe &gt; 0)</option>
                </select>
            </div>
            <button type="submit">Filtrar</button>
        </form>
    </div>

    <div class="card" style="display:flex;gap:30px;flex-wrap:wrap;">
        <div><div style="font-size:12px;color:var(--sanitas-gray);">Enviados</div><div style="font-size:24px;font-weight:700;">{total_enviados}</div></div>
        <div><div style="font-size:12px;color:var(--sanitas-gray);">Leídos</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-teal-dark);">{total_leidos} ({tasa_lectura}%)</div></div>
        <div><div style="font-size:12px;color:var(--sanitas-gray);">Ventas (RUC único)</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-success);">{total_ventas} ({tasa_conversion}%)</div></div>
        <div><div style="font-size:12px;color:var(--sanitas-gray);">Importe total</div><div style="font-size:24px;font-weight:700;color:var(--sanitas-success);">S/ {suma_montos:,.2f}</div></div>
    </div>

    {aviso_limite}
    <div class="table-wrap">
        <table>
        <tr><th>Fecha envío</th><th>RUC</th><th>Empresa</th><th>Celular</th><th>Lectura</th><th>Broker</th><th>Importe</th></tr>
        {filas_html}
        </table>
    </div>
</div>
</body>
</html>"""

    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})


# Ajusta a tus imports reales de conexión/logging del proyecto
# from tu_modulo_db import get_connection
# from tu_modulo_logging import con_log_errores

# ───────────────────────────────────────────────────
# Config
# ───────────────────────────────────────────────────
CHATTIGO_WEBHOOK_BASE_URL = "https://channels.chattigo.com/webhook/chattigo"

CHATTIGO_WEBHOOK_USERNAME = os.environ.get("CHATTIGO_WEBHOOK_USERNAME")
CHATTIGO_WEBHOOK_PASSWORD = os.environ.get("CHATTIGO_WEBHOOK_PASSWORD")

CHATTIGO_AGENT_DID = os.environ.get("CHATTIGO_DID", "51955973936")
CHATTIGO_DASHBOARD_PASSWORD = os.environ.get("CHATTIGO_DASHBOARD_PASSWORD")

LIMA_TZ = pytz.timezone("America/Lima")

STATUS_VALIDOS = ["ONLINE", "LOGOUT", "ALMUERZO", "REUNIÓN", "DESCANSO"]

# Cache en memoria del token (vive mientras la instancia de Cloud Run esté activa)
_token_cache = {"access_token": None, "expira_en": None}


# ───────────────────────────────────────────────────
# Login y consulta a Chattigo
# ───────────────────────────────────────────────────
def _obtener_token_chattigo_webhook():
    """Login contra el servicio de webhook. Cachea el token en memoria ~8h."""
    ahora = datetime.utcnow()
    if _token_cache["access_token"] and _token_cache["expira_en"] and ahora < _token_cache["expira_en"]:
        return _token_cache["access_token"]

    resp = requests.post(
        f"{CHATTIGO_WEBHOOK_BASE_URL}/login",
        json={"username": CHATTIGO_WEBHOOK_USERNAME, "password": CHATTIGO_WEBHOOK_PASSWORD},
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()

    token = data.get("access_token")
    if not token:
        raise Exception(f"Login Chattigo (webhook/agentes) falló: {data}")

    _token_cache["access_token"] = token
    _token_cache["expira_en"] = ahora + timedelta(hours=8) - timedelta(minutes=10)

    return token


def _obtener_agentes():
    """Llama GET /agents/{did} y devuelve la lista de agentes."""
    token = _obtener_token_chattigo_webhook()

    resp = requests.get(
        f"{CHATTIGO_WEBHOOK_BASE_URL}/agents/{CHATTIGO_AGENT_DID}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=10,
    )

    if resp.status_code == 401:
        _token_cache["access_token"] = None
        token = _obtener_token_chattigo_webhook()
        resp = requests.get(
            f"{CHATTIGO_WEBHOOK_BASE_URL}/agents/{CHATTIGO_AGENT_DID}",
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
        )

    resp.raise_for_status()
    return resp.json()


# ───────────────────────────────────────────────────
# Persistencia del histórico
# ───────────────────────────────────────────────────
@con_log_errores
def _guardar_snapshot_agentes(cur, conn, agentes):
    """Inserta una fila por agente en chattigo_agentes_historico."""
    fecha_hora = datetime.now(LIMA_TZ)

    cur.executemany(
        """
        INSERT INTO chattigo_agentes_historico
            (id_agente, nombre_agente, direccion, max_chats, activo, status, id_client, grupos, fecha_hora)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """,
        [
            (
                a.get("id"),
                a.get("name"),
                a.get("address"),
                a.get("maxChats"),
                a.get("active"),
                a.get("status"),
                a.get("idClient"),
                ", ".join(g.get("name", "") for g in (a.get("groups") or [])) or None,
                fecha_hora,
            )
            for a in agentes
        ],
    )
    conn.commit()


def _password_valida(request):
    """
    Password viaja en el BODY del POST (no en query string), para no dejarlo
    expuesto en logs de acceso de Cloud Run ni en el historial del navegador.
    """
    try:
        body = request.get_json(silent=True) or {}
    except Exception:
        body = {}
    password_recibido = body.get("pass") or request.headers.get("X-Dashboard-Password")
    return password_recibido == CHATTIGO_DASHBOARD_PASSWORD


# ───────────────────────────────────────────────────
# Handlers del router
# ───────────────────────────────────────────────────
@con_log_errores
def handle_chattigo_snapshot(request):
    """Snapshot en vivo de agentes (JSON) + guarda fila en histórico."""
    if not _password_valida(request):
        return (json.dumps({"error": "No autorizado"}), 401, {"Content-Type": "application/json"})

    agentes = _obtener_agentes()

    conn = get_connection()
    cur = conn.cursor()
    _guardar_snapshot_agentes(cur, conn, agentes)
    cur.close()
    conn.close()

    total_agentes = len(agentes)
    por_status = {}
    for s in STATUS_VALIDOS:
        por_status[s] = sum(1 for a in agentes if (a.get("status") or "").upper() == s)

    en_linea = por_status.get("ONLINE", 0)

    body = {
        "resumen": {
            "total_agentes": total_agentes,
            "agentes_en_linea": en_linea,
            "por_status": por_status,
        },
        "agentes": agentes,
    }
    return (json.dumps(body), 200, {"Content-Type": "application/json"})


@con_log_errores
def handle_chattigo_historico(request):
    """Serie histórica para graficar tendencias. Body JSON: {"pass": "...", "horas": 24}"""
    if not _password_valida(request):
        return (json.dumps({"error": "No autorizado"}), 401, {"Content-Type": "application/json"})

    body = request.get_json(silent=True) or {}
    horas = int(body.get("horas", 24))

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            date_trunc('minute', fecha_hora) AS minuto,
            COUNT(DISTINCT id_agente) FILTER (WHERE status = 'ONLINE') AS en_linea,
            COUNT(DISTINCT id_agente) AS total_reportado
        FROM chattigo_agentes_historico
        WHERE fecha_hora >= %s
        GROUP BY 1
        ORDER BY 1
        """,
        (datetime.now(LIMA_TZ) - timedelta(hours=horas),),
    )
    filas = cur.fetchall()
    cur.close()
    conn.close()

    serie = [
        {
            "fecha_hora": f[0].isoformat(),
            "agentes_en_linea": int(f[1] or 0),
            "total_reportado": int(f[2] or 0),
        }
        for f in filas
    ]
    return (json.dumps(serie), 200, {"Content-Type": "application/json"})


@con_log_errores
def handle_chattigo_dashboard(request):
    """Sirve el HTML del dashboard."""
    return (_DASHBOARD_HTML, 200, {"Content-Type": "text/html; charset=utf-8"})


_DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Chattigo - Agentes en línea</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  body { font-family: system-ui, sans-serif; background:#f4f5f7; margin:0; padding:24px; color:#1f2430; }
  .login-box { max-width:340px; margin:80px auto; background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,.08); }
  .login-box input { width:100%; padding:10px; margin:10px 0; border:1px solid #ddd; border-radius:6px; box-sizing:border-box; }
  .login-box button { width:100%; padding:10px; background:#2563eb; color:#fff; border:none; border-radius:6px; cursor:pointer; }
  #dashboard { display:none; max-width:1000px; margin:0 auto; }
  h1 { font-size:20px; }
  .cards { display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px; }
  .card { background:#fff; border-radius:12px; padding:16px 20px; box-shadow:0 2px 10px rgba(0,0,0,.06); flex:1; min-width:150px; }
  .card .valor { font-size:28px; font-weight:700; }
  .card .label { font-size:12px; color:#6b7280; text-transform:uppercase; }
  table { width:100%; border-collapse:collapse; background:#fff; border-radius:12px; overflow:hidden; box-shadow:0 2px 10px rgba(0,0,0,.06);}
  th, td { padding:10px 12px; text-align:left; border-bottom:1px solid #eee; font-size:14px; }
  th { background:#fafafa; font-size:12px; text-transform:uppercase; color:#6b7280; }
  .badge { padding:3px 10px; border-radius:999px; font-size:12px; font-weight:600; display:inline-block; }
  .badge-ONLINE { background:#dcfce7; color:#166534; }
  .badge-LOGOUT { background:#f3f4f6; color:#6b7280; }
  .badge-ALMUERZO, .badge-REUNIÓN, .badge-DESCANSO { background:#fef3c7; color:#92400e; }
  canvas { background:#fff; border-radius:12px; padding:12px; box-shadow:0 2px 10px rgba(0,0,0,.06); margin-bottom:24px; }
  #actualizado { font-size:12px; color:#9ca3af; margin-bottom:16px; }
</style>
</head>
<body>

<div class="login-box" id="loginBox">
  <h1>Acceso Dashboard Chattigo</h1>
  <input type="password" id="passInput" placeholder="Password" />
  <button onclick="iniciar()">Entrar</button>
  <div id="loginError" style="color:#dc2626; font-size:13px; margin-top:8px;"></div>
</div>

<div id="dashboard">
  <h1>Agentes en línea — Chattigo</h1>
  <div id="actualizado"></div>
  <div class="cards">
    <div class="card"><div class="valor" id="c_online">-</div><div class="label">En línea</div></div>
    <div class="card"><div class="valor" id="c_total">-</div><div class="label">Total agentes</div></div>
    <div class="card"><div class="valor" id="c_almuerzo">-</div><div class="label">Almuerzo</div></div>
    <div class="card"><div class="valor" id="c_reunion">-</div><div class="label">Reunión</div></div>
    <div class="card"><div class="valor" id="c_descanso">-</div><div class="label">Descanso</div></div>
  </div>
  <canvas id="chartHistorico" height="90"></canvas>
  <table>
    <thead><tr><th>Estado</th><th>Agente</th><th>Ubicación</th><th>Máx. chats</th><th>Grupos</th></tr></thead>
    <tbody id="tablaAgentes"></tbody>
  </table>
</div>

<script>
const URL_SNAPSHOT = "/chattigo/agentes/snapshot";
const URL_HISTORICO = "/chattigo/agentes/historico";

let PASS = "";
let chart = null;

function iniciar() {
  PASS = document.getElementById("passInput").value;
  cargarSnapshot(true);
}

async function cargarSnapshot(esLogin) {
  try {
    const r = await fetch(URL_SNAPSHOT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pass: PASS })
    });
    if (r.status === 401) {
      if (esLogin) document.getElementById("loginError").innerText = "Password incorrecto";
      return;
    }
    const data = await r.json();
    if (esLogin) {
      document.getElementById("loginBox").style.display = "none";
      document.getElementById("dashboard").style.display = "block";
      cargarHistorico();
      setInterval(cargarSnapshot, 30000);
      setInterval(cargarHistorico, 60000);
    }
    pintarSnapshot(data);
  } catch (e) {
    console.error(e);
  }
}

function pintarSnapshot(data) {
  const r = data.resumen;
  document.getElementById("c_online").innerText = r.agentes_en_linea;
  document.getElementById("c_total").innerText = r.total_agentes;
  document.getElementById("c_almuerzo").innerText = r.por_status.ALMUERZO ?? 0;
  document.getElementById("c_reunion").innerText = r.por_status["REUNIÓN"] ?? 0;
  document.getElementById("c_descanso").innerText = r.por_status.DESCANSO ?? 0;
  document.getElementById("actualizado").innerText = "Actualizado: " + new Date().toLocaleTimeString('es-PE');

  const tbody = document.getElementById("tablaAgentes");
  tbody.innerHTML = "";
  data.agentes.forEach(a => {
    const grupos = (a.groups || []).map(g => g.name).join(", ") || "-";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="badge badge-${a.status}">${a.status || '-'}</span></td>
      <td>${a.name || '-'}</td>
      <td>${a.address || '-'}</td>
      <td>${a.maxChats ?? '-'}</td>
      <td>${grupos}</td>
    `;
    tbody.appendChild(tr);
  });

  if (data.agentes.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:#9ca3af;">Sin agentes conectados en este momento</td></tr>';
  }
}

async function cargarHistorico() {
  try {
    const r = await fetch(URL_HISTORICO, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pass: PASS, horas: 24 })
    });
    const serie = await r.json();
    const labels = serie.map(p => new Date(p.fecha_hora).toLocaleTimeString('es-PE', {hour:'2-digit', minute:'2-digit'}));
    const online = serie.map(p => p.agentes_en_linea);

    if (chart) chart.destroy();
    const ctx = document.getElementById('chartHistorico').getContext('2d');
    chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label: 'Agentes en línea', data: online, borderColor: '#22c55e', tension: 0.3 },
        ]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  } catch (e) {
    console.error(e);
  }
}
</script>
</body>
</html>
"""
# ───────────────────────────────────────────────────
# Médico a Domicilio
# ───────────────────────────────────────────────────
@con_log_errores
def rechazar_oferta(oferta_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        UPDATE solicitudes_ofertas
        SET estado = 'rechazada', respondido_en = %s
        WHERE id = %s AND estado = 'pendiente'
        RETURNING solicitud_id, proveedor_id
    """, (datetime.now(LIMA_TZ), oferta_id))

    row = cur.fetchone()
    if not row:
        conn.rollback()
        return {'status': 'error', 'mensaje': 'Oferta ya no disponible'}

    solicitud_id, proveedor_id = row

    cur.execute("""
        UPDATE proveedores_estado SET estado = 'disponible'
        WHERE proveedor_id = %s AND estado = 'en_oferta'
    """, (proveedor_id,))
    conn.commit()
    cur.close(); conn.close()

    resultado = buscar_y_ofrecer_proveedor(solicitud_id)
    return {'status': 'rechazada', 'reasignacion': resultado}

@con_log_errores
def webhook_whatsapp():
    data = get_json(force=True, silent=True)
    entry = data.get('entry', [{}])[0]
    changes = entry.get('changes', [{}])[0]
    value = changes.get('value', {})
    mensajes = value.get('messages', [])

    if not mensajes:
        return (json.dumps({'status': 'ignored'}), 200, HEADERS)

    msg = mensajes[0]
    telefono = msg.get('from')
    tipo = msg.get('type')

    if tipo == 'location':
        loc = msg['location']
        lat, lng = loc['latitude'], loc['longitude']
        # recuperar contexto de conversación: qué servicio pidió (MAD/ambulancia)
        # esto depende de tu máquina de estados conversacional, similar a sofi_agente_estado
        resultado = crear_solicitud_desde_ubicacion(telefono, lat, lng)
        return (json.dumps(resultado), 200, HEADERS)

    elif tipo == 'interactive':
        interactive = msg['interactive']
        # botones de selección de servicio (MAD / Ambulancia) del cliente
        procesar_seleccion_servicio(telefono, msg['interactive'])
        return (json.dumps({'status': 'ok'}), 200, HEADERS)

    elif tipo == 'text':
        procesar_mensaje_texto(telefono, msg['text']['body'])
        return (json.dumps({'status': 'ok'}), 200, HEADERS)

    return (json.dumps({'status': 'ignored'}), 200, HEADERS)

def crear_solicitud(telefono, lat, lng, tipo_servicio, destino_clinica_id=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO solicitudes (cliente_telefono, tipo_servicio, origen, destino_clinica_id, estado)
        VALUES (%s, %s, ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography, %s, 'buscando_proveedor')
        RETURNING id, tracking_token
    """, (telefono, tipo_servicio, lng, lat, destino_clinica_id))
    solicitud_id, tracking_token = cur.fetchone()
    conn.commit()

    # disparar búsqueda de proveedor (síncrono aquí, o encolar)
    buscar_y_ofrecer_proveedor(solicitud_id)

    link_tracking = f"https://sntspe.com/track/{tracking_token}"
    return {'solicitud_id': solicitud_id, 'tracking_url': link_tracking}

def enviar_oferta_wsp(telefono_proveedor, oferta_token, tipo_servicio):
    numero = _normalizar_celular_pe(telefono_proveedor)
    if not numero:
        return
    nombre_servicio = "Médico a domicilio" if tipo_servicio == "MAD" else "Ambulancia"
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link = f"{base_url}/oferta/{oferta_token}"
    mensaje = f"Sanitas: Nueva solicitud de {nombre_servicio} cerca de ti. Revisa y responde aqui: {link}"
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "oferta mad")

def enviar_whatsapp_oferta(telefono_proveedor, oferta_id, solicitud_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT tipo_servicio, origen_direccion_texto,
               ST_Y(origen::geometry) AS lat, ST_X(origen::geometry) AS lng
        FROM solicitudes WHERE id = %s
    """, (solicitud_id,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return

    tipo_servicio, direccion_texto, lat, lng = row
    nombre_servicio = "Médico a domicilio" if tipo_servicio == "MAD" else "Ambulancia"
    ubicacion_texto = direccion_texto or f"{lat:.5f}, {lng:.5f}"

    numero = _normalizar_celular_whatsapp_cloud(telefono_proveedor)
    if not numero:
        return

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "interactive",
        "interactive": {
            "type": "button",
            "body": {
                "text": f"🚨 Nueva solicitud: {nombre_servicio}\n📍 {ubicacion_texto}\n\n¿Aceptas este servicio?"
            },
            "footer": {"text": "Sanitas Perú"},
            "action": {
                "buttons": [
                    {"type": "reply", "reply": {"id": f"oferta_aceptar_{oferta_id}",  "title": "✅ Aceptar"}},
                    {"type": "reply", "reply": {"id": f"oferta_rechazar_{oferta_id}", "title": "❌ Rechazar"}},
                ]
            }
        }
    }

    ok, resultado = _enviar_whatsapp_cloud(payload)
    if not ok:
        logging.error(f"[MAD] No se pudo notificar oferta {oferta_id} a {telefono_proveedor}: {resultado}")

def buscar_y_ofrecer_proveedor(solicitud_id, radio_metros=5000):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT origen, tipo_servicio FROM solicitudes WHERE id = %s", (solicitud_id,))
    origen, tipo_servicio = cur.fetchone()

    cur.execute("""
        SELECT p.id, p.telefono,
               ST_Distance(pe.ubicacion, %s::geography) AS distancia
        FROM proveedores p
        JOIN proveedores_estado pe ON pe.proveedor_id = p.id
        WHERE pe.estado = 'disponible'
          AND p.tipo_servicio IN (%s, 'AMBOS')
          AND ST_DWithin(pe.ubicacion, %s::geography, %s)
        ORDER BY distancia ASC
        LIMIT 1
        FOR UPDATE OF pe SKIP LOCKED
    """, (origen, tipo_servicio, origen, radio_metros))

    row = cur.fetchone()
    if not row:
        # nadie disponible: expandir radio o marcar sin_cobertura
        return {'status': 'sin_proveedores', 'radio_usado': radio_metros}

    proveedor_id, telefono_proveedor, distancia = row

    cur.execute("""
    INSERT INTO solicitudes_ofertas (solicitud_id, proveedor_id, estado)
    VALUES (%s, %s, 'pendiente') RETURNING id, oferta_token
    """, (solicitud_id, proveedor_id))
    oferta_id, oferta_token = cur.fetchone()

    cur.execute("UPDATE proveedores_estado SET estado = 'en_oferta' WHERE proveedor_id = %s", (proveedor_id,))
    conn.commit()

    enviar_oferta_wsp(telefono_proveedor, oferta_token, tipo_servicio)
    return {'status': 'ofrecido', 'oferta_id': oferta_id, 'proveedor_id': proveedor_id}

def notificar_cliente_proveedor_asignado(solicitud_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT s.cliente_telefono, s.tipo_servicio, s.tracking_token, p.nombre
        FROM solicitudes s
        JOIN proveedores p ON p.id = s.proveedor_id
        WHERE s.id = %s
    """, (solicitud_id,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return

    telefono_cliente, tipo_servicio, tracking_token, nombre_proveedor = row
    referencia = "tu médico" if tipo_servicio == "MAD" else "tu ambulancia"
    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    link_tracking = f"{base_url}/track/{tracking_token}"

    numero = _normalizar_celular_pe(telefono_cliente)
    if not numero:
        return

    mensaje = f"Sanitas: {nombre_proveedor} viene en camino, sera {referencia}. Sigue el recorrido: {link_tracking}"
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "mad proveedor asignado")

def aceptar_oferta(oferta_id):
    conn = get_connection()
    cur = conn.cursor()

    # UPDATE atómico condicionado: solo pasa si sigue 'pendiente'
    cur.execute("""
        UPDATE solicitudes_ofertas
        SET estado = 'aceptada', respondido_en = %s
        WHERE id = %s AND estado = 'pendiente'
        RETURNING solicitud_id, proveedor_id
    """, (datetime.now(LIMA_TZ), oferta_id))

    row = cur.fetchone()
    if not row:
        conn.rollback()
        return {'status': 'error', 'mensaje': 'Oferta ya no disponible'}

    solicitud_id, proveedor_id = row

    cur.execute("""
        UPDATE solicitudes SET estado = 'asignado', proveedor_id = %s, asignado_en = %s
        WHERE id = %s
    """, (proveedor_id, datetime.now(LIMA_TZ), solicitud_id))

    cur.execute("UPDATE proveedores_estado SET estado = 'ocupado' WHERE proveedor_id = %s", (proveedor_id,))

    cur.execute("""
        INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
        VALUES (%s, 'buscando_proveedor', 'asignado', 'proveedor')
    """, (solicitud_id,))

    conn.commit()

    notificar_cliente_proveedor_asignado(solicitud_id)
    return {'status': 'ok', 'solicitud_id': solicitud_id}

@con_log_errores
def obtener_estado_tracking(tracking_token):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT s.estado, s.tipo_servicio,
               ST_X(pe.ubicacion::geometry) AS lng, ST_Y(pe.ubicacion::geometry) AS lat,
               ST_X(s.origen::geometry) AS origen_lng, ST_Y(s.origen::geometry) AS origen_lat,
               p.nombre AS proveedor_nombre
        FROM solicitudes s
        LEFT JOIN proveedores p ON p.id = s.proveedor_id
        LEFT JOIN proveedores_estado pe ON pe.proveedor_id = s.proveedor_id
        WHERE s.tracking_token = %s
    """, (tracking_token,))
    row = cur.fetchone()
    if not row:
        return (json.dumps({'error': 'no encontrado'}), 404, HEADERS)

    estado, tipo, plng, plat, olng, olat, nombre = row
    return (json.dumps({
        'estado': estado, 'tipo_servicio': tipo,
        'proveedor_ubicacion': {'lat': plat, 'lng': plng} if plat else None,
        'origen': {'lat': olat, 'lng': olng},
        'proveedor_nombre': nombre
    }), 200, HEADERS)

def handle_servir_tracking(request, tracking_token):
    with open("static/track.html", "r", encoding="utf-8") as f:
        html = f.read()
    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})


@con_log_errores
def handle_crear_solicitud(request):
    data = request.get_json(force=True, silent=True) or {}
    telefono           = data.get("telefono", "")
    lat                = data.get("lat")
    lng                = data.get("lng")
    tipo_servicio      = data.get("tipo_servicio", "")
    destino_clinica_id = data.get("destino_clinica_id")

    if not telefono or lat is None or lng is None or tipo_servicio not in ("MAD", "AMBULANCIA"):
        return (json.dumps({
            "error": "telefono, lat, lng y tipo_servicio (MAD o AMBULANCIA) son obligatorios"
        }), 400, HEADERS)

    resultado = crear_solicitud(telefono, lat, lng, tipo_servicio, destino_clinica_id)
    return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_buscar_proveedor(request, solicitud_id):
    try:
        solicitud_id = int(solicitud_id)
    except ValueError:
        return (json.dumps({"error": "solicitud_id inválido"}), 400, HEADERS)

    resultado = buscar_y_ofrecer_proveedor(solicitud_id)
    return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_aceptar_oferta(request, oferta_id):
    try:
        oferta_id = int(oferta_id)
    except ValueError:
        return (json.dumps({"error": "oferta_id inválido"}), 400, HEADERS)

    resultado = aceptar_oferta(oferta_id)
    status = 200 if resultado.get("status") == "ok" else 409
    return (json.dumps(resultado, ensure_ascii=False), status, HEADERS)


@con_log_errores
def handle_actualizar_ubicacion_proveedor(request, proveedor_id):
    try:
        proveedor_id = int(proveedor_id)
    except ValueError:
        return (json.dumps({"error": "proveedor_id inválido"}), 400, HEADERS)

    data = request.get_json(force=True, silent=True) or {}
    lat, lng = data.get("lat"), data.get("lng")

    if lat is None or lng is None:
        return (json.dumps({"error": "lat y lng son obligatorios"}), 400, HEADERS)

    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("""
        UPDATE proveedores_estado
        SET ubicacion = ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
            actualizado_en = %s
        WHERE proveedor_id = %s
    """, (lng, lat, datetime.now(LIMA_TZ), proveedor_id))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({"status": "ok"}), 200, HEADERS)


def notificar_cliente_sin_cobertura(solicitud_id):
    conn = get_connection()
    cur  = conn.cursor()
    cur.execute("SELECT cliente_telefono, tipo_servicio FROM solicitudes WHERE id = %s", (solicitud_id,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return
    telefono, tipo_servicio = row
    # TODO: reemplazar por tu envío vía Meta Cloud API (WhatsApp Business)
    # mensaje = "No encontramos un proveedor disponible cerca. Un asesor te contactará."
    pass

@con_log_errores
def handle_revisar_timeouts_ofertas(request):
    conn = get_connection()
    cur  = conn.cursor()

    limite = datetime.now(LIMA_TZ) - timedelta(seconds=TIMEOUT_OFERTA_SEGUNDOS)

    # Ofertas pendientes que ya vencieron, con su solicitud aún activa
    cur.execute("""
        SELECT o.id, o.solicitud_id, o.proveedor_id
        FROM solicitudes_ofertas o
        JOIN solicitudes s ON s.id = o.solicitud_id
        WHERE o.estado = 'pendiente'
          AND o.ofrecido_en < %s
          AND s.estado = 'buscando_proveedor'
        FOR UPDATE OF o SKIP LOCKED
    """, (limite,))
    ofertas_vencidas = cur.fetchall()

    if not ofertas_vencidas:
        cur.close(); conn.close()
        return (json.dumps({"mensaje": "Sin ofertas vencidas"}), 200, HEADERS)

    reasignadas = []
    sin_cobertura = []

    for oferta_id, solicitud_id, proveedor_id in ofertas_vencidas:
        # 1. Expira la oferta
        cur.execute("""
            UPDATE solicitudes_ofertas
            SET estado = 'expirada', respondido_en = %s
            WHERE id = %s
        """, (datetime.now(LIMA_TZ), oferta_id))

        # 2. Libera al proveedor (vuelve a estar disponible para otras solicitudes)
        cur.execute("""
            UPDATE proveedores_estado
            SET estado = 'disponible'
            WHERE proveedor_id = %s AND estado = 'en_oferta'
        """, (proveedor_id,))

        # 3. Incrementa intentos y expande el radio de esta solicitud
        cur.execute("""
            UPDATE solicitudes
            SET intentos_asignacion = intentos_asignacion + 1,
                radio_busqueda_metros = LEAST(radio_busqueda_metros + %s, %s)
            WHERE id = %s
            RETURNING intentos_asignacion, radio_busqueda_metros
        """, (INCREMENTO_RADIO_METROS, RADIO_MAXIMO_METROS, solicitud_id))
        intentos, radio_actual = cur.fetchone()

        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, 'buscando_proveedor', 'buscando_proveedor', 'sistema')
        """, (solicitud_id,))

        conn.commit()

        # 4. Si se agotaron los intentos, se marca sin cobertura y se notifica
        if intentos >= MAX_INTENTOS_ASIGNACION:
            cur.execute("""
                UPDATE solicitudes SET estado = 'sin_cobertura' WHERE id = %s
            """, (solicitud_id,))
            conn.commit()
            notificar_cliente_sin_cobertura(solicitud_id)
            sin_cobertura.append(solicitud_id)
        else:
            resultado = buscar_y_ofrecer_proveedor(solicitud_id, radio_metros=radio_actual)
            reasignadas.append({"solicitud_id": solicitud_id, "resultado": resultado})

    cur.close(); conn.close()

    return (json.dumps({
        "mensaje": f"{len(ofertas_vencidas)} ofertas vencidas procesadas",
        "reasignadas": len(reasignadas),
        "sin_cobertura": sin_cobertura,
    }, ensure_ascii=False), 200, HEADERS)

@con_log_errores
def handle_toggle_disponibilidad(request, proveedor_id):
    data = request.get_json(force=True, silent=True) or {}
    estado_nuevo = data.get("estado")

    if estado_nuevo not in ("disponible", "desconectado"):
        return (json.dumps({"error": "estado debe ser 'disponible' o 'desconectado'"}), 400, HEADERS)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        UPDATE proveedores_estado SET estado = %s, actualizado_en = %s
        WHERE proveedor_id = %s AND estado NOT IN ('ocupado', 'en_oferta')
        RETURNING estado
    """, (estado_nuevo, datetime.now(LIMA_TZ), proveedor_id))
    row = cur.fetchone()
    conn.commit()
    cur.close(); conn.close()

    if not row:
        return (json.dumps({
            "error": "No puedes cambiar tu disponibilidad mientras tienes un servicio activo"
        }), 409, HEADERS)

    return (json.dumps({"status": "ok", "estado": row[0]}), 200, HEADERS)


@con_log_errores
def handle_estado_actual_proveedor(request, proveedor_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT p.nombre, pe.estado
        FROM proveedores p
        JOIN proveedores_estado pe ON pe.proveedor_id = p.id
        WHERE p.id = %s
    """, (proveedor_id,))
    row_proveedor = cur.fetchone()
    if not row_proveedor:
        cur.close(); conn.close()
        return (json.dumps({"error": "Proveedor no encontrado"}), 404, HEADERS)
    nombre, disponibilidad = row_proveedor

    # ── Oferta pendiente ──
    cur.execute("""
        SELECT o.id, o.oferta_token, s.tipo_servicio, s.origen_direccion_texto,
            ST_Distance(pe.ubicacion, s.origen) AS distancia_metros
        FROM solicitudes_ofertas o
        JOIN solicitudes s ON s.id = o.solicitud_id
        JOIN proveedores_estado pe ON pe.proveedor_id = o.proveedor_id
        WHERE o.proveedor_id = %s AND o.estado = 'pendiente'
        ORDER BY o.ofrecido_en DESC LIMIT 1
    """, (proveedor_id,))
    row_oferta = cur.fetchone()
    oferta_pendiente = None
    if row_oferta:
        oferta_id, oferta_token, tipo_servicio, direccion, distancia = row_oferta
        oferta_pendiente = {
            "oferta_id": oferta_id,
            "oferta_token": str(oferta_token),
            "tipo_servicio": tipo_servicio,
            "direccion": direccion,
            "distancia_metros": float(distancia) if distancia is not None else None,
        }

    # ── Servicio activo ──
    cur.execute("""
        SELECT id, tipo_servicio, estado, cliente_telefono, cliente_nombre,
               origen_direccion_texto,
               ST_Y(origen::geometry) AS lat, ST_X(origen::geometry) AS lng
        FROM solicitudes
        WHERE proveedor_id = %s AND estado IN ('asignado', 'en_camino', 'llego', 'en_atencion')
        ORDER BY asignado_en DESC LIMIT 1
    """, (proveedor_id,))
    row_servicio = cur.fetchone()
    servicio_activo = None
    if row_servicio:
        sol_id, tipo_servicio, estado, tel, nombre_cliente, direccion, lat, lng = row_servicio
        servicio_activo = {
            "solicitud_id": sol_id,
            "tipo_servicio": tipo_servicio,
            "estado": estado,
            "cliente_telefono": tel,
            "cliente_nombre": nombre_cliente,
            "direccion": direccion,
            "origen_lat": lat,
            "origen_lng": lng,
        }

    cur.close(); conn.close()

    return (json.dumps({
        "nombre": nombre,
        "disponibilidad": disponibilidad,
        "oferta_pendiente": oferta_pendiente,
        "servicio_activo": servicio_activo,
    }, ensure_ascii=False), 200, HEADERS)

def handle_rechazar_oferta(request, oferta_id):
    try:
        oferta_id = int(oferta_id)
    except ValueError:
        return (json.dumps({"error": "oferta_id inválido"}), 400, HEADERS)

    resultado = rechazar_oferta(oferta_id)
    status = 200 if resultado.get("status") == "rechazada" else 409
    return (json.dumps(resultado, ensure_ascii=False), status, HEADERS)


@con_log_errores
def handle_actualizar_estado_solicitud(request, solicitud_id):
    try:
        solicitud_id = int(solicitud_id)
    except ValueError:
        return (json.dumps({"error": "solicitud_id inválido"}), 400, HEADERS)

    data = request.get_json(force=True, silent=True) or {}
    estado_nuevo = data.get("estado")
    proveedor_id = data.get("proveedor_id")

    TRANSICIONES_VALIDAS = {
        "asignado": "en_camino",
        "en_camino": "llego",
        "llego": "en_atencion",
        "en_atencion": "finalizado",
    }

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT estado, proveedor_id FROM solicitudes WHERE id = %s", (solicitud_id,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return (json.dumps({"error": "Solicitud no encontrada"}), 404, HEADERS)

    estado_actual, proveedor_asignado = row

    if str(proveedor_asignado) != str(proveedor_id):
        cur.close(); conn.close()
        return (json.dumps({"error": "Este servicio no está asignado a este proveedor"}), 403, HEADERS)

    if TRANSICIONES_VALIDAS.get(estado_actual) != estado_nuevo:
        cur.close(); conn.close()
        return (json.dumps({
            "error": f"Transición inválida: {estado_actual} → {estado_nuevo}"
        }), 409, HEADERS)

    campos_extra = ""
    valores_extra = []
    if estado_nuevo == "finalizado":
        campos_extra = ", finalizado_en = %s"
        valores_extra = [datetime.now(LIMA_TZ)]

    cur.execute(f"""
        UPDATE solicitudes SET estado = %s{campos_extra} WHERE id = %s
    """, (estado_nuevo, *valores_extra, solicitud_id))

    cur.execute("""
        INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
        VALUES (%s, %s, %s, 'proveedor')
    """, (solicitud_id, estado_actual, estado_nuevo))

    if estado_nuevo == "finalizado":
        cur.execute("""
            UPDATE proveedores_estado SET estado = 'disponible' WHERE proveedor_id = %s
        """, (proveedor_asignado,))

    conn.commit()
    cur.close(); conn.close()

    if estado_nuevo == "en_camino":
        notificar_cliente_proveedor_en_camino(solicitud_id)

    return (json.dumps({"status": "ok", "estado": estado_nuevo}), 200, HEADERS)

def notificar_cliente_proveedor_en_camino(solicitud_id):
    # TODO: mensaje breve por WhatsApp Cloud API — opcional, el cliente ya ve
    # el cambio de estado reflejado en /track/<token>
    pass

def handle_servir_proveedor(request):
    with open("static/proveedor.html", "r", encoding="utf-8") as f:
        html = f.read()
    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})

def _obtener_oferta_por_token(cur, token):
    cur.execute("""
        SELECT o.id, o.estado, s.tipo_servicio, s.origen_direccion_texto,
               ST_Y(s.origen::geometry) AS lat, ST_X(s.origen::geometry) AS lng,
               ST_Distance(pe.ubicacion, s.origen) AS distancia_metros
        FROM solicitudes_ofertas o
        JOIN solicitudes s ON s.id = o.solicitud_id
        JOIN proveedores_estado pe ON pe.proveedor_id = o.proveedor_id
        WHERE o.oferta_token = %s
    """, (token,))
    return cur.fetchone()


@con_log_errores
def handle_estado_oferta_token(request, token):
    conn = get_connection(); cur = conn.cursor()
    row = _obtener_oferta_por_token(cur, token)
    cur.close(); conn.close()

    if not row:
        return (json.dumps({"error": "Oferta no encontrada"}), 404, HEADERS)

    oferta_id, estado, tipo_servicio, direccion, lat, lng, distancia = row
    return (json.dumps({
        "estado": estado,
        "tipo_servicio": tipo_servicio,
        "direccion": direccion,
        "lat": lat, "lng": lng,
        "distancia_metros": float(distancia) if distancia is not None else None,
    }, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_aceptar_oferta_token(request, token):
    conn = get_connection(); cur = conn.cursor()
    cur.execute("SELECT id FROM solicitudes_ofertas WHERE oferta_token = %s", (token,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return (json.dumps({"error": "Oferta no encontrada"}), 404, HEADERS)

    resultado = aceptar_oferta(row[0])
    status = 200 if resultado.get("status") == "ok" else 409
    return (json.dumps(resultado, ensure_ascii=False), status, HEADERS)


@con_log_errores
def handle_rechazar_oferta_token(request, token):
    conn = get_connection(); cur = conn.cursor()
    cur.execute("SELECT id FROM solicitudes_ofertas WHERE oferta_token = %s", (token,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row:
        return (json.dumps({"error": "Oferta no encontrada"}), 404, HEADERS)

    resultado = rechazar_oferta(row[0])
    status = 200 if resultado.get("status") == "rechazada" else 409
    return (json.dumps(resultado, ensure_ascii=False), status, HEADERS)


def handle_servir_oferta(request, token):
    with open("static/oferta.html", "r", encoding="utf-8") as f:
        html = f.read()
    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})

# ============================================================
# Utilidades de password (pbkdf2, sin dependencias externas)
# ============================================================
def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000)
    return f"{salt}${dk.hex()}"


def check_password(password: str, password_hash: str) -> bool:
    try:
        salt, hashed = password_hash.split("$")
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 100_000)
        return hmac.compare_digest(dk.hex(), hashed)
    except Exception:
        return False


def generar_token() -> str:
    return secrets.token_urlsafe(32)


# ============================================================
# AUTENTICACION - validacion de token y decoradores
# ============================================================
def obtener_sesion_activa(request):
    """Lee el header 'Authorization: Bearer <token>' y devuelve el asesor
    autenticado (dict) o None si el token no existe/expiro/el asesor esta INACTIVO."""
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    token = auth_header[7:].strip()
    if not token:
        return None

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT a.id, a.nombre_completo, a.usuario, a.rol, a.estado
           FROM bo_sesiones s
           JOIN bo_asesores a ON a.id = s.asesor_id
           WHERE s.token = %s AND s.fecha_logout IS NULL""",
        (token,),
    )
    row = cur.fetchone()
    cur.close(); conn.close()

    if not row or row[4] != "ACTIVO":
        return None
    return {"id": row[0], "nombre": row[1], "usuario": row[2], "rol": row[3]}


def require_auth(handler):
    """Exige un token de sesion valido (cualquier rol). Inyecta request.asesor_actual."""
    @functools.wraps(handler)
    def wrapper(request, *args, **kwargs):
        asesor = obtener_sesion_activa(request)
        if not asesor:
            return json.dumps({"ok": False, "error": "No autenticado"}), 401, HEADERS
        request.asesor_actual = asesor
        return handler(request, *args, **kwargs)
    return wrapper


def require_admin(handler):
    """Exige un token de sesion valido con rol ADMIN. Inyecta request.asesor_actual."""
    @functools.wraps(handler)
    def wrapper(request, *args, **kwargs):
        asesor = obtener_sesion_activa(request)
        if not asesor:
            return json.dumps({"ok": False, "error": "No autenticado"}), 401, HEADERS
        if asesor["rol"] != "ADMIN":
            return json.dumps({"ok": False, "error": "Requiere rol ADMIN"}), 403, HEADERS
        request.asesor_actual = asesor
        return handler(request, *args, **kwargs)
    return wrapper
@con_log_errores
def bo_login(request):
    body = request.get_json(silent=True) or {}
    usuario = (body.get("usuario") or "").strip()
    password = body.get("password") or ""

    if not usuario or not password:
        return json.dumps({"ok": False, "error": "usuario y password requeridos"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT id, nombre_completo, usuario, password_hash, rol, estado
           FROM bo_asesores WHERE usuario = %s""",
        (usuario,),
    )
    row = cur.fetchone()

    if not row or row[5] != "ACTIVO" or not check_password(password, row[3]):
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "credenciales invalidas"}), 401, HEADERS

    asesor_id = row[0]
    token = generar_token()
    # Registrar inicio de sesion (para productividad), guardar el token y marcar CONECTADO
    cur.execute(
        """INSERT INTO bo_sesiones (asesor_id, token, fecha_login)
           VALUES (%s, %s, now() AT TIME ZONE 'America/Lima') RETURNING id""",
        (asesor_id, token),
    )
    sesion_id = cur.fetchone()[0]
    cur.execute(
        """UPDATE bo_asesores SET estado_conexion = 'CONECTADO',
           ultima_conexion = now() AT TIME ZONE 'America/Lima' WHERE id = %s""",
        (asesor_id,),
    )
    conn.commit()
    cur.close(); conn.close()

    return json.dumps({
        "ok": True,
        "asesor": {"id": asesor_id, "nombre": row[1], "usuario": row[2], "rol": row[4]},
        "sesion_id": sesion_id,
        "token": token,
    }), 200, HEADERS


@con_log_errores
@require_auth
def bo_logout(request):
    auth_header = request.headers.get("Authorization", "")
    token = auth_header[7:].strip()  # ya paso require_auth, sabemos que empieza con "Bearer "
    asesor_id = request.asesor_actual["id"]

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """UPDATE bo_sesiones SET fecha_logout = now() AT TIME ZONE 'America/Lima',
           duracion_segundos = EXTRACT(EPOCH FROM ((now() AT TIME ZONE 'America/Lima') - fecha_login))::int
           WHERE token = %s""",
        (token,),
    )
    cur.execute(
        "UPDATE bo_asesores SET estado_conexion = 'DESCONECTADO' WHERE id = %s",
        (asesor_id,),
    )
    conn.commit()
    cur.close(); conn.close()
    return json.dumps({"ok": True}), 200, HEADERS


# ============================================================
# RECUPERACION DE CLAVE (autoservicio, codigo por SMS)
# ============================================================
def generar_codigo_otp() -> str:
    return f"{random.randint(0, 999999):06d}"


def enviar_sms_reset_password(celular, codigo):
    """Envio DIRECTO a Danaconnect (mismo patron que envia_sms_lead_pases),
    no pasa por la cola/batch -- necesitamos que el codigo llegue al toque.
    Requiere que DANACONNECT_URL, DANACONNECT_AUTH y requests ya esten
    disponibles en el scope de main.py (ya lo estan, porque los usa
    envia_sms_lead_pases)."""
    mensaje = f"Codigo Sanitas"
    payload = {
        "NUM": f" {celular}",
        "TEXT": mensaje,
        "URL": ""
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    try:
        response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS de reset a {celular}: {e}")
        return False

    # Bitacora (igual que en el resto del sistema) -- esto NO envia nada, solo registra
    sms_aviso(celular, mensaje)
    return True


@con_log_errores
def bo_solicitar_reset(request):
    """Paso 1: el asesor pide un codigo. Se envia por SMS al celular registrado."""
    body = request.get_json(silent=True) or {}
    usuario = (body.get("usuario") or "").strip()
    if not usuario:
        return json.dumps({"ok": False, "error": "usuario requerido"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, celular, estado FROM bo_asesores WHERE usuario = %s", (usuario,))
    row = cur.fetchone()

    # Respuesta generica siempre (exista o no el usuario), para no filtrar informacion
    respuesta_generica = (
        json.dumps({"ok": True, "mensaje": "Si el usuario existe y tiene celular registrado, se envio un codigo por SMS"}),
        200, HEADERS,
    )

    if not row or row[2] != "ACTIVO" or not row[1]:
        cur.close(); conn.close()
        return respuesta_generica

    asesor_id, celular, _ = row
    codigo = generar_codigo_otp()
    cur.execute(
        """INSERT INTO bo_reset_password (asesor_id, codigo, fecha_expira)
           VALUES (%s, %s, (now() AT TIME ZONE 'America/Lima') + interval '15 minutes')""",
        (asesor_id, codigo),
    )
    conn.commit()
    cur.close(); conn.close()

    # sofi_sms_avisos.descripcion es varchar(30) -- mensaje corto obligatorio, ya lo maneja
    # enviar_sms_reset_password() (envio directo a Danaconnect + registro en bitacora)
    enviar_sms_reset_password(celular, codigo)

    return respuesta_generica


@con_log_errores
def bo_confirmar_reset(request):
    """Paso 2: el asesor manda usuario + codigo + nueva clave."""
    body = request.get_json(silent=True) or {}
    usuario = (body.get("usuario") or "").strip()
    codigo = (body.get("codigo") or "").strip()
    nueva_password = body.get("nueva_password") or ""

    if not usuario or not codigo or not nueva_password:
        return json.dumps({"ok": False, "error": "usuario, codigo y nueva_password son requeridos"}), 400, HEADERS
    if len(nueva_password) < 6:
        return json.dumps({"ok": False, "error": "la clave debe tener al menos 6 caracteres"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT id FROM bo_asesores WHERE usuario = %s AND estado = 'ACTIVO'", (usuario,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "codigo invalido o expirado"}), 400, HEADERS
    asesor_id = row[0]

    cur.execute(
        """SELECT id FROM bo_reset_password
           WHERE asesor_id = %s AND codigo = %s AND usado = FALSE
             AND fecha_expira > (now() AT TIME ZONE 'America/Lima')
           ORDER BY fecha_creacion DESC LIMIT 1""",
        (asesor_id, codigo),
    )
    reset_row = cur.fetchone()
    if not reset_row:
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "codigo invalido o expirado"}), 400, HEADERS

    cur.execute("UPDATE bo_asesores SET password_hash = %s WHERE id = %s", (hash_password(nueva_password), asesor_id))
    cur.execute("UPDATE bo_reset_password SET usado = TRUE WHERE id = %s", (reset_row[0],))
    # Por seguridad, cerrar cualquier sesion activa de ese asesor al cambiar la clave
    cur.execute(
        """UPDATE bo_sesiones SET fecha_logout = now() AT TIME ZONE 'America/Lima'
           WHERE asesor_id = %s AND fecha_logout IS NULL""",
        (asesor_id,),
    )
    cur.execute("UPDATE bo_asesores SET estado_conexion = 'DESCONECTADO' WHERE id = %s", (asesor_id,))
    conn.commit()
    cur.close(); conn.close()

    return json.dumps({"ok": True}), 200, HEADERS


# ============================================================
# TAREAS
# ============================================================
@con_log_errores
@require_auth
def bo_crear_tarea(request):
    """Primera tarea: INSERT de una nueva tarea en bo_tareas."""
    body = request.get_json(silent=True) or {}

    cliente_nombre = (body.get("cliente_nombre") or "").strip()
    descripcion = (body.get("descripcion") or "").strip()
    if not cliente_nombre or not descripcion:
        return json.dumps({"ok": False, "error": "cliente_nombre y descripcion son requeridos"}), 400, HEADERS

    cliente_dni = body.get("cliente_dni")
    cliente_celular = body.get("cliente_celular")
    cliente_correo = body.get("cliente_correo")
    asesor_id = body.get("asesor_id")  # opcional, puede crearse sin asignar
    estado_inicial = "ASIGNADO" if asesor_id else "NUEVO"

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO bo_tareas
           (asesor_id, estado, cliente_nombre, cliente_dni, cliente_celular, cliente_correo, descripcion)
           VALUES (%s, %s, %s, %s, %s, %s, %s)
           RETURNING id, fecha_creacion""",
        (asesor_id, estado_inicial, cliente_nombre, cliente_dni, cliente_celular, cliente_correo, descripcion),
    )
    tarea_id, fecha_creacion = cur.fetchone()

    cur.execute(
        """INSERT INTO bo_tarea_eventos (tarea_id, asesor_id, tipo, estado_nuevo)
           VALUES (%s, %s, 'CREACION', %s)""",
        (tarea_id, asesor_id, estado_inicial),
    )
    conn.commit()
    cur.close(); conn.close()

    return json.dumps({"ok": True, "id": tarea_id, "estado": estado_inicial}), 201, HEADERS


@con_log_errores
@require_auth
def bo_listar_tareas(request):
    """Listado para el tablero. Filtros opcionales: estado, asesor_id."""
    estado = request.args.get("estado")
    asesor_id = request.args.get("asesor_id")

    query = """SELECT t.id, t.estado, t.cliente_nombre, t.cliente_dni, t.cliente_celular,
                      t.cliente_correo, t.descripcion, t.fecha_creacion, t.fecha_actualizacion,
                      t.asesor_id, a.nombre_completo
               FROM bo_tareas t
               LEFT JOIN bo_asesores a ON a.id = t.asesor_id
               WHERE 1=1"""
    params = []
    if estado:
        query += " AND t.estado = %s"
        params.append(estado)
    if asesor_id:
        query += " AND t.asesor_id = %s"
        params.append(asesor_id)
    query += " ORDER BY t.fecha_creacion DESC"

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(query, tuple(params))
    rows = cur.fetchall()
    cur.close(); conn.close()

    tareas = [{
        "id": r[0], "estado": r[1], "cliente_nombre": r[2], "cliente_dni": r[3],
        "cliente_celular": r[4], "cliente_correo": r[5], "descripcion": r[6],
        "fecha_creacion": r[7].isoformat() if r[7] else None,
        "fecha_actualizacion": r[8].isoformat() if r[8] else None,
        "asesor_id": r[9], "asesor_nombre": r[10],
    } for r in rows]

    return json.dumps({"ok": True, "tareas": tareas}), 200, HEADERS


@con_log_errores
@require_auth
def bo_detalle_tarea(request, tarea_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT t.id, t.estado, t.cliente_nombre, t.cliente_dni, t.cliente_celular,
                  t.cliente_correo, t.descripcion, t.fecha_creacion, t.fecha_actualizacion,
                  t.asesor_id, a.nombre_completo
           FROM bo_tareas t LEFT JOIN bo_asesores a ON a.id = t.asesor_id
           WHERE t.id = %s""",
        (tarea_id,),
    )
    r = cur.fetchone()
    if not r:
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "tarea no encontrada"}), 404, HEADERS

    cur.execute(
        """SELECT e.id, e.tipo, e.comentario, e.estado_anterior, e.estado_nuevo, e.fecha, a.nombre_completo
           FROM bo_tarea_eventos e LEFT JOIN bo_asesores a ON a.id = e.asesor_id
           WHERE e.tarea_id = %s ORDER BY e.fecha DESC""",
        (tarea_id,),
    )
    eventos = [{
        "id": e[0], "tipo": e[1], "comentario": e[2], "estado_anterior": e[3],
        "estado_nuevo": e[4], "fecha": e[5].isoformat() if e[5] else None, "asesor_nombre": e[6],
    } for e in cur.fetchall()]
    cur.close(); conn.close()

    tarea = {
        "id": r[0], "estado": r[1], "cliente_nombre": r[2], "cliente_dni": r[3],
        "cliente_celular": r[4], "cliente_correo": r[5], "descripcion": r[6],
        "fecha_creacion": r[7].isoformat() if r[7] else None,
        "fecha_actualizacion": r[8].isoformat() if r[8] else None,
        "asesor_id": r[9], "asesor_nombre": r[10], "eventos": eventos,
    }
    return json.dumps({"ok": True, "tarea": tarea}), 200, HEADERS


@con_log_errores
@require_auth
def bo_actualizar_tarea(request, tarea_id):
    """Usado por el drag&drop (cambio de columna/estado) y por reasignacion."""
    body = request.get_json(silent=True) or {}
    nuevo_estado = body.get("estado")
    nuevo_asesor_id = body.get("asesor_id")  # a quien se reasigna (puede ser otro asesor)
    asesor_que_actualiza = request.asesor_actual["id"]  # quien hace el cambio: del token, no del body

    if nuevo_estado and nuevo_estado not in ESTADOS_VALIDOS:
        return json.dumps({"ok": False, "error": f"estado invalido, use uno de {sorted(ESTADOS_VALIDOS)}"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT estado, asesor_id FROM bo_tareas WHERE id = %s", (tarea_id,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "tarea no encontrada"}), 404, HEADERS
    estado_anterior, asesor_anterior = row

    campos, valores = [], []
    if nuevo_estado and nuevo_estado != estado_anterior:
        campos.append("estado = %s"); valores.append(nuevo_estado)
    if nuevo_asesor_id is not None and nuevo_asesor_id != asesor_anterior:
        campos.append("asesor_id = %s"); valores.append(nuevo_asesor_id)

    if campos:
        campos.append("fecha_actualizacion = now() AT TIME ZONE 'America/Lima'")
        valores.append(tarea_id)
        cur.execute(f"UPDATE bo_tareas SET {', '.join(campos)} WHERE id = %s", tuple(valores))

        if nuevo_estado and nuevo_estado != estado_anterior:
            cur.execute(
                """INSERT INTO bo_tarea_eventos (tarea_id, asesor_id, tipo, estado_anterior, estado_nuevo)
                   VALUES (%s, %s, 'CAMBIO_ESTADO', %s, %s)""",
                (tarea_id, asesor_que_actualiza, estado_anterior, nuevo_estado),
            )
        if nuevo_asesor_id is not None and nuevo_asesor_id != asesor_anterior:
            cur.execute(
                """INSERT INTO bo_tarea_eventos (tarea_id, asesor_id, tipo, comentario)
                   VALUES (%s, %s, 'ASIGNACION', %s)""",
                (tarea_id, asesor_que_actualiza, f"Reasignada a asesor_id={nuevo_asesor_id}"),
            )
        conn.commit()

    cur.close(); conn.close()
    return json.dumps({"ok": True}), 200, HEADERS


@con_log_errores
@require_auth
def bo_agregar_comentario(request, tarea_id):
    body = request.get_json(silent=True) or {}
    comentario = (body.get("comentario") or "").strip()
    asesor_id = request.asesor_actual["id"]  # quien comenta: del token, no del body
    if not comentario:
        return json.dumps({"ok": False, "error": "comentario requerido"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO bo_tarea_eventos (tarea_id, asesor_id, tipo, comentario)
           VALUES (%s, %s, 'COMENTARIO', %s)""",
        (tarea_id, asesor_id, comentario),
    )
    cur.execute(
        "UPDATE bo_tareas SET fecha_actualizacion = now() AT TIME ZONE 'America/Lima' WHERE id = %s",
        (tarea_id,),
    )
    conn.commit()
    cur.close(); conn.close()
    return json.dumps({"ok": True}), 201, HEADERS


# ============================================================
# ASESORES (mantenimiento)
# ============================================================
@con_log_errores
def bo_crear_asesor(request):
    """Protegido para rol ADMIN, con una excepcion: si bo_asesores esta vacia
    (primer arranque del sistema), se permite crear el primer admin sin token."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM bo_asesores")
    hay_asesores = cur.fetchone()[0] > 0
    cur.close(); conn.close()

    if hay_asesores:
        asesor_actual = obtener_sesion_activa(request)
        if not asesor_actual:
            return json.dumps({"ok": False, "error": "No autenticado"}), 401, HEADERS
        if asesor_actual["rol"] != "ADMIN":
            return json.dumps({"ok": False, "error": "Requiere rol ADMIN"}), 403, HEADERS

    body = request.get_json(silent=True) or {}
    nombre = (body.get("nombre_completo") or "").strip()
    usuario = (body.get("usuario") or "").strip()
    password = body.get("password") or ""
    correo = body.get("correo")
    celular = body.get("celular")
    rol = body.get("rol", "ASESOR") if hay_asesores else "ADMIN"  # el primer usuario siempre es ADMIN

    if not nombre or not usuario or not password:
        return json.dumps({"ok": False, "error": "nombre_completo, usuario y password son requeridos"}), 400, HEADERS

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO bo_asesores (nombre_completo, usuario, password_hash, correo, celular, rol)
               VALUES (%s, %s, %s, %s, %s, %s) RETURNING id""",
            (nombre, usuario, hash_password(password), correo, celular, rol),
        )
        asesor_id = cur.fetchone()[0]
        conn.commit()
    except Exception:
        conn.rollback()
        cur.close(); conn.close()
        return json.dumps({"ok": False, "error": "usuario ya existe"}), 409, HEADERS
    cur.close(); conn.close()
    return json.dumps({"ok": True, "id": asesor_id}), 201, HEADERS


@con_log_errores
@require_admin
def bo_listar_asesores(request):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT a.id, a.nombre_completo, a.usuario, a.correo, a.celular, a.rol, a.estado,
                  a.estado_conexion, a.ultima_conexion,
                  (SELECT COUNT(*) FROM bo_tareas t WHERE t.asesor_id = a.id AND t.estado = 'CERRADO') AS tareas_resueltas
           FROM bo_asesores a ORDER BY a.nombre_completo ASC"""
    )
    rows = cur.fetchall()
    cur.close(); conn.close()
    asesores = [{
        "id": r[0], "nombre_completo": r[1], "usuario": r[2], "correo": r[3], "celular": r[4],
        "rol": r[5], "estado": r[6], "estado_conexion": r[7],
        "ultima_conexion": r[8].isoformat() if r[8] else None, "tareas_resueltas": r[9],
    } for r in rows]
    return json.dumps({"ok": True, "asesores": asesores}), 200, HEADERS


@con_log_errores
@require_admin
def bo_actualizar_asesor(request, asesor_id):
    """Editar datos, o dar de baja (estado=INACTIVO)."""
    body = request.get_json(silent=True) or {}
    campos, valores = [], []

    for campo in ["nombre_completo", "correo", "celular", "rol"]:
        if campo in body:
            campos.append(f"{campo} = %s"); valores.append(body[campo])

    if "password" in body and body["password"]:
        campos.append("password_hash = %s"); valores.append(hash_password(body["password"]))

    if "estado" in body:
        campos.append("estado = %s"); valores.append(body["estado"])
        if body["estado"] == "INACTIVO":
            campos.append("fecha_baja = now() AT TIME ZONE 'America/Lima'")
        elif body["estado"] == "ACTIVO":
            campos.append("fecha_baja = NULL")

    if not campos:
        return json.dumps({"ok": False, "error": "nada para actualizar"}), 400, HEADERS

    valores.append(asesor_id)
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(f"UPDATE bo_asesores SET {', '.join(campos)} WHERE id = %s", tuple(valores))
    conn.commit()
    cur.close(); conn.close()
    return json.dumps({"ok": True}), 200, HEADERS


@con_log_errores
@require_auth
def bo_productividad_asesor(request, asesor_id):
    """Tiempo de conexion acumulado + tareas realizadas, filtrable por fecha (?desde=&hasta=).
    Un asesor puede ver su propia productividad; para ver la de otros se requiere ADMIN."""
    if request.asesor_actual["id"] != int(asesor_id) and request.asesor_actual["rol"] != "ADMIN":
        return json.dumps({"ok": False, "error": "No autorizado"}), 403, HEADERS

    desde = request.args.get("desde")  # YYYY-MM-DD
    hasta = request.args.get("hasta")

    conn = get_connection()
    cur = conn.cursor()

    query_sesiones = """SELECT COALESCE(SUM(duracion_segundos), 0),
                                COALESCE(SUM(CASE WHEN fecha_logout IS NULL
                                    THEN EXTRACT(EPOCH FROM ((now() AT TIME ZONE 'America/Lima') - fecha_login))::int
                                    ELSE 0 END), 0)
                         FROM bo_sesiones WHERE asesor_id = %s"""
    params = [asesor_id]
    if desde:
        query_sesiones += " AND fecha_login >= %s"; params.append(desde)
    if hasta:
        query_sesiones += " AND fecha_login <= %s"; params.append(hasta)
    cur.execute(query_sesiones, tuple(params))
    seg_cerradas, seg_abierta = cur.fetchone()

    query_tareas = """SELECT estado, COUNT(*) FROM bo_tareas WHERE asesor_id = %s"""
    params2 = [asesor_id]
    if desde:
        query_tareas += " AND fecha_creacion >= %s"; params2.append(desde)
    if hasta:
        query_tareas += " AND fecha_creacion <= %s"; params2.append(hasta)
    query_tareas += " GROUP BY estado"
    cur.execute(query_tareas, tuple(params2))
    tareas_por_estado = {r[0]: r[1] for r in cur.fetchall()}

    cur.close(); conn.close()

    total_segundos = seg_cerradas + seg_abierta
    return json.dumps({
        "ok": True,
        "asesor_id": asesor_id,
        "tiempo_conectado_segundos": total_segundos,
        "tiempo_conectado_horas": round(total_segundos / 3600, 2),
        "tareas_por_estado": tareas_por_estado,
        "total_tareas": sum(tareas_por_estado.values()),
    }), 200, HEADERS

IMAGEN_OG_PROGRAMAS_SALUD = os.environ.get("IMAGEN_OG_PROGRAMAS_SALUD")

@con_log_errores
def handle_redirect_programa_salud(code, request):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT whatsapp_destino FROM programa_salud_links WHERE code = %s
            """, (code,))
            row = cur.fetchone()
            if not row:
                return "Enlace no encontrado", 404
            whatsapp_destino = row[0]

            cur.execute("""
                UPDATE programa_salud_links SET
                    raw_hits = raw_hits + 1,
                    ip_address = %s,
                    user_agent = %s,
                    entregado = TRUE,
                    entregado_at = COALESCE(entregado_at, NOW())
                WHERE code = %s
            """, (request.remote_addr, request.headers.get("User-Agent", ""), code))
            conn.commit()
    finally:
        conn.close()

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    url_canonica = f"{base_url}/ps/{code}"

    texto_wa = quote("Quiero inscribirme a programas")
    url_destino = f"https://wa.me/{whatsapp_destino}?text={texto_wa}"

    if IMAGEN_OG_PROGRAMAS_SALUD:
        img_url_meta = f"{IMAGEN_OG_PROGRAMAS_SALUD}?v={int(time.time())}"
        imagen_meta = f"""<meta property="og:image" content="{img_url_meta}">
  <meta property="og:image:width" content="1200">
  <meta property="og:image:height" content="630">
  <meta name="twitter:image" content="{img_url_meta}">"""
    else:
        imagen_meta = ""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta property="og:type" content="website">
  <meta property="og:url" content="{url_canonica}">
  <meta property="og:site_name" content="Sanitas Perú">
  <meta property="og:title" content="Programas de Salud Sanitas">
  <meta property="og:description" content="Te acompañamos en el cuidado de tu salud">
  {imagen_meta}
  <meta name="twitter:card" content="summary_large_image">
  <title>Programas de Salud Sanitas</title>
</head>
<body>
  <script>
    if (navigator.sendBeacon) {{
      navigator.sendBeacon("/ps/confirm/{code}");
    }} else {{
      fetch("/ps/confirm/{code}", {{ method: "POST", keepalive: true }});
    }}
    window.location.href = "{url_destino}";
  </script>
  <p>Redirigiendo a WhatsApp... <a href="{url_destino}">Haz clic aquí si no eres redirigido.</a></p>
</body>
</html>"""
    return html, 200


@con_log_errores
def handle_confirmar_click_programa_salud(code, request):
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM programa_salud_links WHERE code = %s", (code,))
            if not cur.fetchone():
                return "", 204

            cur.execute("""
                UPDATE programa_salud_links SET
                    clicked_at = COALESCE(clicked_at, NOW()),
                    click_count = click_count + 1
                WHERE code = %s
            """, (code,))
            conn.commit()
    finally:
        conn.close()

    return "", 204