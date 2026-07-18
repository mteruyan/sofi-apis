"""
admision_forms.py  –  Endpoints para formularios de admisión de clínicas
Integrar en main.py igual que sctr_forms.py

Endpoints a registrar en main.py:
  POST /admision/sctr_nuevo
  POST /admision/sctr_continuidad
  POST /admision/eps_nuevo
"""

import json
import os
import psycopg2
import pytz
from datetime import datetime
import base64
import logging
import requests

# ─────────────────────────────────────────────
# VARIABLES
# ─────────────────────────────────────────────
CORREO_DESTINO = "mteruya@sanitasperu.com"

HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
}

# ──────────────────────────────────────────────
# DB (mismo patrón que main.py)
# ──────────────────────────────────────────────
def get_conn():
    return psycopg2.connect(
        dbname="sofi_whatsapp",
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        host=f"/cloudsql/{os.environ['CLOUD_SQL_CONNECTION_NAME']}"
    )

def u(data, campo):
    val = data.get(campo, "")
    return val.strip().upper() if val else ""

# ──────────────────────────────────────────────
# Helper: insertar en tabla principal
# ──────────────────────────────────────────────
def insertar_formulario_principal(cur, tipo, data):
    lima  = pytz.timezone("America/Lima")
    ahora = datetime.now(lima)

    cur.execute("""
        INSERT INTO admision_formularios
            (tipo_formulario, celular_registro, clinica, telefono_clinica,
             nombre_admisionista, correo_clinica, dni_paciente,
             nombre_paciente, telefono_paciente, creado_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id
    """, (
        tipo,
        u(data, "celular_registro"),
        u(data, "clinica"),
		u(data, "telefono_clinica"),
        u(data, "nombre_admisionista"),
		u(data, "correo_clinica"),
        u(data, "dni_paciente"),
		u(data, "nombre_paciente"),
        u(data, "telefono_paciente"),
        ahora
    ))
    return cur.fetchone()[0]

# ──────────────────────────────────────────────
# Helper: validar campos requeridos
# ──────────────────────────────────────────────
def validar_campos(data, campos):
    for campo in campos:
        if not data.get(campo):
            return campo
    return None

# ──────────────────────────────────────────────
# POST /admision/sctr_nuevo
# ──────────────────────────────────────────────
def handle_admision_sctr_nuevo(request):
    data = request.get_json(force=True, silent=True) or {}

    campos_requeridos = [
        "celular_registro", "clinica", "nombre_admisionista",
        "dni_paciente", "nombre_paciente"
    ]
    campo_faltante = validar_campos(data, campos_requeridos)
    if campo_faltante:
        return (json.dumps({
            "ok": False,
            "error": f"Campo requerido: {campo_faltante}"
        }), 400, HEADERS)

    try:
        conn = get_conn()
        cur  = conn.cursor()

        formulario_id = insertar_formulario_principal(cur, "SCTR_NUEVO", data)

        cur.execute("""
            INSERT INTO admision_sctr_nuevo
                (formulario_id, ruc_empresa, razon_social, formato_sctr_firmado,
                 fecha_hora_accidente, horario_trabajo, hace_horas_extras,
                 cargo_paciente, area_trabajo, direccion_accidente,
                 implementos_seguridad, detalle_accidente)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            formulario_id,
            u(data, "ruc_empresa"),
            u(data, "razon_social"),
            u(data, "formato_sctr_firmado"),
            u(data, "fecha_hora_accidente"),
            u(data, "horario_trabajo"),
            u(data, "hace_horas_extras"),
            u(data, "cargo_paciente"),
            u(data, "area_trabajo"),
            u(data, "direccion_accidente"),
            u(data, "implementos_seguridad"),
            u(data, "detalle_accidente")
        ))

        conn.commit()
        cur.close()
        conn.close()

        if request.args.get("origen") == "web":
            correo_destino = CORREO_DESTINO
            _enviar_correo_admision("SCTR_NUEVO", _payload_sctr_nuevo(data, correo_destino))

        return (json.dumps({
            "ok":            True,
            "formulario_id": formulario_id,
            "tipo":          "SCTR_NUEVO",
            "mensaje":       "Formulario SCTR Nuevo registrado correctamente."
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)


# ──────────────────────────────────────────────
# POST /admision/sctr_continuidad
# ──────────────────────────────────────────────
def handle_admision_sctr_continuidad(request):
    data = request.get_json(force=True, silent=True) or {}

    campos_requeridos = [
        "celular_registro", "clinica", "nombre_admisionista",
        "dni_paciente", "nombre_paciente"
    ]
    campo_faltante = validar_campos(data, campos_requeridos)
    if campo_faltante:
        return (json.dumps({
            "ok": False,
            "error": f"Campo requerido: {campo_faltante}"
        }), 400, HEADERS)

    try:
        conn = get_conn()
        cur  = conn.cursor()

        formulario_id = insertar_formulario_principal(cur, "SCTR_CONTINUIDAD", data)

        cur.execute("""
            INSERT INTO admision_sctr_continuidad
                (formulario_id, ruc_empresa, razon_social, formato_sctr_firmado,
                 fecha_hora_accidente, fecha_ultima_atencion,
                 especialidad, diagnostico, lateralidad)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            formulario_id,
            u(data, "ruc_empresa"),
            u(data, "razon_social"),
            u(data, "formato_sctr_firmado"),
            u(data, "fecha_hora_accidente"),
			u(data, "fecha_ultima_atencion"),
            u(data, "especialidad"),
            u(data, "diagnostico"),
            u(data, "lateralidad")
        ))

        conn.commit()
        cur.close()
        conn.close()

        if request.args.get("origen") == "web":
            correo_destino = CORREO_DESTINO
            _enviar_correo_admision("SCTR_CONTINUIDAD", _payload_sctr_nuevo(data, correo_destino))

        return (json.dumps({
            "ok":            True,
            "formulario_id": formulario_id,
            "tipo":          "SCTR_CONTINUIDAD",
            "mensaje":       "Formulario SCTR Continuidad registrado correctamente."
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)


# ──────────────────────────────────────────────
# POST /admision/eps_nuevo
# ──────────────────────────────────────────────
def handle_admision_eps_nuevo(request):
    data = request.get_json(force=True, silent=True) or {}

    campos_requeridos = [
        "celular_registro", "clinica", "nombre_admisionista",
        "dni_paciente", "nombre_paciente"
    ]
    campo_faltante = validar_campos(data, campos_requeridos)
    if campo_faltante:
        return (json.dumps({
            "ok": False,
            "error": f"Campo requerido: {campo_faltante}"
        }), 400, HEADERS)

    try:
        conn = get_conn()
        cur  = conn.cursor()

        formulario_id = insertar_formulario_principal(cur, "EPS_NUEVO", data)

        cur.execute("""
            INSERT INTO admision_eps_nuevo
                (formulario_id, celular_paciente, especialidad)
            VALUES (%s, %s, %s)
        """, (
            formulario_id,
            u(data, "celular_paciente"),
            u(data, "especialidad")
        ))

        conn.commit()
        cur.close()
        conn.close()

        if request.args.get("origen") == "web":
            correo_destino = CORREO_DESTINO
            _enviar_correo_admision("EPS_NUEVO", _payload_sctr_nuevo(data, correo_destino))

        return (json.dumps({
            "ok":            True,
            "formulario_id": formulario_id,
            "tipo":          "EPS_NUEVO",
            "mensaje":       "Formulario EPS Nuevo registrado correctamente."
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)

# ──────────────────────────────────────────────
# POST /admision/consulta
# ──────────────────────────────────────────────
def handle_admision_consulta(request):
    data         = request.get_json(force=True, silent=True) or {}
    formulario_id = data.get("formulario_id")
    tipo          = (data.get("tipo") or "").strip().upper()

    if not formulario_id:
        return (json.dumps({"ok": False, "error": "Campo requerido: formulario_id"}), 400, HEADERS)
    if tipo not in ("SCTR_NUEVO", "SCTR_CONTINUIDAD", "EPS_NUEVO"):
        return (json.dumps({"ok": False, "error": "tipo debe ser SCTR_NUEVO, SCTR_CONTINUIDAD o EPS_NUEVO"}), 400, HEADERS)

    try:
        conn = get_conn()
        cur  = conn.cursor()

        # Datos comunes
        cur.execute("""
            SELECT tipo_formulario, celular_registro, clinica, telefono_clinica,
                   nombre_admisionista, correo_clinica, dni_paciente,
                   nombre_paciente, telefono_paciente, creado_at
            FROM admision_formularios
            WHERE id = %s AND tipo_formulario = %s
        """, (formulario_id, tipo))

        row = cur.fetchone()
        if not row:
            cur.close(); conn.close()
            return (json.dumps({"ok": False, "error": "Formulario no encontrado"}), 404, HEADERS)

        resultado = {
            "formulario_id":      formulario_id,
            "tipo_formulario":    row[0],
            "celular_registro":   row[1],
            "clinica":            row[2],
            "telefono_clinica":   row[3],
            "nombre_admisionista":row[4],
            "correo_clinica":     row[5],
            "dni_paciente":       row[6],
            "nombre_paciente":    row[7],
            "telefono_paciente":  row[8],
            "creado_at":          str(row[9]),
        }

        # Datos específicos según tipo
        if tipo == "SCTR_NUEVO":
            cur.execute("""
                SELECT ruc_empresa, razon_social, formato_sctr_firmado,
                       fecha_hora_accidente, horario_trabajo, hace_horas_extras,
                       cargo_paciente, area_trabajo, direccion_accidente,
                       implementos_seguridad, detalle_accidente
                FROM admision_sctr_nuevo
                WHERE formulario_id = %s
            """, (formulario_id,))
            det = cur.fetchone()
            if det:
                resultado.update({
                    "ruc_empresa":           det[0],
                    "razon_social":          det[1],
                    "formato_sctr_firmado":  det[2],
                    "fecha_hora_accidente":  det[3],
                    "horario_trabajo":       det[4],
                    "hace_horas_extras":     det[5],
                    "cargo_paciente":        det[6],
                    "area_trabajo":          det[7],
                    "direccion_accidente":   det[8],
                    "implementos_seguridad": det[9],
                    "detalle_accidente":     det[10],
                })

        elif tipo == "SCTR_CONTINUIDAD":
            cur.execute("""
                SELECT ruc_empresa, razon_social, formato_sctr_firmado,
                       fecha_hora_accidente, fecha_ultima_atencion,
                       especialidad, diagnostico, lateralidad
                FROM admision_sctr_continuidad
                WHERE formulario_id = %s
            """, (formulario_id,))
            det = cur.fetchone()
            if det:
                resultado.update({
                    "ruc_empresa":          det[0],
                    "razon_social":         det[1],
                    "formato_sctr_firmado": det[2],
                    "fecha_hora_accidente": det[3],
                    "fecha_ultima_atencion":det[4],
                    "especialidad":         det[5],
                    "diagnostico":          det[6],
                    "lateralidad":          det[7],
                })

        elif tipo == "EPS_NUEVO":
            cur.execute("""
                SELECT celular_paciente, especialidad
                FROM admision_eps_nuevo
                WHERE formulario_id = %s
            """, (formulario_id,))
            det = cur.fetchone()
            if det:
                resultado.update({
                    "celular_paciente": det[0],
                    "especialidad":    det[1],
                })

        cur.close(); conn.close()
        return (json.dumps({"ok": True, "data": resultado}, ensure_ascii=False), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"ok": False, "error": str(e)}), 500, HEADERS)

def _payload_eps_nuevo(data, correo_destino):
    return {
        "NOMBRE_CLINICA": data.get("clinica", ""),
        "TELEFONO_CLINICA": data.get("telefono_clinica", ""),
        "EMAIL_CLINICA": data.get("correo_clinica", ""),
        "NOMBRE_ADMISIONISTA": data.get("nombre_admisionista", ""),
        "DOI_PACIENTE": data.get("dni_paciente", ""),
        "NOMBRE_PACIENTE": data.get("nombre_paciente", ""),
        "CELULAR_PACIENTE": data.get("celular_paciente", ""),
        "ESPECIALIDAD": data.get("especialidad", ""),
        "CORREO_DESTINO": correo_destino,
    }

def _payload_sctr_nuevo(data, correo_destino):
    return {
        "NOMBRE_CLINICA": data.get("clinica", ""),
        "TELEFONO_CLINICA": data.get("telefono_clinica", ""),
        "EMAIL_CLINICA": data.get("correo_clinica", ""),
        "NOMBRE_ADMISIONISTA": data.get("nombre_admisionista", ""),
        "DNI_PACIENTE": data.get("dni_paciente", ""),
        "NOMBRE_PACIENTE": data.get("nombre_paciente", ""),
        "CELULAR_PACIENTE": data.get("celular_registro", ""),
        "RUC": data.get("ruc_empresa", ""),
        "RAZON_SOCIAL": data.get("razon_social", ""),
        "FORMATO_SCTR_FIRMADO": data.get("formato_sctr_firmado", ""),
        "FECHA_HORA_ACCIDENTE": data.get("fecha_hora_accidente", ""),
        "HORARIO_TRABAJADOR": data.get("horario_trabajo", ""),
        "HORAS_EXTRAS": data.get("hace_horas_extras", ""),
        "CARGO_TRABAJADOR": data.get("cargo_paciente", ""),
        "AREA_TRABAJADOR": data.get("area_trabajo", ""),
        "IMPLEMENTOS_SEGURIDAD": data.get("implementos_seguridad", ""),
        "DIRECCION_ACCIDENTE": data.get("direccion_accidente", ""),
        "DESCRIPCION_ACCIDENTE": data.get("detalle_accidente", ""),
        "CORREO_DESTINO": correo_destino,
    }

def _payload_sctr_continuidad(data, correo_destino):
    return {
        "NOMBRE_CLINICA": data.get("clinica", ""),
        "TELEFONO_CLINICA": data.get("telefono_clinica", ""),
        "EMAIL_CLINICA": data.get("correo_clinica", ""),
        "NOMBRE_ADMISIONISTA": data.get("nombre_admisionista", ""),
        "DOI_PACIENTE": data.get("dni_paciente", ""),
        "NOMBRE_PACIENTE": data.get("nombre_paciente", ""),
        "CELULAR_PACIENTE": data.get("celular_registro", ""),
        "RUC": data.get("ruc_empresa", ""),
        "RAZON_SOCIAL": data.get("razon_social", ""),
        "FORMATO_SCTR_FIRMADO": data.get("formato_sctr_firmado", ""),
        "FECHA_HORA_ACCIDENTE": data.get("fecha_hora_accidente", ""),
        "FECHA_ULT_ATENCION": data.get("fecha_ultima_atencion", ""),
        "ESPECIALIDAD": data.get("especialidad", ""),
        "DIAGNOSTICO": data.get("diagnostico", ""),
        "LATERALIDAD": data.get("lateralidad", ""),
        "CORREO_DESTINO": correo_destino,
    }

# ─────────────────────────────────────────────
# Credenciales DANAConnect (guardar en Secret Manager)
# ─────────────────────────────────────────────

DANACONNECT_AUTH = os.environ.get("DANACONNECT_AUTH")

CONVERSATION_IDS = {
    "EPS_NUEVO": 592884,
    "SCTR_NUEVO": 592886,
    "SCTR_CONTINUIDAD": 592888,
}

# API: DANAConnect - start/data
def _enviar_correo_admision(tipo, payload):
    """
    Envía la notificación transaccional a DANAConnect luego de un insert exitoso.
    No lanza excepción si falla (para no romper el flujo principal de admisión);
    solo loguea el error.
    """
    conversation_id = CONVERSATION_IDS.get(tipo)
    if not conversation_id:
        logging.error(f"[DANAConnect] Tipo de admisión no reconocido: {tipo}")
        return False

    url = f"https://appserv.danaconnect.com/api/1.0/rest/conversation/{conversation_id}/start/data"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"{DANACONNECT_AUTH}",
        "X-DEBUG": "0",
    }

    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=10)
        resp.raise_for_status()
        logging.info(f"[DANAConnect] Notificación '{tipo}' enviada OK: {resp.text}")
        return True
    except requests.exceptions.RequestException as e:
        detalle = getattr(e.response, "text", str(e))
        logging.error(f"[DANAConnect] Error enviando '{tipo}': {detalle}")
        return False