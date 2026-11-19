"""
main_mad.py — Médico a domicilio (MAD) y ambulancia · sofi-apis
================================================================
Módulo autónomo: trae su propia conexión a la base de datos y sus utilidades.
Todas las rutas MAD se resuelven en mad_router(request, path).

INTEGRACIÓN EN main.py (2 cambios):

    1) Arriba, con los demás imports:
           from main_mad import mad_router

    2) Dentro de api(request), justo después de:   path = request.path.rstrip("/")
           respuesta_mad = mad_router(request, path)
           if respuesta_mad is not None:
               return respuesta_mad

El código MAD viejo de main.py queda sin uso (nunca se alcanza) y se puede borrar con calma.

IMPORTANTE: el filtro de ApiKey de main.py debe seguir dejando públicas las rutas
/track/, /oferta/, /proveedor, /proveedores/, /ofertas/, /solicitudes/ y /mad/
(se protegen con sesión) y NO debe incluir /mad-bootstrap/ (esa exige ApiKey).

Variables de entorno: DB_USER, DB_PASSWORD, CLOUD_SQL_CONNECTION_NAME, DANACONNECT_AUTH,
API_KEY, BASE_URL y, opcional, MAD_ALERTA_CELULARES (celulares de coordinación, separados por coma).
Dependencias: psycopg2, pytz, pyotp, requests, openpyxl y crypto_utils (ya existente en el proyecto).
"""
import os
import re
import json
import uuid
import time
import hmac
import string
import hashlib
import logging
import secrets
import traceback
from datetime import datetime, timedelta
from functools import wraps
from io import BytesIO

import psycopg2
import pytz
import pyotp
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from crypto_utils import cifrar, descifrar

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIMA_TZ = pytz.timezone("America/Lima")

HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json",
}

DANACONNECT_URL = "https://appserv.danaconnect.com/api/1.0/rest/conversation/517078/start/data"
DANACONNECT_AUTH = os.environ.get("DANACONNECT_AUTH", "")
WHATSAPP_URL = "sanitasperu.com/?ac=1"

# Despacho
TIMEOUT_OFERTA_SEGUNDOS = 120      # tiempo que tiene el doctor para aceptar o rechazar
MAX_INTENTOS_ASIGNACION = 6        # después de esto la solicitud pasa a sin_cobertura
INCREMENTO_RADIO_METROS = 7500     # cada reintento amplía el radio de búsqueda
RADIO_MAXIMO_METROS = 40000     # radios de búsqueda: 5, 12.5, 20, 27.5, 35 y 40 km
ESPERA_REINTENTO_SIN_PROVEEDOR_SEGUNDOS = 45   # menor que el intervalo del scheduler (60 s)
MAD_ALERTA_CELULARES = [c.strip() for c in os.environ.get("MAD_ALERTA_CELULARES", "").split(",") if c.strip()]
# Cambiar este valor (y APP_VERSION en proveedor.html) cada vez que se publique una versión nueva del panel
MAD_APP_VERSION = "2026-10-29.1"
# Ubicación vieja tolerada: un doctor "disponible" recibe ofertas según su ÚLTIMA posición conocida mientras
# no pase de este tiempo; pasado esto se desconecta solo. Con avisos push el doctor no necesita tener la app abierta.
MAD_APK_VERSION = os.environ.get("MAD_APK_VERSION", "1.0")      # versión vigente de la app Android (APK)
MAD_REABRIR_MIN = int(os.environ.get("MAD_REABRIR_MIN", "60"))   # minutos en que un servicio del call center sin cobertura se reabre si se conecta un doctor
MAD_AUTODESCONEXION_MIN = int(os.environ.get("MAD_AUTODESCONEXION_MIN", "45"))
MAD_SENAL_FRESCA_MIN = 5           # hasta aquí la posición se considera "en vivo"; los doctores con señal fresca van primero
MAD_OFERTA_LINK_AL_PANEL = True    # True: el SMS de oferta lleva al panel (/proveedor); False: a /oferta/<token>

# Tiempo estimado de llegada
VELOCIDAD_DEFAULT_KMH = 22
MINUTOS_PREPARACION = 5
MIN_MUESTRAS_CALIBRACION = 5
MAX_CAMBIOS_MEDICO = 1

# Cancelación y liberación
MOTIVOS_CANCELACION = {
    "demora":         "Tardará mucho en llegar",
    "medico":         "No quiero que me atienda este médico",
    "ya_no_necesito": "Ya no necesito el servicio",
    "otro_lugar":     "Me atendí en otro lugar",
    "error":          "Lo solicité por error",
    "reprogramar":    "Necesito cambiar la fecha, hora o lugar",
    "otro":           "Otro motivo",
}
ESTADOS_CANCELABLES = ("programada", "buscando_proveedor", "asignado", "en_camino")

MOTIVOS_LIBERACION = {
    "sin_senal":       "El doctor perdió señal o no responde",
    "no_acude":        "El doctor no acude al domicilio",
    "doctor_solicito": "El doctor pidió que lo liberen",
    "otro":            "Otro motivo",
}
ESTADOS_LIBERABLES = ("asignado", "en_camino", "llego")
MAD_LIB_ALERTA_7D = 2
MAD_LIB_ALERTA_30D = 4

# Servicios programados
MAD_PROG_MIN_MIN = 60
MAD_PROG_MAX_DIAS = 14
MAD_PROG_ACTIVAR_MIN = 45
MAD_PROG_MAX_ACTIVAS = 3
PERU_LAT = (-18.5, 0.1)
PERU_LNG = (-81.5, -68.5)
MAD_COBERTURA_HORAS_1 = 3
MAD_COBERTURA_HORAS_2 = 1
MAD_COBERTURA_SENAL_MIN = MAD_AUTODESCONEXION_MIN   # la cobertura cuenta a quien puede recibir ofertas
MAD_MAX_REPROGRAMACIONES = 2
MAD_REPROG_DISTANCIA_M = 150
ESTADOS_REPROGRAMABLES = ("programada", "buscando_proveedor", "asignado")

# Sesiones y seguridad
MAD_SESION_HORAS = 12
MAD_PARCIAL_MINUTOS = 5
MAD_MAX_INTENTOS = 5
MAD_BLOQUEO_MINUTOS = 15
MAD_TOTP_EMISOR = "Sanitas MAD"

# Empresas
MAD_MAX_DOCTORES_EMPRESA = 50
MAD_MAX_USUARIOS_EMPRESA = 5
MAD_UBICACION_MAX_MIN = 30
MAD_EXCEL_MAX_FILAS = 20000

# Día (hora de Lima) en que se finalizó el servicio.
# Si solicitudes.finalizado_en es "timestamp without time zone", usar: "s.finalizado_en::date"
FECHA_SERVICIO_SQL = "(s.finalizado_en AT TIME ZONE 'America/Lima')::date"


# ─────────────────────────────────────────────
# UTILIDADES PROPIAS (el módulo no depende de main.py)
# ─────────────────────────────────────────────
def get_connection():
    return psycopg2.connect(
        dbname="sofi_whatsapp",
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        host=f"/cloudsql/{os.environ['CLOUD_SQL_CONNECTION_NAME']}",
    )


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


def _mad_json(data, status=200):
    return (json.dumps(data, ensure_ascii=False, default=str), status, HEADERS)


def _iso(v):
    return v.isoformat() if v else None


def _a_bytes(v):
    """psycopg2 devuelve las columnas BYTEA como memoryview; descifrar() espera bytes."""
    return v.tobytes() if isinstance(v, memoryview) else v


def validar_api_key(request):
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("ApiKey "):
        return False
    esperada = os.environ.get("API_KEY", "")
    return bool(esperada) and hmac.compare_digest(auth.split(" ", 1)[1], esperada)


def enviar_danaconnect(numero, mensaje):
    payload = {"NUM": f" {numero}", "TEXT": mensaje, "URL": WHATSAPP_URL}
    headers = {"Content-Type": "application/json", "Authorization": DANACONNECT_AUTH}
    try:
        response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Error al enviar SMS a {numero}: {e}")


def sms_aviso(celular, descripcion):
    """Bitácora de SMS enviados. sofi_sms_avisos.descripcion es varchar(30); nunca debe romper el flujo."""
    try:
        conn = get_connection(); cur = conn.cursor()
        try:
            cur.execute("INSERT INTO sofi_sms_avisos (celular, descripcion, fecha) VALUES (%s, %s, %s)",
                        (celular, (descripcion or "")[:30], datetime.now(LIMA_TZ).strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
        finally:
            cur.close(); conn.close()
    except Exception as e:
        logging.warning(f"[sms_aviso] no se pudo registrar: {e}")


def _normalizar_celular_pe(celular):
    """Deja el número en formato +51XXXXXXXXX. None si no se puede usar."""
    if not celular:
        return None
    if re.match(r"^[A-Z]{2}\.\d+$", celular):      # wa_id ofuscado de WhatsApp
        return None
    limpio = re.sub(r"[^\d]", "", celular)
    if len(limpio) == 9 and limpio.startswith("9"):
        return f"+51{limpio}"
    if len(limpio) == 11 and limpio.startswith("51"):
        return f"+{limpio}"
    if celular.startswith("+"):
        return celular
    logging.warning(f"[MAD] Número con formato inesperado, se envía tal cual: {celular}")
    return f"+{limpio}"


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


def _hora_en_rango(hora, inicio, fin):
    if inicio <= fin:
        return inicio <= hora < fin
    return hora >= inicio or hora < fin   # rango que cruza medianoche


# ═════════════════════════════════════════════
# AUTENTICACIÓN (clave + Google Authenticator)
# ═════════════════════════════════════════════
_mad_hash_senuelo = None
_MSG_BLOQUEO = f"Cuenta bloqueada temporalmente por intentos fallidos. Intenta en {MAD_BLOQUEO_MINUTOS} minutos."


def _mad_hash_token(token):
    return hashlib.sha256(token.encode()).hexdigest()


def _mad_check_password(password, password_hash):
    """Compara en tiempo similar exista o no el usuario (evita adivinar usuarios por tiempo de respuesta)."""
    global _mad_hash_senuelo
    if password_hash:
        return check_password(password, password_hash)
    if _mad_hash_senuelo is None:
        _mad_hash_senuelo = hash_password(secrets.token_hex(8))
    check_password(password, _mad_hash_senuelo)
    return False


def _mad_password_valida(p):
    return len(p) >= 10 and bool(re.search(r"[A-Za-z]", p)) and bool(re.search(r"\d", p))


def _mad_password_temporal():
    alfabeto = string.ascii_letters + string.digits
    while True:
        p = "".join(secrets.choice(alfabeto) for _ in range(12))
        if _mad_password_valida(p):
            return p


def _mad_crear_sesion(cur, usuario_id, tipo):
    token = secrets.token_urlsafe(32)
    duracion = timedelta(hours=MAD_SESION_HORAS) if tipo == "completa" else timedelta(minutes=MAD_PARCIAL_MINUTOS)
    cur.execute("""
        INSERT INTO mad_sesiones (token_hash, usuario_id, tipo, expira_en)
        VALUES (%s, %s, %s, NOW() + %s)
    """, (_mad_hash_token(token), usuario_id, tipo, duracion))
    return token


MAD_DISPOSITIVO_DIAS = int(os.environ.get("MAD_DISPOSITIVO_DIAS", "30"))
MAD_MAX_DISPOSITIVOS = 3          # celulares de confianza por usuario


def _mad_crear_dispositivo(cur, usuario_id, etiqueta=None, tipo="confianza"):
    """Celular de confianza ('confianza': entra sin código de Authenticator tras usuario y clave) o
    ingreso con huella ('biometrico': reemplaza la clave en ese celular). Ambos duran MAD_DISPOSITIVO_DIAS días.
    Solo se entrega con sesión completa. Guarda únicamente el hash del token."""
    token = secrets.token_urlsafe(32)
    cur.execute("""
        UPDATE mad_dispositivos SET revocado = TRUE
        WHERE usuario_id = %s AND tipo = %s AND NOT revocado
          AND id NOT IN (SELECT id FROM mad_dispositivos WHERE usuario_id = %s AND tipo = %s AND NOT revocado
                         AND expira_en > NOW() ORDER BY id DESC LIMIT %s)
    """, (usuario_id, tipo, usuario_id, tipo, MAD_MAX_DISPOSITIVOS - 1))
    cur.execute(f"""
        INSERT INTO mad_dispositivos (usuario_id, token_hash, etiqueta, tipo, expira_en)
        VALUES (%s, %s, %s, %s, NOW() + INTERVAL '{MAD_DISPOSITIVO_DIAS} days')
    """, (usuario_id, _mad_hash_token(token), (etiqueta or "")[:120], tipo))
    return token


def _mad_token_de(request):
    auth = request.headers.get("Authorization", "")
    return auth[7:].strip() if auth.startswith("Bearer ") else ""


def _mad_leer_sesion(request, tipo="completa"):
    token = _mad_token_de(request)
    if not token:
        return None
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT u.id, u.usuario, u.nombre, u.rol, u.proveedor_id,
                   u.debe_cambiar_password, u.totp_activo, u.empresa_id
            FROM mad_sesiones s
            JOIN mad_usuarios u ON u.id = s.usuario_id
            WHERE s.token_hash = %s AND s.tipo = %s AND NOT s.revocado
              AND s.expira_en > NOW() AND u.activo
        """, (_mad_hash_token(token), tipo))
        r = cur.fetchone()
    finally:
        cur.close(); conn.close()
    if not r:
        return None
    return {"id": r[0], "usuario": r[1], "nombre": r[2], "rol": r[3], "proveedor_id": r[4],
            "debe_cambiar_password": r[5], "totp_activo": r[6], "empresa_id": r[7],
            "token_hash": _mad_hash_token(token)}


def _mad_exigir(request, rol=None, tipo="completa"):
    s = _mad_leer_sesion(request, tipo)
    if not s:
        return None, _mad_json({"error": "No autenticado"}, 401)
    if rol and s["rol"] != rol:
        return None, _mad_json({"error": "No autorizado"}, 403)
    return s, None


def _exigir_proveedor(request, proveedor_id_ruta=None):
    """Sesión de doctor. Si se indica proveedor_id_ruta, debe ser el de la sesión (nunca el de la URL)."""
    s, err = _mad_exigir(request, "proveedor")
    if err:
        return None, err
    if proveedor_id_ruta is not None and str(s["proveedor_id"]) != str(proveedor_id_ruta):
        return None, _mad_json({"error": "No autorizado"}, 403)
    return s, None


def _oferta_es_del_proveedor(proveedor_id, oferta_id=None, oferta_token=None):
    conn = get_connection(); cur = conn.cursor()
    try:
        if oferta_token is not None:
            try:
                uuid.UUID(oferta_token)
            except ValueError:
                return False
            cur.execute("SELECT 1 FROM solicitudes_ofertas WHERE oferta_token = %s AND proveedor_id = %s",
                        (oferta_token, proveedor_id))
        else:
            cur.execute("SELECT 1 FROM solicitudes_ofertas WHERE id = %s AND proveedor_id = %s",
                        (oferta_id, proveedor_id))
        return cur.fetchone() is not None
    finally:
        cur.close(); conn.close()


def _mad_registrar_fallo(cur, usuario_id):
    cur.execute("""
        UPDATE mad_usuarios
        SET bloqueado_hasta   = CASE WHEN intentos_fallidos + 1 >= %s THEN NOW() + %s ELSE bloqueado_hasta END,
            intentos_fallidos = CASE WHEN intentos_fallidos + 1 >= %s THEN 0 ELSE intentos_fallidos + 1 END
        WHERE id = %s
    """, (MAD_MAX_INTENTOS, timedelta(minutes=MAD_BLOQUEO_MINUTOS), MAD_MAX_INTENTOS, usuario_id))


def _mad_audit(cur, usuario_id, accion, objetivo, detalle=None):
    cur.execute("""
        INSERT INTO mad_auditoria (usuario_id, accion, objetivo, detalle)
        VALUES (%s, %s, %s, %s)
    """, (usuario_id, accion, objetivo,
          json.dumps(detalle, ensure_ascii=False, default=str) if detalle is not None else None))


def _mad_enviar_avisos(avisos, descripcion):
    for tel, texto in avisos:
        numero = _normalizar_celular_pe(tel)
        if numero:
            enviar_danaconnect(numero, texto)
            sms_aviso(numero, descripcion)


# ── Paso 1: usuario + clave → sesión parcial (5 min) ──
@con_log_errores
def handle_mad_login(request):
    data = request.get_json(force=True, silent=True) or {}
    usuario = (data.get("usuario") or "").strip().lower()
    password = data.get("password") or ""
    if not usuario or not password:
        return _mad_json({"error": "Usuario y clave son obligatorios"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT id, password_hash, activo, COALESCE(bloqueado_hasta > NOW(), FALSE),
                   debe_cambiar_password, totp_activo
            FROM mad_usuarios WHERE usuario = %s
            FOR UPDATE
        """, (usuario,))
        r = cur.fetchone()
        if not r or not r[2]:
            _mad_check_password(password, None)
            return _mad_json({"error": "Usuario o clave incorrectos"}, 401)

        uid, phash, _, bloqueado, debe_cambiar, totp_activo = r
        if bloqueado:
            return _mad_json({"error": _MSG_BLOQUEO}, 423)
        if not _mad_check_password(password, phash):
            _mad_registrar_fallo(cur, uid)
            conn.commit()
            return _mad_json({"error": "Usuario o clave incorrectos"}, 401)

        # Celular de confianza (solo doctores): si el token es válido, no se pide el código de Authenticator
        dispositivo_invalido = False
        tok_disp = str(data.get("token_dispositivo") or "")
        if tok_disp:
            cur.execute("""
                SELECT d.id, u.nombre, u.rol, u.proveedor_id
                FROM mad_dispositivos d JOIN mad_usuarios u ON u.id = d.usuario_id
                WHERE d.token_hash = %s AND d.usuario_id = %s AND NOT d.revocado AND d.expira_en > NOW()
                  AND d.tipo = 'confianza'
                  AND u.rol = 'proveedor' AND u.totp_activo AND NOT u.debe_cambiar_password
            """, (_mad_hash_token(tok_disp), uid))
            d = cur.fetchone()
            if d:
                cur.execute("UPDATE mad_dispositivos SET ultimo_uso = NOW() WHERE id = %s", (d[0],))
                cur.execute("""UPDATE mad_usuarios SET intentos_fallidos = 0, bloqueado_hasta = NULL, ultimo_login = NOW()
                               WHERE id = %s""", (uid,))
                token_completo = _mad_crear_sesion(cur, uid, "completa")
                _mad_audit(cur, uid, "login_dispositivo", f"usuario:{uid}")
                conn.commit()
                return _mad_json({"token": token_completo, "siguiente": "ok", "usuario": {
                    "nombre": d[1], "usuario": usuario, "rol": d[2], "proveedor_id": d[3]}})
            dispositivo_invalido = True

        # El contador de fallos NO se reinicia aquí, solo tras el código TOTP correcto
        token = _mad_crear_sesion(cur, uid, "parcial")
        conn.commit()
    finally:
        cur.close(); conn.close()

    siguiente = "cambiar_password" if debe_cambiar else ("totp" if totp_activo else "enrolar_totp")
    return _mad_json({"token_parcial": token, "siguiente": siguiente, "dispositivo_invalido": dispositivo_invalido})


# ── Ingreso con huella (solo doctores, app Android) ──
# El celular guarda el token cifrado con una llave que solo se abre con la huella. El servidor guarda su hash.
@con_log_errores
def handle_mad_biometrico_activar(request):
    """Con la sesión completa abierta (entró con clave), entrega el token para guardar con huella."""
    s, err = _mad_exigir(request, "proveedor")
    if err:
        return err
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT totp_activo, debe_cambiar_password FROM mad_usuarios WHERE id = %s", (s["id"],))
        r = cur.fetchone()
        if not r or not r[0] or r[1]:
            return _mad_json({"error": "Completa primero la configuración de tu cuenta"}, 409)
        token = _mad_crear_dispositivo(cur, s["id"], "huella app", tipo="biometrico")
        _mad_audit(cur, s["id"], "biometrico_activar", f"usuario:{s['id']}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"token_biometrico": token, "usuario": s["usuario"], "dias": MAD_DISPOSITIVO_DIAS})


@con_log_errores
def handle_mad_biometrico_login(request):
    data = request.get_json(force=True, silent=True) or {}
    usuario = (data.get("usuario") or "").strip().lower()
    tok = str(data.get("token") or "")
    if not usuario or not tok:
        return _mad_json({"error": "Datos incompletos"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT d.id, u.id, u.nombre, u.rol, u.proveedor_id
            FROM mad_dispositivos d JOIN mad_usuarios u ON u.id = d.usuario_id
            WHERE u.usuario = %s AND d.token_hash = %s AND d.tipo = 'biometrico'
              AND NOT d.revocado AND d.expira_en > NOW()
              AND u.activo AND u.rol = 'proveedor' AND u.totp_activo AND NOT u.debe_cambiar_password
              AND NOT COALESCE(u.bloqueado_hasta > NOW(), FALSE)
            FOR UPDATE OF d
        """, (usuario, _mad_hash_token(tok)))
        d = cur.fetchone()
        if not d:
            return _mad_json({"error": "El ingreso con huella ya no es válido. Ingresa con tu clave."}, 401)
        # La vigencia se renueva con cada uso
        cur.execute(f"""UPDATE mad_dispositivos SET ultimo_uso = NOW(), expira_en = NOW() + INTERVAL '{MAD_DISPOSITIVO_DIAS} days'
                        WHERE id = %s""", (d[0],))
        cur.execute("UPDATE mad_usuarios SET ultimo_login = NOW() WHERE id = %s", (d[1],))
        token_completo = _mad_crear_sesion(cur, d[1], "completa")
        _mad_audit(cur, d[1], "login_biometrico", f"usuario:{d[1]}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"token": token_completo, "siguiente": "ok", "usuario": {
        "nombre": d[2], "usuario": usuario, "rol": d[3], "proveedor_id": d[4]}})



# ── Primer ingreso: cambiar la clave temporal ──
@con_log_errores
def handle_mad_cambiar_password(request):
    s, err = _mad_exigir(request, tipo="parcial")
    if err:
        return err
    nueva = (request.get_json(force=True, silent=True) or {}).get("nueva_password") or ""
    if not _mad_password_valida(nueva):
        return _mad_json({"error": "La clave debe tener al menos 10 caracteres, con letras y números"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("UPDATE mad_usuarios SET password_hash = %s, debe_cambiar_password = FALSE WHERE id = %s",
                    (hash_password(nueva), s["id"]))
        cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (s["id"],))
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"siguiente": "totp" if s["totp_activo"] else "enrolar_totp"})


# ── Primer ingreso: generar secreto y QR de Google Authenticator ──
@con_log_errores
def handle_mad_totp_iniciar(request):
    s, err = _mad_exigir(request, tipo="parcial")
    if err:
        return err
    if s["debe_cambiar_password"]:
        return _mad_json({"error": "Primero debes cambiar tu clave"}, 409)
    if s["totp_activo"]:
        return _mad_json({"error": "Google Authenticator ya está configurado"}, 409)

    secreto = pyotp.random_base32()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("UPDATE mad_usuarios SET totp_secret_enc = %s, totp_ultimo_paso = NULL WHERE id = %s",
                    (cifrar(secreto), s["id"]))
        conn.commit()
    finally:
        cur.close(); conn.close()

    uri = pyotp.TOTP(secreto).provisioning_uri(name=s["usuario"], issuer_name=MAD_TOTP_EMISOR)
    return _mad_json({"otpauth_uri": uri, "secreto": secreto})


# ── Paso 2: código de 6 dígitos → sesión completa (12 h) ──
@con_log_errores
def handle_mad_totp_verificar(request):
    s, err = _mad_exigir(request, tipo="parcial")
    if err:
        return err
    if s["debe_cambiar_password"]:
        return _mad_json({"error": "Primero debes cambiar tu clave"}, 409)

    codigo = re.sub(r"\D", "", str((request.get_json(force=True, silent=True) or {}).get("codigo") or ""))
    if len(codigo) != 6:
        return _mad_json({"error": "Ingresa los 6 dígitos de Google Authenticator"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT totp_secret_enc, totp_ultimo_paso, COALESCE(bloqueado_hasta > NOW(), FALSE)
            FROM mad_usuarios WHERE id = %s
            FOR UPDATE
        """, (s["id"],))
        secreto_enc, ultimo_paso, bloqueado = cur.fetchone()
        if bloqueado:
            return _mad_json({"error": _MSG_BLOQUEO}, 423)
        if not secreto_enc:
            return _mad_json({"error": "Primero configura Google Authenticator"}, 409)

        totp = pyotp.TOTP(descifrar(_a_bytes(secreto_enc)))
        paso_actual = int(time.time()) // 30
        paso_valido = None
        for delta in (0, -1, 1):   # tolera ±30 s de desfase del reloj del celular
            if hmac.compare_digest(totp.at((paso_actual + delta) * 30), codigo):
                paso_valido = paso_actual + delta
                break

        # Rechaza códigos incorrectos y códigos ya usados (anti-repetición)
        if paso_valido is None or (ultimo_paso is not None and paso_valido <= ultimo_paso):
            _mad_registrar_fallo(cur, s["id"])
            conn.commit()
            return _mad_json({"error": "Código incorrecto o ya utilizado"}, 400)

        cur.execute("""
            UPDATE mad_usuarios
            SET totp_activo = TRUE, totp_ultimo_paso = %s, intentos_fallidos = 0,
                bloqueado_hasta = NULL, ultimo_login = NOW()
            WHERE id = %s
        """, (paso_valido, s["id"]))
        cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE token_hash = %s", (s["token_hash"],))
        token = _mad_crear_sesion(cur, s["id"], "completa")
        token_dispositivo = None
        if (request.get_json(force=True, silent=True) or {}).get("recordar") and s["rol"] == "proveedor":
            token_dispositivo = _mad_crear_dispositivo(cur, s["id"], request.headers.get("User-Agent"))
            _mad_audit(cur, s["id"], "dispositivo_confiable", f"usuario:{s['id']}")
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"token": token, "token_dispositivo": token_dispositivo, "usuario": {
        "nombre": s["nombre"], "usuario": s["usuario"], "rol": s["rol"], "proveedor_id": s["proveedor_id"]
    }})


@con_log_errores
def handle_mad_logout(request):
    token = _mad_token_de(request)
    if token:
        conn = get_connection(); cur = conn.cursor()
        try:
            cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE token_hash = %s", (_mad_hash_token(token),))
            conn.commit()
        finally:
            cur.close(); conn.close()
    return _mad_json({"ok": True})


@con_log_errores
def handle_mad_yo(request):
    s, err = _mad_exigir(request)
    if err:
        return err
    return _mad_json({k: s[k] for k in ("nombre", "usuario", "rol", "proveedor_id")})


# ── Solo para crear coordinadores. Exige la ApiKey de sofi-apis ──
@con_log_errores
def handle_mad_crear_coordinador(request):
    if not validar_api_key(request):
        return _mad_json({"error": "API Key invalida"}, 401)
    data = request.get_json(force=True, silent=True) or {}
    usuario = (data.get("usuario") or "").strip().lower()
    nombre = (data.get("nombre") or "").strip()[:120]
    if not nombre or not re.fullmatch(r"[a-z0-9._-]{4,30}", usuario):
        return _mad_json({"error": "usuario (4 a 30 caracteres) y nombre son obligatorios"}, 400)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_usuarios WHERE usuario = %s", (usuario,))
        if cur.fetchone():
            return _mad_json({"error": "Ese usuario ya existe"}, 409)
        cur.execute("""
            INSERT INTO mad_usuarios (usuario, nombre, rol, password_hash)
            VALUES (%s, %s, 'coordinador', %s)
        """, (usuario, nombre, hash_password(temporal)))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "usuario": usuario, "password_temporal": temporal}, 201)


# ═════════════════════════════════════════════════════════════
# DESPACHO — validaciones compartidas
# ═════════════════════════════════════════════════════════════
def _uuid_valido(valor):
    try:
        uuid.UUID(str(valor))
        return True
    except (ValueError, AttributeError, TypeError):
        return False


def _parse_programada(valor):
    """'2026-10-18T15:30' (hora de Lima) -> datetime con zona. ValueError si es inválido."""
    texto = str(valor).strip().replace(" ", "T")[:16]
    return LIMA_TZ.localize(datetime.strptime(texto, "%Y-%m-%dT%H:%M"))


def _horarios_programacion(cur):
    cur.execute("SELECT hora_inicio, hora_fin FROM sofi_horarios WHERE clave = 'mad_programado' AND activo = TRUE")
    return cur.fetchall()


def _validar_lugar(lat, lng, direccion, distrito):
    if lat is None or lng is None:
        return "Falta la ubicación (lat y lng)"
    if not (PERU_LAT[0] <= lat <= PERU_LAT[1] and PERU_LNG[0] <= lng <= PERU_LNG[1]):
        return "La ubicación debe estar dentro de Perú"
    if len(direccion) < 5 or not distrito:
        return "direccion (mínimo 5 caracteres) y distrito son obligatorios"
    return None


def _normalizar_contacto(valor):
    """Devuelve (telefono, error). Vacío es válido (el contacto es opcional)."""
    tel = re.sub(r"\D", "", valor or "")
    if not tel:
        return None, None
    if len(tel) == 9 and tel.startswith("9"):
        tel = "51" + tel
    if not re.fullmatch(r"51\d{9}", tel):
        return None, "El celular de contacto no es válido (9 dígitos)"
    return tel, None


def _validar_hora_programada(cur, programada_para):
    ahora = datetime.now(LIMA_TZ)
    if programada_para < ahora + timedelta(minutes=MAD_PROG_MIN_MIN):
        return f"La hora debe ser al menos {MAD_PROG_MIN_MIN} minutos desde ahora. Si lo necesitas antes, solicita el servicio ahora."
    if programada_para > ahora + timedelta(days=MAD_PROG_MAX_DIAS):
        return f"Solo puedes programar hasta {MAD_PROG_MAX_DIAS} días hacia adelante"
    rangos = _horarios_programacion(cur)
    if rangos and not any(_hora_en_rango(programada_para.time(), ini, fin) for ini, fin in rangos):
        txt = " y ".join(f"{i.strftime('%H:%M')} a {f.strftime('%H:%M')}" for i, f in rangos)
        return f"Solo se puede programar entre las {txt}"
    return None


def _base_url():
    return os.environ.get("BASE_URL", "https://sntspe.com")


def _err(msg, code=400, **extra):
    return (json.dumps({"error": msg, **extra}, ensure_ascii=False), code, HEADERS)


# ═════════════════════════════════════════════════════════════
# DESPACHO — crear solicitud
# ═════════════════════════════════════════════════════════════
ESTADOS_ABIERTOS = ('programada', 'buscando_proveedor', 'asignado', 'en_camino', 'llego', 'en_atencion')


ESTADOS_EN_CURSO = ('buscando_proveedor', 'asignado', 'en_camino', 'llego', 'en_atencion')   # abiertos y no programados


class DniConServicioAbierto(Exception):
    mensaje = "Ya existe un servicio en curso para este DNI. No se puede generar otro hasta que termine o se cancele."

    def __init__(self, solicitud_id, estado):
        super().__init__(f"Servicio abierto #{solicitud_id} ({estado})")
        self.solicitud_id, self.estado = solicitud_id, estado


class CelularConServicioAbierto(DniConServicioAbierto):
    mensaje = "Este celular ya tiene un servicio en curso. No se puede generar otro hasta que termine o se cancele."


MSG_DNI_ABIERTO = DniConServicioAbierto.mensaje


def crear_solicitud(telefono, lat, lng, tipo_servicio, destino_clinica_id=None,
                    direccion=None, distrito=None, referencia=None, cliente_nombre=None,
                    programada_para=None, contacto_nombre=None, contacto_telefono=None,
                    extra=None, notificar_cliente=True):
    """extra (opcional, servicios digitados por el call center): paciente_dni, paciente_edad,
    sintomas, antecedentes, alergias, canal, creado_por_usuario_id."""
    extra = extra or {}
    estado_inicial = 'programada' if programada_para else 'buscando_proveedor'
    conn = get_connection()
    cur = conn.cursor()
    try:
        dni = extra.get("paciente_dni")
        if dni:
            # Un DNI no puede tener dos servicios abiertos a la vez (webchat y call center comparten esta regla).
            # El candado serializa pedidos simultáneos del mismo DNI.
            cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s))", ("mad_dni:" + dni,))
            cur.execute("""SELECT id, estado FROM solicitudes
                           WHERE paciente_dni = %s AND estado IN %s LIMIT 1""", (dni, ESTADOS_ABIERTOS))
            ya = cur.fetchone()
            if ya:
                conn.rollback()
                raise DniConServicioAbierto(ya[0], ya[1])
        if not programada_para:
            # Segunda barrera: un mismo celular no tiene dos servicios en curso a la vez (con o sin DNI).
            # Los servicios programados a futuro no cuentan, para poder reservar y pedir otro ahora.
            cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s))", ("mad_tel:" + str(telefono),))
            cur.execute("""SELECT id, estado FROM solicitudes
                           WHERE cliente_telefono = %s AND estado IN %s LIMIT 1""", (telefono, ESTADOS_EN_CURSO))
            ya = cur.fetchone()
            if ya:
                conn.rollback()
                raise CelularConServicioAbierto(ya[0], ya[1])
        cur.execute("""
            INSERT INTO solicitudes (cliente_telefono, cliente_nombre, tipo_servicio, origen,
                                     origen_direccion_texto, origen_distrito, origen_referencia,
                                     destino_clinica_id, programada_para, contacto_nombre,
                                     contacto_telefono, estado,
                                     paciente_dni, paciente_edad, sintomas, antecedentes, alergias,
                                     canal, creado_por_usuario_id, copago_monto, forma_pago)
            VALUES (%s, %s, %s, ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
                    %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id, tracking_token
        """, (telefono, cliente_nombre, tipo_servicio, lng, lat,
              direccion, distrito, referencia, destino_clinica_id,
              programada_para, contacto_nombre, contacto_telefono, estado_inicial,
              extra.get("paciente_dni"), extra.get("paciente_edad"), extra.get("sintomas"),
              extra.get("antecedentes"), extra.get("alergias"),
              extra.get("canal") or "webchat", extra.get("creado_por_usuario_id"),
              extra.get("copago_monto"), extra.get("forma_pago")))
        solicitud_id, tracking_token = cur.fetchone()
        if extra.get("creado_por_usuario_id"):
            _mad_audit(cur, extra["creado_por_usuario_id"], "servicio_crear", f"solicitud:{solicitud_id}",
                       {"canal": extra.get("canal"), "tipo": tipo_servicio, "distrito": distrito,
                        "copago": extra.get("copago_monto"), "forma_pago": extra.get("forma_pago"),
                        "programada": bool(programada_para)})
        conn.commit()
    finally:
        cur.close(); conn.close()

    link = f"{_base_url()}/track/{tracking_token}"

    if programada_para:
        # No se busca doctor ahora: lo activa _activar_programadas() poco antes de la hora
        numero = _normalizar_celular_pe(telefono) if notificar_cliente else None
        if numero:
            cuando = programada_para.astimezone(LIMA_TZ).strftime("%d/%m %H:%M")
            enviar_danaconnect(numero, f"Sanitas: Servicio programado para {cuando} en {(distrito or '').title()}. Sigue o cancela aqui: {link}")
            sms_aviso(numero, "mad programado")
        return {'solicitud_id': solicitud_id, 'tracking_url': link, 'programada': True,
                'programada_para': programada_para.astimezone(LIMA_TZ).isoformat()}

    buscar_y_ofrecer_proveedor(solicitud_id)
    if extra.get("canal") == "callcenter" and notificar_cliente:
        numero = _normalizar_celular_pe(telefono)
        if numero:
            enviar_danaconnect(numero, f"Sanitas: Recibimos tu solicitud. Sigue a tu doctor aqui: {link}")
            sms_aviso(numero, "mad callcenter")
    return {'solicitud_id': solicitud_id, 'tracking_url': link, 'programada': False}


@con_log_errores
def handle_crear_solicitud(request):
    data = request.get_json(force=True, silent=True) or {}
    telefono           = (data.get("telefono") or "").strip()
    tipo_servicio      = str(data.get("tipo_servicio") or "").strip().upper()
    destino_clinica_id = data.get("destino_clinica_id")
    direccion          = (data.get("direccion") or "").strip()[:200]
    distrito           = (data.get("distrito") or "").strip().upper()[:80]
    referencia         = (data.get("referencia") or "").strip()[:200] or None
    cliente_nombre     = (data.get("cliente_nombre") or "").strip()[:120] or None
    contacto_nombre    = (data.get("contacto_nombre") or "").strip()[:120] or None

    try:
        lat = float(data.get("lat")); lng = float(data.get("lng"))
    except (TypeError, ValueError):
        lat = lng = None

    if not telefono or lat is None or lng is None or tipo_servicio not in ("MAD", "AMBULANCIA"):
        return _err("telefono, lat, lng y tipo_servicio (MAD o AMBULANCIA) son obligatorios")

    error = _validar_lugar(lat, lng, direccion, distrito)
    if error:
        return _err(error)

    contacto_tel, error = _normalizar_contacto(data.get("contacto_telefono"))
    if error:
        return _err(error)

    programada_para = None
    if data.get("programada_para"):
        try:
            programada_para = _parse_programada(data.get("programada_para"))
        except ValueError:
            return _err("Fecha y hora inválidas (formato AAAA-MM-DDTHH:MM, hora de Lima)")

        conn = get_connection(); cur = conn.cursor()
        try:
            error = _validar_hora_programada(cur, programada_para)
            if error:
                return _err(error)
            cur.execute("SELECT COUNT(*) FROM solicitudes WHERE cliente_telefono = %s AND estado = 'programada'", (telefono,))
            if cur.fetchone()[0] >= MAD_PROG_MAX_ACTIVAS:
                return _err(f"Ya tienes {MAD_PROG_MAX_ACTIVAS} servicios programados. Cancela uno para programar otro.", 409)
        finally:
            cur.close(); conn.close()

    dni = re.sub(r"\D", "", str(data.get("paciente_dni") or data.get("dni") or ""))
    if dni and not re.fullmatch(r"\d{8}", dni):
        return _err("El DNI debe tener 8 dígitos")
    extra = {"paciente_dni": dni, "canal": "webchat"} if dni else None
    try:
        resultado = crear_solicitud(telefono, lat, lng, tipo_servicio, destino_clinica_id,
                                    direccion, distrito, referencia, cliente_nombre,
                                    programada_para, contacto_nombre, contacto_tel, extra=extra)
    except DniConServicioAbierto as e:
        return _err(e.mensaje, 409, solicitud_id=e.solicitud_id, estado=e.estado)
    return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_mad_admin_crear_solicitud(request):
    """Coordinador: digita un servicio que llegó por el call center. Entra al mismo despacho
    que los del webchat (los doctores lo reciben como oferta)."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    data = request.get_json(force=True, silent=True) or {}

    def _txt(clave, maximo):
        return re.sub(r"[ \t]+", " ", str(data.get(clave) or "")).strip()[:maximo]

    tipo_servicio = _txt("tipo_servicio", 20).upper()
    if tipo_servicio not in ("MAD", "AMBULANCIA"):
        return _err("Elige el tipo de servicio (médico a domicilio o ambulancia)")

    nombre = _txt("paciente_nombre", 120)
    if len(nombre) < 3:
        return _err("Escribe el nombre del paciente")
    dni = re.sub(r"\D", "", str(data.get("paciente_dni") or ""))
    if not re.fullmatch(r"\d{8}", dni):
        return _err("El DNI debe tener 8 dígitos")
    try:
        edad = int(data.get("paciente_edad"))
    except (TypeError, ValueError):
        return _err("Escribe la edad del paciente")
    if not (0 <= edad <= 120):
        return _err("La edad no es válida")
    cel = re.sub(r"\D", "", str(data.get("paciente_celular") or ""))
    if len(cel) == 9 and cel.startswith("9"):
        cel = "51" + cel
    if not re.fullmatch(r"51\d{9}", cel):
        return _err("El celular del paciente no es válido (9 dígitos, empieza con 9)")

    direccion  = _txt("direccion", 200)
    distrito   = _txt("distrito", 80).upper()
    referencia = _txt("referencia", 200) or None
    sintomas   = _txt("sintomas", 500)
    antecedentes = _txt("antecedentes", 500) or None
    alergias   = _txt("alergias", 300) or None
    if len(sintomas) < 3:
        return _err("Describe los síntomas del paciente")
    # Para no dejar el dato en blanco por descuido: si no hay, se escribe "Ninguna" / "Niega"
    if not antecedentes:
        return _err("Indica los antecedentes médicos (escribe «Ninguno» si no tiene)")
    if not alergias:
        return _err("Indica las alergias a medicamentos (escribe «Niega» si no tiene)")

    try:
        lat = float(data.get("lat")); lng = float(data.get("lng"))
    except (TypeError, ValueError):
        return _err("Marca la ubicación del paciente en el mapa")
    error = _validar_lugar(lat, lng, direccion, distrito)
    if error:
        return _err(error)

    programada_para = None
    if data.get("programada_para"):
        try:
            programada_para = _parse_programada(data.get("programada_para"))
        except ValueError:
            return _err("Fecha y hora inválidas")
        conn = get_connection(); cur = conn.cursor()
        try:
            error = _validar_hora_programada(cur, programada_para)
            if error:
                return _err(error)
        finally:
            cur.close(); conn.close()

    try:
        copago = round(float(str(data.get("copago_monto")).replace(",", ".")), 2)
    except (TypeError, ValueError):
        return _err("Escribe el monto del copago (0 si no corresponde)")
    if not (0 <= copago <= 9999):
        return _err("El monto del copago no es válido")
    forma_pago = _txt("forma_pago", 20)
    if forma_pago not in ("efectivo", "tarjeta", "yape_plin"):
        return _err("Elige la forma de pago (efectivo, tarjeta o Yape/Plin)")

    extra = {"copago_monto": copago, "forma_pago": forma_pago,
             "paciente_dni": dni, "paciente_edad": edad, "sintomas": sintomas,
             "antecedentes": antecedentes, "alergias": alergias,
             "canal": "callcenter", "creado_por_usuario_id": s["id"]}
    try:
        resultado = crear_solicitud(cel, lat, lng, tipo_servicio, None, direccion, distrito, referencia,
                                    nombre, programada_para, None, None,
                                    extra=extra, notificar_cliente=bool(data.get("enviar_sms", True)))
    except DniConServicioAbierto as e:
        return _mad_json({"error": f"{e.mensaje} (servicio #{e.solicitud_id}, estado: {e.estado})"}, 409)
    return _mad_json(resultado, 201)


# ═════════════════════════════════════════════════════════════
# DESPACHO — avisos por SMS
# ═════════════════════════════════════════════════════════════
def enviar_oferta_wsp(telefono_proveedor, oferta_token, tipo_servicio, distrito=None, programada_para=None):
    numero = _normalizar_celular_pe(telefono_proveedor)
    if not numero:
        return
    nombre_servicio = "Medico a domicilio" if tipo_servicio == "MAD" else "Ambulancia"
    zona = f" en {distrito.title()}" if distrito else " cerca de ti"
    cuando = f" PROGRAMADA {programada_para.astimezone(LIMA_TZ).strftime('%d/%m %H:%M')}" if programada_para else ""
    # Con PWA, lo más cómodo es llevar al panel (ya muestra la oferta con Aceptar / Rechazar)
    link = f"{_base_url()}/proveedor" if MAD_OFERTA_LINK_AL_PANEL else f"{_base_url()}/oferta/{oferta_token}"
    enviar_danaconnect(numero, f"Sanitas: Nueva solicitud{cuando} de {nombre_servicio}{zona}. Revisa y responde aqui: {link}")
    sms_aviso(numero, "oferta mad")


# ═════════════════════════════════════════════════════════════
# AVISOS PUSH AL DOCTOR (Web Push / VAPID)
# Tabla: proveedores_push. Variables: VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT (opcional).
# Nunca llevan datos del paciente: solo tipo de servicio, distrito y hora.
# ═════════════════════════════════════════════════════════════
VAPID_PUBLIC_KEY = os.environ.get("VAPID_PUBLIC_KEY", "")
VAPID_PRIVATE_KEY = os.environ.get("VAPID_PRIVATE_KEY", "")
MAX_DISPOSITIVOS_PUSH = 5


def _push_configurado():
    return bool(VAPID_PUBLIC_KEY and VAPID_PRIVATE_KEY)


def enviar_push_proveedor(proveedor_id, titulo, cuerpo, url="/proveedor", tag="mad", ttl=TIMEOUT_OFERTA_SEGUNDOS):
    """Envía el aviso a todos los dispositivos del doctor. Nunca rompe el flujo. Devuelve cuántos recibieron."""
    if not _push_configurado():
        return 0
    try:
        from pywebpush import webpush, WebPushException
    except ImportError:
        logging.warning("[MAD][push] pywebpush no está instalado")
        return 0

    try:
        conn = get_connection(); cur = conn.cursor()
        try:
            cur.execute("SELECT id, endpoint, p256dh, auth FROM proveedores_push WHERE proveedor_id = %s", (proveedor_id,))
            subs = cur.fetchall()
        finally:
            cur.close(); conn.close()
    except Exception as e:
        logging.warning(f"[MAD][push] no se pudo leer suscripciones: {e}")
        return 0

    payload = json.dumps({"titulo": titulo, "cuerpo": cuerpo, "url": url, "tag": tag}, ensure_ascii=False)
    enviados, vencidas, ok_ids = 0, [], []
    for sub_id, endpoint, p256dh, auth in subs:
        try:
            webpush(subscription_info={"endpoint": endpoint, "keys": {"p256dh": p256dh, "auth": auth}},
                    data=payload, vapid_private_key=VAPID_PRIVATE_KEY,
                    vapid_claims={"sub": os.environ.get("VAPID_SUBJECT") or _base_url()},
                    ttl=ttl, headers={"Urgency": "high"}, timeout=6)
            enviados += 1; ok_ids.append(sub_id)
        except WebPushException as e:
            codigo = getattr(getattr(e, "response", None), "status_code", None)
            if codigo in (404, 410):
                vencidas.append(sub_id)        # el dispositivo ya no existe (desinstaló o revocó el permiso)
            else:
                logging.warning(f"[MAD][push] proveedor={proveedor_id} falló ({codigo}): {e}")
        except Exception as e:
            logging.warning(f"[MAD][push] proveedor={proveedor_id} error: {e}")

    if vencidas or ok_ids:
        try:
            conn = get_connection(); cur = conn.cursor()
            try:
                if vencidas:
                    cur.execute("DELETE FROM proveedores_push WHERE id = ANY(%s)", (vencidas,))
                if ok_ids:
                    cur.execute("UPDATE proveedores_push SET ultimo_ok = NOW() WHERE id = ANY(%s)", (ok_ids,))
                conn.commit()
            finally:
                cur.close(); conn.close()
        except Exception as e:
            logging.warning(f"[MAD][push] no se pudo actualizar suscripciones: {e}")
    return enviados


@con_log_errores
def handle_push_clave(request, proveedor_id):
    """Clave pública para suscribirse + si el doctor ya tiene algún dispositivo registrado."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT COUNT(*) FROM proveedores_push WHERE proveedor_id = %s", (proveedor_id,))
        n = cur.fetchone()[0]
    finally:
        cur.close(); conn.close()
    return _mad_json({"public_key": VAPID_PUBLIC_KEY if _push_configurado() else "", "dispositivos": n})


@con_log_errores
def handle_push_suscribir(request, proveedor_id):
    data = request.get_json(force=True, silent=True) or {}
    endpoint = str(data.get("endpoint") or "").strip()
    keys = data.get("keys") or {}
    p256dh, auth = str(keys.get("p256dh") or ""), str(keys.get("auth") or "")
    if not endpoint.startswith("https://") or len(endpoint) > 1000 or not p256dh or not auth \
            or len(p256dh) > 200 or len(auth) > 100:
        return _mad_json({"error": "Suscripción inválida"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        # Un mismo dispositivo (endpoint) pertenece a un solo doctor: si otro inicia sesión en él, se reasigna
        cur.execute("""
            INSERT INTO proveedores_push (proveedor_id, endpoint, p256dh, auth, user_agent)
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (endpoint) DO UPDATE
               SET proveedor_id = EXCLUDED.proveedor_id, p256dh = EXCLUDED.p256dh,
                   auth = EXCLUDED.auth, user_agent = EXCLUDED.user_agent
        """, (proveedor_id, endpoint, p256dh, auth, (request.headers.get("User-Agent") or "")[:200]))
        # Máximo de dispositivos por doctor: se conservan los más recientes
        cur.execute("""
            DELETE FROM proveedores_push WHERE proveedor_id = %s AND id NOT IN (
                SELECT id FROM proveedores_push WHERE proveedor_id = %s ORDER BY id DESC LIMIT %s)
        """, (proveedor_id, proveedor_id, MAX_DISPOSITIVOS_PUSH))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"status": "ok"})


@con_log_errores
def handle_push_baja(request, proveedor_id):
    data = request.get_json(force=True, silent=True) or {}
    endpoint = str(data.get("endpoint") or "").strip()
    if not endpoint:
        return _mad_json({"error": "endpoint requerido"}, 400)
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("DELETE FROM proveedores_push WHERE proveedor_id = %s AND endpoint = %s", (proveedor_id, endpoint))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"status": "ok"})


@con_log_errores
def handle_push_probar(request, proveedor_id):
    if not _push_configurado():
        return _mad_json({"error": "Los avisos aún no están configurados en el servidor"}, 503)
    n = enviar_push_proveedor(proveedor_id, "Prueba de aviso", "Así te avisaremos cuando llegue una solicitud.", tag="prueba", ttl=60)
    if not n:
        return _mad_json({"error": "No se pudo enviar. Vuelve a activar los avisos en este dispositivo."}, 502)
    return _mad_json({"status": "ok", "dispositivos": n})


def notificar_cliente_proveedor_asignado(solicitud_id):
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT s.cliente_telefono, s.tipo_servicio, s.tracking_token, p.nombre, s.programada_para
            FROM solicitudes s JOIN proveedores p ON p.id = s.proveedor_id
            WHERE s.id = %s
        """, (solicitud_id,))
        row = cur.fetchone()
    finally:
        cur.close(); conn.close()
    if not row:
        return

    telefono_cliente, tipo_servicio, tracking_token, nombre_proveedor, programada_para = row
    numero = _normalizar_celular_pe(telefono_cliente)
    if not numero:
        return
    link = f"{_base_url()}/track/{tracking_token}"

    if programada_para:
        cuando = programada_para.astimezone(LIMA_TZ).strftime("%d/%m %H:%M")
        mensaje = f"Sanitas: {nombre_proveedor} atendera tu servicio programado ({cuando}). Sigue el recorrido: {link}"
    else:
        referencia = "tu médico" if tipo_servicio == "MAD" else "tu ambulancia"
        mensaje = f"Sanitas: {nombre_proveedor} viene en camino, sera {referencia}. Sigue el recorrido: {link}"
    enviar_danaconnect(numero, mensaje)
    sms_aviso(numero, "mad proveedor asignado")


def notificar_cliente_proveedor_en_camino(solicitud_id):
    # El cliente ya ve el cambio de estado en /track/<token>; se deja como gancho opcional.
    pass


def notificar_cliente_sin_cobertura(solicitud_id):
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT cliente_telefono, cliente_nombre, tipo_servicio, origen_direccion_texto, origen_distrito
            FROM solicitudes WHERE id = %s
        """, (solicitud_id,))
        row = cur.fetchone()
    finally:
        cur.close(); conn.close()
    if not row:
        return

    telefono, nombre, tipo_servicio, direccion, distrito = row
    nombre_servicio = "medico a domicilio" if tipo_servicio == "MAD" else "ambulancia"

    # 1) Al cliente
    numero = _normalizar_celular_pe(telefono)
    if numero:
        mensaje = (f"Sanitas: Lo sentimos, en este momento no encontramos {nombre_servicio} "
                   f"disponible cerca. Nuestro equipo te contactara. Si es una emergencia llama al 219-1919.")
        enviar_danaconnect(numero, mensaje)
        sms_aviso(numero, "mad sin cobertura")

    # 2) Al equipo de coordinación
    for cel in MAD_ALERTA_CELULARES:
        alerta = (f"ALERTA MAD: solicitud #{solicitud_id} SIN COBERTURA. {nombre_servicio}, "
                  f"{distrito or ''}. Cliente {nombre or ''} {telefono}")
        enviar_danaconnect(_normalizar_celular_pe(cel) or cel, alerta[:160])


# ═════════════════════════════════════════════════════════════
# DESPACHO — tiempo estimado de llegada (calibrado)
# ═════════════════════════════════════════════════════════════
def _calcular_eta_min(cur, distancia_m, estado):
    """ETA en minutos usando la velocidad calibrada para la hora actual de Lima.
    Para cambiar luego a la API de rutas de Google, solo se reemplaza esta función."""
    if distancia_m is None:
        return None
    cur.execute("SELECT kmh FROM mad_velocidad_hora WHERE hora = %s", (datetime.now(LIMA_TZ).hour,))
    r = cur.fetchone()
    kmh = float(r[0]) if r and r[0] else VELOCIDAD_DEFAULT_KMH
    minutos = (float(distancia_m) / 1000.0) / kmh * 60
    if estado == "asignado":
        minutos += MINUTOS_PREPARACION
    return max(1, round(minutos))


@con_log_errores
def handle_recalibrar_eta(request):
    """Velocidad efectiva (línea recta / tiempo real) por hora del día, mediana de 60 días."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            WITH viajes AS (
                SELECT EXTRACT(HOUR FROM en_camino_en AT TIME ZONE 'America/Lima')::int AS hora,
                       ST_Distance(en_camino_ubicacion, origen) / 1000.0            AS km,
                       EXTRACT(EPOCH FROM (llego_en - en_camino_en)) / 3600.0        AS horas
                FROM solicitudes
                WHERE en_camino_en IS NOT NULL AND llego_en IS NOT NULL
                  AND en_camino_ubicacion IS NOT NULL
                  AND en_camino_en > NOW() - INTERVAL '60 days'
            ),
            validos AS (
                SELECT hora, km / horas AS kmh
                FROM viajes
                WHERE km >= 0.3
                  AND horas BETWEEN (2.0 / 60) AND 2
                  AND km / horas BETWEEN 3 AND 80
            )
            SELECT hora,
                   PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY kmh) AS kmh_mediana,
                   COUNT(*) AS muestras
            FROM validos
            GROUP BY hora
        """)
        filas = cur.fetchall()

        actualizadas = []
        for hora, kmh, muestras in filas:
            if muestras >= MIN_MUESTRAS_CALIBRACION:
                cur.execute("""
                    UPDATE mad_velocidad_hora
                    SET kmh = %s, muestras = %s, actualizado_en = NOW()
                    WHERE hora = %s
                """, (round(float(kmh), 1), muestras, hora))
                actualizadas.append({"hora": hora, "kmh": round(float(kmh), 1), "muestras": muestras})
        conn.commit()
    finally:
        cur.close(); conn.close()

    return (json.dumps({"horas_actualizadas": actualizadas}, ensure_ascii=False), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# DESPACHO — búsqueda y ofertas
# ═════════════════════════════════════════════════════════════
def buscar_y_ofrecer_proveedor(solicitud_id, radio_metros=None):
    """Ofrece la solicitud a la EMPRESA que corresponde según la tabla de pesos (una tabla para médico a
    domicilio y otra para ambulancia). La empresa acepta y luego asigna a uno de sus doctores.
    Solo participan empresas con al menos un doctor conectado (señal reciente) dentro de RADIO_MAXIMO_METROS
    y que no hayan rechazado o dejado vencer esta solicitud. Solo actúa si la solicitud sigue en
    'buscando_proveedor', sin oferta pendiente y sin empresa que ya la aceptó. (radio_metros ya no se usa.)"""
    ahora = datetime.now(LIMA_TZ).replace(tzinfo=None)
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT origen, tipo_servicio, origen_distrito, programada_para, estado, empresa_aceptada_en
            FROM solicitudes WHERE id = %s
            FOR UPDATE
        """, (solicitud_id,))
        row_sol = cur.fetchone()
        if not row_sol:
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Solicitud no encontrada'}
        origen, tipo_servicio, distrito, programada_para, estado, empresa_aceptada_en = row_sol

        if estado != 'buscando_proveedor':
            conn.rollback()
            return {'status': 'no_aplica', 'estado': estado}
        if empresa_aceptada_en is not None:
            conn.rollback()
            return {'status': 'esperando_asignacion'}

        cur.execute("""
            SELECT 1 FROM solicitudes_ofertas_empresa WHERE solicitud_id = %s AND estado = 'pendiente'
            UNION ALL
            SELECT 1 FROM solicitudes_ofertas WHERE solicitud_id = %s AND estado = 'pendiente'
            LIMIT 1
        """, (solicitud_id, solicitud_id))
        if cur.fetchone():
            conn.rollback()
            return {'status': 'ya_ofrecido'}

        cur.execute("UPDATE solicitudes SET ultima_busqueda_en = %s WHERE id = %s", (ahora, solicitud_id))

        cur.execute(f"""
            SELECT e.id, e.nombre, e.celular_alertas, pr.peso,
                   MIN(ST_Distance(pe.ubicacion, %s::geography)) AS distancia
            FROM mad_empresas e
            JOIN mad_empresas_prioridad pr ON pr.empresa_id = e.id AND pr.tipo_servicio = %s AND pr.peso > 0
            JOIN proveedores p ON p.empresa_id = e.id AND p.tipo_servicio IN (%s, 'AMBOS')
            JOIN mad_usuarios u ON u.proveedor_id = p.id AND u.activo
            JOIN proveedores_estado pe ON pe.proveedor_id = p.id
            WHERE e.activo
              AND pe.estado = 'disponible' AND pe.ubicacion IS NOT NULL
              AND pe.actualizado_en > NOW() - INTERVAL '{MAD_AUTODESCONEXION_MIN} minutes'
              AND ST_DWithin(pe.ubicacion, %s::geography, {RADIO_MAXIMO_METROS})
              AND p.id NOT IN (SELECT proveedor_id FROM solicitudes_proveedores_excluidos WHERE solicitud_id = %s)
              AND e.id NOT IN (SELECT empresa_id FROM solicitudes_ofertas_empresa
                               WHERE solicitud_id = %s AND estado IN ('rechazada', 'expirada'))
            GROUP BY e.id, e.nombre, e.celular_alertas, pr.peso
            ORDER BY pr.peso DESC, distancia ASC
            LIMIT 1
        """, (origen, tipo_servicio, tipo_servicio, origen, solicitud_id, solicitud_id))
        row = cur.fetchone()

        if not row:
            conn.commit()   # guarda ultima_busqueda_en; el reintento lo hace revisar-timeouts
            return {'status': 'sin_proveedores'}

        empresa_id, empresa_nombre, celular_alertas, peso, distancia = row
        cur.execute("""
            INSERT INTO solicitudes_ofertas_empresa (solicitud_id, empresa_id, estado)
            VALUES (%s, %s, 'pendiente') RETURNING id
        """, (solicitud_id, empresa_id))
        oferta_emp_id = cur.fetchone()[0]
        conn.commit()
    finally:
        cur.close(); conn.close()

    avisar_empresa_oferta(celular_alertas, tipo_servicio, distrito, programada_para)
    return {'status': 'ofrecido', 'oferta_empresa_id': oferta_emp_id, 'empresa_id': empresa_id}


# ═════════════════════════════════════════════════════════════
# OFERTAS A EMPRESAS: la empresa acepta y asigna a uno de sus doctores
# Tablas: solicitudes_ofertas_empresa, mad_empresas_prioridad (pesos), mad_empresas.celular_alertas
# ═════════════════════════════════════════════════════════════
MAD_OFERTA_EMPRESA_SEG = int(os.environ.get("MAD_OFERTA_EMPRESA_SEG", "120"))   # plazo para aceptar o rechazar
MAD_ASIGNAR_MIN = int(os.environ.get("MAD_ASIGNAR_MIN", "5"))                   # plazo para asignar doctor tras aceptar


def avisar_empresa_oferta(celular, tipo_servicio, distrito, programada_para=None):
    numero = _normalizar_celular_pe(celular) if celular else None
    if not numero:
        return
    nombre_serv = "Medico a domicilio" if tipo_servicio == "MAD" else "Ambulancia"
    zona = f" en {distrito.title()}" if distrito else ""
    cuando = f" PROGRAMADA {programada_para.astimezone(LIMA_TZ).strftime('%d/%m %H:%M')}" if programada_para else ""
    enviar_danaconnect(numero, f"Sanitas: Nueva solicitud{cuando} de {nombre_serv}{zona}. Ingresa a tu portal para aceptarla: {_base_url()}/mad/empresa")
    sms_aviso(numero, "oferta empresa")


def aceptar_oferta_empresa(oferta_id, empresa_id, usuario_id):
    """La empresa se compromete con el servicio; falta que asigne un doctor (plazo MAD_ASIGNAR_MIN)."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT solicitud_id FROM solicitudes_ofertas_empresa WHERE id = %s AND empresa_id = %s",
                    (oferta_id, empresa_id))
        r = cur.fetchone()
        if not r:
            return {'status': 'error', 'mensaje': 'Oferta no encontrada'}
        solicitud_id = r[0]
        cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
        if cur.fetchone()[0] != 'buscando_proveedor':
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Esta solicitud ya no está disponible'}
        cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'aceptada', respondido_en = NOW(), usuario_id = %s
                       WHERE id = %s AND estado = 'pendiente' RETURNING id""", (usuario_id, oferta_id))
        if not cur.fetchone():
            conn.rollback()
            return {'status': 'error', 'mensaje': 'La oferta venció o ya fue respondida'}
        cur.execute("UPDATE solicitudes SET empresa_id = %s, empresa_aceptada_en = NOW() WHERE id = %s",
                    (empresa_id, solicitud_id))
        cur.execute("""INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                       VALUES (%s, 'buscando_proveedor', 'buscando_proveedor', 'empresa_acepto')""", (solicitud_id,))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return {'status': 'ok', 'solicitud_id': solicitud_id}


def rechazar_oferta_empresa(oferta_id, empresa_id, usuario_id):
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT solicitud_id FROM solicitudes_ofertas_empresa WHERE id = %s AND empresa_id = %s",
                    (oferta_id, empresa_id))
        r = cur.fetchone()
        if not r:
            return {'status': 'error', 'mensaje': 'Oferta no encontrada'}
        solicitud_id = r[0]
        cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
        estado_sol = cur.fetchone()[0]
        cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'rechazada', respondido_en = NOW(), usuario_id = %s
                       WHERE id = %s AND estado = 'pendiente' RETURNING id""", (usuario_id, oferta_id))
        if not cur.fetchone():
            conn.rollback()
            return {'status': 'error', 'mensaje': 'La oferta venció o ya fue respondida'}
        conn.commit()
    finally:
        cur.close(); conn.close()
    resultado = buscar_y_ofrecer_proveedor(solicitud_id) if estado_sol == 'buscando_proveedor' else None
    return {'status': 'rechazada', 'siguiente': resultado}


def asignar_doctor_empresa(solicitud_id, proveedor_id, empresa_id, usuario_id):
    """La empresa asigna uno de sus doctores disponibles. El doctor recibe el servicio ya asignado."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""SELECT estado, empresa_id, empresa_aceptada_en IS NOT NULL, tipo_servicio, origen_distrito, programada_para
                       FROM solicitudes WHERE id = %s FOR UPDATE""", (solicitud_id,))
        r = cur.fetchone()
        if not r:
            return {'status': 'error', 'mensaje': 'Solicitud no encontrada'}
        estado, emp, aceptada, tipo, distrito, programada_para = r
        if estado != 'buscando_proveedor' or emp != empresa_id or not aceptada:
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Esta solicitud ya no está pendiente de asignación en tu empresa (puede haber vencido el plazo)'}

        cur.execute("""SELECT p.tipo_servicio, p.telefono, pe.estado
                       FROM proveedores p JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                       WHERE p.id = %s AND p.empresa_id = %s FOR UPDATE OF pe""", (proveedor_id, empresa_id))
        d = cur.fetchone()
        if not d:
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Ese doctor no pertenece a tu empresa'}
        tipo_doc, tel_doc, estado_doc = d
        if tipo_doc not in (tipo, 'AMBOS'):
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Ese doctor no atiende este tipo de servicio'}
        if estado_doc not in ('disponible', 'desconectado'):
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Ese doctor está atendiendo u ocupado con otra solicitud'}

        cur.execute("""UPDATE solicitudes
                       SET estado = 'asignado', proveedor_id = %s, asignado_en = NOW(), empresa_aceptada_en = NULL
                       WHERE id = %s""", (proveedor_id, solicitud_id))
        cur.execute("UPDATE proveedores_estado SET estado = 'ocupado' WHERE proveedor_id = %s", (proveedor_id,))
        cur.execute("""INSERT INTO solicitudes_ofertas (solicitud_id, proveedor_id, estado, respondido_en)
                       VALUES (%s, %s, 'aceptada', NOW())""", (solicitud_id, proveedor_id))
        cur.execute("""INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                       VALUES (%s, 'buscando_proveedor', 'asignado', 'empresa')""", (solicitud_id,))
        _mad_audit(cur, usuario_id, "empresa_asigna", f"solicitud:{solicitud_id}", {"proveedor_id": proveedor_id})
        conn.commit()
    finally:
        cur.close(); conn.close()

    notificar_cliente_proveedor_asignado(solicitud_id)
    nombre_serv = "Médico a domicilio" if tipo == "MAD" else "Ambulancia"
    enviar_push_proveedor(proveedor_id, f"Te asignaron un servicio: {nombre_serv}",
                          f"{(distrito or 'Zona no especificada').title()}. Abre la app para ver los datos.", tag="asignado")
    numero = _normalizar_celular_pe(tel_doc)
    if numero:
        enviar_danaconnect(numero, f"Sanitas: Te asignaron un servicio de {nombre_serv.replace('é','e')} en {(distrito or '').title()}. Abre la app: {_base_url()}/proveedor")
        sms_aviso(numero, "servicio asignado")
    return {'status': 'ok', 'solicitud_id': solicitud_id}


def _vencer_ofertas_empresa():
    """(1) Ofertas a empresa sin respuesta en MAD_OFERTA_EMPRESA_SEG → pasan a la siguiente empresa.
    (2) Servicios aceptados por una empresa que no asignó doctor en MAD_ASIGNAR_MIN → vuelven a la búsqueda."""
    reofrecer = []
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute(f"""
            SELECT o.id, o.solicitud_id FROM solicitudes_ofertas_empresa o
            JOIN solicitudes s ON s.id = o.solicitud_id
            WHERE o.estado = 'pendiente' AND o.ofrecido_en < NOW() - INTERVAL '{MAD_OFERTA_EMPRESA_SEG} seconds'
        """)
        for oid, sid in cur.fetchall():
            cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (sid,))
            est = cur.fetchone()
            cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'expirada', respondido_en = NOW()
                           WHERE id = %s AND estado = 'pendiente' RETURNING id""", (oid,))
            if cur.fetchone() and est and est[0] == 'buscando_proveedor':
                reofrecer.append(sid)

        cur.execute(f"""
            SELECT s.id, s.empresa_id FROM solicitudes s
            WHERE s.estado = 'buscando_proveedor' AND s.empresa_aceptada_en IS NOT NULL
              AND s.empresa_aceptada_en < NOW() - INTERVAL '{MAD_ASIGNAR_MIN} minutes'
        """)
        for sid, emp in cur.fetchall():
            cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (sid,))
            est = cur.fetchone()
            if not est or est[0] != 'buscando_proveedor':
                continue
            # La oferta aceptada pasa a 'expirada' para que esa empresa no vuelva a recibir este servicio
            cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'expirada'
                           WHERE solicitud_id = %s AND empresa_id = %s AND estado = 'aceptada'""", (sid, emp))
            cur.execute("UPDATE solicitudes SET empresa_aceptada_en = NULL, empresa_id = NULL WHERE id = %s", (sid,))
            cur.execute("""INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                           VALUES (%s, 'buscando_proveedor', 'buscando_proveedor', 'sistema_sin_asignar')""", (sid,))
            reofrecer.append(sid)
        conn.commit()
    finally:
        cur.close(); conn.close()
    resultados = []
    for sid in reofrecer:
        resultados.append({"solicitud_id": sid, "resultado": buscar_y_ofrecer_proveedor(sid)})
    return resultados


def aceptar_oferta(oferta_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT solicitud_id FROM solicitudes_ofertas WHERE id = %s", (oferta_id,))
        r = cur.fetchone()
        if not r:
            return {'status': 'error', 'mensaje': 'Oferta no encontrada'}
        solicitud_id = r[0]

        # Orden de bloqueo: solicitud → oferta → doctor (igual que cancelar, reprogramar y liberar)
        cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
        estado_sol = cur.fetchone()[0]
        if estado_sol != 'buscando_proveedor':
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Oferta ya no disponible'}

        # UPDATE atómico condicionado: solo pasa si sigue 'pendiente'
        cur.execute("""
            UPDATE solicitudes_ofertas
            SET estado = 'aceptada', respondido_en = %s
            WHERE id = %s AND estado = 'pendiente'
            RETURNING proveedor_id
        """, (datetime.now(LIMA_TZ), oferta_id))
        row = cur.fetchone()
        if not row:
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Oferta ya no disponible'}
        proveedor_id = row[0]

        # La empresa del doctor se guarda AHORA: el historial por empresa queda exacto
        cur.execute("""
            UPDATE solicitudes
            SET estado = 'asignado', proveedor_id = %s, asignado_en = %s,
                empresa_id = (SELECT empresa_id FROM proveedores WHERE id = %s)
            WHERE id = %s
        """, (proveedor_id, datetime.now(LIMA_TZ), proveedor_id, solicitud_id))

        cur.execute("UPDATE proveedores_estado SET estado = 'ocupado' WHERE proveedor_id = %s", (proveedor_id,))

        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, 'buscando_proveedor', 'asignado', 'proveedor')
        """, (solicitud_id,))
        conn.commit()
    finally:
        cur.close(); conn.close()

    notificar_cliente_proveedor_asignado(solicitud_id)
    return {'status': 'ok', 'solicitud_id': solicitud_id}


def rechazar_oferta(oferta_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT solicitud_id FROM solicitudes_ofertas WHERE id = %s", (oferta_id,))
        r = cur.fetchone()
        if not r:
            return {'status': 'error', 'mensaje': 'Oferta no encontrada'}
        solicitud_id = r[0]

        cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
        estado_sol = cur.fetchone()[0]

        cur.execute("""
            UPDATE solicitudes_ofertas
            SET estado = 'rechazada', respondido_en = %s
            WHERE id = %s AND estado = 'pendiente'
            RETURNING proveedor_id
        """, (datetime.now(LIMA_TZ), oferta_id))
        row = cur.fetchone()
        if not row:
            conn.rollback()
            return {'status': 'error', 'mensaje': 'Oferta ya no disponible'}
        proveedor_id = row[0]

        cur.execute("""
            UPDATE proveedores_estado SET estado = 'disponible'
            WHERE proveedor_id = %s AND estado = 'en_oferta'
        """, (proveedor_id,))
        conn.commit()
    finally:
        cur.close(); conn.close()

    resultado = buscar_y_ofrecer_proveedor(solicitud_id) if estado_sol == 'buscando_proveedor' else None
    return {'status': 'rechazada', 'reasignacion': resultado}


def _oferta_id_por_token(token):
    if not _uuid_valido(token):
        return None
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT id FROM solicitudes_ofertas WHERE oferta_token = %s", (token,))
        r = cur.fetchone()
    finally:
        cur.close(); conn.close()
    return r[0] if r else None


def _obtener_oferta_por_token(cur, token):
    cur.execute("""
        SELECT o.id, o.estado, s.tipo_servicio, s.origen_distrito,
               ST_Y(s.origen::geometry) AS lat, ST_X(s.origen::geometry) AS lng,
               ST_Distance(pe.ubicacion, s.origen) AS distancia_metros, s.programada_para
        FROM solicitudes_ofertas o
        JOIN solicitudes s ON s.id = o.solicitud_id
        JOIN proveedores_estado pe ON pe.proveedor_id = o.proveedor_id
        WHERE o.oferta_token = %s
    """, (token,))
    return cur.fetchone()


@con_log_errores
def handle_estado_oferta_token(request, token):
    if not _uuid_valido(token):
        return _err("Oferta no encontrada", 404)
    conn = get_connection(); cur = conn.cursor()
    try:
        row = _obtener_oferta_por_token(cur, token)
    finally:
        cur.close(); conn.close()
    if not row:
        return _err("Oferta no encontrada", 404)

    oferta_id, estado, tipo_servicio, distrito, lat, lng, distancia, programada_para = row
    # Antes de aceptar solo se muestra distrito, distancia y hora: nunca coordenadas ni dirección
    return (json.dumps({
        "estado": estado,
        "tipo_servicio": tipo_servicio,
        "distrito": distrito,
        "distancia_metros": float(distancia) if distancia is not None else None,
        "programada_para": programada_para.isoformat() if programada_para else None,
    }, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_aceptar_oferta_token(request, token):
    oferta_id = _oferta_id_por_token(token)
    if not oferta_id:
        return _err("Oferta no encontrada", 404)
    resultado = aceptar_oferta(oferta_id)
    return (json.dumps(resultado, ensure_ascii=False), 200 if resultado.get("status") == "ok" else 409, HEADERS)


@con_log_errores
def handle_rechazar_oferta_token(request, token):
    oferta_id = _oferta_id_por_token(token)
    if not oferta_id:
        return _err("Oferta no encontrada", 404)
    resultado = rechazar_oferta(oferta_id)
    return (json.dumps(resultado, ensure_ascii=False), 200 if resultado.get("status") == "rechazada" else 409, HEADERS)


@con_log_errores
def handle_aceptar_oferta(request, oferta_id):
    try:
        oferta_id = int(oferta_id)
    except ValueError:
        return _err("oferta_id inválido")
    resultado = aceptar_oferta(oferta_id)
    return (json.dumps(resultado, ensure_ascii=False), 200 if resultado.get("status") == "ok" else 409, HEADERS)


@con_log_errores
def handle_rechazar_oferta(request, oferta_id):
    try:
        oferta_id = int(oferta_id)
    except ValueError:
        return _err("oferta_id inválido")
    resultado = rechazar_oferta(oferta_id)
    return (json.dumps(resultado, ensure_ascii=False), 200 if resultado.get("status") == "rechazada" else 409, HEADERS)


@con_log_errores
def handle_buscar_proveedor(request, solicitud_id):
    try:
        solicitud_id = int(solicitud_id)
    except ValueError:
        return _err("solicitud_id inválido")
    resultado = buscar_y_ofrecer_proveedor(solicitud_id)
    return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# PÁGINAS HTML
# ═════════════════════════════════════════════════════════════
def _servir_html(nombre):
    with open(os.path.join(BASE_DIR, "static", nombre), "r", encoding="utf-8") as f:
        html = f.read().replace("__BASE_URL__", _base_url())
    return (html, 200, {"Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-cache"})


def handle_servir_tracking(request, tracking_token=None):
    return _servir_html("track.html")


# ═════════════════════════════════════════════════════════════
# UBICACIÓN CONFIRMADA POR EL CLIENTE
# El monitor pide la ubicación (SMS con enlace), el cliente marca su domicilio en un mapa
# y el monitor la recibe para usarla en el servicio. Tabla: mad_ubicaciones_pedidas.
# Rutas públicas bajo /track/u/<token> (el token funciona como llave, igual que /track/).
# ═════════════════════════════════════════════════════════════
MAD_UBICACION_VIGENCIA_MIN = 60
MAD_UBICACION_MAX_PEDIDOS_HORA = 4      # por celular
MAD_UBICACION_MAX_ENVIOS = 5            # veces que el cliente puede volver a marcar


@con_log_errores
def handle_mad_admin_pedir_ubicacion(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    data = request.get_json(force=True, silent=True) or {}
    tel = re.sub(r"\D", "", str(data.get("telefono") or ""))
    if len(tel) == 9 and tel.startswith("9"):
        tel = "51" + tel
    if not re.fullmatch(r"51\d{9}", tel):
        return _mad_json({"error": "Escribe el celular del cliente (9 dígitos, empieza con 9)"}, 400)
    direccion = re.sub(r"\s+", " ", str(data.get("direccion") or "")).strip()[:200] or None
    try:
        clat, clng = float(data.get("centro_lat")), float(data.get("centro_lng"))
        if not (PERU_LAT[0] <= clat <= PERU_LAT[1] and PERU_LNG[0] <= clng <= PERU_LNG[1]):
            clat = clng = None
    except (TypeError, ValueError):
        clat = clng = None

    token = str(uuid.uuid4())
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""SELECT COUNT(*) FROM mad_ubicaciones_pedidas
                       WHERE telefono = %s AND creado_en > NOW() - INTERVAL '1 hour'""", (tel,))
        if cur.fetchone()[0] >= MAD_UBICACION_MAX_PEDIDOS_HORA:
            return _mad_json({"error": "Ya se envió el enlace varias veces a este celular. Espera un momento."}, 429)
        cur.execute(f"""
            INSERT INTO mad_ubicaciones_pedidas (token, telefono, creado_por, expira_en, direccion_texto, centro_lat, centro_lng)
            VALUES (%s, %s, %s, NOW() + INTERVAL '{MAD_UBICACION_VIGENCIA_MIN} minutes', %s, %s, %s)
            RETURNING id
        """, (token, tel, s["id"], direccion, clat, clng))
        pedido_id = cur.fetchone()[0]
        _mad_audit(cur, s["id"], "ubicacion_pedir", f"ubicacion:{pedido_id}", {"destino_final": tel[-3:]})
        conn.commit()
    finally:
        cur.close(); conn.close()

    numero = _normalizar_celular_pe(tel)
    enviar_danaconnect(numero, f"Sanitas: Para enviarte al medico, marca en el mapa donde estas (toma 1 minuto): {_base_url()}/track/u/{token}")
    sms_aviso(numero, "mad pedir ubicacion")
    return _mad_json({"ok": True, "pedido_id": pedido_id, "vigencia_min": MAD_UBICACION_VIGENCIA_MIN}, 201)


@con_log_errores
def handle_mad_admin_estado_ubicacion(request, pedido_id):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(pedido_id).isdigit():
        return _mad_json({"error": "Pedido inválido"}, 400)
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""SELECT lat, lng, precision_m, respondida_en, expira_en < NOW()
                       FROM mad_ubicaciones_pedidas WHERE id = %s""", (int(pedido_id),))
        row = cur.fetchone()
    finally:
        cur.close(); conn.close()
    if not row:
        return _mad_json({"error": "Pedido no encontrado"}, 404)
    lat, lng, prec, resp, expirada = row
    if lat is not None:
        return _mad_json({"estado": "respondida", "lat": lat, "lng": lng,
                          "precision_m": round(prec) if prec is not None else None, "respondida_en": _iso(resp)})
    return _mad_json({"estado": "expirada" if expirada else "pendiente"})


def _ubicacion_por_token(cur, token):
    if not _uuid_valido(token):
        return None
    cur.execute("""SELECT id, direccion_texto, centro_lat, centro_lng, intentos, expira_en < NOW(), lat IS NOT NULL
                   FROM mad_ubicaciones_pedidas WHERE token = %s""", (token,))
    return cur.fetchone()


def handle_ubicacion_info(token):
    conn = get_connection(); cur = conn.cursor()
    try:
        row = _ubicacion_por_token(cur, token)
    finally:
        cur.close(); conn.close()
    if not row:
        return _err("Enlace inválido", 404)
    _id, direccion, clat, clng, intentos, expirada, ya = row
    return (json.dumps({"direccion": direccion, "centro": {"lat": clat, "lng": clng} if clat is not None else None,
                        "expirada": bool(expirada), "ya_enviada": bool(ya)}, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_ubicacion_enviar(request, token):
    data = request.get_json(force=True, silent=True) or {}
    try:
        lat, lng = float(data.get("lat")), float(data.get("lng"))
    except (TypeError, ValueError):
        return _err("Marca tu ubicación en el mapa")
    if not (PERU_LAT[0] <= lat <= PERU_LAT[1] and PERU_LNG[0] <= lng <= PERU_LNG[1]):
        return _err("La ubicación debe estar dentro de Perú")
    try:
        prec = float(data.get("precision_m"))
        prec = prec if 0 <= prec <= 100000 else None
    except (TypeError, ValueError):
        prec = None

    conn = get_connection(); cur = conn.cursor()
    try:
        if not _uuid_valido(token):
            return _err("Enlace inválido", 404)
        cur.execute("SELECT id, intentos, expira_en < NOW() FROM mad_ubicaciones_pedidas WHERE token = %s FOR UPDATE", (token,))
        row = cur.fetchone()
        if not row:
            return _err("Enlace inválido", 404)
        if row[2]:
            return _err("Este enlace venció. Pide a Sanitas que te envíe uno nuevo.", 410)
        if row[1] >= MAD_UBICACION_MAX_ENVIOS:
            return _err("Ya enviaste tu ubicación varias veces. Llama al 219-1919 si necesitas corregirla.", 429)
        cur.execute("""UPDATE mad_ubicaciones_pedidas
                       SET lat = %s, lng = %s, precision_m = %s, respondida_en = NOW(), intentos = intentos + 1
                       WHERE id = %s""", (lat, lng, prec, row[0]))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return (json.dumps({"ok": True}), 200, HEADERS)


def handle_servir_ubicacion(request, token=None):
    return _servir_html("ubicacion.html")


def handle_servir_proveedor(request):
    return _servir_html("proveedor.html")


def handle_servir_oferta(request, token=None):
    return _servir_html("oferta.html")


def handle_mad_monitor(request):
    # La página es pública; los DATOS exigen sesión de coordinador
    return _servir_html("mad_monitor.html")


def handle_mad_empresa_pagina(request):
    return _servir_html("mad_empresa.html")


def handle_pwa_manifest(request):
    manifest = {
        "name": "Sanitas Proveedor", "short_name": "Sanitas MAD",
        "start_url": "/proveedor", "scope": "/",
        "display": "standalone", "orientation": "portrait",
        "background_color": "#003876", "theme_color": "#003876",
        "icons": [
            {"src": "/static/pwa/proveedor-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/static/pwa/proveedor-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"},
        ],
    }
    return (json.dumps(manifest, ensure_ascii=False), 200,
            {"Content-Type": "application/manifest+json; charset=utf-8", "Cache-Control": "no-cache"})


def handle_pwa_sw(request):
    with open(os.path.join(BASE_DIR, "static", "proveedor-sw.js"), encoding="utf-8") as f:
        return (f.read(), 200, {"Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "no-cache"})


# ═════════════════════════════════════════════════════════════
# SEGUIMIENTO DEL CLIENTE (/track)
# ═════════════════════════════════════════════════════════════
ESTADOS_NOVEDAD = ("asignado", "en_camino", "llego")
MAD_NOVEDADES_MAX_POR_SERVICIO = 30
MAD_NOVEDAD_INTERVALO_SEG = 20

# Mensajes predefinidos que el doctor puede enviar: texto pensado para el cliente.
NOVEDADES_PRESET = {
    "trafico":   ("🚦", "Hay tráfico en la ruta; podría demorar un poco más."),
    "problema":  ("⚠️", "Tengo dificultades para llegar al punto; estoy resolviéndolo."),
    "cerca":     ("📍", "Estoy cerca de tu ubicación."),
    "direccion": ("🏠", "No encuentro la dirección; por favor mantente atento al teléfono."),
    "otro":      ("💬", None),
}


def _listar_novedades(cur, solicitud_id, limite=8):
    cur.execute("""
        SELECT tipo, texto, creado_en FROM solicitudes_novedades
        WHERE solicitud_id = %s ORDER BY id DESC LIMIT %s
    """, (solicitud_id, limite))
    return [{"tipo": t, "icono": NOVEDADES_PRESET.get(t, ("💬", None))[0], "texto": x, "creado_en": _iso(c)}
            for t, x, c in cur.fetchall()]


@con_log_errores
def handle_novedad_doctor(request, solicitud_id):
    """El doctor asignado informa al cliente y al monitor (tráfico, problemas para llegar, etc.)."""
    s, err = _exigir_proveedor(request, None)
    if err:
        return err
    if not str(solicitud_id).isdigit():
        return _err("Solicitud inválida", 400)
    data = request.get_json(force=True, silent=True) or {}
    tipo = (data.get("tipo") or "").strip()
    if tipo not in NOVEDADES_PRESET:
        return _err("Tipo de aviso inválido", 400)
    icono, texto = NOVEDADES_PRESET[tipo]
    if tipo == "otro" or data.get("texto"):
        extra = re.sub(r"\s+", " ", str(data.get("texto") or "")).strip()[:140]
        if tipo == "otro" and len(extra) < 3:
            return _err("Escribe el mensaje para el cliente", 400)
        if tipo == "otro":
            texto = extra
        elif extra:
            texto = f"{texto} {extra}"[:200]

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT proveedor_id, estado FROM solicitudes WHERE id = %s FOR UPDATE", (int(solicitud_id),))
        row = cur.fetchone()
        if not row or row[0] != s["proveedor_id"]:
            return _err("Este servicio no es tuyo", 403)
        if row[1] not in ESTADOS_NOVEDAD:
            return _err("Ya no puedes enviar avisos en este servicio", 409)
        cur.execute("""
            SELECT COUNT(*), COALESCE(EXTRACT(EPOCH FROM (NOW() - MAX(creado_en))), 9999)
            FROM solicitudes_novedades WHERE solicitud_id = %s
        """, (int(solicitud_id),))
        total, seg = cur.fetchone()
        if total >= MAD_NOVEDADES_MAX_POR_SERVICIO:
            return _err("Llegaste al límite de avisos para este servicio", 429)
        if seg < MAD_NOVEDAD_INTERVALO_SEG:
            return _err("Espera unos segundos antes de enviar otro aviso", 429)
        cur.execute("""
            INSERT INTO solicitudes_novedades (solicitud_id, proveedor_id, tipo, texto)
            VALUES (%s, %s, %s, %s)
        """, (int(solicitud_id), s["proveedor_id"], tipo, texto))
        conn.commit()
    finally:
        cur.close(); conn.close()
    return (json.dumps({"ok": True, "icono": icono, "texto": texto}, ensure_ascii=False), 200, HEADERS)


# ── Mensajes del cliente al doctor (desde la página de seguimiento) ──
MAD_MSG_CLIENTE_MAX = 20            # por servicio
MAD_MSG_CLIENTE_INTERVALO_SEG = 10
MAD_MSG_CLIENTE_LARGO = 200


def _listar_mensajes_cliente(cur, solicitud_id, limite=10):
    cur.execute("""
        SELECT id, texto, creado_en FROM (
            SELECT id, texto, creado_en FROM solicitudes_mensajes_cliente
            WHERE solicitud_id = %s ORDER BY id DESC LIMIT %s
        ) t ORDER BY id
    """, (solicitud_id, limite))
    return [{"id": i, "texto": t, "creado_en": _iso(c)} for i, t, c in cur.fetchall()]


@con_log_errores
def handle_mensaje_cliente(request, tracking_token):
    """El cliente escribe al doctor asignado mientras el servicio está en curso."""
    if not _uuid_valido(tracking_token):
        return _err("Enlace inválido", 404)
    data = request.get_json(force=True, silent=True) or {}
    texto = re.sub(r"[\x00-\x1f\x7f]+", " ", str(data.get("texto") or ""))
    texto = re.sub(r"\s+", " ", texto).strip()[:MAD_MSG_CLIENTE_LARGO]
    if len(texto) < 2:
        return _err("Escribe tu mensaje")

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT id, estado, proveedor_id, tipo_servicio FROM solicitudes WHERE tracking_token = %s FOR UPDATE",
                    (tracking_token,))
        row = cur.fetchone()
        if not row:
            return _err("No encontrado", 404)
        sid, estado, pid, tipo = row
        if estado not in ESTADOS_NOVEDAD or not pid:
            return _err("Ya no puedes escribir al doctor en este servicio", 409)
        cur.execute("""SELECT COUNT(*), COALESCE(EXTRACT(EPOCH FROM (NOW() - MAX(creado_en))), 9999)
                       FROM solicitudes_mensajes_cliente WHERE solicitud_id = %s""", (sid,))
        total, seg = cur.fetchone()
        if total >= MAD_MSG_CLIENTE_MAX:
            return _err("Llegaste al límite de mensajes para este servicio. Puedes llamar al doctor.", 429)
        if seg < MAD_MSG_CLIENTE_INTERVALO_SEG:
            return _err("Espera unos segundos antes de enviar otro mensaje", 429)
        cur.execute("INSERT INTO solicitudes_mensajes_cliente (solicitud_id, texto) VALUES (%s, %s)", (sid, texto))
        conn.commit()
    finally:
        cur.close(); conn.close()

    # El aviso no lleva el contenido (puede incluir datos del paciente)
    enviar_push_proveedor(pid, "Mensaje del cliente", "Abre la app para leerlo.", tag="mensaje")
    return (json.dumps({"ok": True}), 200, HEADERS)


@con_log_errores
def obtener_estado_tracking(tracking_token):
    if not _uuid_valido(tracking_token):
        return (json.dumps({'error': 'no encontrado'}), 404, HEADERS)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT s.estado, s.tipo_servicio,
                   ST_X(pe.ubicacion::geometry), ST_Y(pe.ubicacion::geometry),
                   ST_X(s.origen::geometry), ST_Y(s.origen::geometry),
                   p.nombre, s.origen_direccion_texto, s.origen_distrito,
                   ST_Distance(pe.ubicacion, s.origen), COALESCE(s.cambios_medico, 0),
                   s.programada_para, s.origen_referencia, s.contacto_nombre, s.contacto_telefono,
                   COALESCE(s.reprogramaciones, 0), s.creado_en,
                   p.cmp, em.nombre, s.id, (s.empresa_aceptada_en IS NOT NULL)
            FROM solicitudes s
            LEFT JOIN proveedores p ON p.id = s.proveedor_id
            LEFT JOIN mad_empresas em ON em.id = p.empresa_id
            LEFT JOIN proveedores_estado pe ON pe.proveedor_id = s.proveedor_id
            WHERE s.tracking_token = %s
        """, (tracking_token,))
        row = cur.fetchone()
        if not row:
            return (json.dumps({'error': 'no encontrado'}), 404, HEADERS)

        (estado, tipo, plng, plat, olng, olat, nombre, direccion, distrito, dist_m, cambios_medico,
         programada_para, referencia, contacto_nombre, contacto_telefono, n_repro, creado_en,
         prov_cmp, prov_empresa, sol_id, empresa_confirmo) = row

        eta_min = None
        if estado in ('asignado', 'en_camino') and dist_m is not None:
            eta_min = _calcular_eta_min(cur, dist_m, estado)
        novedades = _listar_novedades(cur, sol_id) if estado in ESTADOS_NOVEDAD else []
        mis_mensajes = _listar_mensajes_cliente(cur, sol_id) if estado in ESTADOS_NOVEDAD else []
    finally:
        cur.close(); conn.close()

    # La posición del doctor solo se comparte mientras hay un servicio asignado en curso
    hay_doctor = estado in ('asignado', 'en_camino', 'llego', 'en_atencion')
    return (json.dumps({
        'estado': estado,
        'tipo_servicio': tipo,
        'proveedor_ubicacion': {'lat': plat, 'lng': plng} if (plat is not None and hay_doctor) else None,
        'origen': {'lat': olat, 'lng': olng},
        'direccion': direccion, 'distrito': distrito, 'referencia': referencia,
        'proveedor_nombre': nombre if hay_doctor else None,
        'proveedor_cmp': prov_cmp if hay_doctor else None,
        'proveedor_empresa': prov_empresa if hay_doctor else None,
        'novedades': novedades,
        'mis_mensajes': mis_mensajes,
        'puede_escribir': estado in ESTADOS_NOVEDAD,
        'empresa_confirmo': bool(empresa_confirmo) and estado == 'buscando_proveedor',
        'eta_min': eta_min,
        'distancia_km': round(float(dist_m) / 1000, 1) if (dist_m is not None and hay_doctor) else None,
        'cambios_medico': cambios_medico,
        'puede_cambiar_medico': estado in ('asignado', 'en_camino') and cambios_medico < MAX_CAMBIOS_MEDICO,
        'programada_para': programada_para.isoformat() if programada_para else None,
        'contacto_nombre': contacto_nombre, 'contacto_telefono': contacto_telefono,
        'reprogramaciones': n_repro,
        'creado_en': _iso(creado_en),
        'puede_reprogramar': (programada_para is not None and estado in ESTADOS_REPROGRAMABLES
                              and n_repro < MAD_MAX_REPROGRAMACIONES),
    }, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_cancelar_solicitud_cliente(request, tracking_token):
    if not _uuid_valido(tracking_token):
        return _err("Enlace inválido", 404)

    data = request.get_json(force=True, silent=True) or {}
    motivo = (data.get("motivo") or "").strip()
    comentario = (data.get("comentario") or "").strip()[:300] or None

    if motivo not in MOTIVOS_CANCELACION:
        return _err("Selecciona un motivo de cancelación")
    if motivo == "otro" and not comentario:
        return _err("Cuéntanos brevemente el motivo")

    proveedores_a_avisar = []
    distrito = None
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT id, estado, proveedor_id, origen_distrito
            FROM solicitudes WHERE tracking_token = %s
            FOR UPDATE
        """, (tracking_token,))
        row = cur.fetchone()
        if not row:
            return _err("Solicitud no encontrada", 404)

        solicitud_id, estado_anterior, proveedor_asignado, distrito = row
        if estado_anterior not in ESTADOS_CANCELABLES:
            conn.rollback()
            return _err("Este servicio ya no se puede cancelar desde aquí", 409, estado=estado_anterior)

        cur.execute("""
            UPDATE solicitudes
            SET estado = 'cancelado', motivo_cancelacion = %s, comentario_cancelacion = %s,
                cancelado_en = NOW(), cancelado_por = 'cliente'
            WHERE id = %s
        """, (motivo, comentario, solicitud_id))

        cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'cancelada', respondido_en = NOW()
                       WHERE solicitud_id = %s AND estado IN ('pendiente', 'aceptada')""", (solicitud_id,))
        # Ofertas pendientes de esta solicitud → canceladas
        cur.execute("""
            UPDATE solicitudes_ofertas SET estado = 'cancelada', respondido_en = NOW()
            WHERE solicitud_id = %s AND estado = 'pendiente'
            RETURNING proveedor_id
        """, (solicitud_id,))
        proveedores_liberar = {r[0] for r in cur.fetchall()}
        if proveedor_asignado:
            proveedores_liberar.add(proveedor_asignado)

        # Liberar proveedores (vuelven a estar disponibles); la hora de su última señal no se toca
        for pid in proveedores_liberar:
            cur.execute("""
                UPDATE proveedores_estado SET estado = 'disponible'
                WHERE proveedor_id = %s AND estado IN ('en_oferta', 'ocupado')
            """, (pid,))

        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, %s, 'cancelado', 'cliente')
        """, (solicitud_id, estado_anterior))

        if proveedores_liberar:
            cur.execute("SELECT telefono FROM proveedores WHERE id = ANY(%s)", (list(proveedores_liberar),))
            proveedores_a_avisar = [r[0] for r in cur.fetchall()]

        conn.commit()
    finally:
        cur.close(); conn.close()

    # Aviso al/los proveedor(es) — fuera de la transacción
    for telefono in proveedores_a_avisar:
        numero = _normalizar_celular_pe(telefono)
        if numero:
            zona = f" en {distrito.title()}" if distrito else ""
            enviar_danaconnect(numero, f"Sanitas: El cliente cancelo el servicio #{solicitud_id}{zona}. Ya estas disponible para nuevas solicitudes.")
            sms_aviso(numero, "mad cancelado cliente")

    return (json.dumps({"ok": True, "estado": "cancelado"}), 200, HEADERS)


@con_log_errores
def handle_cambiar_medico_cliente(request, tracking_token):
    """El cliente no quiere al médico asignado: se busca otro (el anterior queda excluido)."""
    if not _uuid_valido(tracking_token):
        return _err("Enlace inválido", 404)

    telefono_anterior = None
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT id, estado, proveedor_id, COALESCE(cambios_medico, 0)
            FROM solicitudes WHERE tracking_token = %s
            FOR UPDATE
        """, (tracking_token,))
        row = cur.fetchone()
        if not row:
            return _err("Solicitud no encontrada", 404)

        solicitud_id, estado_anterior, proveedor_anterior, cambios = row

        if estado_anterior not in ("asignado", "en_camino") or not proveedor_anterior:
            conn.rollback()
            return _err("Solo puedes pedir otro médico antes de que llegue a tu domicilio", 409)
        if cambios >= MAX_CAMBIOS_MEDICO:
            conn.rollback()
            return _err("Ya solicitaste un cambio de médico para este servicio. Si lo deseas, puedes cancelarlo.", 409)

        cur.execute("""
            INSERT INTO solicitudes_proveedores_excluidos (solicitud_id, proveedor_id, motivo)
            VALUES (%s, %s, 'cambio_solicitado_por_cliente')
            ON CONFLICT DO NOTHING
        """, (solicitud_id, proveedor_anterior))

        cur.execute("""
            UPDATE proveedores_estado SET estado = 'disponible'
            WHERE proveedor_id = %s AND estado = 'ocupado'
        """, (proveedor_anterior,))

        cur.execute("""
            UPDATE solicitudes
            SET estado = 'buscando_proveedor', proveedor_id = NULL, empresa_id = NULL, asignado_en = NULL,
                en_camino_en = NULL, en_camino_ubicacion = NULL, en_camino_eta_min = NULL, llego_en = NULL,
                cambios_medico = COALESCE(cambios_medico, 0) + 1,
                intentos_asignacion = 0, radio_busqueda_metros = 5000, ultima_busqueda_en = NULL
            WHERE id = %s
        """, (solicitud_id,))

        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, %s, 'buscando_proveedor', 'cliente')
        """, (solicitud_id, estado_anterior))

        cur.execute("SELECT telefono FROM proveedores WHERE id = %s", (proveedor_anterior,))
        r = cur.fetchone()
        telefono_anterior = r[0] if r else None

        conn.commit()
    finally:
        cur.close(); conn.close()

    numero = _normalizar_celular_pe(telefono_anterior)
    if numero:
        # "reasignado" y no "el cliente no lo quiso": evita fricción. El motivo real queda en la tabla de exclusiones.
        enviar_danaconnect(numero, f"Sanitas: El servicio #{solicitud_id} fue reasignado a otro profesional. Ya estas disponible para nuevas solicitudes.")
        sms_aviso(numero, "mad reasignado cliente")

    resultado = buscar_y_ofrecer_proveedor(solicitud_id)
    return (json.dumps({"ok": True, "estado": "buscando_proveedor", "busqueda": resultado},
                       ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_reprogramar_solicitud_cliente(request, tracking_token):
    if not _uuid_valido(tracking_token):
        return _err("Enlace inválido", 404)

    data = request.get_json(force=True, silent=True) or {}
    notificar_doctores, tel_cliente, resultado = [], None, {}
    material, nueva_prog, distrito, sid = False, None, "", None

    conn = get_connection(); cur = conn.cursor()
    try:
        # 1) Lectura sin bloqueo: valores actuales
        cur.execute("""
            SELECT id, estado, proveedor_id, programada_para, origen_direccion_texto, origen_distrito,
                   origen_referencia, contacto_nombre, contacto_telefono,
                   ST_Y(origen::geometry), ST_X(origen::geometry),
                   COALESCE(reprogramaciones, 0), cliente_telefono
            FROM solicitudes WHERE tracking_token = %s
        """, (tracking_token,))
        row = cur.fetchone()
        if not row:
            return _err("Solicitud no encontrada", 404)
        (sid, estado, prov, prog_act, dir_act, dist_act, ref_act,
         cn_act, ct_act, lat_act, lng_act, n_repro, tel_cliente) = row

        if prog_act is None:
            return _err("Este servicio no es programado. Si ya no lo necesitas, puedes cancelarlo.", 409)
        if estado not in ESTADOS_REPROGRAMABLES:
            return _err("El profesional ya está en camino o el servicio ya terminó, no se puede reprogramar. Puedes cancelarlo.", 409)

        # 2) Valores nuevos (lo que no viene, se conserva)
        direccion = str(data["direccion"] if "direccion" in data else (dir_act or "")).strip()[:200]
        distrito  = str(data["distrito"] if "distrito" in data else (dist_act or "")).strip().upper()[:80]
        referencia = ((data.get("referencia") or "").strip()[:200] or None) if "referencia" in data else ref_act
        cn = ((data.get("contacto_nombre") or "").strip()[:120] or None) if "contacto_nombre" in data else cn_act
        if "contacto_telefono" in data:
            ct, e = _normalizar_contacto(data.get("contacto_telefono"))
            if e:
                return _err(e)
        else:
            ct = ct_act

        if "lat" in data or "lng" in data:
            try:
                lat = float(data.get("lat")); lng = float(data.get("lng"))
            except (TypeError, ValueError):
                return _err("Ubicación inválida")
        else:
            lat, lng = lat_act, lng_act

        e = _validar_lugar(lat, lng, direccion, distrito)
        if e:
            return _err(e)

        nueva_prog = prog_act
        if data.get("programada_para"):
            try:
                nueva_prog = _parse_programada(data["programada_para"])
            except ValueError:
                return _err("Fecha y hora inválidas (formato AAAA-MM-DDTHH:MM, hora de Lima)")
        cambia_hora = abs((nueva_prog - prog_act).total_seconds()) >= 60
        if cambia_hora:
            e = _validar_hora_programada(cur, nueva_prog)   # la anticipación solo aplica si cambia la hora
            if e:
                return _err(e)
        else:
            nueva_prog = prog_act

        cur.execute("""
            SELECT ST_Distance(ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
                               ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography)
        """, (lng_act, lat_act, lng, lat))
        dist_m = float(cur.fetchone()[0] or 0)

        material = cambia_hora or dist_m > MAD_REPROG_DISTANCIA_M or distrito != (dist_act or "").upper()
        if material and n_repro >= MAD_MAX_REPROGRAMACIONES:
            return _err(f"Ya reprogramaste este servicio {MAD_MAX_REPROGRAMACIONES} veces. Cancélalo y crea uno nuevo.", 409)

        # 3) Bloqueos (orden: solicitud → ofertas → doctor), igual que aceptar_oferta y cancelar
        cur.execute("SELECT estado, proveedor_id FROM solicitudes WHERE id = %s FOR UPDATE", (sid,))
        estado, prov = cur.fetchone()
        if estado not in ESTADOS_REPROGRAMABLES:       # cambió mientras tanto
            conn.rollback()
            return _err("El profesional ya está en camino, no se puede reprogramar. Puedes cancelar el servicio.", 409)

        liberar = set()
        doctor_liberado = None
        if material:
            cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'cancelada', respondido_en = NOW()
                           WHERE solicitud_id = %s AND estado IN ('pendiente', 'aceptada')""", (sid,))
            cur.execute("UPDATE solicitudes SET empresa_aceptada_en = NULL WHERE id = %s AND estado = 'buscando_proveedor'", (sid,))
            cur.execute("""
                UPDATE solicitudes_ofertas SET estado = 'cancelada', respondido_en = NOW()
                WHERE solicitud_id = %s AND estado = 'pendiente' RETURNING proveedor_id
            """, (sid,))
            liberar = {r[0] for r in cur.fetchall()}

            if estado == "asignado" and prov:
                liberar.add(prov); doctor_liberado = prov
            for pid in liberar:
                cur.execute("""
                    UPDATE proveedores_estado SET estado = 'disponible'
                    WHERE proveedor_id = %s AND estado IN ('en_oferta', 'ocupado')
                """, (pid,))
            if liberar:
                cur.execute("SELECT telefono FROM proveedores WHERE id = ANY(%s)", (list(liberar),))
                notificar_doctores = [r[0] for r in cur.fetchall()]

            cur.execute("""
                UPDATE solicitudes
                SET estado = 'programada', programada_para = %s,
                    origen = ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
                    origen_direccion_texto = %s, origen_distrito = %s, origen_referencia = %s,
                    contacto_nombre = %s, contacto_telefono = %s,
                    proveedor_id = NULL, empresa_id = NULL, asignado_en = NULL,
                    en_camino_en = NULL, en_camino_ubicacion = NULL, en_camino_eta_min = NULL, llego_en = NULL,
                    intentos_asignacion = 0, radio_busqueda_metros = 5000, ultima_busqueda_en = NULL,
                    alerta_cobertura_nivel = 0, reprogramaciones = COALESCE(reprogramaciones, 0) + 1
                WHERE id = %s
            """, (nueva_prog, lng, lat, direccion, distrito, referencia, cn, ct, sid))
            if estado != "programada":
                cur.execute("""
                    INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                    VALUES (%s, %s, 'programada', 'cliente')
                """, (sid, estado))
        else:
            # Cambio menor: se actualiza sin tocar al doctor asignado
            cur.execute("""
                UPDATE solicitudes
                SET origen = ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
                    origen_direccion_texto = %s, origen_referencia = %s,
                    contacto_nombre = %s, contacto_telefono = %s
                WHERE id = %s
            """, (lng, lat, direccion, referencia, cn, ct, sid))

        cur.execute("""
            INSERT INTO mad_reprogramaciones (solicitud_id, material, programada_anterior, programada_nueva,
                direccion_anterior, direccion_nueva, distrito_anterior, distrito_nuevo, distancia_m, doctor_liberado)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (sid, material, prog_act, nueva_prog, dir_act, direccion, dist_act, distrito, round(dist_m), doctor_liberado))
        conn.commit()

        resultado = {"ok": True, "material": material, "doctor_liberado": doctor_liberado is not None,
                     "estado": "programada" if material else estado,
                     "programada_para": nueva_prog.astimezone(LIMA_TZ).isoformat()}
    finally:
        cur.close(); conn.close()

    # Avisos fuera de la transacción
    for tel in notificar_doctores:
        numero = _normalizar_celular_pe(tel)
        if numero:
            enviar_danaconnect(numero, f"Sanitas: El servicio #{sid} fue reprogramado por el cliente. Ya estas disponible; se te avisara si se te asigna de nuevo.")
            sms_aviso(numero, "mad reprogramado doctor")
    if material:
        numero = _normalizar_celular_pe(tel_cliente)
        if numero:
            cuando = nueva_prog.astimezone(LIMA_TZ).strftime("%d/%m %H:%M")
            enviar_danaconnect(numero, f"Sanitas: Tu servicio fue reprogramado para {cuando} en {distrito.title()}. Sigue o cancela aqui: {_base_url()}/track/{tracking_token}")
            sms_aviso(numero, "mad reprogramado cliente")

    return (json.dumps(resultado, ensure_ascii=False), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# PANEL DEL PROVEEDOR
# ═════════════════════════════════════════════════════════════
@con_log_errores
def handle_actualizar_ubicacion_proveedor(request, proveedor_id):
    try:
        proveedor_id = int(proveedor_id)
    except ValueError:
        return _err("proveedor_id inválido")

    data = request.get_json(force=True, silent=True) or {}
    try:
        lat, lng = float(data.get("lat")), float(data.get("lng"))
    except (TypeError, ValueError):
        return _err("lat y lng son obligatorios")
    if not (PERU_LAT[0] <= lat <= PERU_LAT[1] and PERU_LNG[0] <= lng <= PERU_LNG[1]):
        return _err("La ubicación debe estar dentro de Perú")

    conn = get_connection()
    cur = conn.cursor()
    recien_conectado = False
    try:
        cur.execute("""SELECT estado, ubicacion IS NULL,
                              COALESCE(EXTRACT(EPOCH FROM (NOW() - actualizado_en)), 99999)
                       FROM proveedores_estado WHERE proveedor_id = %s""", (proveedor_id,))
        previo = cur.fetchone()
        cur.execute("""
            UPDATE proveedores_estado
            SET ubicacion = ST_SetSRID(ST_MakePoint(%s, %s), 4326)::geography,
                actualizado_en = %s
            WHERE proveedor_id = %s
        """, (lng, lat, datetime.now(LIMA_TZ), proveedor_id))
        conn.commit()
        # Primera señal tras estar desconectado o sin reportar (>90 s): puede haber servicios esperando
        recien_conectado = bool(previo and previo[0] == 'disponible' and (previo[1] or previo[2] > 90))
    finally:
        cur.close(); conn.close()

    if recien_conectado:
        try:
            _reabrir_sin_cobertura(proveedor_id)
            _ofrecer_en_espera_a_proveedor(proveedor_id)
        except Exception as e:
            logging.warning(f"[MAD] revisión de servicios en espera falló (proveedor={proveedor_id}): {e}")

    return (json.dumps({"status": "ok"}), 200, HEADERS)


@con_log_errores
def handle_toggle_disponibilidad(request, proveedor_id):
    data = request.get_json(force=True, silent=True) or {}
    estado_nuevo = data.get("estado")

    if estado_nuevo not in ("disponible", "desconectado"):
        return _err("estado debe ser 'disponible' o 'desconectado'")

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE proveedores_estado SET estado = %s, actualizado_en = %s
            WHERE proveedor_id = %s AND estado NOT IN ('ocupado', 'en_oferta')
            RETURNING estado
        """, (estado_nuevo, datetime.now(LIMA_TZ), proveedor_id))
        row = cur.fetchone()
        conn.commit()
    finally:
        cur.close(); conn.close()

    if not row:
        return _err("No puedes cambiar tu disponibilidad mientras tienes un servicio activo", 409)
    logging.info(f"[MAD] disponibilidad proveedor={proveedor_id} → {row[0]}")
    return (json.dumps({"status": "ok", "estado": row[0]}), 200, HEADERS)


@con_log_errores
def handle_estado_actual_proveedor(request, proveedor_id):
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT p.nombre, pe.estado
            FROM proveedores p
            JOIN proveedores_estado pe ON pe.proveedor_id = p.id
            WHERE p.id = %s
        """, (proveedor_id,))
        row_proveedor = cur.fetchone()
        if not row_proveedor:
            return _err("Proveedor no encontrado", 404)
        nombre, disponibilidad = row_proveedor

        # ── Oferta pendiente: solo distrito, distancia y hora (nunca dirección ni coordenadas) ──
        cur.execute("""
            SELECT o.id, o.oferta_token, s.tipo_servicio, s.origen_distrito,
                   ST_Distance(pe.ubicacion, s.origen) AS distancia_metros, s.programada_para,
                   EXTRACT(EPOCH FROM (NOW() - pe.actualizado_en)) / 60.0
            FROM solicitudes_ofertas o
            JOIN solicitudes s ON s.id = o.solicitud_id
            JOIN proveedores_estado pe ON pe.proveedor_id = o.proveedor_id
            WHERE o.proveedor_id = %s AND o.estado = 'pendiente'
            ORDER BY o.ofrecido_en DESC LIMIT 1
        """, (proveedor_id,))
        row_oferta = cur.fetchone()
        oferta_pendiente = None
        if row_oferta:
            oferta_id, oferta_token, tipo_servicio, distrito, distancia, programada_para, ubic_min = row_oferta
            oferta_pendiente = {
                "oferta_id": oferta_id,
                "oferta_token": str(oferta_token),
                "tipo_servicio": tipo_servicio,
                "distrito": distrito,
                "distancia_metros": float(distancia) if distancia is not None else None,
                "programada_para": programada_para.isoformat() if programada_para else None,
                "ubicacion_hace_min": int(ubic_min) if ubic_min is not None else None,   # antigüedad de la posición usada para la distancia
            }

        # ── Servicio activo: ya aceptado, ve dirección, referencia, cliente y contacto ──
        cur.execute("""
            SELECT id, tipo_servicio, estado, cliente_telefono, cliente_nombre,
                   origen_direccion_texto, origen_distrito, origen_referencia,
                   ST_Y(origen::geometry) AS lat, ST_X(origen::geometry) AS lng,
                   programada_para, contacto_nombre, contacto_telefono,
                   paciente_dni, paciente_edad, sintomas, antecedentes, alergias,
                   copago_monto, forma_pago
            FROM solicitudes
            WHERE proveedor_id = %s AND estado IN ('asignado', 'en_camino', 'llego', 'en_atencion')
            ORDER BY asignado_en DESC LIMIT 1
        """, (proveedor_id,))
        row_servicio = cur.fetchone()
        servicio_activo = None
        if row_servicio:
            (sol_id, tipo_servicio, estado, tel, nombre_cliente, direccion, distrito, referencia,
             lat, lng, programada_para, contacto_nombre, contacto_telefono,
             p_dni, p_edad, p_sintomas, p_antec, p_alergias, p_copago, p_forma) = row_servicio
            servicio_activo = {
                "solicitud_id": sol_id, "tipo_servicio": tipo_servicio, "estado": estado,
                "cliente_telefono": tel, "cliente_nombre": nombre_cliente,
                "direccion": direccion, "distrito": distrito, "referencia": referencia,
                "origen_lat": lat, "origen_lng": lng,
                "programada_para": programada_para.isoformat() if programada_para else None,
                "contacto_nombre": contacto_nombre, "contacto_telefono": contacto_telefono,
                "paciente_dni": p_dni, "paciente_edad": p_edad,
                "sintomas": p_sintomas, "antecedentes": p_antec, "alergias": p_alergias,
                "copago_monto": float(p_copago) if p_copago is not None else None, "forma_pago": p_forma,
                "mensajes_cliente": _listar_mensajes_cliente(cur, sol_id) if estado in ESTADOS_NOVEDAD else [],
            }
    finally:
        cur.close(); conn.close()

    return (json.dumps({
        "nombre": nombre,
        "disponibilidad": disponibilidad,
        "app_version": MAD_APP_VERSION,
        "apk_version": MAD_APK_VERSION,
        "autodesconexion_min": MAD_AUTODESCONEXION_MIN,
        "oferta_pendiente": oferta_pendiente,
        "servicio_activo": servicio_activo,
    }, ensure_ascii=False), 200, HEADERS)


@con_log_errores
def handle_actualizar_estado_solicitud(request, solicitud_id, proveedor_id_sesion):
    """El proveedor sale de la sesión (token), nunca del body."""
    try:
        solicitud_id = int(solicitud_id)
    except ValueError:
        return _err("solicitud_id inválido")

    data = request.get_json(force=True, silent=True) or {}
    estado_nuevo = data.get("estado")
    proveedor_id = proveedor_id_sesion

    TRANSICIONES_VALIDAS = {
        "asignado": "en_camino",
        "en_camino": "llego",
        "llego": "en_atencion",
        "en_atencion": "finalizado",
    }

    conn = get_connection()
    cur = conn.cursor()
    try:
        # FOR UPDATE: si coordinación libera el servicio a la vez, una de las dos operaciones espera a la otra
        cur.execute("SELECT estado, proveedor_id FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
        row = cur.fetchone()
        if not row:
            conn.rollback()
            return _err("Solicitud no encontrada", 404)

        estado_actual, proveedor_asignado = row

        if str(proveedor_asignado) != str(proveedor_id):
            conn.rollback()
            return _err("Este servicio no está asignado a este proveedor", 403)

        if TRANSICIONES_VALIDAS.get(estado_actual) != estado_nuevo:
            conn.rollback()
            return _err(f"Transición inválida: {estado_actual} → {estado_nuevo}", 409)

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

        # ── Datos para calibrar el ETA ──
        if estado_nuevo == "en_camino":
            cur.execute("""
                UPDATE solicitudes s
                SET en_camino_en = NOW(), en_camino_ubicacion = pe.ubicacion
                FROM proveedores_estado pe
                WHERE s.id = %s AND pe.proveedor_id = s.proveedor_id
                RETURNING ST_Distance(pe.ubicacion, s.origen)
            """, (solicitud_id,))
            r = cur.fetchone()
            if r and r[0] is not None:
                cur.execute("UPDATE solicitudes SET en_camino_eta_min = %s WHERE id = %s",
                            (_calcular_eta_min(cur, r[0], "en_camino"), solicitud_id))
        elif estado_nuevo == "llego":
            cur.execute("UPDATE solicitudes SET llego_en = NOW() WHERE id = %s", (solicitud_id,))

        if estado_nuevo == "finalizado":
            cur.execute("""
                UPDATE proveedores_estado SET estado = 'disponible' WHERE proveedor_id = %s
            """, (proveedor_asignado,))

        conn.commit()
    finally:
        cur.close(); conn.close()

    if estado_nuevo == "en_camino":
        notificar_cliente_proveedor_en_camino(solicitud_id)

    return (json.dumps({"status": "ok", "estado": estado_nuevo}), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# TAREAS PERIÓDICAS (Cloud Scheduler cada minuto → /solicitudes/revisar-timeouts)
# ═════════════════════════════════════════════════════════════
def _autodesconectar_proveedores():
    """Doctores 'disponible' sin ubicación nueva por más de MAD_AUTODESCONEXION_MIN → desconectado."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE proveedores_estado SET estado = 'desconectado'
            WHERE estado = 'disponible' AND actualizado_en < NOW() - %s
            RETURNING proveedor_id, actualizado_en
        """, (timedelta(minutes=MAD_AUTODESCONEXION_MIN),))
        filas = cur.fetchall()
        ids = [r[0] for r in filas]
        conn.commit()
        for pid, ultima in filas:
            logging.warning(f"[MAD] autodesconexión proveedor={pid} última señal={ultima}")
    finally:
        cur.close(); conn.close()
    return ids


def _reabrir_sin_cobertura(proveedor_id=None):
    """Servicios digitados por el call center que quedaron 'sin_cobertura' hace poco: si ahora hay un
    doctor conectado (con señal fresca) a su alcance, se reabre la búsqueda y se le ofrece.
    proveedor_id: restringe a ese doctor (cuando acaba de conectarse). Devuelve los ids reabiertos."""
    reabiertas = []
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute(f"""
            SELECT s.id FROM solicitudes s
            WHERE s.estado = 'sin_cobertura' AND s.canal = 'callcenter'
              AND s.creado_en > NOW() - INTERVAL '{MAD_REABRIR_MIN} minutes'
              AND EXISTS (
                SELECT 1 FROM proveedores p
                JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                WHERE pe.estado = 'disponible' AND pe.ubicacion IS NOT NULL
                  AND pe.actualizado_en > NOW() - INTERVAL '{MAD_AUTODESCONEXION_MIN} minutes'
                  AND p.tipo_servicio IN (s.tipo_servicio, 'AMBOS')
                  AND ST_DWithin(pe.ubicacion, s.origen, {RADIO_MAXIMO_METROS})
                  AND (%s::int IS NULL OR p.id = %s::int)
                  AND EXISTS (SELECT 1 FROM mad_empresas_prioridad pr
                              WHERE pr.empresa_id = p.empresa_id AND pr.tipo_servicio = s.tipo_servicio AND pr.peso > 0)
                  AND NOT EXISTS (SELECT 1 FROM solicitudes_ofertas_empresa oe
                                  WHERE oe.solicitud_id = s.id AND oe.empresa_id = p.empresa_id
                                    AND oe.estado IN ('rechazada','expirada'))
              )
            ORDER BY s.creado_en LIMIT 5
        """, (proveedor_id, proveedor_id))
        ids = [r[0] for r in cur.fetchall()]
        for sid in ids:
            cur.execute("""
                UPDATE solicitudes
                SET estado = 'buscando_proveedor', intentos_asignacion = 0,
                    radio_busqueda_metros = 5000, ultima_busqueda_en = NULL
                WHERE id = %s AND estado = 'sin_cobertura' RETURNING id
            """, (sid,))
            if cur.fetchone():
                cur.execute("""
                    INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                    VALUES (%s, 'sin_cobertura', 'buscando_proveedor', 'sistema_doctor_conectado')
                """, (sid,))
                reabiertas.append(sid)
        conn.commit()
    finally:
        cur.close(); conn.close()

    for sid in reabiertas:
        try:
            r = buscar_y_ofrecer_proveedor(sid, radio_metros=RADIO_MAXIMO_METROS)
            logging.info(f"[MAD] servicio #{sid} reabierto por doctor conectado → {r.get('status')}")
        except Exception as e:
            logging.warning(f"[MAD] no se pudo reabrir #{sid}: {e}")
    return reabiertas


def _ofrecer_en_espera_a_proveedor(proveedor_id):
    """Un doctor acaba de conectarse: si hay servicios buscando empresa (sin oferta pendiente ni empresa que
    ya lo aceptó) y la empresa del doctor puede recibirlos, se dispara la búsqueda ya, sin esperar al minuto."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute(f"""
            SELECT s.id FROM solicitudes s
            JOIN proveedores_estado pe ON pe.proveedor_id = %s
            JOIN proveedores p ON p.id = pe.proveedor_id
            JOIN mad_empresas_prioridad pr ON pr.empresa_id = p.empresa_id AND pr.tipo_servicio = s.tipo_servicio AND pr.peso > 0
            WHERE s.estado = 'buscando_proveedor' AND s.programada_para IS NULL
              AND s.empresa_aceptada_en IS NULL
              AND pe.estado = 'disponible' AND pe.ubicacion IS NOT NULL
              AND p.tipo_servicio IN (s.tipo_servicio, 'AMBOS')
              AND ST_DWithin(pe.ubicacion, s.origen, {RADIO_MAXIMO_METROS})
              AND NOT EXISTS (SELECT 1 FROM solicitudes_ofertas_empresa oe WHERE oe.solicitud_id = s.id AND oe.estado = 'pendiente')
              AND NOT EXISTS (SELECT 1 FROM solicitudes_ofertas_empresa oe
                              WHERE oe.solicitud_id = s.id AND oe.empresa_id = p.empresa_id AND oe.estado IN ('rechazada','expirada'))
            ORDER BY s.creado_en LIMIT 3
        """, (proveedor_id,))
        ids = [r[0] for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
    for sid in ids:
        r = buscar_y_ofrecer_proveedor(sid, radio_metros=RADIO_MAXIMO_METROS)
        if r.get('status') == 'ofrecido':
            break


def _marcar_sin_cobertura(solicitud_id):
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE solicitudes SET estado = 'sin_cobertura'
            WHERE id = %s AND estado = 'buscando_proveedor'
            RETURNING id
        """, (solicitud_id,))
        if not cur.fetchone():
            conn.rollback()
            return False
        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, 'buscando_proveedor', 'sin_cobertura', 'sistema')
        """, (solicitud_id,))
        cur.execute("""SELECT COUNT(*) FROM solicitudes_historial_estado
                       WHERE solicitud_id = %s AND estado_nuevo = 'sin_cobertura'""", (solicitud_id,))
        primera_vez = cur.fetchone()[0] <= 1
        conn.commit()
    finally:
        cur.close(); conn.close()

    if primera_vez:       # si el servicio se reabrió y volvió a fallar, no se vuelve a molestar al cliente
        notificar_cliente_sin_cobertura(solicitud_id)
    return True


def _reintentar_solicitudes_sin_oferta():
    """Solicitudes en buscando_proveedor que NO tienen oferta pendiente
    (no había nadie disponible en el último intento). Amplía el radio y
    reintenta; al llegar a MAX_INTENTOS_ASIGNACION, pasan a sin_cobertura."""
    limite = datetime.now(LIMA_TZ).replace(tzinfo=None) - timedelta(seconds=ESPERA_REINTENTO_SIN_PROVEEDOR_SEGUNDOS)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT s.id FROM solicitudes s
            WHERE s.estado = 'buscando_proveedor'
              AND s.empresa_aceptada_en IS NULL
              AND NOT EXISTS (SELECT 1 FROM solicitudes_ofertas_empresa oe
                              WHERE oe.solicitud_id = s.id AND oe.estado = 'pendiente')
              AND NOT EXISTS (SELECT 1 FROM solicitudes_ofertas o
                              WHERE o.solicitud_id = s.id AND o.estado = 'pendiente')
              AND COALESCE(s.ultima_busqueda_en, '-infinity'::timestamp) < %s
        """, (limite,))
        ids = [r[0] for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()

    reintentadas, sin_cobertura = [], []
    for solicitud_id in ids:
        conn = get_connection(); cur = conn.cursor()
        try:
            cur.execute("""
                UPDATE solicitudes
                SET intentos_asignacion = COALESCE(intentos_asignacion, 0) + 1,
                    radio_busqueda_metros = LEAST(COALESCE(radio_busqueda_metros, 5000) + %s, %s)
                WHERE id = %s AND estado = 'buscando_proveedor'
                RETURNING intentos_asignacion, radio_busqueda_metros
            """, (INCREMENTO_RADIO_METROS, RADIO_MAXIMO_METROS, solicitud_id))
            fila = cur.fetchone()
            conn.commit()
        finally:
            cur.close(); conn.close()
        if not fila:
            continue

        intentos, radio = fila
        if intentos >= MAX_INTENTOS_ASIGNACION:
            if _marcar_sin_cobertura(solicitud_id):
                sin_cobertura.append(solicitud_id)
        else:
            resultado = buscar_y_ofrecer_proveedor(solicitud_id, radio_metros=radio)
            reintentadas.append({"solicitud_id": solicitud_id, "resultado": resultado})

    return reintentadas, sin_cobertura


def _activar_programadas():
    """Pasa a 'buscando_proveedor' las solicitudes programadas cuya hora está a
    MAD_PROG_ACTIVAR_MIN minutos o menos, y lanza la búsqueda de doctor."""
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            WITH vencidas AS (
                SELECT id FROM solicitudes
                WHERE estado = 'programada' AND programada_para <= NOW() + %s
                FOR UPDATE SKIP LOCKED
            )
            UPDATE solicitudes s
            SET estado = 'buscando_proveedor', intentos_asignacion = 0,
                radio_busqueda_metros = 5000, ultima_busqueda_en = NULL
            FROM vencidas v WHERE s.id = v.id
            RETURNING s.id
        """, (timedelta(minutes=MAD_PROG_ACTIVAR_MIN),))
        ids = [r[0] for r in cur.fetchall()]
        for sid in ids:
            cur.execute("""
                INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                VALUES (%s, 'programada', 'buscando_proveedor', 'sistema')
            """, (sid,))
        conn.commit()
    finally:
        cur.close(); conn.close()

    return [{"solicitud_id": sid, "resultado": buscar_y_ofrecer_proveedor(sid)} for sid in ids]


def _alertar_cobertura_programadas():
    """Revisa los servicios programados que entran en la ventana de 3 h (y luego 1 h) y
    avisa a coordinación si no hay doctores conectados cerca del lugar."""
    conn = get_connection(); cur = conn.cursor()
    avisos = []
    try:
        cur.execute("""
            SELECT s.id, s.tipo_servicio, s.origen_distrito, s.programada_para,
                   (s.programada_para <= NOW() + %(h2)s) AS nivel2, c.conectados
            FROM solicitudes s
            CROSS JOIN LATERAL (
                SELECT COUNT(*) AS conectados
                FROM proveedores p
                JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                JOIN mad_usuarios u ON u.proveedor_id = p.id AND u.activo
                WHERE pe.estado IN ('disponible', 'en_oferta', 'ocupado')
                  AND p.tipo_servicio IN (s.tipo_servicio, 'AMBOS')
                  AND pe.ubicacion IS NOT NULL
                  AND pe.actualizado_en > NOW() - %(senal)s
                  AND ST_DWithin(pe.ubicacion, s.origen, %(radio)s)
            ) c
            WHERE s.estado = 'programada' AND s.programada_para > NOW()
              AND ((s.alerta_cobertura_nivel < 1 AND s.programada_para <= NOW() + %(h1)s)
                OR (s.alerta_cobertura_nivel < 2 AND s.programada_para <= NOW() + %(h2)s))
            FOR UPDATE OF s SKIP LOCKED
        """, {"h1": timedelta(hours=MAD_COBERTURA_HORAS_1), "h2": timedelta(hours=MAD_COBERTURA_HORAS_2),
              "senal": timedelta(minutes=MAD_COBERTURA_SENAL_MIN), "radio": RADIO_MAXIMO_METROS})
        for sid, tipo, distrito, prog, nivel2, conectados in cur.fetchall():
            cur.execute("UPDATE solicitudes SET alerta_cobertura_nivel = %s WHERE id = %s",
                        (2 if nivel2 else 1, sid))
            if conectados == 0:
                avisos.append((sid, tipo, distrito, prog, bool(nivel2)))
        conn.commit()
    finally:
        cur.close(); conn.close()

    for sid, tipo, distrito, prog, urgente in avisos:
        cuando = prog.astimezone(LIMA_TZ).strftime("%d/%m %H:%M")
        texto = (f"ALERTA MAD{' URGENTE' if urgente else ''}: #{sid} programado {cuando} en "
                 f"{(distrito or '').title()}. Sin doctores conectados cerca. Contacta doctores de la zona.")[:160]
        for cel in MAD_ALERTA_CELULARES:
            numero = _normalizar_celular_pe(cel) or cel
            enviar_danaconnect(numero, texto)
            sms_aviso(numero, "mad alerta cobertura")
    return avisos


@con_log_errores
def handle_revisar_timeouts_ofertas(request):
    # 0) Doctores sin señal reciente dejan de recibir ofertas
    desconectados = _autodesconectar_proveedores()

    limite = datetime.now(LIMA_TZ) - timedelta(seconds=TIMEOUT_OFERTA_SEGUNDOS)

    # 0b) Ofertas a empresas vencidas y servicios aceptados que la empresa no asignó a tiempo
    empresas_vencidas = _vencer_ofertas_empresa()

    # 1) Ofertas pendientes que ya vencieron, con su solicitud aún activa
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT o.id, o.solicitud_id, o.proveedor_id
            FROM solicitudes_ofertas o
            JOIN solicitudes s ON s.id = o.solicitud_id
            WHERE o.estado = 'pendiente'
              AND o.ofrecido_en < %s
              AND s.estado = 'buscando_proveedor'
        """, (limite,))
        candidatas = cur.fetchall()
    finally:
        cur.close(); conn.close()

    reasignadas, sin_cobertura, vencidas = [], [], 0
    for oferta_id, solicitud_id, proveedor_id in candidatas:
        conn = get_connection(); cur = conn.cursor()
        try:
            # Orden de bloqueo: solicitud → oferta → doctor
            cur.execute("SELECT estado FROM solicitudes WHERE id = %s FOR UPDATE", (solicitud_id,))
            r = cur.fetchone()
            if not r or r[0] != 'buscando_proveedor':
                conn.rollback()
                continue

            # a) Expira la oferta (si el doctor justo respondió, ya no está pendiente y se omite)
            cur.execute("""
                UPDATE solicitudes_ofertas SET estado = 'expirada', respondido_en = %s
                WHERE id = %s AND estado = 'pendiente'
                RETURNING id
            """, (datetime.now(LIMA_TZ), oferta_id))
            if not cur.fetchone():
                conn.rollback()
                continue

            # b) Libera al proveedor (vuelve a estar disponible para otras solicitudes)
            cur.execute("""
                UPDATE proveedores_estado SET estado = 'disponible'
                WHERE proveedor_id = %s AND estado = 'en_oferta'
            """, (proveedor_id,))

            # c) Incrementa intentos y expande el radio de esta solicitud
            cur.execute("""
                UPDATE solicitudes
                SET intentos_asignacion = COALESCE(intentos_asignacion, 0) + 1,
                    radio_busqueda_metros = LEAST(COALESCE(radio_busqueda_metros, 5000) + %s, %s)
                WHERE id = %s
                RETURNING intentos_asignacion, radio_busqueda_metros
            """, (INCREMENTO_RADIO_METROS, RADIO_MAXIMO_METROS, solicitud_id))
            intentos, radio_actual = cur.fetchone()

            cur.execute("""
                INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                VALUES (%s, 'buscando_proveedor', 'buscando_proveedor', 'sistema')
            """, (solicitud_id,))
            conn.commit()
        finally:
            cur.close(); conn.close()
        vencidas += 1

        # d) Si se agotaron los intentos → sin cobertura (marca + notifica cliente y coordinación)
        if intentos >= MAX_INTENTOS_ASIGNACION:
            if _marcar_sin_cobertura(solicitud_id):
                sin_cobertura.append(solicitud_id)
        else:
            resultado = buscar_y_ofrecer_proveedor(solicitud_id, radio_metros=radio_actual)
            reasignadas.append({"solicitud_id": solicitud_id, "resultado": resultado})

    # 2) Programadas que entran en su ventana de búsqueda
    programadas_activadas = _activar_programadas()

    # 3) Alerta temprana de cobertura para programadas
    alertas_cobertura = _alertar_cobertura_programadas()

    # 4) Solicitudes huérfanas: buscando_proveedor SIN oferta pendiente
    huerfanas_reintentadas, huerfanas_sin_cobertura = _reintentar_solicitudes_sin_oferta()
    sin_cobertura.extend(huerfanas_sin_cobertura)

    # 5) Servicios del call center sin cobertura que ya tienen un doctor conectado a su alcance
    reabiertas = _reabrir_sin_cobertura()

    return (json.dumps({
        "reabiertas": reabiertas,
        "ofertas_empresa_vencidas": len(empresas_vencidas),
        "autodesconectados": len(desconectados),
        "ofertas_vencidas": vencidas,
        "reasignadas": len(reasignadas),
        "programadas_activadas": len(programadas_activadas),
        "alertas_cobertura": len(alertas_cobertura),
        "huerfanas_reintentadas": len(huerfanas_reintentadas),
        "sin_cobertura": sin_cobertura,
    }, ensure_ascii=False), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# ADMINISTRACIÓN (solo coordinadores)
# ═════════════════════════════════════════════════════════════
def _mad_doctor_en_servicio(cur, proveedor_id):
    cur.execute("""
        SELECT 1 FROM solicitudes
        WHERE proveedor_id = %s AND estado IN ('asignado', 'en_camino', 'llego', 'en_atencion')
        LIMIT 1
    """, (proveedor_id,))
    return cur.fetchone() is not None


def _mad_desactivar_usuario(cur, usuario_id, proveedor_id):
    """Desactiva el usuario. NO hace commit. Devuelve False si el doctor tiene un servicio
    en curso: en ese caso el llamador debe hacer rollback."""
    if proveedor_id:
        # 1) Primero las ofertas pendientes. Si el doctor justo la estaba aceptando, ese servicio
        #    ya cuenta como en curso en el paso 2.
        cur.execute("""
            UPDATE solicitudes_ofertas SET estado = 'cancelada', respondido_en = NOW()
            WHERE proveedor_id = %s AND estado = 'pendiente'
        """, (proveedor_id,))
        if _mad_doctor_en_servicio(cur, proveedor_id):
            return False

    cur.execute("UPDATE mad_usuarios SET activo = FALSE WHERE id = %s", (usuario_id,))
    cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE usuario_id = %s", (usuario_id,))
    cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (usuario_id,))
    if proveedor_id:
        cur.execute("DELETE FROM proveedores_push WHERE proveedor_id = %s", (proveedor_id,))
        cur.execute("""
            UPDATE proveedores_estado SET estado = 'desconectado'
            WHERE proveedor_id = %s AND estado IN ('disponible', 'en_oferta')
        """, (proveedor_id,))
    return True


_MSG_SERVICIO_EN_CURSO = "Este doctor tiene un servicio en curso. Podrás desactivarlo cuando lo termine."


@con_log_errores
def handle_mad_admin_listar(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT u.id, u.usuario, u.nombre, u.rol, u.proveedor_id, p.telefono, p.tipo_servicio,
                   u.activo, u.totp_activo, u.debe_cambiar_password,
                   COALESCE(u.bloqueado_hasta > NOW(), FALSE), u.ultimo_login,
                   COALESCE(p.empresa_id, u.empresa_id) AS empresa_id, e.nombre, p.cmp
            FROM mad_usuarios u
            LEFT JOIN proveedores p ON p.id = u.proveedor_id
            LEFT JOIN mad_empresas e ON e.id = COALESCE(p.empresa_id, u.empresa_id)
            ORDER BY u.rol, e.nombre NULLS LAST, u.nombre
        """)
        usuarios = [{
            "id": r[0], "usuario": r[1], "nombre": r[2], "rol": r[3], "proveedor_id": r[4],
            "telefono": r[5], "tipo_servicio": r[6], "activo": r[7], "totp_activo": r[8],
            "debe_cambiar_password": r[9], "bloqueado": r[10],
            "ultimo_login": _iso(r[11]),
            "empresa_id": r[12], "empresa": r[13], "cmp": r[14],
        } for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
    return _mad_json({"usuarios": usuarios})


def _normalizar_cmp(valor, obligatorio):
    """CMP (Colegio Médico del Perú): solo dígitos, 4 a 6. Devuelve (cmp|None, error)."""
    cmp_ = re.sub(r"\D", "", str(valor or ""))
    if not cmp_:
        return (None, "El código CMP es obligatorio para doctores de médico a domicilio") if obligatorio else (None, None)
    if not re.fullmatch(r"\d{4,6}", cmp_):
        return None, "El CMP debe tener entre 4 y 6 dígitos"
    return cmp_, None


def _datos_nuevo_doctor(data):
    """Valida los datos comunes para crear un doctor. Devuelve (datos, error)."""
    nombre   = (data.get("nombre") or "").strip()[:120]
    usuario  = (data.get("usuario") or "").strip().lower()
    telefono = re.sub(r"\D", "", data.get("telefono") or "")
    if len(telefono) == 9 and telefono.startswith("9"):
        telefono = "51" + telefono
    tipo = (data.get("tipo_servicio") or "").strip().upper()
    if (not nombre or tipo not in ("MAD", "AMBULANCIA", "AMBOS")
            or not re.fullmatch(r"[a-z0-9._-]{4,30}", usuario)
            or not re.fullmatch(r"51\d{9}", telefono)):
        return None, "Revisa los datos: nombre, usuario (4 a 30 caracteres: letras, números, punto, guion), celular peruano y tipo de servicio"
    cmp_, err_cmp = _normalizar_cmp(data.get("cmp"), obligatorio=(tipo in ("MAD", "AMBOS")))
    if err_cmp:
        return None, err_cmp
    return (nombre, usuario, telefono, tipo, cmp_), None


@con_log_errores
def handle_mad_admin_crear_proveedor(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err

    data = request.get_json(force=True, silent=True) or {}
    datos, error = _datos_nuevo_doctor(data)
    if error:
        return _mad_json({"error": error}, 400)
    nombre, usuario, telefono, tipo, cmp_ = datos
    empresa_id = data.get("empresa_id")
    empresa_id = int(empresa_id) if str(empresa_id or "").isdigit() else None

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_usuarios WHERE usuario = %s", (usuario,))
        if cur.fetchone():
            return _mad_json({"error": "Ese usuario ya existe"}, 409)
        cur.execute("SELECT 1 FROM proveedores WHERE telefono = %s", (telefono,))
        if cur.fetchone():
            return _mad_json({"error": "Ese celular ya está registrado en otro doctor"}, 409)
        if cmp_:
            cur.execute("SELECT 1 FROM proveedores WHERE cmp = %s", (cmp_,))
            if cur.fetchone():
                return _mad_json({"error": "Ese CMP ya está registrado en otro doctor"}, 409)

        if empresa_id is not None:
            cur.execute("SELECT 1 FROM mad_empresas WHERE id = %s AND activo", (empresa_id,))
            if not cur.fetchone():
                return _mad_json({"error": "La empresa seleccionada no existe o está inactiva"}, 400)

        cur.execute("""
            INSERT INTO proveedores (nombre, telefono, tipo_servicio, empresa_id, cmp)
            VALUES (%s, %s, %s, %s, %s) RETURNING id
        """, (nombre, telefono, tipo, empresa_id, cmp_))
        proveedor_id = cur.fetchone()[0]

        cur.execute("""
            INSERT INTO proveedores_estado (proveedor_id, estado, actualizado_en)
            VALUES (%s, 'desconectado', NOW())
        """, (proveedor_id,))

        cur.execute("""
            INSERT INTO mad_usuarios (usuario, nombre, rol, proveedor_id, password_hash)
            VALUES (%s, %s, 'proveedor', %s, %s)
        """, (usuario, nombre, proveedor_id, hash_password(temporal)))
        _mad_audit(cur, s["id"], "doctor_crear", f"proveedor:{proveedor_id}", {"usuario": usuario, "empresa_id": empresa_id, "cmp": cmp_})
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "proveedor_id": proveedor_id, "usuario": usuario,
                      "password_temporal": temporal}, 201)


@con_log_errores
def handle_mad_admin_reset(request, usuario_id):
    """Para cuando el usuario olvidó su clave o perdió el celular:
    nueva clave temporal + debe volver a configurar Google Authenticator."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(usuario_id).isdigit():
        return _mad_json({"error": "Usuario inválido"}, 400)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE mad_usuarios
            SET password_hash = %s, debe_cambiar_password = TRUE,
                totp_secret_enc = NULL, totp_activo = FALSE, totp_ultimo_paso = NULL,
                intentos_fallidos = 0, bloqueado_hasta = NULL
            WHERE id = %s
            RETURNING usuario
        """, (hash_password(temporal), int(usuario_id)))
        r = cur.fetchone()
        if not r:
            return _mad_json({"error": "Usuario no encontrado"}, 404)
        cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE usuario_id = %s", (int(usuario_id),))
        cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (int(usuario_id),))
        _mad_audit(cur, s["id"], "usuario_reset", f"usuario:{int(usuario_id)}")
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "usuario": r[0], "password_temporal": temporal})


@con_log_errores
def handle_mad_admin_estado(request, usuario_id):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(usuario_id).isdigit():
        return _mad_json({"error": "Usuario inválido"}, 400)
    uid = int(usuario_id)
    if uid == s["id"]:
        return _mad_json({"error": "No puedes desactivar tu propio usuario"}, 400)

    activo = bool((request.get_json(force=True, silent=True) or {}).get("activo"))
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT proveedor_id FROM mad_usuarios WHERE id = %s FOR UPDATE", (uid,))
        r = cur.fetchone()
        if not r:
            return _mad_json({"error": "Usuario no encontrado"}, 404)

        if activo:
            cur.execute("UPDATE mad_usuarios SET activo = TRUE WHERE id = %s", (uid,))
        elif not _mad_desactivar_usuario(cur, uid, r[0]):
            conn.rollback()
            return _mad_json({"error": _MSG_SERVICIO_EN_CURSO}, 409)
        _mad_audit(cur, s["id"], "usuario_activar" if activo else "usuario_desactivar", f"usuario:{uid}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "activo": activo})


@con_log_errores
def handle_mad_admin_empresas_listar(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT e.id, e.nombre, e.ruc, e.activo,
                   (SELECT COUNT(*) FROM proveedores p WHERE p.empresa_id = e.id)
            FROM mad_empresas e ORDER BY e.nombre
        """)
        empresas = [{"id": r[0], "nombre": r[1], "ruc": r[2], "activo": r[3], "doctores": r[4]}
                    for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
    return _mad_json({"empresas": empresas})


@con_log_errores
def handle_mad_admin_crear_empresa(request):
    """Crea la empresa y su primer usuario de acceso (clave temporal + 2FA en el primer ingreso)."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err

    data = request.get_json(force=True, silent=True) or {}
    nombre   = (data.get("nombre") or "").strip()[:120]
    ruc      = re.sub(r"\D", "", data.get("ruc") or "")
    usuario  = (data.get("usuario") or "").strip().lower()
    contacto = (data.get("contacto") or "").strip()[:120] or nombre

    if not nombre or not re.fullmatch(r"[a-z0-9._-]{4,30}", usuario) or (ruc and len(ruc) != 11):
        return _mad_json({"error": "Revisa los datos: nombre, RUC (11 dígitos, opcional) y usuario (4 a 30 caracteres)"}, 400)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_usuarios WHERE usuario = %s", (usuario,))
        if cur.fetchone():
            return _mad_json({"error": "Ese usuario ya existe"}, 409)

        cur.execute("INSERT INTO mad_empresas (nombre, ruc) VALUES (%s, %s) RETURNING id",
                    (nombre, ruc or None))
        empresa_id = cur.fetchone()[0]

        cur.execute("""
            INSERT INTO mad_usuarios (usuario, nombre, rol, empresa_id, password_hash)
            VALUES (%s, %s, 'empresa', %s, %s)
        """, (usuario, contacto, empresa_id, hash_password(temporal)))
        _mad_audit(cur, s["id"], "empresa_crear", f"empresa:{empresa_id}", {"nombre": nombre})
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "empresa_id": empresa_id, "usuario": usuario,
                      "password_temporal": temporal}, 201)


@con_log_errores
def handle_mad_admin_asignar_empresa(request, proveedor_id):
    """Asigna un doctor a una empresa (o lo deja independiente con empresa_id = null).
    Solo si no tiene una oferta ni un servicio en curso."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Proveedor inválido"}, 400)
    pid = int(proveedor_id)

    empresa_id = (request.get_json(force=True, silent=True) or {}).get("empresa_id")
    empresa_id = int(empresa_id) if str(empresa_id or "").isdigit() else None

    conn = get_connection(); cur = conn.cursor()
    try:
        if empresa_id is not None:
            cur.execute("SELECT 1 FROM mad_empresas WHERE id = %s AND activo", (empresa_id,))
            if not cur.fetchone():
                return _mad_json({"error": "La empresa no existe o está inactiva"}, 400)

        cur.execute("SELECT 1 FROM proveedores WHERE id = %s FOR UPDATE", (pid,))
        if not cur.fetchone():
            return _mad_json({"error": "Proveedor no encontrado"}, 404)

        cur.execute("SELECT estado FROM proveedores_estado WHERE proveedor_id = %s", (pid,))
        r = cur.fetchone()
        if _mad_doctor_en_servicio(cur, pid) or (r and r[0] == "en_oferta"):
            return _mad_json({"error": "Este doctor tiene una oferta o un servicio en curso. Cámbialo de empresa cuando termine."}, 409)

        cur.execute("UPDATE proveedores SET empresa_id = %s WHERE id = %s", (empresa_id, pid))
        _mad_audit(cur, s["id"], "doctor_empresa", f"proveedor:{pid}", {"empresa_id": empresa_id})
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "empresa_id": empresa_id})


def _mad_editar_doctor(cur, pid, data, actor_id, empresa_id=None):
    """Edita nombre, celular y tipo de servicio. NO hace commit (lo hace el llamador).
    Si empresa_id viene, el doctor debe pertenecer a esa empresa.
    Devuelve (status, payload, avisos); avisos = [(telefono, texto)]."""
    if empresa_id is not None:
        cur.execute("""SELECT nombre, telefono, tipo_servicio, cmp FROM proveedores
                       WHERE id = %s AND empresa_id = %s FOR UPDATE""", (pid, empresa_id))
    else:
        cur.execute("SELECT nombre, telefono, tipo_servicio, cmp FROM proveedores WHERE id = %s FOR UPDATE", (pid,))
    row = cur.fetchone()
    if not row:
        return 404, {"error": "Doctor no encontrado"}, []
    nombre_act, tel_act, tipo_act, cmp_act = row
    tel_act_digitos = re.sub(r"\D", "", tel_act or "")

    nombre = nombre_act
    if "nombre" in data:
        nombre = str(data.get("nombre") or "").strip()[:120]
        if len(nombre) < 3:
            return 400, {"error": "Escribe el nombre del doctor"}, []

    telefono = tel_act_digitos
    if "telefono" in data:
        t = re.sub(r"\D", "", str(data.get("telefono") or ""))
        if len(t) == 9 and t.startswith("9"):
            t = "51" + t
        if not re.fullmatch(r"51\d{9}", t):
            return 400, {"error": "El celular no es válido (9 dígitos, empieza con 9)"}, []
        telefono = t

    tipo = tipo_act
    if "tipo_servicio" in data:
        tipo = str(data.get("tipo_servicio") or "").strip().upper()
        if tipo not in ("MAD", "AMBULANCIA", "AMBOS"):
            return 400, {"error": "Tipo de servicio inválido"}, []

    cmp_ = cmp_act
    if "cmp" in data:
        cmp_, err_cmp = _normalizar_cmp(data.get("cmp"), obligatorio=(tipo in ("MAD", "AMBOS")))
        if err_cmp:
            return 400, {"error": err_cmp}, []

    cambios = []
    if (cmp_ or None) != (cmp_act or None): cambios.append("cmp")
    if nombre != nombre_act:            cambios.append("nombre")
    if telefono != tel_act_digitos:     cambios.append("telefono")
    if tipo != tipo_act:                cambios.append("tipo_servicio")
    if not cambios:
        return 200, {"ok": True, "cambios": []}, []

    if "tipo_servicio" in cambios:
        cur.execute("SELECT estado FROM proveedores_estado WHERE proveedor_id = %s", (pid,))
        est = cur.fetchone()
        if _mad_doctor_en_servicio(cur, pid) or (est and est[0] == "en_oferta"):
            return 409, {"error": "Este doctor tiene una oferta o un servicio en curso. Cambia su tipo de servicio cuando termine."}, []

    if "telefono" in cambios:
        cur.execute("SELECT 1 FROM proveedores WHERE telefono = %s AND id <> %s", (telefono, pid))
        if cur.fetchone():
            return 409, {"error": "Ese celular ya está registrado en otro doctor"}, []

    if "cmp" in cambios and cmp_:
        cur.execute("SELECT 1 FROM proveedores WHERE cmp = %s AND id <> %s", (cmp_, pid))
        if cur.fetchone():
            return 409, {"error": "Ese CMP ya está registrado en otro doctor"}, []

    cur.execute("UPDATE proveedores SET nombre = %s, telefono = %s, tipo_servicio = %s, cmp = %s WHERE id = %s",
                (nombre, telefono, tipo, cmp_, pid))
    if "nombre" in cambios:
        cur.execute("UPDATE mad_usuarios SET nombre = %s WHERE proveedor_id = %s AND rol = 'proveedor'", (nombre, pid))

    detalle = {}
    if "nombre" in cambios:        detalle["nombre"] = [nombre_act, nombre]
    if "telefono" in cambios:      detalle["telefono"] = [tel_act_digitos, telefono]
    if "tipo_servicio" in cambios: detalle["tipo_servicio"] = [tipo_act, tipo]
    if "cmp" in cambios:           detalle["cmp"] = [cmp_act, cmp_]
    _mad_audit(cur, actor_id, "doctor_editar", f"proveedor:{pid}", detalle)

    avisos = []
    if "telefono" in cambios:
        avisos.append((telefono, "Sanitas: Este celular fue registrado para recibir solicitudes de servicio como profesional. Si no lo reconoces, avisa a coordinacion."))
        if tel_act_digitos:
            avisos.append((tel_act_digitos, "Sanitas: Tu celular de contacto fue cambiado. Ya no recibiras solicitudes en este numero."))
    return 200, {"ok": True, "cambios": cambios}, avisos


@con_log_errores
def handle_mad_admin_doctor_editar(request, proveedor_id):
    """Coordinador: edita cualquier doctor (útil para independientes, que no tienen empresa)."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Doctor inválido"}, 400)
    data = request.get_json(force=True, silent=True) or {}

    conn = get_connection(); cur = conn.cursor()
    try:
        status, payload, avisos = _mad_editar_doctor(cur, int(proveedor_id), data, s["id"], None)
        if status == 200:
            conn.commit()
        else:
            conn.rollback()
    finally:
        cur.close(); conn.close()

    _mad_enviar_avisos(avisos, "mad doctor celular")
    return _mad_json(payload, status)


@con_log_errores
def handle_mad_admin_liberar_servicio(request, solicitud_id):
    """Quita el servicio al doctor asignado y lo devuelve a la búsqueda, excluyéndolo.
    Solo coordinadores. Pensado para doctores que desaparecen en pleno servicio."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(solicitud_id).isdigit():
        return _mad_json({"error": "Solicitud inválida"}, 400)
    sid = int(solicitud_id)

    data = request.get_json(force=True, silent=True) or {}
    motivo = (data.get("motivo") or "").strip()
    comentario = (data.get("comentario") or "").strip()[:300] or None
    avisar_paciente = bool(data.get("avisar_paciente", True))

    if motivo not in MOTIVOS_LIBERACION:
        return _mad_json({"error": "Selecciona un motivo"}, 400)
    if motivo == "otro" and not comentario:
        return _mad_json({"error": "Describe brevemente el motivo"}, 400)

    telefono_doctor = telefono_cliente = None
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT estado, proveedor_id, cliente_telefono
            FROM solicitudes WHERE id = %s
            FOR UPDATE
        """, (sid,))
        row = cur.fetchone()
        if not row:
            return _mad_json({"error": "Solicitud no encontrada"}, 404)

        estado_anterior, proveedor_anterior, telefono_cliente = row

        if estado_anterior == "en_atencion":
            conn.rollback()
            return _mad_json({"error": "El servicio ya está en atención: no se puede liberar. Si quedó abierto por error, contacta al doctor para que lo finalice."}, 409)
        if estado_anterior not in ESTADOS_LIBERABLES or not proveedor_anterior:
            conn.rollback()
            return _mad_json({"error": "Solo se puede liberar un servicio asignado a un doctor", "estado": estado_anterior}, 409)

        # El doctor queda excluido de ESTA solicitud
        cur.execute("""
            INSERT INTO solicitudes_proveedores_excluidos (solicitud_id, proveedor_id, motivo)
            VALUES (%s, %s, 'liberado_por_coordinador')
            ON CONFLICT DO NOTHING
        """, (sid, proveedor_anterior))

        # Auditoría
        cur.execute("""
            INSERT INTO mad_liberaciones (solicitud_id, proveedor_id, coordinador_id, estado_anterior, motivo, comentario)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (sid, proveedor_anterior, s["id"], estado_anterior, motivo, comentario))

        # No recibe ofertas hasta que él mismo se reconecte
        cur.execute("""
            UPDATE proveedores_estado SET estado = 'desconectado' WHERE proveedor_id = %s
        """, (proveedor_anterior,))

        # La solicitud vuelve a la búsqueda, sin rastro del viaje anterior
        cur.execute("""
            UPDATE solicitudes
            SET estado = 'buscando_proveedor', proveedor_id = NULL, empresa_id = NULL, asignado_en = NULL,
                en_camino_en = NULL, en_camino_ubicacion = NULL, en_camino_eta_min = NULL, llego_en = NULL,
                intentos_asignacion = 0, radio_busqueda_metros = 5000, ultima_busqueda_en = NULL
            WHERE id = %s
        """, (sid,))

        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, %s, 'buscando_proveedor', 'coordinador')
        """, (sid, estado_anterior))

        cur.execute("SELECT telefono FROM proveedores WHERE id = %s", (proveedor_anterior,))
        r = cur.fetchone()
        telefono_doctor = r[0] if r else None

        conn.commit()
    finally:
        cur.close(); conn.close()

    # Avisos fuera de la transacción
    numero = _normalizar_celular_pe(telefono_doctor)
    if numero:
        enviar_danaconnect(numero, f"Sanitas: El servicio #{sid} fue reasignado por coordinacion. No acudas al domicilio. Ante dudas comunicate con coordinacion.")
        sms_aviso(numero, "mad liberado doctor")
    if avisar_paciente:
        numero = _normalizar_celular_pe(telefono_cliente)
        if numero:
            enviar_danaconnect(numero, "Sanitas: Estamos asignando otro profesional a tu servicio. Puedes seguir el avance en tu enlace de seguimiento.")
            sms_aviso(numero, "mad liberado cliente")

    resultado = buscar_y_ofrecer_proveedor(sid)
    return _mad_json({"ok": True, "estado": "buscando_proveedor", "busqueda": resultado})


@con_log_errores
def handle_mad_admin_liberaciones_proveedor(request, proveedor_id):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Doctor inválido"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT nombre FROM proveedores WHERE id = %s", (int(proveedor_id),))
        p = cur.fetchone()
        if not p:
            return _mad_json({"error": "Doctor no encontrado"}, 404)
        cur.execute("""
            SELECT l.creado_en, l.solicitud_id, l.estado_anterior, l.motivo, l.comentario, u.nombre
            FROM mad_liberaciones l
            JOIN mad_usuarios u ON u.id = l.coordinador_id
            WHERE l.proveedor_id = %s AND l.creado_en > NOW() - INTERVAL '90 days'
            ORDER BY l.creado_en DESC LIMIT 50
        """, (int(proveedor_id),))
        filas = cur.fetchall()
    finally:
        cur.close(); conn.close()

    return _mad_json({
        "doctor": p[0],
        "liberaciones": [{
            "fecha": _iso(f), "solicitud_id": sid, "estado_anterior": ea,
            "motivo": MOTIVOS_LIBERACION.get(m, m), "comentario": c, "coordinador": coord,
        } for f, sid, ea, m, c, coord in filas],
    })


MAD_MAX_SMS_TRACKING = 5     # reenvíos por servicio


@con_log_errores
def handle_mad_admin_ofrecer_a_proveedor(request, solicitud_id):
    """Coordinador: ofrece un servicio a la EMPRESA de un doctor en línea (sin importar peso ni distancia).
    La empresa acepta y asigna. Cancela cualquier oferta/aceptación previa de la solicitud."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    data = request.get_json(force=True, silent=True) or {}
    if not str(solicitud_id).isdigit() or not str(data.get("proveedor_id") or "").isdigit():
        return _mad_json({"error": "Datos inválidos"}, 400)
    sid, pid = int(solicitud_id), int(data["proveedor_id"])

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""SELECT estado, tipo_servicio, origen_distrito, programada_para,
                              ST_Distance(origen, (SELECT ubicacion FROM proveedores_estado WHERE proveedor_id = %s))
                       FROM solicitudes WHERE id = %s FOR UPDATE""", (pid, sid))
        row = cur.fetchone()
        if not row:
            return _mad_json({"error": "Solicitud no encontrada"}, 404)
        estado, tipo, distrito, programada_para, dist_m = row
        if estado not in ("buscando_proveedor", "sin_cobertura"):
            return _mad_json({"error": "Solo se puede ofrecer un servicio que esté buscando doctor o sin cobertura"}, 409)

        cur.execute("""SELECT p.empresa_id, p.tipo_servicio, pe.estado, e.celular_alertas
                       FROM proveedores p JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                       LEFT JOIN mad_empresas e ON e.id = p.empresa_id AND e.activo
                       WHERE p.id = %s""", (pid,))
        d = cur.fetchone()
        if not d:
            conn.rollback()
            return _mad_json({"error": "Doctor no encontrado"}, 404)
        empresa_id, tipo_doc, estado_doc, celular_alertas = d
        if not empresa_id:
            conn.rollback()
            return _mad_json({"error": "El doctor no tiene empresa asignada"}, 409)
        if estado_doc != "disponible":
            conn.rollback()
            return _mad_json({"error": "El doctor no está disponible en línea (puede estar ocupado o desconectado)"}, 409)
        if tipo_doc not in (tipo, "AMBOS"):
            conn.rollback()
            return _mad_json({"error": "El doctor no atiende este tipo de servicio"}, 409)

        cur.execute("""UPDATE solicitudes_ofertas_empresa SET estado = 'cancelada', respondido_en = NOW()
                       WHERE solicitud_id = %s AND estado IN ('pendiente', 'aceptada')""", (sid,))
        if estado == "sin_cobertura":
            cur.execute("""UPDATE solicitudes SET estado = 'buscando_proveedor', intentos_asignacion = 0,
                               radio_busqueda_metros = 5000 WHERE id = %s""", (sid,))
            cur.execute("""INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
                           VALUES (%s, 'sin_cobertura', 'buscando_proveedor', 'coordinador')""", (sid,))
        cur.execute("""UPDATE solicitudes SET empresa_aceptada_en = NULL, ultima_busqueda_en = %s WHERE id = %s""",
                    (datetime.now(LIMA_TZ).replace(tzinfo=None), sid))
        cur.execute("""INSERT INTO solicitudes_ofertas_empresa (solicitud_id, empresa_id, estado)
                       VALUES (%s, %s, 'pendiente')""", (sid, empresa_id))
        _mad_audit(cur, s["id"], "oferta_manual", f"solicitud:{sid}",
                   {"proveedor_id": pid, "empresa_id": empresa_id,
                    "distancia_km": round(dist_m / 1000, 1) if dist_m is not None else None})
        conn.commit()
    finally:
        cur.close(); conn.close()

    avisar_empresa_oferta(celular_alertas, tipo, distrito, programada_para)
    return _mad_json({"ok": True, "distancia_km": round(dist_m / 1000, 1) if dist_m is not None else None})


@con_log_errores
def handle_mad_admin_enviar_tracking(request, solicitud_id):
    """Coordinador: envía (o reenvía) por SMS el enlace de seguimiento al cliente."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(solicitud_id).isdigit():
        return _mad_json({"error": "Solicitud inválida"}, 400)
    sid = int(solicitud_id)
    data = request.get_json(force=True, silent=True) or {}

    tel_alt = re.sub(r"\D", "", str(data.get("telefono") or ""))
    if tel_alt:
        if len(tel_alt) == 9 and tel_alt.startswith("9"):
            tel_alt = "51" + tel_alt
        if not re.fullmatch(r"51\d{9}", tel_alt):
            return _mad_json({"error": "El celular no es válido (9 dígitos, empieza con 9)"}, 400)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT estado, tracking_token, cliente_telefono, tipo_servicio FROM solicitudes WHERE id = %s", (sid,))
        row = cur.fetchone()
        if not row:
            return _mad_json({"error": "Solicitud no encontrada"}, 404)
        estado, token, tel_cliente, tipo = row
        if estado not in ESTADOS_ABIERTOS:
            return _mad_json({"error": "El servicio ya terminó; no tiene sentido enviar el seguimiento"}, 409)
        cur.execute("SELECT COUNT(*) FROM mad_auditoria WHERE accion = 'tracking_sms' AND objetivo = %s", (f"solicitud:{sid}",))
        if cur.fetchone()[0] >= MAD_MAX_SMS_TRACKING:
            return _mad_json({"error": f"Ya se envió el seguimiento {MAD_MAX_SMS_TRACKING} veces para este servicio"}, 429)
        numero = _normalizar_celular_pe(tel_alt or tel_cliente)
        if not numero:
            return _mad_json({"error": "El servicio no tiene un celular válido. Escribe uno."}, 400)
        _mad_audit(cur, s["id"], "tracking_sms", f"solicitud:{sid}", {"destino_final": numero[-3:], "alterno": bool(tel_alt)})
        conn.commit()
    finally:
        cur.close(); conn.close()

    link = f"{_base_url()}/track/{token}"
    nombre = "Medico a domicilio" if tipo == "MAD" else "Ambulancia"
    enviar_danaconnect(numero, f"Sanitas: Sigue tu servicio de {nombre} y la ubicacion del profesional aqui: {link}")
    sms_aviso(numero, "mad tracking monitor")
    return _mad_json({"ok": True, "enviado_a": "***" + numero[-3:]})


@con_log_errores
def handle_mad_admin_activar_programada(request, solicitud_id):
    """Coordinador: empezar a buscar doctor ya, sin esperar a la hora de activación."""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    if not str(solicitud_id).isdigit():
        return _mad_json({"error": "Solicitud inválida"}, 400)
    sid = int(solicitud_id)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE solicitudes
            SET estado = 'buscando_proveedor', intentos_asignacion = 0,
                radio_busqueda_metros = 5000, ultima_busqueda_en = NULL
            WHERE id = %s AND estado IN ('programada', 'sin_cobertura')
            RETURNING id
        """, (sid,))
        if not cur.fetchone():
            return _mad_json({"error": "La solicitud no está programada ni sin cobertura"}, 409)
        cur.execute("""
            INSERT INTO solicitudes_historial_estado (solicitud_id, estado_anterior, estado_nuevo, cambiado_por)
            VALUES (%s, 'programada', 'buscando_proveedor', 'coordinador')
        """, (sid,))
        conn.commit()
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "busqueda": buscar_y_ofrecer_proveedor(sid)})


# ═════════════════════════════════════════════════════════════
# MONITOR (coordinadores)
# ═════════════════════════════════════════════════════════════
@con_log_errores
def handle_mad_monitor_datos(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err

    conn = get_connection(); cur = conn.cursor()
    try:
        # f-string: usa MAD_COBERTURA_SENAL_MIN y RADIO_MAXIMO_METROS (constantes numéricas del código)
        cur.execute(f"""
            SELECT s.id, s.estado, s.tipo_servicio, s.cliente_nombre, s.cliente_telefono,
                   s.origen_direccion_texto, s.origen_distrito, s.origen_referencia,
                   ST_Y(s.origen::geometry) AS lat, ST_X(s.origen::geometry) AS lng,
                   s.creado_en, s.asignado_en, s.intentos_asignacion, s.radio_busqueda_metros,
                   p.nombre AS proveedor_asignado, s.tracking_token,
                   (SELECT CASE WHEN o.estado = 'aceptada' THEN e2.nombre || ' (aceptó, falta asignar doctor)'
                                ELSE e2.nombre || ' (esperando respuesta)' END
                      FROM solicitudes_ofertas_empresa o
                      JOIN mad_empresas e2 ON e2.id = o.empresa_id
                    WHERE o.solicitud_id = s.id AND o.estado IN ('pendiente', 'aceptada') AND s.estado = 'buscando_proveedor'
                    ORDER BY o.id DESC LIMIT 1) AS ofrecido_a,
                   s.proveedor_id,
                   s.programada_para, s.contacto_nombre, s.contacto_telefono,
                   COALESCE(s.reprogramaciones, 0),
                   p.cmp, em.nombre,
                   (SELECT json_build_object('tipo', n.tipo, 'texto', n.texto, 'creado_en', n.creado_en)
                      FROM solicitudes_novedades n WHERE n.solicitud_id = s.id ORDER BY n.id DESC LIMIT 1),
                   s.canal, s.paciente_dni, s.paciente_edad, s.sintomas, s.antecedentes, s.alergias,
                   s.copago_monto, s.forma_pago,
                   (SELECT COUNT(*) FROM proveedores p3
                      JOIN proveedores_estado pe3 ON pe3.proveedor_id = p3.id
                      JOIN mad_usuarios u3 ON u3.proveedor_id = p3.id AND u3.activo
                    WHERE s.estado = 'programada'
                      AND pe3.estado IN ('disponible', 'en_oferta', 'ocupado')
                      AND p3.tipo_servicio IN (s.tipo_servicio, 'AMBOS')
                      AND pe3.ubicacion IS NOT NULL
                      AND pe3.actualizado_en > NOW() - INTERVAL '{MAD_COBERTURA_SENAL_MIN} minutes'
                      AND ST_DWithin(pe3.ubicacion, s.origen, {RADIO_MAXIMO_METROS})) AS conectados_cerca
            FROM solicitudes s
            LEFT JOIN proveedores p ON p.id = s.proveedor_id
            LEFT JOIN mad_empresas em ON em.id = p.empresa_id
            WHERE s.estado IN ('programada','buscando_proveedor','asignado','en_camino','llego','en_atencion')
               OR (s.estado = 'sin_cobertura' AND s.creado_en > NOW() - INTERVAL '24 hours')
            ORDER BY s.id DESC
        """)
        ahora = datetime.now(LIMA_TZ)
        solicitudes = []
        for r in cur.fetchall():
            programada_para = r[18]
            riesgo = bool(
                r[1] == "programada" and programada_para is not None and r[33] == 0
                and programada_para - ahora <= timedelta(hours=MAD_COBERTURA_HORAS_1)
            )
            solicitudes.append({
                "id": r[0], "estado": r[1], "tipo_servicio": r[2],
                "cliente_nombre": r[3], "cliente_telefono": r[4],
                "direccion": r[5], "distrito": r[6], "referencia": r[7],
                "lat": r[8], "lng": r[9],
                "creado_en": _iso(r[10]), "asignado_en": _iso(r[11]),
                "intentos": r[12] or 0, "radio": r[13],
                "proveedor_asignado": r[14],
                "tracking_token": str(r[15]) if r[15] else None,
                "ofrecido_a": r[16], "proveedor_id": r[17],
                "programada_para": _iso(programada_para),
                "contacto_nombre": r[19], "contacto_telefono": r[20],
                "reprogramaciones": r[21],
                "cobertura_riesgo": riesgo,
                "proveedor_cmp": r[22], "proveedor_empresa": r[23],
                "novedad": ({"tipo": r[24]["tipo"], "texto": r[24]["texto"], "creado_en": r[24]["creado_en"]}
                            if r[24] else None),
                "canal": r[25], "paciente_dni": r[26], "paciente_edad": r[27],
                "sintomas": r[28], "antecedentes": r[29], "alergias": r[30],
                "copago_monto": float(r[31]) if r[31] is not None else None, "forma_pago": r[32],
            })

        cur.execute("""
            SELECT p.id, p.nombre, p.telefono, p.tipo_servicio, pe.estado,
                   ST_Y(pe.ubicacion::geometry) AS lat, ST_X(pe.ubicacion::geometry) AS lng,
                   pe.actualizado_en,
                   COALESCE(lib.l7, 0), COALESCE(lib.l30, 0), lib.ultima
            FROM proveedores p
            JOIN proveedores_estado pe ON pe.proveedor_id = p.id
            LEFT JOIN (
                SELECT proveedor_id,
                       COUNT(*) FILTER (WHERE creado_en > NOW() - INTERVAL '7 days')  AS l7,
                       COUNT(*) FILTER (WHERE creado_en > NOW() - INTERVAL '30 days') AS l30,
                       MAX(creado_en) AS ultima
                FROM mad_liberaciones
                GROUP BY proveedor_id
            ) lib ON lib.proveedor_id = p.id
            ORDER BY p.nombre
        """)
        proveedores = [{
            "id": r[0], "nombre": r[1], "telefono": r[2], "tipo_servicio": r[3],
            "estado": r[4], "lat": r[5], "lng": r[6], "actualizado_en": _iso(r[7]),
            "lib_7": int(r[8]), "lib_30": int(r[9]), "lib_ultima": _iso(r[10]),
            "alerta_lib": int(r[8]) >= MAD_LIB_ALERTA_7D or int(r[9]) >= MAD_LIB_ALERTA_30D,
        } for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()

    return (json.dumps({
        "ahora": datetime.now(LIMA_TZ).isoformat(),
        "umbral_senal_min": MAD_AUTODESCONEXION_MIN,
        "solicitudes": solicitudes,
        "proveedores": proveedores,
    }, ensure_ascii=False, default=str), 200, HEADERS)


# ═════════════════════════════════════════════════════════════
# ESTADÍSTICAS DE SERVICIOS
# ═════════════════════════════════════════════════════════════
def _mad_rango(request):
    """Rango de fechas (hora Lima). Por defecto: del 1 del mes a hoy. Máximo 1 año."""
    hoy = datetime.now(LIMA_TZ).date()

    def parse(v, defecto):
        try:
            return datetime.strptime(v, "%Y-%m-%d").date()
        except (TypeError, ValueError):
            return defecto

    desde = parse(request.args.get("desde"), hoy.replace(day=1))
    hasta = parse(request.args.get("hasta"), hoy)
    if desde > hasta:
        desde, hasta = hasta, desde
    if (hasta - desde).days > 366:
        desde = hasta - timedelta(days=366)
    return hoy, desde, hasta


# ── Prioridades de empresas (monitor) ──
@con_log_errores
def handle_mad_admin_prioridades_listar(request):
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT e.id, e.nombre, e.activo, e.celular_alertas,
                   COALESCE(m.peso, 0), COALESCE(a.peso, 0),
                   (SELECT COUNT(*) FROM proveedores p WHERE p.empresa_id = e.id)
            FROM mad_empresas e
            LEFT JOIN mad_empresas_prioridad m ON m.empresa_id = e.id AND m.tipo_servicio = 'MAD'
            LEFT JOIN mad_empresas_prioridad a ON a.empresa_id = e.id AND a.tipo_servicio = 'AMBULANCIA'
            ORDER BY e.nombre
        """)
        filas = [{"empresa_id": r[0], "nombre": r[1], "activo": r[2], "celular_alertas": r[3],
                  "peso_mad": r[4], "peso_ambulancia": r[5], "doctores": r[6]} for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
    return _mad_json({"empresas": filas})


@con_log_errores
def handle_mad_admin_prioridades_guardar(request):
    """Body: {"empresas": [{"empresa_id":1,"peso_mad":10,"peso_ambulancia":0,"celular_alertas":"9XXXXXXXX"}, ...]}"""
    s, err = _mad_exigir(request, "coordinador")
    if err:
        return err
    data = request.get_json(force=True, silent=True) or {}
    items = data.get("empresas")
    if not isinstance(items, list) or not items:
        return _err("Lista de empresas vacía")
    conn = get_connection(); cur = conn.cursor()
    try:
        for it in items:
            try:
                eid = int(it.get("empresa_id"))
                pm = max(0, min(int(it.get("peso_mad") or 0), 1000))
                pa = max(0, min(int(it.get("peso_ambulancia") or 0), 1000))
            except (TypeError, ValueError):
                conn.rollback()
                return _err("Datos inválidos")
            cel_raw = str(it.get("celular_alertas") or "").strip()
            cel = _normalizar_celular_pe(cel_raw) if cel_raw else None
            if cel_raw and not cel:
                conn.rollback()
                return _err(f"Celular de alertas inválido: {cel_raw}")
            cur.execute("UPDATE mad_empresas SET celular_alertas = %s WHERE id = %s", (cel, eid))
            for tipo, peso in (("MAD", pm), ("AMBULANCIA", pa)):
                cur.execute("""
                    INSERT INTO mad_empresas_prioridad (empresa_id, tipo_servicio, peso) VALUES (%s, %s, %s)
                    ON CONFLICT (empresa_id, tipo_servicio) DO UPDATE SET peso = EXCLUDED.peso, actualizado_en = NOW()
                """, (eid, tipo, peso))
        _mad_audit(cur, s["id"], "prioridades_guardar", "empresas", {"n": len(items)})
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True})


# ── Ofertas en el portal de la empresa ──
@con_log_errores
def handle_mad_empresa_ofertas(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    eid = s["empresa_id"]
    conn = get_connection(); cur = conn.cursor()
    try:
        # Pendientes de aceptar: solo zona/tipo/distancia (sin datos del paciente)
        cur.execute("""
            SELECT o.id, s.id, s.tipo_servicio, s.origen_distrito, s.programada_para,
                   EXTRACT(EPOCH FROM (o.ofrecido_en + make_interval(secs => %s) - NOW()))::int,
                   (SELECT MIN(ST_Distance(pe.ubicacion, s.origen::geography)) FROM proveedores p
                      JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                     WHERE p.empresa_id = %s AND pe.estado = 'disponible' AND pe.ubicacion IS NOT NULL)
            FROM solicitudes_ofertas_empresa o JOIN solicitudes s ON s.id = o.solicitud_id
            WHERE o.empresa_id = %s AND o.estado = 'pendiente' AND s.estado = 'buscando_proveedor'
            ORDER BY o.ofrecido_en
        """, (MAD_OFERTA_EMPRESA_SEG, eid, eid))
        pendientes = [{"oferta_id": r[0], "solicitud_id": r[1], "tipo": r[2], "distrito": r[3],
                       "programada_para": r[4].isoformat() if r[4] else None,
                       "segundos_restantes": max(0, r[5] or 0),
                       "doctor_mas_cercano_km": round(r[6] / 1000, 1) if r[6] is not None else None}
                      for r in cur.fetchall()]

        # Aceptadas por la empresa, a la espera de asignar doctor: ya ve dirección y referencia (no datos clínicos)
        cur.execute("""
            SELECT s.id, s.tipo_servicio, s.origen_direccion_texto, s.origen_distrito, s.origen_referencia,
                   s.programada_para,
                   EXTRACT(EPOCH FROM (s.empresa_aceptada_en + make_interval(mins => %s) - NOW()))::int
            FROM solicitudes s
            WHERE s.empresa_id = %s AND s.estado = 'buscando_proveedor' AND s.empresa_aceptada_en IS NOT NULL
            ORDER BY s.empresa_aceptada_en
        """, (MAD_ASIGNAR_MIN, eid))
        por_asignar = []
        for r in cur.fetchall():
            cur.execute("""
                SELECT p.id, p.nombre, pe.estado,
                       CASE WHEN pe.ubicacion IS NULL OR pe.estado <> 'disponible' THEN NULL
                            ELSE ST_Distance(pe.ubicacion, (SELECT origen::geography FROM solicitudes WHERE id = %s)) END
                FROM proveedores p LEFT JOIN proveedores_estado pe ON pe.proveedor_id = p.id
                WHERE p.empresa_id = %s AND p.tipo_servicio IN (%s, 'AMBOS')
                  AND pe.estado IN ('disponible', 'desconectado')
                ORDER BY (pe.estado = 'disponible') DESC, 4 NULLS LAST
            """, (r[0], eid, r[1]))
            docs = [{"id": d[0], "nombre": d[1], "en_linea": d[2] == 'disponible',
                     "km": round(d[3] / 1000, 1) if d[3] is not None else None}
                    for d in cur.fetchall()]
            por_asignar.append({"solicitud_id": r[0], "tipo": r[1], "direccion": r[2], "distrito": r[3],
                                "referencia": r[4], "programada_para": r[5].isoformat() if r[5] else None,
                                "segundos_restantes": max(0, r[6] or 0), "doctores": docs})
    finally:
        cur.close(); conn.close()
    return _mad_json({"pendientes": pendientes, "por_asignar": por_asignar})


@con_log_errores
def handle_mad_empresa_oferta_accion(request, oferta_id, accion):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if accion == "aceptar":
        r = aceptar_oferta_empresa(oferta_id, s["empresa_id"], s["id"])
    else:
        r = rechazar_oferta_empresa(oferta_id, s["empresa_id"], s["id"])
    if r.get("status") == "error":
        return _err(r.get("mensaje") or "No se pudo", 409)
    return _mad_json({"ok": True, **{k: v for k, v in r.items() if k != "siguiente"}})


@con_log_errores
def handle_mad_empresa_asignar(request, solicitud_id):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    try:
        pid = int((request.get_json(force=True, silent=True) or {}).get("proveedor_id"))
    except (TypeError, ValueError):
        return _err("Elige un doctor")
    r = asignar_doctor_empresa(solicitud_id, pid, s["empresa_id"], s["id"])
    if r.get("status") != "ok":
        return _err(r.get("mensaje") or "No se pudo asignar", 409)
    return _mad_json({"ok": True})



def _mad_conteos(cur, hoy, desde, hasta, proveedor_ids=None, empresa_id=None):
    """{proveedor_id: (hoy, mes, rango, rango_mad, rango_ambulancia)}
    - proveedor_ids: servicios brindados por esos doctores, sin importar la empresa.
    - empresa_id: servicios atribuidos a esa empresa (la vigente al aceptar)."""
    if empresa_id is not None:
        filtro, valor = "s.empresa_id = %(valor)s", empresa_id
    elif proveedor_ids:
        filtro, valor = "s.proveedor_id = ANY(%(valor)s)", proveedor_ids
    else:
        return {}
    cur.execute(f"""
        SELECT proveedor_id,
               COUNT(*) FILTER (WHERE f = %(hoy)s),
               COUNT(*) FILTER (WHERE f BETWEEN %(mes)s AND %(hoy)s),
               COUNT(*) FILTER (WHERE f BETWEEN %(desde)s AND %(hasta)s),
               COUNT(*) FILTER (WHERE f BETWEEN %(desde)s AND %(hasta)s AND tipo_servicio = 'MAD'),
               COUNT(*) FILTER (WHERE f BETWEEN %(desde)s AND %(hasta)s AND tipo_servicio = 'AMBULANCIA')
        FROM (
            SELECT s.proveedor_id, s.tipo_servicio, {FECHA_SERVICIO_SQL} AS f
            FROM solicitudes s
            WHERE s.estado = 'finalizado' AND {filtro}
        ) x
        GROUP BY proveedor_id
    """, {"hoy": hoy, "mes": hoy.replace(day=1), "desde": desde, "hasta": hasta, "valor": valor})
    return {r[0]: tuple(int(v) for v in r[1:]) for r in cur.fetchall()}


def _mad_por_dia(cur, desde, hasta, proveedor_ids=None, empresa_id=None):
    if empresa_id is not None:
        filtro, valor = "s.empresa_id = %s", empresa_id
    elif proveedor_ids:
        filtro, valor = "s.proveedor_id = ANY(%s)", proveedor_ids
    else:
        return []
    cur.execute(f"""
        SELECT f, COUNT(*),
               COUNT(*) FILTER (WHERE tipo_servicio = 'MAD'),
               COUNT(*) FILTER (WHERE tipo_servicio = 'AMBULANCIA')
        FROM (
            SELECT s.tipo_servicio, {FECHA_SERVICIO_SQL} AS f
            FROM solicitudes s
            WHERE s.estado = 'finalizado' AND {filtro}
        ) x
        WHERE f BETWEEN %s AND %s
        GROUP BY f ORDER BY f DESC
    """, (valor, desde, hasta))
    return [{"fecha": r[0].isoformat(), "total": r[1], "mad": r[2], "ambulancia": r[3]} for r in cur.fetchall()]


# ── Doctor: sus propios servicios ──
@con_log_errores
def handle_estadisticas_proveedor(request, proveedor_id):
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Proveedor inválido"}, 400)
    pid = int(proveedor_id)
    hoy, desde, hasta = _mad_rango(request)

    conn = get_connection(); cur = conn.cursor()
    try:
        c = _mad_conteos(cur, hoy, desde, hasta, proveedor_ids=[pid]).get(pid, (0, 0, 0, 0, 0))
        dias = _mad_por_dia(cur, desde, hasta, proveedor_ids=[pid])
        cur.execute(f"""
            SELECT s.finalizado_en, s.tipo_servicio, s.origen_distrito
            FROM solicitudes s
            WHERE s.estado = 'finalizado' AND s.proveedor_id = %s
              AND {FECHA_SERVICIO_SQL} BETWEEN %s AND %s
            ORDER BY s.finalizado_en DESC LIMIT 200
        """, (pid, desde, hasta))
        servicios = [{"finalizado_en": _iso(r[0]), "tipo": r[1], "distrito": r[2]}
                     for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()

    return _mad_json({
        "hoy": c[0], "mes": c[1],
        "desde": desde.isoformat(), "hasta": hasta.isoformat(),
        "rango": c[2], "rango_mad": c[3], "rango_ambulancia": c[4],
        "por_dia": dias, "servicios": servicios,
    })


# ═════════════════════════════════════════════════════════════
# PORTAL DE LA EMPRESA PRESTADORA
# ═════════════════════════════════════════════════════════════
def _mad_empresa_ctx(request):
    """Sesión de empresa válida + empresa activa. Devuelve (sesion, error)."""
    s, err = _mad_exigir(request, "empresa")
    if err:
        return None, err
    if not s["empresa_id"]:
        return None, _mad_json({"error": "Usuario sin empresa asignada"}, 403)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT nombre FROM mad_empresas WHERE id = %s AND activo", (s["empresa_id"],))
        r = cur.fetchone()
    finally:
        cur.close(); conn.close()
    if not r:
        return None, _mad_json({"error": "Empresa inactiva o inexistente"}, 403)

    s["empresa_nombre"] = r[0]
    return s, None


@con_log_errores
def handle_mad_empresa_resumen(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    hoy, desde, hasta = _mad_rango(request)

    conn = get_connection(); cur = conn.cursor()
    try:
        # Doctores actuales de la empresa (el empresa_id sale de la sesión, nunca del request)
        cur.execute("""
            SELECT p.id, p.nombre, p.tipo_servicio, pe.estado,
                   ROUND(EXTRACT(EPOCH FROM (NOW() - pe.actualizado_en)) / 60),
                   COALESCE(u.activo, FALSE), u.usuario, COALESCE(u.totp_activo, FALSE), p.telefono, p.cmp
            FROM proveedores p
            LEFT JOIN proveedores_estado pe ON pe.proveedor_id = p.id
            LEFT JOIN mad_usuarios u ON u.proveedor_id = p.id
            WHERE p.empresa_id = %s
            ORDER BY p.nombre
        """, (s["empresa_id"],))
        filas = cur.fetchall()

        # Servicios atribuidos a la empresa (incluye los de doctores que ya no están)
        conteos = _mad_conteos(cur, hoy, desde, hasta, empresa_id=s["empresa_id"])
        dias = _mad_por_dia(cur, desde, hasta, empresa_id=s["empresa_id"])

        vigentes = {r[0] for r in filas}
        extra = [pid for pid, c in conteos.items() if pid not in vigentes and (c[1] > 0 or c[2] > 0)]
        ex = []
        if extra:
            cur.execute("SELECT id, nombre, tipo_servicio, cmp FROM proveedores WHERE id = ANY(%s) ORDER BY nombre", (extra,))
            ex = cur.fetchall()
    finally:
        cur.close(); conn.close()

    doctores, tot = [], [0, 0, 0, 0, 0]
    for pid, nombre, tipo, estado, min_senal, usuario_activo, usuario, totp_activo, telefono, cmp_doc in filas:
        c = conteos.get(pid, (0, 0, 0, 0, 0))
        tot = [a + b for a, b in zip(tot, c)]
        doctores.append({
            "id": pid, "nombre": nombre, "tipo_servicio": tipo, "usuario": usuario, "vigente": True,
            "telefono": telefono, "cmp": cmp_doc,
            "estado": estado or "desconectado",
            "minutos_senal": int(min_senal) if min_senal is not None else None,
            "usuario_activo": usuario_activo, "totp_activo": totp_activo,
            "hoy": c[0], "mes": c[1], "rango": c[2], "rango_mad": c[3], "rango_ambulancia": c[4],
        })
    for pid, nombre, tipo, cmp_doc in ex:
        c = conteos[pid]
        tot = [a + b for a, b in zip(tot, c)]
        doctores.append({
            "id": pid, "nombre": nombre, "tipo_servicio": tipo, "usuario": None, "vigente": False,
            "telefono": None, "cmp": cmp_doc,
            "estado": None, "minutos_senal": None, "usuario_activo": False, "totp_activo": False,
            "hoy": c[0], "mes": c[1], "rango": c[2], "rango_mad": c[3], "rango_ambulancia": c[4],
        })

    return _mad_json({
        "empresa": s["empresa_nombre"], "max_doctores": MAD_MAX_DOCTORES_EMPRESA,
        "desde": desde.isoformat(), "hasta": hasta.isoformat(),
        "totales": {"hoy": tot[0], "mes": tot[1], "rango": tot[2], "rango_mad": tot[3], "rango_ambulancia": tot[4]},
        "doctores": doctores, "por_dia": dias,
    })


def _xl(ws, fila, col, valor):
    c = ws.cell(row=fila, column=col, value=valor)
    if isinstance(valor, str):
        c.data_type = "s"      # nunca interpretar un texto como fórmula (ej. un nombre que empiece con "=")
    return c


@con_log_errores
def handle_mad_empresa_exportar(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    hoy, desde, hasta = _mad_rango(request)

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT id, nombre FROM proveedores WHERE empresa_id = %s ORDER BY nombre",
                    (s["empresa_id"],))
        doctores = cur.fetchall()
        cur.execute(f"""
            SELECT s.id, s.finalizado_en, {FECHA_SERVICIO_SQL} AS f,
                   p.id, p.nombre, s.tipo_servicio, s.origen_distrito
            FROM solicitudes s
            JOIN proveedores p ON p.id = s.proveedor_id
            WHERE s.estado = 'finalizado' AND s.empresa_id = %s
              AND {FECHA_SERVICIO_SQL} BETWEEN %s AND %s
            ORDER BY s.finalizado_en
            LIMIT %s
        """, (s["empresa_id"], desde, hasta, MAD_EXCEL_MAX_FILAS + 1))
        filas = cur.fetchall()
    finally:
        cur.close(); conn.close()

    if len(filas) > MAD_EXCEL_MAX_FILAS:
        return _mad_json({"error": "El rango tiene demasiados servicios para exportar. Reduce las fechas."}, 400)

    vigentes = {pid for pid, _ in doctores}
    nombres = {pid: n for pid, n in doctores}
    por_doctor = {pid: [0, 0] for pid, _ in doctores}     # [MAD, ambulancia]
    por_dia, detalle = {}, []
    for sid, fin, f, pid, nombre, tipo, distrito in filas:
        if pid not in vigentes:
            nombre = f"{nombre} (ya no pertenece a la empresa)"
        i = 1 if tipo == "AMBULANCIA" else 0
        por_doctor.setdefault(pid, [0, 0])[i] += 1
        nombres.setdefault(pid, nombre)
        por_dia.setdefault((f, pid), [0, 0])[i] += 1
        if fin is not None and getattr(fin, "tzinfo", None):
            fin = fin.astimezone(LIMA_TZ).replace(tzinfo=None)
        detalle.append((sid, f, fin, nombre, tipo, distrito))

    wb = openpyxl.Workbook()
    fill = PatternFill(start_color="003654", end_color="003654", fill_type="solid")
    blanco = Font(color="FFFFFF", bold=True)

    def encabezado(ws, fila, titulos):
        for i, t in enumerate(titulos, start=1):
            c = _xl(ws, fila, i, t)
            c.fill = fill; c.font = blanco
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def anchos(ws, lista):
        for i, w in enumerate(lista, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # ── Hoja 1: resumen por doctor ──
    ws = wb.active; ws.title = "Resumen"
    _xl(ws, 1, 1, f"Empresa: {s['empresa_nombre']}").font = Font(bold=True, size=13)
    _xl(ws, 2, 1, f"Período: {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}")
    _xl(ws, 3, 1, f"Generado: {datetime.now(LIMA_TZ).strftime('%d/%m/%Y %H:%M')} (hora de Lima) · solo servicios finalizados")
    encabezado(ws, 5, ["Doctor", "Médico a domicilio", "Ambulancia", "Total"])
    fila, tot_m, tot_a = 6, 0, 0
    for pid in sorted(por_doctor, key=lambda k: nombres.get(k, "")):
        m, a = por_doctor[pid]
        _xl(ws, fila, 1, nombres.get(pid, "")); _xl(ws, fila, 2, m)
        _xl(ws, fila, 3, a); _xl(ws, fila, 4, m + a)
        tot_m += m; tot_a += a; fila += 1
    _xl(ws, fila, 1, "TOTAL").font = Font(bold=True)
    for col, v in ((2, tot_m), (3, tot_a), (4, tot_m + tot_a)):
        _xl(ws, fila, col, v).font = Font(bold=True)
    anchos(ws, [34, 20, 14, 10])

    # ── Hoja 2: por día y doctor ──
    ws2 = wb.create_sheet("Por día")
    encabezado(ws2, 1, ["Fecha", "Doctor", "Médico a domicilio", "Ambulancia", "Total"])
    fila = 2
    for (f, pid) in sorted(por_dia, key=lambda k: (k[0], nombres.get(k[1], ""))):
        m, a = por_dia[(f, pid)]
        _xl(ws2, fila, 1, f).number_format = "DD/MM/YYYY"
        _xl(ws2, fila, 2, nombres.get(pid, "")); _xl(ws2, fila, 3, m)
        _xl(ws2, fila, 4, a); _xl(ws2, fila, 5, m + a)
        fila += 1
    anchos(ws2, [13, 34, 20, 14, 10]); ws2.freeze_panes = "A2"

    # ── Hoja 3: detalle (sin datos del paciente) ──
    ws3 = wb.create_sheet("Detalle")
    encabezado(ws3, 1, ["N° servicio", "Fecha", "Hora", "Doctor", "Tipo", "Distrito"])
    for fila, (sid, f, fin, nombre, tipo, distrito) in enumerate(detalle, start=2):
        _xl(ws3, fila, 1, sid)
        _xl(ws3, fila, 2, f).number_format = "DD/MM/YYYY"
        _xl(ws3, fila, 3, fin.strftime("%H:%M") if fin else "")
        _xl(ws3, fila, 4, nombre)
        _xl(ws3, fila, 5, "Ambulancia" if tipo == "AMBULANCIA" else "Médico a domicilio")
        _xl(ws3, fila, 6, (distrito or "").title())
    anchos(ws3, [12, 13, 8, 34, 20, 22]); ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:F{max(2, len(detalle) + 1)}"

    buf = BytesIO(); wb.save(buf)
    return (buf.getvalue(), 200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": f'attachment; filename="servicios_{desde.isoformat()}_{hasta.isoformat()}.xlsx"',
        "Cache-Control": "no-store",
    })


@con_log_errores
def handle_mad_empresa_ubicaciones(request):
    """Mapa de doctores de la empresa, con reglas de privacidad: no se muestra a desconectados,
    ni a quien lleva más de MAD_UBICACION_MAX_MIN sin señal, ni a quien está en el domicilio de un paciente."""
    s, err = _mad_empresa_ctx(request)
    if err:
        return err

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT p.id, p.nombre, COALESCE(pe.estado, 'desconectado'),
                   ST_Y(pe.ubicacion::geometry), ST_X(pe.ubicacion::geometry),
                   ROUND(EXTRACT(EPOCH FROM (NOW() - pe.actualizado_en)) / 60),
                   EXISTS (SELECT 1 FROM solicitudes sl
                           WHERE sl.proveedor_id = p.id AND sl.estado IN ('llego', 'en_atencion'))
            FROM proveedores p
            JOIN mad_usuarios u ON u.proveedor_id = p.id AND u.activo
            LEFT JOIN proveedores_estado pe ON pe.proveedor_id = p.id
            WHERE p.empresa_id = %s
            ORDER BY p.nombre
        """, (s["empresa_id"],))
        filas = cur.fetchall()
    finally:
        cur.close(); conn.close()

    doctores = []
    for pid, nombre, estado, lat, lng, minutos, en_domicilio in filas:
        minutos = int(minutos) if minutos is not None else None
        if estado == "desconectado":
            oculto = "desconectado"        # su última posición podría ser su casa
        elif en_domicilio:
            oculto = "en_atencion"         # estaría en el domicilio de un paciente
        elif lat is None or minutos is None or minutos > MAD_UBICACION_MAX_MIN:
            oculto = "sin_senal"
        else:
            oculto = None
        doctores.append({
            "id": pid, "nombre": nombre, "estado": estado, "minutos_senal": minutos, "oculto": oculto,
            "lat": lat if oculto is None else None, "lng": lng if oculto is None else None,
        })
    return _mad_json({"doctores": doctores})


@con_log_errores
def handle_mad_empresa_crear_doctor(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err

    datos, error = _datos_nuevo_doctor(request.get_json(force=True, silent=True) or {})
    if error:
        return _mad_json({"error": error}, 400)
    nombre, usuario, telefono, tipo, cmp_ = datos

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_empresas WHERE id = %s FOR UPDATE", (s["empresa_id"],))  # serializa altas
        cur.execute("SELECT COUNT(*) FROM proveedores WHERE empresa_id = %s", (s["empresa_id"],))
        if cur.fetchone()[0] >= MAD_MAX_DOCTORES_EMPRESA:
            return _mad_json({"error": f"Alcanzaste el máximo de {MAD_MAX_DOCTORES_EMPRESA} doctores. Contacta a Sanitas."}, 409)

        cur.execute("SELECT 1 FROM mad_usuarios WHERE usuario = %s", (usuario,))
        if cur.fetchone():
            return _mad_json({"error": "Ese usuario ya existe, elige otro"}, 409)

        cur.execute("SELECT 1 FROM proveedores WHERE telefono = %s", (telefono,))
        if cur.fetchone():
            return _mad_json({"error": "Ese celular ya está registrado en otro doctor"}, 409)
        if cmp_:
            cur.execute("SELECT 1 FROM proveedores WHERE cmp = %s", (cmp_,))
            if cur.fetchone():
                return _mad_json({"error": "Ese CMP ya está registrado en otro doctor"}, 409)

        cur.execute("""
            INSERT INTO proveedores (nombre, telefono, tipo_servicio, empresa_id, cmp)
            VALUES (%s, %s, %s, %s, %s) RETURNING id
        """, (nombre, telefono, tipo, s["empresa_id"], cmp_))
        proveedor_id = cur.fetchone()[0]

        cur.execute("""
            INSERT INTO proveedores_estado (proveedor_id, estado, actualizado_en)
            VALUES (%s, 'desconectado', NOW())
        """, (proveedor_id,))

        cur.execute("""
            INSERT INTO mad_usuarios (usuario, nombre, rol, proveedor_id, password_hash)
            VALUES (%s, %s, 'proveedor', %s, %s)
        """, (usuario, nombre, proveedor_id, hash_password(temporal)))
        _mad_audit(cur, s["id"], "doctor_crear", f"proveedor:{proveedor_id}", {"usuario": usuario, "cmp": cmp_})
        conn.commit()
    except psycopg2.IntegrityError:
        conn.rollback()
        return _mad_json({"error": "Ese usuario o celular ya existe"}, 409)
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "proveedor_id": proveedor_id, "usuario": usuario,
                      "password_temporal": temporal}, 201)


@con_log_errores
def handle_mad_empresa_doctor_estado(request, proveedor_id):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Doctor inválido"}, 400)
    pid = int(proveedor_id)
    activo = bool((request.get_json(force=True, silent=True) or {}).get("activo"))

    conn = get_connection(); cur = conn.cursor()
    try:
        # Solo doctores de ESTA empresa (si es de otra, responde igual que si no existiera)
        cur.execute("""
            SELECT u.id FROM proveedores p
            JOIN mad_usuarios u ON u.proveedor_id = p.id AND u.rol = 'proveedor'
            WHERE p.id = %s AND p.empresa_id = %s
            FOR UPDATE OF u
        """, (pid, s["empresa_id"]))
        u = cur.fetchone()
        if not u:
            return _mad_json({"error": "Doctor no encontrado"}, 404)

        if activo:
            cur.execute("UPDATE mad_usuarios SET activo = TRUE WHERE id = %s", (u[0],))
        elif not _mad_desactivar_usuario(cur, u[0], pid):
            conn.rollback()
            return _mad_json({"error": _MSG_SERVICIO_EN_CURSO}, 409)
        _mad_audit(cur, s["id"], "doctor_activar" if activo else "doctor_desactivar", f"proveedor:{pid}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "activo": activo})


@con_log_errores
def handle_mad_empresa_doctor_reset(request, proveedor_id):
    """Nueva clave temporal + obliga a configurar Google Authenticator de nuevo
    (doctor que olvidó su clave o cambió de celular)."""
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Doctor inválido"}, 400)
    pid = int(proveedor_id)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM proveedores WHERE id = %s AND empresa_id = %s", (pid, s["empresa_id"]))
        if not cur.fetchone():
            return _mad_json({"error": "Doctor no encontrado"}, 404)

        cur.execute("""
            UPDATE mad_usuarios
            SET password_hash = %s, debe_cambiar_password = TRUE,
                totp_secret_enc = NULL, totp_activo = FALSE, totp_ultimo_paso = NULL,
                intentos_fallidos = 0, bloqueado_hasta = NULL
            WHERE proveedor_id = %s AND rol = 'proveedor'
            RETURNING id, usuario
        """, (hash_password(temporal), pid))
        u = cur.fetchone()
        if not u:
            return _mad_json({"error": "El doctor no tiene usuario de acceso"}, 404)
        cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE usuario_id = %s", (u[0],))
        cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (u[0],))
        _mad_audit(cur, s["id"], "usuario_reset", f"usuario:{u[0]}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "usuario": u[1], "password_temporal": temporal})


@con_log_errores
def handle_mad_empresa_doctor_editar(request, proveedor_id):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if not str(proveedor_id).isdigit():
        return _mad_json({"error": "Doctor inválido"}, 400)
    data = request.get_json(force=True, silent=True) or {}

    conn = get_connection(); cur = conn.cursor()
    try:
        status, payload, avisos = _mad_editar_doctor(cur, int(proveedor_id), data, s["id"], s["empresa_id"])
        if status == 200:
            conn.commit()
        else:
            conn.rollback()
    finally:
        cur.close(); conn.close()

    _mad_enviar_avisos(avisos, "mad doctor celular")
    return _mad_json(payload, status)


# ── Usuarios (encargados) de la empresa ──
@con_log_errores
def handle_mad_empresa_usuarios_listar(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT id, usuario, nombre, activo, totp_activo, debe_cambiar_password,
                   COALESCE(bloqueado_hasta > NOW(), FALSE), ultimo_login
            FROM mad_usuarios
            WHERE rol = 'empresa' AND empresa_id = %s
            ORDER BY activo DESC, nombre
        """, (s["empresa_id"],))
        usuarios = [{
            "id": r[0], "usuario": r[1], "nombre": r[2], "activo": r[3], "totp_activo": r[4],
            "debe_cambiar_password": r[5], "bloqueado": r[6],
            "ultimo_login": _iso(r[7]), "yo": r[0] == s["id"],
        } for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
    return _mad_json({"usuarios": usuarios, "max_usuarios": MAD_MAX_USUARIOS_EMPRESA})


@con_log_errores
def handle_mad_empresa_crear_usuario(request):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    data = request.get_json(force=True, silent=True) or {}
    nombre  = (data.get("nombre") or "").strip()[:120]
    usuario = (data.get("usuario") or "").strip().lower()
    if len(nombre) < 3 or not re.fullmatch(r"[a-z0-9._-]{4,30}", usuario):
        return _mad_json({"error": "Revisa los datos: nombre y usuario (4 a 30 caracteres: letras, números, punto, guion)"}, 400)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_empresas WHERE id = %s FOR UPDATE", (s["empresa_id"],))  # serializa altas y bajas
        cur.execute("SELECT COUNT(*) FROM mad_usuarios WHERE rol = 'empresa' AND empresa_id = %s AND activo",
                    (s["empresa_id"],))
        if cur.fetchone()[0] >= MAD_MAX_USUARIOS_EMPRESA:
            return _mad_json({"error": f"Alcanzaste el máximo de {MAD_MAX_USUARIOS_EMPRESA} usuarios activos. Desactiva uno para crear otro."}, 409)
        cur.execute("SELECT 1 FROM mad_usuarios WHERE usuario = %s", (usuario,))
        if cur.fetchone():
            return _mad_json({"error": "Ese usuario ya existe, elige otro"}, 409)

        cur.execute("""
            INSERT INTO mad_usuarios (usuario, nombre, rol, empresa_id, password_hash)
            VALUES (%s, %s, 'empresa', %s, %s) RETURNING id
        """, (usuario, nombre, s["empresa_id"], hash_password(temporal)))
        nuevo_id = cur.fetchone()[0]
        _mad_audit(cur, s["id"], "usuario_crear", f"usuario:{nuevo_id}", {"usuario": usuario})
        conn.commit()
    except psycopg2.IntegrityError:
        conn.rollback()
        return _mad_json({"error": "Ese usuario ya existe, elige otro"}, 409)
    finally:
        cur.close(); conn.close()

    return _mad_json({"ok": True, "usuario": usuario, "password_temporal": temporal}, 201)


@con_log_errores
def handle_mad_empresa_usuario_estado(request, usuario_id):
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if not str(usuario_id).isdigit():
        return _mad_json({"error": "Usuario inválido"}, 400)
    uid = int(usuario_id)
    if uid == s["id"]:
        return _mad_json({"error": "No puedes desactivar tu propio usuario"}, 400)
    activo = bool((request.get_json(force=True, silent=True) or {}).get("activo"))

    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("SELECT 1 FROM mad_empresas WHERE id = %s FOR UPDATE", (s["empresa_id"],))
        # Quien actúa debe seguir activo: otro encargado pudo desactivarlo mientras esperaba el bloqueo.
        # Así siempre queda al menos un encargado activo.
        cur.execute("SELECT activo FROM mad_usuarios WHERE id = %s", (s["id"],))
        yo = cur.fetchone()
        if not yo or not yo[0]:
            return _mad_json({"error": "No autenticado"}, 401)

        cur.execute("""SELECT activo FROM mad_usuarios
                       WHERE id = %s AND rol = 'empresa' AND empresa_id = %s FOR UPDATE""", (uid, s["empresa_id"]))
        t = cur.fetchone()
        if not t:
            return _mad_json({"error": "Usuario no encontrado"}, 404)

        if activo:
            if not t[0]:
                cur.execute("SELECT COUNT(*) FROM mad_usuarios WHERE rol = 'empresa' AND empresa_id = %s AND activo",
                            (s["empresa_id"],))
                if cur.fetchone()[0] >= MAD_MAX_USUARIOS_EMPRESA:
                    return _mad_json({"error": f"Alcanzaste el máximo de {MAD_MAX_USUARIOS_EMPRESA} usuarios activos."}, 409)
            cur.execute("UPDATE mad_usuarios SET activo = TRUE WHERE id = %s", (uid,))
        else:
            cur.execute("UPDATE mad_usuarios SET activo = FALSE WHERE id = %s", (uid,))
            cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE usuario_id = %s", (uid,))
            cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (uid,))
        _mad_audit(cur, s["id"], "usuario_activar" if activo else "usuario_desactivar", f"usuario:{uid}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "activo": activo})


@con_log_errores
def handle_mad_empresa_usuario_reset(request, usuario_id):
    """Clave temporal nueva + obliga a configurar Google Authenticator de nuevo."""
    s, err = _mad_empresa_ctx(request)
    if err:
        return err
    if not str(usuario_id).isdigit():
        return _mad_json({"error": "Usuario inválido"}, 400)
    uid = int(usuario_id)
    if uid == s["id"]:
        return _mad_json({"error": "No puedes resetear tu propio acceso. Pídelo a otro encargado o a Sanitas."}, 400)

    temporal = _mad_password_temporal()
    conn = get_connection(); cur = conn.cursor()
    try:
        cur.execute("""
            UPDATE mad_usuarios
            SET password_hash = %s, debe_cambiar_password = TRUE,
                totp_secret_enc = NULL, totp_activo = FALSE, totp_ultimo_paso = NULL,
                intentos_fallidos = 0, bloqueado_hasta = NULL
            WHERE id = %s AND rol = 'empresa' AND empresa_id = %s
            RETURNING usuario
        """, (hash_password(temporal), uid, s["empresa_id"]))
        r = cur.fetchone()
        if not r:
            return _mad_json({"error": "Usuario no encontrado"}, 404)
        cur.execute("UPDATE mad_sesiones SET revocado = TRUE WHERE usuario_id = %s", (uid,))
        cur.execute("UPDATE mad_dispositivos SET revocado = TRUE WHERE usuario_id = %s", (uid,))
        _mad_audit(cur, s["id"], "usuario_reset", f"usuario:{uid}")
        conn.commit()
    finally:
        cur.close(); conn.close()
    return _mad_json({"ok": True, "usuario": r[0], "password_temporal": temporal})


# ═════════════════════════════════════════════════════════════
# ROUTER MAD
# ═════════════════════════════════════════════════════════════
def _entre(path, prefijo, sufijo):
    """'/oferta/<x>/aceptar' → '<x>' (None si no coincide o <x> trae más '/')."""
    if not (path.startswith(prefijo) and path.endswith(sufijo)):
        return None
    medio = path[len(prefijo):len(path) - len(sufijo)]
    return medio if medio and "/" not in medio else None


def _no_autorizado():
    return _mad_json({"error": "No autorizado"}, 403)


def _exigir_apikey(request):
    """None si la ApiKey es válida; si no, la respuesta 401."""
    if validar_api_key(request):
        return None
    return _mad_json({"error": "API Key invalida"}, 401)


def mad_router(request, path):
    """Resuelve todas las rutas de Médico a domicilio / ambulancia.
    Devuelve la respuesta, o None si la ruta no es de este módulo (main.py sigue con las suyas)."""
    m = request.method

    # ── Autenticación ──
    if path == "/mad/auth/login" and m == "POST":
        return handle_mad_login(request)
    if path == "/mad/auth/cambiar-password" and m == "POST":
        return handle_mad_cambiar_password(request)
    if path == "/mad/auth/totp/iniciar" and m == "POST":
        return handle_mad_totp_iniciar(request)
    if path == "/mad/auth/totp/verificar" and m == "POST":
        return handle_mad_totp_verificar(request)
    if path == "/mad/auth/biometrico/activar" and m == "POST":
        return handle_mad_biometrico_activar(request)
    if path == "/mad/auth/biometrico" and m == "POST":
        return handle_mad_biometrico_login(request)
    if path == "/mad/auth/logout" and m == "POST":
        return handle_mad_logout(request)
    if path == "/mad/auth/yo" and m == "GET":
        return handle_mad_yo(request)
    if path == "/mad-bootstrap/coordinador" and m == "POST":
        return handle_mad_crear_coordinador(request)      # exige ApiKey (no es ruta pública)

    # ── Monitor y administración (coordinador) ──
    if path == "/mad/monitor" and m == "GET":
        return handle_mad_monitor(request)
    if path == "/mad/monitor/datos" and m == "GET":
        return handle_mad_monitor_datos(request)
    if path == "/mad/admin/usuarios" and m == "GET":
        return handle_mad_admin_listar(request)
    if path == "/mad/admin/proveedores" and m == "POST":
        return handle_mad_admin_crear_proveedor(request)
    if path == "/mad/admin/ubicaciones/pedir" and m == "POST":
        return handle_mad_admin_pedir_ubicacion(request)
    if m == "GET":
        x = _entre(path, "/mad/admin/ubicaciones/", "")
        if x and x.isdigit():
            return handle_mad_admin_estado_ubicacion(request, x)
    if path == "/mad/admin/solicitudes" and m == "POST":
        return handle_mad_admin_crear_solicitud(request)
    if path == "/mad/admin/prioridades" and m == "GET":
        return handle_mad_admin_prioridades_listar(request)
    if path == "/mad/admin/prioridades" and m == "POST":
        return handle_mad_admin_prioridades_guardar(request)
    if path == "/mad/admin/empresas" and m == "GET":
        return handle_mad_admin_empresas_listar(request)
    if path == "/mad/admin/empresas" and m == "POST":
        return handle_mad_admin_crear_empresa(request)
    if m == "POST":
        x = _entre(path, "/mad/admin/usuarios/", "/reset")
        if x: return handle_mad_admin_reset(request, x)
        x = _entre(path, "/mad/admin/usuarios/", "/estado")
        if x: return handle_mad_admin_estado(request, x)
        x = _entre(path, "/mad/admin/proveedores/", "/empresa")
        if x: return handle_mad_admin_asignar_empresa(request, x)
        x = _entre(path, "/mad/admin/proveedores/", "/editar")
        if x: return handle_mad_admin_doctor_editar(request, x)
        x = _entre(path, "/mad/admin/solicitudes/", "/liberar")
        if x: return handle_mad_admin_liberar_servicio(request, x)
        x = _entre(path, "/mad/admin/solicitudes/", "/ofrecer")
        if x: return handle_mad_admin_ofrecer_a_proveedor(request, x)
        x = _entre(path, "/mad/admin/solicitudes/", "/enviar-tracking")
        if x: return handle_mad_admin_enviar_tracking(request, x)
        x = _entre(path, "/mad/admin/solicitudes/", "/activar")
        if x: return handle_mad_admin_activar_programada(request, x)
    if m == "GET":
        x = _entre(path, "/mad/admin/proveedores/", "/liberaciones")
        if x: return handle_mad_admin_liberaciones_proveedor(request, x)

    # ── Portal de la empresa ──
    if path == "/mad/empresa" and m == "GET":
        return handle_mad_empresa_pagina(request)
    if path == "/mad/empresa/ofertas" and m == "GET":
        return handle_mad_empresa_ofertas(request)
    if path == "/mad/empresa/resumen" and m == "GET":
        return handle_mad_empresa_resumen(request)
    if path == "/mad/empresa/exportar" and m == "GET":
        return handle_mad_empresa_exportar(request)
    if path == "/mad/empresa/ubicaciones" and m == "GET":
        return handle_mad_empresa_ubicaciones(request)
    if path == "/mad/empresa/doctores" and m == "POST":
        return handle_mad_empresa_crear_doctor(request)
    if path == "/mad/empresa/usuarios" and m == "GET":
        return handle_mad_empresa_usuarios_listar(request)
    if path == "/mad/empresa/usuarios" and m == "POST":
        return handle_mad_empresa_crear_usuario(request)
    if m == "POST":
        x = _entre(path, "/mad/empresa/ofertas/", "/aceptar")
        if x and x.isdigit(): return handle_mad_empresa_oferta_accion(request, int(x), "aceptar")
        x = _entre(path, "/mad/empresa/ofertas/", "/rechazar")
        if x and x.isdigit(): return handle_mad_empresa_oferta_accion(request, int(x), "rechazar")
        x = _entre(path, "/mad/empresa/solicitudes/", "/asignar")
        if x and x.isdigit(): return handle_mad_empresa_asignar(request, int(x))
        x = _entre(path, "/mad/empresa/doctores/", "/estado")
        if x: return handle_mad_empresa_doctor_estado(request, x)
        x = _entre(path, "/mad/empresa/doctores/", "/reset")
        if x: return handle_mad_empresa_doctor_reset(request, x)
        x = _entre(path, "/mad/empresa/doctores/", "/editar")
        if x: return handle_mad_empresa_doctor_editar(request, x)
        x = _entre(path, "/mad/empresa/usuarios/", "/estado")
        if x: return handle_mad_empresa_usuario_estado(request, x)
        x = _entre(path, "/mad/empresa/usuarios/", "/reset")
        if x: return handle_mad_empresa_usuario_reset(request, x)

    # ── Oferta por enlace (el doctor debe haber iniciado sesión y ser el dueño de la oferta) ──
    if path.startswith("/oferta/"):
        for sufijo, metodo, manejador in (("/estado", "GET", handle_estado_oferta_token),
                                          ("/aceptar", "POST", handle_aceptar_oferta_token),
                                          ("/rechazar", "POST", handle_rechazar_oferta_token)):
            token = _entre(path, "/oferta/", sufijo)
            if token and m == metodo:
                s, err = _exigir_proveedor(request)
                if err:
                    return err
                if not _oferta_es_del_proveedor(s["proveedor_id"], oferta_token=token):
                    return _no_autorizado()
                return manejador(request, token)
        if m == "GET":
            return handle_servir_oferta(request, path.split("/oferta/")[1])

    # ── Panel del proveedor ──
    if path == "/proveedor" and m == "GET":
        return handle_servir_proveedor(request)
    if path == "/proveedor.webmanifest" and m == "GET":
        return handle_pwa_manifest(request)
    if path == "/proveedor-sw.js" and m == "GET":
        return handle_pwa_sw(request)

    if path.startswith("/proveedores/"):
        for sufijo, metodo, manejador in (("/disponibilidad", "POST", handle_toggle_disponibilidad),
                                          ("/estado-actual", "GET", handle_estado_actual_proveedor),
                                          ("/ubicacion", "POST", handle_actualizar_ubicacion_proveedor),
                                          ("/estadisticas", "GET", handle_estadisticas_proveedor),
                                          ("/push/clave", "GET", handle_push_clave),
                                          ("/push/suscribir", "POST", handle_push_suscribir),
                                          ("/push/baja", "POST", handle_push_baja),
                                          ("/push/probar", "POST", handle_push_probar)):
            pid = _entre(path, "/proveedores/", sufijo)
            if pid and m == metodo:
                s, err = _exigir_proveedor(request, pid)
                if err:
                    return err
                return manejador(request, pid)

    if path.startswith("/ofertas/") and m == "POST":
        for sufijo, manejador in (("/aceptar", handle_aceptar_oferta), ("/rechazar", handle_rechazar_oferta)):
            oferta_id = _entre(path, "/ofertas/", sufijo)
            if oferta_id:
                s, err = _exigir_proveedor(request)
                if err:
                    return err
                if not oferta_id.isdigit() or not _oferta_es_del_proveedor(s["proveedor_id"], oferta_id=int(oferta_id)):
                    return _no_autorizado()
                return manejador(request, oferta_id)

    # ── Solicitudes ──
    if path == "/solicitudes/crear" and m == "POST":
        err = _exigir_apikey(request)
        if err:
            return err
        return handle_crear_solicitud(request)
    if path == "/solicitudes/revisar-timeouts" and m == "GET":
        err = _exigir_apikey(request)
        if err:
            return err
        return handle_revisar_timeouts_ofertas(request)
    if path == "/solicitudes/recalibrar-eta" and m == "GET":
        err = _exigir_apikey(request)
        if err:
            return err
        return handle_recalibrar_eta(request)
    if path.startswith("/solicitudes/"):
        x = _entre(path, "/solicitudes/", "/estado")
        if x and m == "POST":
            s, err = _exigir_proveedor(request)
            if err:
                return err
            return handle_actualizar_estado_solicitud(request, x, s["proveedor_id"])
        x = _entre(path, "/solicitudes/", "/novedad")
        if x and m == "POST":
            return handle_novedad_doctor(request, x)
        x = _entre(path, "/solicitudes/", "/buscar-proveedor")
        if x and m == "GET":
            s, err = _mad_exigir(request, "coordinador")
            if err:
                return err
            return handle_buscar_proveedor(request, x)

    # ── Seguimiento del cliente (el token del enlace funciona como llave) ──
    if path.startswith("/track/u/"):
        x = _entre(path, "/track/u/", "/enviar")
        if x and m == "POST":
            return handle_ubicacion_enviar(request, x)
        x = _entre(path, "/track/u/", "/info")
        if x and m == "GET":
            return handle_ubicacion_info(x)
        x = path.split("/track/u/")[1]
        if x and "/" not in x and m == "GET":
            return handle_servir_ubicacion(request, x)
    if path.startswith("/track/"):
        if m == "POST":
            x = _entre(path, "/track/", "/cancelar")
            if x: return handle_cancelar_solicitud_cliente(request, x)
            x = _entre(path, "/track/", "/mensaje")
            if x: return handle_mensaje_cliente(request, x)
            x = _entre(path, "/track/", "/cambiar-medico")
            if x: return handle_cambiar_medico_cliente(request, x)
            x = _entre(path, "/track/", "/reprogramar")
            if x: return handle_reprogramar_solicitud_cliente(request, x)
        if m == "GET":
            x = _entre(path, "/track/", "/estado")
            if x: return obtener_estado_tracking(x)
            x = path.split("/track/")[1]
            if x and "/" not in x:
                return handle_servir_tracking(request, x)

    return None
