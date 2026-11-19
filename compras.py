"""
compras.py — Portal de Proveedores + Módulo de Compras (sofi-apis)

Rutas (todas bajo /compras, exentas de ApiKey; usan sesión propia Bearer):
  Páginas
    GET  /compras/portal                                  → static/compras/portal.html (proveedor)
    GET  /compras/admin                                   → static/compras/admin.html  (equipo de compras)

  Proveedor (Authorization: Bearer <token proveedor>)
    POST /compras/api/proveedor/login                     {ruc, password}
    POST /compras/api/proveedor/cambiar-password          {actual, nueva}
    POST /compras/api/proveedor/logout
    GET  /compras/api/proveedor/yo
    POST /compras/api/proveedor/datos                     {contacto_nombre, correo, telefono}
    POST /compras/api/proveedor/documentos                multipart: tipo_id, archivo, fecha_emision?, fecha_vencimiento?
    GET  /compras/api/proveedor/documentos/<id>/descargar

  Compras (Authorization: Bearer <token admin>)
    POST /compras/api/admin/login                         {usuario, password}
    POST /compras/api/admin/cambiar-password              {actual, nueva}
    POST /compras/api/admin/logout
    GET  /compras/api/admin/yo
    GET  /compras/api/admin/consulta-ruc?ruc=
    GET  /compras/api/admin/empresas?q=&filtro=
    POST /compras/api/admin/empresas                      {ruc, razon_social, ...}
    GET  /compras/api/admin/empresas/<ruc>
    POST /compras/api/admin/empresas/<ruc>/editar
    POST /compras/api/admin/empresas/<ruc>/reset
    POST /compras/api/admin/empresas/<ruc>/estado         {activo}
    GET  /compras/api/admin/pendientes
    POST /compras/api/admin/documentos/<id>/revisar       {estado, observacion?, fecha_vencimiento?}
    GET  /compras/api/admin/documentos/<id>/descargar
    GET  /compras/api/admin/tipos-documento
    POST /compras/api/admin/tipos-documento               (rol admin)
    GET  /compras/api/admin/usuarios                      (rol admin)
    POST /compras/api/admin/usuarios                      (rol admin)
    POST /compras/api/admin/usuarios/<id>/reset           (rol admin)
    POST /compras/api/admin/usuarios/<id>/estado          (rol admin)

  Bootstrap (requiere ApiKey, lo valida main.py porque no está exenta)
    POST /compras-bootstrap/admin                         {usuario, nombre, password?}
"""

import os
import re
import json
import hmac
import hashlib
import secrets
import logging
import traceback
from decimal import Decimal
from datetime import datetime, date, timedelta

import pytz
import psycopg2
import requests
import google.auth
from google.auth.transport import requests as google_auth_requests
from google.cloud import storage

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
LIMA_TZ = pytz.timezone("America/Lima")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_COMPRAS = os.path.join(BASE_DIR, "static", "compras")

COMPRAS_BUCKET = os.environ.get("COMPRAS_BUCKET", "sanitas-proveedores-docs")
SESION_HORAS_ADMIN = 8
SESION_HORAS_PROVEEDOR = 4
MAX_INTENTOS_LOGIN = 5
BLOQUEO_MINUTOS = 15
MAX_ARCHIVO_MB = 10
MINUTOS_URL_FIRMADA = 15
ROLES_VALIDOS = ("admin", "analista")
ESTADOS_REVISION = ("aprobado", "observado")
DIAS_POR_VENCER = 30

TIPOS_ARCHIVO = {
    "application/pdf": "pdf",
    "image/jpeg": "jpg",
    "image/png": "png",
}

HOY_LIMA = "(now() AT TIME ZONE 'America/Lima')::date"

HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json",
    "Cache-Control": "no-store",
}

log = logging.getLogger("compras")


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
# UTILIDADES
# ─────────────────────────────────────────────
class ErrorCompras(Exception):
    def __init__(self, mensaje, status=400):
        super().__init__(mensaje)
        self.mensaje = mensaje
        self.status = status


def _serializar(v):
    if isinstance(v, datetime):
        return v.astimezone(LIMA_TZ).isoformat() if v.tzinfo else v.isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, Decimal):
        return float(v)
    return str(v)


def _json(data, status=200):
    return (json.dumps(data, default=_serializar, ensure_ascii=False), status, HEADERS)


def _error(mensaje, status=400):
    return _json({"ok": False, "error": mensaje}, status)


def _body(request):
    return request.get_json(force=True, silent=True) or {}


def _texto(valor, maximo=250):
    if valor is None:
        return None
    t = str(valor).strip()
    return t[:maximo] if t else None


def _bool(valor):
    if isinstance(valor, bool):
        return valor
    return str(valor).strip().lower() in ("1", "true", "si", "sí", "on", "yes")


def _fecha(valor, campo):
    if valor in (None, ""):
        return None
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
    except ValueError:
        raise ErrorCompras(f"Fecha inválida en {campo} (use AAAA-MM-DD)")


def _hoy_lima():
    return datetime.now(LIMA_TZ).date()


def ruc_valido(ruc):
    """RUC peruano: 11 dígitos, prefijo válido y dígito verificador (módulo 11)."""
    if not ruc or not re.fullmatch(r"\d{11}", ruc):
        return False
    if ruc[:2] not in ("10", "15", "16", "17", "20"):
        return False
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(d) * p for d, p in zip(ruc[:10], pesos))
    dv = (11 - suma % 11) % 10
    return dv == int(ruc[10])


def _correo_valido(correo):
    return bool(correo) and bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo))


# ─────────────────────────────────────────────
# CONTRASEÑAS Y SESIONES
# ─────────────────────────────────────────────
_ITERACIONES = 200_000


def hash_password(password):
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), _ITERACIONES)
    return f"pbkdf2_sha256${_ITERACIONES}${salt}${dk.hex()}"


def check_password(password, password_hash):
    try:
        _, iteraciones, salt, esperado = password_hash.split("$")
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), int(iteraciones))
        return hmac.compare_digest(dk.hex(), esperado)
    except Exception:
        return False


_HASH_FICTICIO = hash_password(secrets.token_hex(8))


def password_temporal():
    letras = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz"
    digitos = "23456789"
    while True:
        p = "".join(secrets.choice(letras + digitos) for _ in range(10))
        if any(c in digitos for c in p) and any(c in letras for c in p):
            return p


def _validar_password_nueva(nueva, actual=None):
    if not nueva or len(nueva) < 8:
        raise ErrorCompras("La contraseña debe tener al menos 8 caracteres")
    if not re.search(r"[A-Za-z]", nueva) or not re.search(r"\d", nueva):
        raise ErrorCompras("La contraseña debe combinar letras y números")
    if actual and nueva == actual:
        raise ErrorCompras("La nueva contraseña debe ser distinta a la actual")


def _hash_token(token):
    return hashlib.sha256(token.encode()).hexdigest()


def _token_de(request):
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:].strip() or None
    return None


def _crear_sesion(cur, tipo, usuario_id=None, ruc=None):
    token = secrets.token_urlsafe(32)
    horas = SESION_HORAS_ADMIN if tipo == "admin" else SESION_HORAS_PROVEEDOR
    cur.execute(
        """INSERT INTO compras_sesiones (token_hash, tipo, usuario_id, ruc, expira_en)
           VALUES (%s, %s, %s, %s, now() + %s * interval '1 hour')""",
        (_hash_token(token), tipo, usuario_id, ruc, horas),
    )
    return token


def _sesion_admin(request):
    token = _token_de(request)
    if not token:
        return None
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT u.id, u.usuario, u.nombre, u.rol, u.debe_cambiar
                 FROM compras_sesiones s
                 JOIN compras_usuarios u ON u.id = s.usuario_id
                WHERE s.token_hash = %s AND s.tipo = 'admin'
                  AND s.expira_en > now() AND u.activo""",
            (_hash_token(token),),
        )
        r = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    if not r:
        return None
    return {"id": r[0], "usuario": r[1], "nombre": r[2], "rol": r[3],
            "debe_cambiar": r[4], "token_hash": _hash_token(token)}


def _sesion_proveedor(request):
    token = _token_de(request)
    if not token:
        return None
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT e.ruc, e.razon_social, e.debe_cambiar
                 FROM compras_sesiones s
                 JOIN compras_empresas e ON e.ruc = s.ruc
                WHERE s.token_hash = %s AND s.tipo = 'proveedor'
                  AND s.expira_en > now() AND e.activo""",
            (_hash_token(token),),
        )
        r = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    if not r:
        return None
    return {"ruc": r[0], "razon_social": r[1], "debe_cambiar": r[2],
            "token_hash": _hash_token(token)}


def _exigir_admin(request, rol=None, permitir_cambio_pendiente=False):
    s = _sesion_admin(request)
    if not s:
        return None, _error("Sesión no válida o expirada", 401)
    if s["debe_cambiar"] and not permitir_cambio_pendiente:
        return None, _error("Debes cambiar tu contraseña antes de continuar", 403)
    if rol and s["rol"] != rol:
        return None, _error("No tienes permisos para esta acción", 403)
    return s, None


def _exigir_proveedor(request, permitir_cambio_pendiente=False):
    s = _sesion_proveedor(request)
    if not s:
        return None, _error("Sesión no válida o expirada", 401)
    if s["debe_cambiar"] and not permitir_cambio_pendiente:
        return None, _error("Debes cambiar tu contraseña antes de continuar", 403)
    return s, None


def _audit(cur, actor, accion, ruc=None, detalle=None):
    cur.execute(
        "INSERT INTO compras_auditoria (actor, accion, ruc, detalle) VALUES (%s, %s, %s, %s)",
        (actor, accion, ruc, json.dumps(detalle, default=_serializar) if detalle else None),
    )


def _login_generico(cur, tabla, campo_llave, valor_llave, password):
    """
    Lógica de login con bloqueo por intentos (tabla compras_usuarios o compras_empresas).
    Devuelve la fila (id/ruc, activo, debe_cambiar) o lanza ErrorCompras.
    """
    cur.execute(
        f"""SELECT {campo_llave}, password_hash, activo, intentos_fallidos,
                   bloqueado_hasta, debe_cambiar, bloqueado_hasta > now()
              FROM {tabla}
             WHERE {"lower(usuario) = lower(%s)" if tabla == "compras_usuarios" else "ruc = %s"}""",
        (valor_llave,),
    )
    r = cur.fetchone()
    if not r:
        check_password(password, _HASH_FICTICIO)   # mismo tiempo de respuesta
        raise ErrorCompras("Credenciales incorrectas", 401)
    llave, pwd_hash, activo, intentos, _, debe_cambiar, bloqueado = r
    if bloqueado:
        raise ErrorCompras(
            f"Cuenta bloqueada temporalmente por intentos fallidos. Intenta en {BLOQUEO_MINUTOS} minutos.", 423)
    if not check_password(password, pwd_hash):
        intentos = (intentos or 0) + 1
        if intentos >= MAX_INTENTOS_LOGIN:
            cur.execute(
                f"""UPDATE {tabla} SET intentos_fallidos = 0,
                           bloqueado_hasta = now() + %s * interval '1 minute'
                     WHERE {campo_llave} = %s""",
                (BLOQUEO_MINUTOS, llave),
            )
        else:
            cur.execute(f"UPDATE {tabla} SET intentos_fallidos = %s WHERE {campo_llave} = %s",
                        (intentos, llave))
        return None
    if not activo:
        raise ErrorCompras("La cuenta está inactiva. Comunícate con el equipo de Compras.", 403)
    cur.execute(
        f"""UPDATE {tabla} SET intentos_fallidos = 0, bloqueado_hasta = NULL, ultimo_acceso = now()
             WHERE {campo_llave} = %s""",
        (llave,),
    )
    return {"llave": llave, "debe_cambiar": debe_cambiar}


# ─────────────────────────────────────────────
# STORAGE (bucket privado + URL firmada)
# ─────────────────────────────────────────────
_storage_client = None
_credenciales_firma = None


def _storage():
    global _storage_client
    if _storage_client is None:
        _storage_client = storage.Client()
    return _storage_client


def _credenciales():
    """Firma vía IAM SignBlob (la cuenta de servicio necesita roles/iam.serviceAccountTokenCreator)."""
    global _credenciales_firma
    if _credenciales_firma is None:
        _credenciales_firma, _ = google.auth.default()
    if not _credenciales_firma.valid:
        _credenciales_firma.refresh(google_auth_requests.Request())
    return _credenciales_firma


def _subir_archivo(ruta, contenido, content_type):
    blob = _storage().bucket(COMPRAS_BUCKET).blob(ruta)
    blob.upload_from_string(contenido, content_type=content_type)
    return f"gs://{COMPRAS_BUCKET}/{ruta}"


def _borrar_archivo(ruta_gs):
    try:
        bucket_name, path = ruta_gs.replace("gs://", "", 1).split("/", 1)
        _storage().bucket(bucket_name).blob(path).delete()
    except Exception:
        log.warning("No se pudo borrar %s", ruta_gs)


def _url_firmada(ruta_gs, nombre_archivo=None):
    bucket_name, path = ruta_gs.replace("gs://", "", 1).split("/", 1)
    blob = _storage().bucket(bucket_name).blob(path)
    cred = _credenciales()
    nombre = re.sub(r'[^\w.\- ]', "_", nombre_archivo or path.rsplit("/", 1)[-1])
    return blob.generate_signed_url(
        version="v4",
        expiration=timedelta(minutes=MINUTOS_URL_FIRMADA),
        method="GET",
        service_account_email=cred.service_account_email,
        access_token=cred.token,
        response_disposition=f'inline; filename="{nombre}"',
    )


def _detectar_tipo(contenido):
    if contenido[:5] == b"%PDF-":
        return "application/pdf"
    if contenido[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if contenido[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    return None


# ─────────────────────────────────────────────
# CHECKLIST DE DOCUMENTOS DE UNA EMPRESA
# ─────────────────────────────────────────────
def _checklist(cur, ruc):
    cur.execute(
        f"""SELECT t.id, t.codigo, t.nombre, t.descripcion, t.obligatorio, t.requiere_vencimiento,
                   d.id, d.estado, d.observacion, d.nombre_original, d.fecha_emision,
                   d.fecha_vencimiento, d.subido_en, d.revisado_en,
                   (d.fecha_vencimiento IS NOT NULL AND d.fecha_vencimiento < {HOY_LIMA}),
                   (d.fecha_vencimiento IS NOT NULL
                      AND d.fecha_vencimiento >= {HOY_LIMA}
                      AND d.fecha_vencimiento < {HOY_LIMA} + %s)
              FROM compras_tipos_documento t
              LEFT JOIN compras_documentos d
                     ON d.tipo_id = t.id AND d.ruc = %s AND d.vigente
             WHERE t.activo
             ORDER BY t.orden, t.nombre""",
        (DIAS_POR_VENCER, ruc),
    )
    items = []
    for r in cur.fetchall():
        if r[6] is None:
            estado = "pendiente"
        elif r[14]:
            estado = "vencido"
        else:
            estado = r[7]
        items.append({
            "tipo_id": r[0], "codigo": r[1], "nombre": r[2], "descripcion": r[3],
            "obligatorio": r[4], "requiere_vencimiento": r[5],
            "documento_id": r[6], "estado": estado, "observacion": r[8],
            "nombre_archivo": r[9], "fecha_emision": r[10], "fecha_vencimiento": r[11],
            "subido_en": r[12], "revisado_en": r[13], "por_vencer": bool(r[15]),
        })
    obligatorios = [i for i in items if i["obligatorio"]]
    aprobados = [i for i in obligatorios if i["estado"] == "aprobado"]
    resumen = {
        "obligatorios": len(obligatorios),
        "aprobados": len(aprobados),
        "en_revision": sum(1 for i in items if i["estado"] == "en_revision"),
        "observados": sum(1 for i in items if i["estado"] == "observado"),
        "vencidos": sum(1 for i in items if i["estado"] == "vencido"),
        "pendientes": sum(1 for i in obligatorios if i["estado"] == "pendiente"),
        "porcentaje": round(100 * len(aprobados) / len(obligatorios)) if obligatorios else 100,
    }
    resumen["completo"] = resumen["aprobados"] == resumen["obligatorios"]
    return items, resumen


# ═════════════════════════════════════════════
# PROVEEDOR
# ═════════════════════════════════════════════
def handle_proveedor_login(request):
    data = _body(request)
    ruc = re.sub(r"\D", "", str(data.get("ruc", "")))
    password = str(data.get("password", ""))
    if not ruc or not password:
        return _error("Ingresa RUC y contraseña")
    conn = get_connection()
    try:
        cur = conn.cursor()
        try:
            r = _login_generico(cur, "compras_empresas", "ruc", ruc, password)
        except ErrorCompras:
            conn.commit()
            raise
        if not r:
            conn.commit()
            return _error("Credenciales incorrectas", 401)
        token = _crear_sesion(cur, "proveedor", ruc=ruc)
        _audit(cur, f"proveedor:{ruc}", "login", ruc)
        conn.commit()
        cur.close()
        return _json({"ok": True, "token": token, "debe_cambiar": r["debe_cambiar"]})
    finally:
        conn.close()


def handle_proveedor_cambiar_password(request):
    s, err = _exigir_proveedor(request, permitir_cambio_pendiente=True)
    if err:
        return err
    data = _body(request)
    actual, nueva = str(data.get("actual", "")), str(data.get("nueva", ""))
    _validar_password_nueva(nueva, actual)
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT password_hash FROM compras_empresas WHERE ruc = %s", (s["ruc"],))
        r = cur.fetchone()
        if not r or not check_password(actual, r[0]):
            return _error("La contraseña actual no es correcta")
        cur.execute(
            """UPDATE compras_empresas SET password_hash = %s, debe_cambiar = FALSE, actualizado_en = now()
                WHERE ruc = %s""",
            (hash_password(nueva), s["ruc"]),
        )
        cur.execute("DELETE FROM compras_sesiones WHERE ruc = %s AND token_hash <> %s",
                    (s["ruc"], s["token_hash"]))
        _audit(cur, f"proveedor:{s['ruc']}", "cambiar_password", s["ruc"])
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


def handle_logout(request):
    token = _token_de(request)
    if token:
        conn = get_connection()
        try:
            cur = conn.cursor()
            cur.execute("DELETE FROM compras_sesiones WHERE token_hash = %s", (_hash_token(token),))
            conn.commit()
            cur.close()
        finally:
            conn.close()
    return _json({"ok": True})


def handle_proveedor_yo(request):
    s, err = _exigir_proveedor(request, permitir_cambio_pendiente=True)
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT ruc, razon_social, nombre_comercial, direccion, contacto_nombre, correo, telefono
                 FROM compras_empresas WHERE ruc = %s""",
            (s["ruc"],),
        )
        e = cur.fetchone()
        items, resumen = ([], None) if s["debe_cambiar"] else _checklist(cur, s["ruc"])
        cur.close()
        return _json({
            "ok": True,
            "debe_cambiar": s["debe_cambiar"],
            "empresa": {
                "ruc": e[0], "razon_social": e[1], "nombre_comercial": e[2], "direccion": e[3],
                "contacto_nombre": e[4], "correo": e[5], "telefono": e[6],
            },
            "documentos": items,
            "resumen": resumen,
            "max_archivo_mb": MAX_ARCHIVO_MB,
        })
    finally:
        conn.close()


def handle_proveedor_datos(request):
    s, err = _exigir_proveedor(request)
    if err:
        return err
    data = _body(request)
    contacto = _texto(data.get("contacto_nombre"), 150)
    correo = _texto(data.get("correo"), 150)
    telefono = _texto(data.get("telefono"), 30)
    if correo and not _correo_valido(correo):
        return _error("Correo electrónico inválido")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """UPDATE compras_empresas
                  SET contacto_nombre = %s, correo = %s, telefono = %s, actualizado_en = now()
                WHERE ruc = %s""",
            (contacto, correo, telefono, s["ruc"]),
        )
        _audit(cur, f"proveedor:{s['ruc']}", "actualizar_contacto", s["ruc"],
               {"contacto_nombre": contacto, "correo": correo, "telefono": telefono})
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


def handle_proveedor_subir_documento(request):
    s, err = _exigir_proveedor(request)
    if err:
        return err

    tipo_id = request.form.get("tipo_id", "")
    archivo = request.files.get("archivo")
    if not tipo_id.isdigit():
        return _error("Tipo de documento inválido")
    if not archivo or not archivo.filename:
        return _error("Adjunta un archivo")

    contenido = archivo.read(MAX_ARCHIVO_MB * 1024 * 1024 + 1)
    if len(contenido) == 0:
        return _error("El archivo está vacío")
    if len(contenido) > MAX_ARCHIVO_MB * 1024 * 1024:
        return _error(f"El archivo supera el máximo de {MAX_ARCHIVO_MB} MB")
    content_type = _detectar_tipo(contenido)
    if not content_type:
        return _error("Formato no permitido. Sube PDF, JPG o PNG.")

    fecha_emision = _fecha(request.form.get("fecha_emision"), "fecha de emisión")
    fecha_venc = _fecha(request.form.get("fecha_vencimiento"), "fecha de vencimiento")
    if fecha_emision and fecha_emision > _hoy_lima():
        return _error("La fecha de emisión no puede ser futura")

    conn = get_connection()
    ruta_gs = None
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT id, codigo, nombre, requiere_vencimiento
                 FROM compras_tipos_documento WHERE id = %s AND activo""",
            (int(tipo_id),),
        )
        t = cur.fetchone()
        if not t:
            return _error("Tipo de documento no disponible")
        if t[3]:
            if not fecha_venc:
                return _error(f"Indica la fecha de vencimiento de: {t[2]}")
            if fecha_venc < _hoy_lima():
                return _error("El documento ya está vencido; sube uno vigente")

        nombre_original = re.sub(r"[\r\n\"]", "", archivo.filename)[:255]
        sello = datetime.now(LIMA_TZ).strftime("%Y%m%d-%H%M%S")
        ruta = f"{s['ruc']}/{t[1]}/{sello}_{secrets.token_hex(4)}.{TIPOS_ARCHIVO[content_type]}"
        ruta_gs = _subir_archivo(ruta, contenido, content_type)

        cur.execute(
            "UPDATE compras_documentos SET vigente = FALSE WHERE ruc = %s AND tipo_id = %s AND vigente",
            (s["ruc"], t[0]),
        )
        cur.execute(
            """INSERT INTO compras_documentos
                   (ruc, tipo_id, ruta_gcs, nombre_original, content_type, tamano_bytes, sha256,
                    fecha_emision, fecha_vencimiento, estado, vigente)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, 'en_revision', TRUE)
               RETURNING id""",
            (s["ruc"], t[0], ruta_gs, nombre_original, content_type, len(contenido),
             hashlib.sha256(contenido).hexdigest(), fecha_emision, fecha_venc),
        )
        doc_id = cur.fetchone()[0]
        _audit(cur, f"proveedor:{s['ruc']}", "subir_documento", s["ruc"],
               {"documento_id": doc_id, "tipo": t[1], "archivo": nombre_original})
        conn.commit()
        cur.close()
        return _json({"ok": True, "documento_id": doc_id})
    except psycopg2.errors.UniqueViolation:
        conn.rollback()
        if ruta_gs:
            _borrar_archivo(ruta_gs)
        return _error("Se está subiendo otro archivo del mismo tipo. Intenta nuevamente.", 409)
    except Exception:
        conn.rollback()
        if ruta_gs:
            _borrar_archivo(ruta_gs)
        raise
    finally:
        conn.close()


def _descargar(ruc_dueno, documento_id):
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT ruc, ruta_gcs, nombre_original FROM compras_documentos WHERE id = %s",
                    (documento_id,))
        r = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    if not r or (ruc_dueno and r[0] != ruc_dueno):
        return _error("Documento no encontrado", 404)
    return _json({"ok": True, "url": _url_firmada(r[1], r[2])})


def handle_proveedor_descargar(request, documento_id):
    s, err = _exigir_proveedor(request)
    if err:
        return err
    return _descargar(s["ruc"], documento_id)


# ═════════════════════════════════════════════
# COMPRAS — AUTENTICACIÓN Y USUARIOS
# ═════════════════════════════════════════════
def handle_admin_login(request):
    data = _body(request)
    usuario = str(data.get("usuario", "")).strip()
    password = str(data.get("password", ""))
    if not usuario or not password:
        return _error("Ingresa usuario y contraseña")
    conn = get_connection()
    try:
        cur = conn.cursor()
        try:
            r = _login_generico(cur, "compras_usuarios", "id", usuario, password)
        except ErrorCompras:
            conn.commit()
            raise
        if not r:
            conn.commit()
            return _error("Credenciales incorrectas", 401)
        token = _crear_sesion(cur, "admin", usuario_id=r["llave"])
        cur.execute("SELECT usuario, nombre, rol FROM compras_usuarios WHERE id = %s", (r["llave"],))
        u = cur.fetchone()
        _audit(cur, f"admin:{u[0]}", "login")
        conn.commit()
        cur.close()
        return _json({"ok": True, "token": token, "debe_cambiar": r["debe_cambiar"],
                      "usuario": {"usuario": u[0], "nombre": u[1], "rol": u[2]}})
    finally:
        conn.close()


def handle_admin_cambiar_password(request):
    s, err = _exigir_admin(request, permitir_cambio_pendiente=True)
    if err:
        return err
    data = _body(request)
    actual, nueva = str(data.get("actual", "")), str(data.get("nueva", ""))
    _validar_password_nueva(nueva, actual)
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT password_hash FROM compras_usuarios WHERE id = %s", (s["id"],))
        r = cur.fetchone()
        if not r or not check_password(actual, r[0]):
            return _error("La contraseña actual no es correcta")
        cur.execute("UPDATE compras_usuarios SET password_hash = %s, debe_cambiar = FALSE WHERE id = %s",
                    (hash_password(nueva), s["id"]))
        cur.execute("DELETE FROM compras_sesiones WHERE usuario_id = %s AND token_hash <> %s",
                    (s["id"], s["token_hash"]))
        _audit(cur, f"admin:{s['usuario']}", "cambiar_password")
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


def handle_admin_yo(request):
    s, err = _exigir_admin(request, permitir_cambio_pendiente=True)
    if err:
        return err
    return _json({"ok": True, "usuario": {k: s[k] for k in ("id", "usuario", "nombre", "rol", "debe_cambiar")}})


def handle_bootstrap_admin(request):
    """Crea el PRIMER usuario admin. main.py exige ApiKey para esta ruta."""
    data = _body(request)
    usuario = _texto(data.get("usuario"), 60)
    nombre = _texto(data.get("nombre"), 150)
    if not usuario or not nombre:
        return _error("usuario y nombre son obligatorios")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT count(*) FROM compras_usuarios WHERE rol = 'admin'")
        if cur.fetchone()[0] > 0:
            return _error("Ya existe un administrador; crea usuarios desde el módulo", 409)
        temporal = password_temporal()
        cur.execute(
            """INSERT INTO compras_usuarios (usuario, nombre, rol, password_hash, debe_cambiar)
               VALUES (%s, %s, 'admin', %s, TRUE) RETURNING id""",
            (usuario, nombre, hash_password(temporal)),
        )
        uid = cur.fetchone()[0]
        _audit(cur, "sistema", "bootstrap_admin", None, {"usuario": usuario})
        conn.commit()
        cur.close()
        return _json({"ok": True, "id": uid, "usuario": usuario, "password_temporal": temporal})
    finally:
        conn.close()


def handle_admin_usuarios_listar(request):
    s, err = _exigir_admin(request, rol="admin")
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT id, usuario, nombre, rol, activo, debe_cambiar, ultimo_acceso, creado_en,
                      bloqueado_hasta > now()
                 FROM compras_usuarios ORDER BY activo DESC, nombre"""
        )
        filas = [{"id": r[0], "usuario": r[1], "nombre": r[2], "rol": r[3], "activo": r[4],
                  "debe_cambiar": r[5], "ultimo_acceso": r[6], "creado_en": r[7], "bloqueado": bool(r[8])}
                 for r in cur.fetchall()]
        cur.close()
        return _json({"ok": True, "usuarios": filas})
    finally:
        conn.close()


def handle_admin_usuarios_crear(request):
    s, err = _exigir_admin(request, rol="admin")
    if err:
        return err
    data = _body(request)
    usuario = _texto(data.get("usuario"), 60)
    nombre = _texto(data.get("nombre"), 150)
    rol = str(data.get("rol", "analista")).strip().lower()
    if not usuario or not re.fullmatch(r"[A-Za-z0-9._-]{3,60}", usuario):
        return _error("Usuario inválido (3-60 caracteres: letras, números, punto, guion)")
    if not nombre:
        return _error("El nombre es obligatorio")
    if rol not in ROLES_VALIDOS:
        return _error("Rol inválido")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM compras_usuarios WHERE lower(usuario) = lower(%s)", (usuario,))
        if cur.fetchone():
            return _error("Ese usuario ya existe", 409)
        temporal = password_temporal()
        cur.execute(
            """INSERT INTO compras_usuarios (usuario, nombre, rol, password_hash, debe_cambiar)
               VALUES (%s, %s, %s, %s, TRUE) RETURNING id""",
            (usuario, nombre, rol, hash_password(temporal)),
        )
        uid = cur.fetchone()[0]
        _audit(cur, f"admin:{s['usuario']}", "crear_usuario", None, {"usuario": usuario, "rol": rol})
        conn.commit()
        cur.close()
        return _json({"ok": True, "id": uid, "password_temporal": temporal})
    finally:
        conn.close()


def handle_admin_usuario_reset(request, usuario_id):
    s, err = _exigir_admin(request, rol="admin")
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        temporal = password_temporal()
        cur.execute(
            """UPDATE compras_usuarios
                  SET password_hash = %s, debe_cambiar = TRUE, intentos_fallidos = 0, bloqueado_hasta = NULL
                WHERE id = %s RETURNING usuario""",
            (hash_password(temporal), usuario_id),
        )
        r = cur.fetchone()
        if not r:
            return _error("Usuario no encontrado", 404)
        cur.execute("DELETE FROM compras_sesiones WHERE usuario_id = %s", (usuario_id,))
        _audit(cur, f"admin:{s['usuario']}", "reset_usuario", None, {"usuario": r[0]})
        conn.commit()
        cur.close()
        return _json({"ok": True, "password_temporal": temporal})
    finally:
        conn.close()


def handle_admin_usuario_estado(request, usuario_id):
    s, err = _exigir_admin(request, rol="admin")
    if err:
        return err
    activo = _bool(_body(request).get("activo"))
    if usuario_id == s["id"] and not activo:
        return _error("No puedes desactivar tu propio usuario")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("UPDATE compras_usuarios SET activo = %s WHERE id = %s RETURNING usuario",
                    (activo, usuario_id))
        r = cur.fetchone()
        if not r:
            return _error("Usuario no encontrado", 404)
        if not activo:
            cur.execute("DELETE FROM compras_sesiones WHERE usuario_id = %s", (usuario_id,))
        _audit(cur, f"admin:{s['usuario']}", "estado_usuario", None, {"usuario": r[0], "activo": activo})
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


# ═════════════════════════════════════════════
# COMPRAS — EMPRESAS
# ═════════════════════════════════════════════
def handle_admin_consulta_ruc(request):
    s, err = _exigir_admin(request)
    if err:
        return err
    ruc = re.sub(r"\D", "", request.args.get("ruc", ""))
    if not ruc_valido(ruc):
        return _error("RUC inválido")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT razon_social FROM compras_empresas WHERE ruc = %s", (ruc,))
        existe = cur.fetchone()
        cur.close()
    finally:
        conn.close()
    if existe:
        return _json({"ok": True, "existe": True, "razon_social": existe[0]})

    token = os.environ.get("APISPERU_TOKEN")
    if not token:
        return _json({"ok": True, "existe": False, "sunat": None})
    try:
        resp = requests.get(f"https://dniruc.apisperu.com/api/v1/ruc/{ruc}",
                            params={"token": token}, timeout=10)
        obj = resp.json() if resp.ok else {}
    except Exception:
        obj = {}
    if not obj.get("razonSocial"):
        return _json({"ok": True, "existe": False, "sunat": None})
    partes_dir = [obj.get("direccion"), obj.get("distrito"), obj.get("provincia"), obj.get("departamento")]
    return _json({"ok": True, "existe": False, "sunat": {
        "razon_social": obj.get("razonSocial"),
        "nombre_comercial": obj.get("nombreComercial") if obj.get("nombreComercial") not in (None, "-") else None,
        "direccion": " - ".join(p for p in partes_dir if p and p != "-") or None,
        "estado": obj.get("estado"),
        "condicion": obj.get("condicion"),
    }})


def handle_admin_empresas_listar(request):
    s, err = _exigir_admin(request)
    if err:
        return err
    q = request.args.get("q", "").strip()
    filtro = request.args.get("filtro", "todas")
    patron = f"%{q}%"
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            f"""WITH req AS (SELECT id FROM compras_tipos_documento WHERE activo AND obligatorio),
                     doc AS (SELECT d.ruc, d.tipo_id, d.estado,
                                    (d.fecha_vencimiento IS NOT NULL AND d.fecha_vencimiento < {HOY_LIMA}) AS vencido
                               FROM compras_documentos d
                               JOIN compras_tipos_documento t ON t.id = d.tipo_id AND t.activo
                              WHERE d.vigente)
                SELECT e.ruc, e.razon_social, e.contacto_nombre, e.correo, e.activo,
                       e.ultimo_acceso, e.creado_en, e.debe_cambiar,
                       (SELECT count(*) FROM req),
                       count(*) FILTER (WHERE doc.estado = 'aprobado' AND NOT doc.vencido
                                          AND doc.tipo_id IN (SELECT id FROM req)),
                       count(*) FILTER (WHERE doc.estado = 'en_revision' AND NOT doc.vencido),
                       count(*) FILTER (WHERE doc.estado = 'observado' AND NOT doc.vencido),
                       count(*) FILTER (WHERE doc.vencido)
                  FROM compras_empresas e
                  LEFT JOIN doc ON doc.ruc = e.ruc
                 WHERE (%s = '' OR e.ruc LIKE %s OR e.razon_social ILIKE %s OR e.nombre_comercial ILIKE %s)
                 GROUP BY e.ruc
                 ORDER BY e.razon_social
                 LIMIT 1000""",
            (q, patron, patron, patron),
        )
        empresas = []
        for r in cur.fetchall():
            obligatorios, aprobados = r[8], r[9]
            empresas.append({
                "ruc": r[0], "razon_social": r[1], "contacto_nombre": r[2], "correo": r[3],
                "activo": r[4], "ultimo_acceso": r[5], "creado_en": r[6], "nunca_ingreso": r[7],
                "obligatorios": obligatorios, "aprobados": aprobados,
                "en_revision": r[10], "observados": r[11], "vencidos": r[12],
                "completo": aprobados >= obligatorios,
                "porcentaje": round(100 * aprobados / obligatorios) if obligatorios else 100,
            })
        cur.close()
    finally:
        conn.close()

    if filtro == "por_revisar":
        empresas = [e for e in empresas if e["en_revision"] > 0 and e["activo"]]
    elif filtro == "incompletas":
        empresas = [e for e in empresas if not e["completo"] and e["activo"]]
    elif filtro == "completas":
        empresas = [e for e in empresas if e["completo"] and e["activo"]]
    elif filtro == "inactivas":
        empresas = [e for e in empresas if not e["activo"]]
    return _json({"ok": True, "empresas": empresas})


def _campos_empresa(data):
    campos = {
        "razon_social": _texto(data.get("razon_social"), 250),
        "nombre_comercial": _texto(data.get("nombre_comercial"), 250),
        "direccion": _texto(data.get("direccion"), 400),
        "rubro": _texto(data.get("rubro"), 150),
        "contacto_nombre": _texto(data.get("contacto_nombre"), 150),
        "correo": _texto(data.get("correo"), 150),
        "telefono": _texto(data.get("telefono"), 30),
        "notas": _texto(data.get("notas"), 2000),
    }
    if not campos["razon_social"]:
        raise ErrorCompras("La razón social es obligatoria")
    if campos["correo"] and not _correo_valido(campos["correo"]):
        raise ErrorCompras("Correo electrónico inválido")
    return campos


def _url_portal(request):
    return f"https://{request.host}/compras/portal"


def handle_admin_empresa_crear(request):
    s, err = _exigir_admin(request)
    if err:
        return err
    data = _body(request)
    ruc = re.sub(r"\D", "", str(data.get("ruc", "")))
    if not ruc_valido(ruc):
        return _error("RUC inválido (11 dígitos y dígito verificador correcto)")
    campos = _campos_empresa(data)
    temporal = password_temporal()
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO compras_empresas
                   (ruc, razon_social, nombre_comercial, direccion, rubro, contacto_nombre, correo,
                    telefono, notas, password_hash, debe_cambiar, creado_por)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, TRUE, %s)
               ON CONFLICT (ruc) DO NOTHING""",
            (ruc, campos["razon_social"], campos["nombre_comercial"], campos["direccion"], campos["rubro"],
             campos["contacto_nombre"], campos["correo"], campos["telefono"], campos["notas"],
             hash_password(temporal), s["id"]),
        )
        if cur.rowcount == 0:
            return _error("Ya existe una empresa registrada con ese RUC", 409)
        _audit(cur, f"admin:{s['usuario']}", "crear_empresa", ruc, {"razon_social": campos["razon_social"]})
        conn.commit()
        cur.close()
        return _json({"ok": True, "ruc": ruc, "password_temporal": temporal, "url_portal": _url_portal(request)})
    finally:
        conn.close()


def handle_admin_empresa_detalle(request, ruc):
    s, err = _exigir_admin(request)
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT ruc, razon_social, nombre_comercial, direccion, rubro, contacto_nombre, correo,
                      telefono, notas, activo, debe_cambiar, ultimo_acceso, creado_en, actualizado_en,
                      bloqueado_hasta > now()
                 FROM compras_empresas WHERE ruc = %s""",
            (ruc,),
        )
        e = cur.fetchone()
        if not e:
            return _error("Empresa no encontrada", 404)
        items, resumen = _checklist(cur, ruc)
        cur.execute(
            """SELECT d.id, t.nombre, d.nombre_original, d.estado, d.observacion, d.vigente,
                      d.fecha_vencimiento, d.subido_en, d.revisado_en, u.nombre, d.tamano_bytes
                 FROM compras_documentos d
                 JOIN compras_tipos_documento t ON t.id = d.tipo_id
                 LEFT JOIN compras_usuarios u ON u.id = d.revisado_por
                WHERE d.ruc = %s
                ORDER BY d.subido_en DESC
                LIMIT 200""",
            (ruc,),
        )
        historial = [{"id": r[0], "tipo": r[1], "nombre_archivo": r[2], "estado": r[3], "observacion": r[4],
                      "vigente": r[5], "fecha_vencimiento": r[6], "subido_en": r[7], "revisado_en": r[8],
                      "revisado_por": r[9], "tamano_bytes": r[10]} for r in cur.fetchall()]
        cur.execute(
            """SELECT actor, accion, detalle, creado_en FROM compras_auditoria
                WHERE ruc = %s ORDER BY creado_en DESC LIMIT 50""",
            (ruc,),
        )
        auditoria = [{"actor": r[0], "accion": r[1], "detalle": r[2], "creado_en": r[3]} for r in cur.fetchall()]
        cur.close()
        return _json({
            "ok": True,
            "empresa": {
                "ruc": e[0], "razon_social": e[1], "nombre_comercial": e[2], "direccion": e[3], "rubro": e[4],
                "contacto_nombre": e[5], "correo": e[6], "telefono": e[7], "notas": e[8], "activo": e[9],
                "nunca_ingreso": e[10], "ultimo_acceso": e[11], "creado_en": e[12], "actualizado_en": e[13],
                "bloqueado": bool(e[14]),
            },
            "documentos": items, "resumen": resumen, "historial": historial, "auditoria": auditoria,
            "url_portal": _url_portal(request),
        })
    finally:
        conn.close()


def handle_admin_empresa_editar(request, ruc):
    s, err = _exigir_admin(request)
    if err:
        return err
    campos = _campos_empresa(_body(request))
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """UPDATE compras_empresas
                  SET razon_social = %s, nombre_comercial = %s, direccion = %s, rubro = %s,
                      contacto_nombre = %s, correo = %s, telefono = %s, notas = %s, actualizado_en = now()
                WHERE ruc = %s""",
            (campos["razon_social"], campos["nombre_comercial"], campos["direccion"], campos["rubro"],
             campos["contacto_nombre"], campos["correo"], campos["telefono"], campos["notas"], ruc),
        )
        if cur.rowcount == 0:
            return _error("Empresa no encontrada", 404)
        _audit(cur, f"admin:{s['usuario']}", "editar_empresa", ruc, campos)
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


def handle_admin_empresa_reset(request, ruc):
    s, err = _exigir_admin(request)
    if err:
        return err
    temporal = password_temporal()
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """UPDATE compras_empresas
                  SET password_hash = %s, debe_cambiar = TRUE, intentos_fallidos = 0,
                      bloqueado_hasta = NULL, actualizado_en = now()
                WHERE ruc = %s""",
            (hash_password(temporal), ruc),
        )
        if cur.rowcount == 0:
            return _error("Empresa no encontrada", 404)
        cur.execute("DELETE FROM compras_sesiones WHERE ruc = %s", (ruc,))
        _audit(cur, f"admin:{s['usuario']}", "reset_password_empresa", ruc)
        conn.commit()
        cur.close()
        return _json({"ok": True, "password_temporal": temporal, "url_portal": _url_portal(request)})
    finally:
        conn.close()


def handle_admin_empresa_estado(request, ruc):
    s, err = _exigir_admin(request)
    if err:
        return err
    activo = _bool(_body(request).get("activo"))
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("UPDATE compras_empresas SET activo = %s, actualizado_en = now() WHERE ruc = %s",
                    (activo, ruc))
        if cur.rowcount == 0:
            return _error("Empresa no encontrada", 404)
        if not activo:
            cur.execute("DELETE FROM compras_sesiones WHERE ruc = %s", (ruc,))
        _audit(cur, f"admin:{s['usuario']}", "activar_empresa" if activo else "desactivar_empresa", ruc)
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


# ═════════════════════════════════════════════
# COMPRAS — REVISIÓN DE DOCUMENTOS
# ═════════════════════════════════════════════
def handle_admin_pendientes(request):
    s, err = _exigir_admin(request)
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT d.id, d.ruc, e.razon_social, t.nombre, t.obligatorio, d.nombre_original,
                      d.fecha_emision, d.fecha_vencimiento, d.subido_en, t.requiere_vencimiento
                 FROM compras_documentos d
                 JOIN compras_empresas e ON e.ruc = d.ruc AND e.activo
                 JOIN compras_tipos_documento t ON t.id = d.tipo_id AND t.activo
                WHERE d.vigente AND d.estado = 'en_revision'
                ORDER BY d.subido_en
                LIMIT 500"""
        )
        filas = [{"id": r[0], "ruc": r[1], "razon_social": r[2], "tipo": r[3], "obligatorio": r[4],
                  "nombre_archivo": r[5], "fecha_emision": r[6], "fecha_vencimiento": r[7],
                  "subido_en": r[8], "requiere_vencimiento": r[9]} for r in cur.fetchall()]
        cur.close()
        return _json({"ok": True, "pendientes": filas})
    finally:
        conn.close()


def handle_admin_revisar_documento(request, documento_id):
    s, err = _exigir_admin(request)
    if err:
        return err
    data = _body(request)
    estado = str(data.get("estado", "")).strip().lower()
    observacion = _texto(data.get("observacion"), 1000)
    if estado not in ESTADOS_REVISION:
        return _error("Estado inválido (aprobado u observado)")
    if estado == "observado" and not observacion:
        return _error("Escribe la observación para el proveedor")
    nueva_fecha = _fecha(data.get("fecha_vencimiento"), "fecha de vencimiento")
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT d.ruc, d.vigente, t.nombre, t.requiere_vencimiento, d.fecha_vencimiento
                 FROM compras_documentos d JOIN compras_tipos_documento t ON t.id = d.tipo_id
                WHERE d.id = %s FOR UPDATE OF d""",
            (documento_id,),
        )
        d = cur.fetchone()
        if not d:
            return _error("Documento no encontrado", 404)
        if not d[1]:
            return _error("Este documento fue reemplazado por una versión más reciente", 409)
        fecha_final = nueva_fecha or d[4]
        if estado == "aprobado" and d[3] and not fecha_final:
            return _error("Indica la fecha de vencimiento antes de aprobar")
        cur.execute(
            """UPDATE compras_documentos
                  SET estado = %s, observacion = %s, fecha_vencimiento = %s,
                      revisado_por = %s, revisado_en = now()
                WHERE id = %s""",
            (estado, observacion, fecha_final, s["id"], documento_id),
        )
        _audit(cur, f"admin:{s['usuario']}", f"documento_{estado}", d[0],
               {"documento_id": documento_id, "tipo": d[2], "observacion": observacion})
        conn.commit()
        cur.close()
        return _json({"ok": True})
    finally:
        conn.close()


def handle_admin_descargar(request, documento_id):
    s, err = _exigir_admin(request)
    if err:
        return err
    return _descargar(None, documento_id)


# ═════════════════════════════════════════════
# COMPRAS — CATÁLOGO DE DOCUMENTOS
# ═════════════════════════════════════════════
def handle_admin_tipos_listar(request):
    s, err = _exigir_admin(request)
    if err:
        return err
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT id, codigo, nombre, descripcion, obligatorio, requiere_vencimiento, orden, activo
                 FROM compras_tipos_documento ORDER BY activo DESC, orden, nombre"""
        )
        filas = [{"id": r[0], "codigo": r[1], "nombre": r[2], "descripcion": r[3], "obligatorio": r[4],
                  "requiere_vencimiento": r[5], "orden": r[6], "activo": r[7]} for r in cur.fetchall()]
        cur.close()
        return _json({"ok": True, "tipos": filas})
    finally:
        conn.close()


def handle_admin_tipos_guardar(request):
    s, err = _exigir_admin(request, rol="admin")
    if err:
        return err
    data = _body(request)
    tipo_id = data.get("id")
    nombre = _texto(data.get("nombre"), 150)
    descripcion = _texto(data.get("descripcion"), 1000)
    codigo = re.sub(r"[^A-Z0-9_]", "_", str(data.get("codigo") or nombre or "").strip().upper())[:40].strip("_")
    obligatorio = _bool(data.get("obligatorio", True))
    requiere_venc = _bool(data.get("requiere_vencimiento", False))
    activo = _bool(data.get("activo", True))
    try:
        orden = int(data.get("orden", 100))
    except (TypeError, ValueError):
        return _error("Orden inválido")
    if not nombre or not codigo:
        return _error("El nombre es obligatorio")
    conn = get_connection()
    try:
        cur = conn.cursor()
        try:
            if tipo_id:
                cur.execute(
                    """UPDATE compras_tipos_documento
                          SET codigo = %s, nombre = %s, descripcion = %s, obligatorio = %s,
                              requiere_vencimiento = %s, orden = %s, activo = %s
                        WHERE id = %s""",
                    (codigo, nombre, descripcion, obligatorio, requiere_venc, orden, activo, int(tipo_id)),
                )
                if cur.rowcount == 0:
                    return _error("Tipo de documento no encontrado", 404)
                accion = "editar_tipo_documento"
            else:
                cur.execute(
                    """INSERT INTO compras_tipos_documento
                           (codigo, nombre, descripcion, obligatorio, requiere_vencimiento, orden, activo)
                       VALUES (%s, %s, %s, %s, %s, %s, %s) RETURNING id""",
                    (codigo, nombre, descripcion, obligatorio, requiere_venc, orden, activo),
                )
                tipo_id = cur.fetchone()[0]
                accion = "crear_tipo_documento"
        except psycopg2.errors.UniqueViolation:
            conn.rollback()
            return _error(f"Ya existe un documento con el código {codigo}", 409)
        _audit(cur, f"admin:{s['usuario']}", accion, None, {"id": tipo_id, "codigo": codigo, "nombre": nombre})
        conn.commit()
        cur.close()
        return _json({"ok": True, "id": tipo_id})
    finally:
        conn.close()


# ═════════════════════════════════════════════
# PÁGINAS
# ═════════════════════════════════════════════
def _servir_html(nombre):
    ruta = os.path.join(STATIC_COMPRAS, nombre)
    if not os.path.exists(ruta):
        return ("Página no encontrada", 404, {"Content-Type": "text/plain; charset=utf-8"})
    with open(ruta, encoding="utf-8") as f:
        return (f.read(), 200, {
            "Content-Type": "text/html; charset=utf-8",
            "Cache-Control": "no-store",
            "X-Frame-Options": "DENY",
            "Referrer-Policy": "no-referrer",
        })


# ═════════════════════════════════════════════
# ROUTER
# ═════════════════════════════════════════════
_R_EMPRESA = re.compile(r"^/compras/api/admin/empresas/(\d{11})(?:/(editar|reset|estado))?$")
_R_DOC_ADMIN = re.compile(r"^/compras/api/admin/documentos/(\d+)/(revisar|descargar)$")
_R_DOC_PROV = re.compile(r"^/compras/api/proveedor/documentos/(\d+)/descargar$")
_R_USUARIO = re.compile(r"^/compras/api/admin/usuarios/(\d+)/(reset|estado)$")


def compras_router(request, path):
    """Devuelve la respuesta si la ruta es del módulo de compras, o None si no le corresponde."""
    if path == "/compras-bootstrap/admin":
        if request.method != "POST":
            return _error("Método no permitido", 405)
        return _ejecutar(handle_bootstrap_admin, request)
    if path != "/compras" and not path.startswith("/compras/"):
        return None

    m = request.method

    # Páginas
    if m == "GET" and path in ("/compras", "/compras/portal"):
        return _servir_html("portal.html")
    if m == "GET" and path == "/compras/admin":
        return _servir_html("admin.html")

    rutas = {
        ("POST", "/compras/api/proveedor/login"): handle_proveedor_login,
        ("POST", "/compras/api/proveedor/cambiar-password"): handle_proveedor_cambiar_password,
        ("POST", "/compras/api/proveedor/logout"): handle_logout,
        ("GET", "/compras/api/proveedor/yo"): handle_proveedor_yo,
        ("POST", "/compras/api/proveedor/datos"): handle_proveedor_datos,
        ("POST", "/compras/api/proveedor/documentos"): handle_proveedor_subir_documento,
        ("POST", "/compras/api/admin/login"): handle_admin_login,
        ("POST", "/compras/api/admin/cambiar-password"): handle_admin_cambiar_password,
        ("POST", "/compras/api/admin/logout"): handle_logout,
        ("GET", "/compras/api/admin/yo"): handle_admin_yo,
        ("GET", "/compras/api/admin/consulta-ruc"): handle_admin_consulta_ruc,
        ("GET", "/compras/api/admin/empresas"): handle_admin_empresas_listar,
        ("POST", "/compras/api/admin/empresas"): handle_admin_empresa_crear,
        ("GET", "/compras/api/admin/pendientes"): handle_admin_pendientes,
        ("GET", "/compras/api/admin/tipos-documento"): handle_admin_tipos_listar,
        ("POST", "/compras/api/admin/tipos-documento"): handle_admin_tipos_guardar,
        ("GET", "/compras/api/admin/usuarios"): handle_admin_usuarios_listar,
        ("POST", "/compras/api/admin/usuarios"): handle_admin_usuarios_crear,
    }
    handler = rutas.get((m, path))
    if handler:
        return _ejecutar(handler, request)

    r = _R_EMPRESA.match(path)
    if r:
        ruc, accion = r.group(1), r.group(2)
        if accion is None and m == "GET":
            return _ejecutar(handle_admin_empresa_detalle, request, ruc)
        if accion and m == "POST":
            fn = {"editar": handle_admin_empresa_editar, "reset": handle_admin_empresa_reset,
                  "estado": handle_admin_empresa_estado}[accion]
            return _ejecutar(fn, request, ruc)

    r = _R_DOC_ADMIN.match(path)
    if r:
        doc_id, accion = int(r.group(1)), r.group(2)
        if accion == "revisar" and m == "POST":
            return _ejecutar(handle_admin_revisar_documento, request, doc_id)
        if accion == "descargar" and m == "GET":
            return _ejecutar(handle_admin_descargar, request, doc_id)

    r = _R_DOC_PROV.match(path)
    if r and m == "GET":
        return _ejecutar(handle_proveedor_descargar, request, int(r.group(1)))

    r = _R_USUARIO.match(path)
    if r and m == "POST":
        uid, accion = int(r.group(1)), r.group(2)
        fn = handle_admin_usuario_reset if accion == "reset" else handle_admin_usuario_estado
        return _ejecutar(fn, request, uid)

    return _error(f"Ruta no encontrada: {path}", 404)


def _ejecutar(handler, request, *args):
    try:
        return handler(request, *args)
    except ErrorCompras as e:
        return _error(e.mensaje, e.status)
    except Exception as e:
        log.error("Error en %s: %s\n%s", handler.__name__, e, traceback.format_exc())
        return _error("Error interno. Intenta nuevamente.", 500)
