"""
agente_consulta.py

Agente virtual de preguntas frecuentes para clientes de Sanitas Perú.
Responde por WhatsApp usando Gemini (Vertex AI). La base de conocimiento
(RAG simple, sin vectores por ahora) vive en PostgreSQL en la tabla
sofi_agente_kb, para poder editarla sin necesidad de hacer redeploy.
El historial de conversación por celular vive en sofi_agente_historial,
para que el agente pueda mantener hilo entre preguntas seguidas (memoria
de corto plazo, ventana de HISTORIAL_VENTANA_MINUTOS). Cuando un hilo se
cierra (pasó la ventana sin mensajes nuevos), su contenido se resume y
se guarda en sofi_agente_perfil_cliente (memoria de largo plazo, para que
el agente "recuerde" al cliente aunque vuelva semanas después).

Además del RAG sobre conocimiento estático, el agente detecta ciertas
intenciones transaccionales (por ahora: "estado de mi carta de garantía")
y responde con datos REALES leídos en vivo de la tabla cartas_rechazadas
(resolviendo primero el DOI del cliente a partir de su celular), en vez
de dejar que el modelo intente responder con la base de conocimiento
genérica, que no tiene esa información y podría inventarla.

Integrar en main.py siguiendo el mismo patrón if/elif que usas para las
demás rutas (ver bloque al final de este archivo).

Tablas necesarias (crear una sola vez):

    CREATE TABLE sofi_agente_kb (
        id SERIAL PRIMARY KEY,
        titulo TEXT UNIQUE NOT NULL,
        contenido TEXT NOT NULL,
        activo BOOLEAN NOT NULL DEFAULT TRUE,
        fecha_actualizacion TIMESTAMP NOT NULL DEFAULT NOW()
    );

    CREATE TABLE sofi_agente_historial (
        id SERIAL PRIMARY KEY,
        celular TEXT NOT NULL,
        pregunta TEXT NOT NULL,
        respuesta TEXT NOT NULL,
        fecha TIMESTAMP NOT NULL DEFAULT NOW()
    );
    CREATE INDEX idx_agente_historial_celular_fecha
        ON sofi_agente_historial (celular, fecha);

    CREATE TABLE sofi_agente_perfil_cliente (
        celular TEXT PRIMARY KEY,
        resumen TEXT NOT NULL DEFAULT '',
        fecha_actualizacion TIMESTAMP NOT NULL DEFAULT NOW()
    );
"""

import os
import json
import random
import psycopg2
import vertexai
from vertexai.generative_models import GenerativeModel
from datetime import datetime
import pytz
import re

PROJECT_ID = os.environ.get("GCP_PROJECT", "sanitas-recetas-prod")
LOCATION = "us-central1"
MODEL_NAME = "gemini-2.5-flash-lite"
LIMA_TZ = pytz.timezone("America/Lima")

# Cuántos intercambios previos (pregunta+respuesta) se incluyen como hilo
HISTORIAL_LIMITE = 6
# Si el último mensaje del cliente fue hace más de esto, se considera una
# conversación nueva y no se arrastra el historial anterior
HISTORIAL_VENTANA_MINUTOS = 30

_model = None


def _clave_kb_valida(request) -> bool:
    """Valida el parámetro ?clave= contra la variable de entorno
    AGENTE_KB_CLAVE, para proteger el formulario de edición de la base
    de conocimiento (que no usa la ApiKey normal)."""
    clave_esperada = os.environ.get("AGENTE_KB_CLAVE", "")
    if not clave_esperada:
        # Si no se configuró la variable de entorno, bloquear todo por defecto
        return False
    clave_recibida = request.args.get("clave", "")
    return clave_recibida == clave_esperada


# ─────────────────────────────────────────────
# CONEXIÓN (mismo patrón que main.py)
# ─────────────────────────────────────────────
def get_connection():
    return psycopg2.connect(
        dbname="sofi_whatsapp",
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        host=f"/cloudsql/{os.environ['CLOUD_SQL_CONNECTION_NAME']}"
    )


def _get_model():
    global _model
    if _model is None:
        vertexai.init(project=PROJECT_ID, location=LOCATION)
        _model = GenerativeModel(MODEL_NAME)
    return _model

def _extraer_documento_de_texto(texto: str) -> str:
    """Busca un número de 8 o 9 dígitos contiguos dentro de un texto libre
    (ej. 'busca mi carta con el DNI 47164411'), útil cuando el cliente ya
    incluyó su documento en la misma pregunta. No valida contra
    RENIEC/Migraciones, solo formato."""
    coincidencias = re.findall(r'\d{8,9}', texto or "")
    for c in coincidencias:
        if len(c) in (8, 9):
            return c
    return ""

_PREGUNTA_DNI_CARTA = (
    "Para consultarte el estado de tu carta de garantía, indícame tu número "
    "de DNI (8 dígitos) o Carné de Extranjería (9 dígitos)."
)


def _es_documento_valido(texto: str) -> str:
    """Devuelve el documento limpio (solo dígitos) si el texto luce como un
    DNI peruano (8 dígitos) o un Carné de Extranjería (9 dígitos), o ''
    si no aplica. No valida contra RENIEC/Migraciones, solo formato."""
    limpio = "".join(c for c in (texto or "") if c.isdigit())
    if len(limpio) in (8, 9):
        return limpio
    return ""


def _es_consulta_estado_carta(pregunta: str) -> bool:
    """Distingue una consulta PERSONAL de estado ('¿cómo va mi carta?')
    de una pregunta GENÉRICA de FAQ ('¿cuánto dura una carta de garantía?'),
    que sí debe responder la base de conocimiento normal."""
    pl = pregunta.lower()
    if "carta" not in pl or "garant" not in pl:
        return False
    disparadores = [
        "estado", "como va", "cómo va", "que paso", "qué pasó",
        "avance", "seguimiento", "mi carta", "ya sali", "ya salió",
        "fue aprobada", "fue rechazada", "revisar mi", "busca", 
        "buscar", "consulta", "consultar", "dni", "documento"
    ]
    return any(d in pl for d in disparadores)


def _consultar_estado_carta(doi: str):
    """Misma consulta que usa handle_busca_carta_rechazada_tabla en
    main.py, sobre la tabla cartas_rechazadas (sincronizada por
    handle_sync_cartas_rechazadas)."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT tipo_carta, afiliado, ipress, plan, cobertura, fecha_prog,
               estado_intermedio, estado_final, observacion
        FROM cartas_rechazadas
        WHERE afiliado LIKE %s
        ORDER BY id_origen DESC
        LIMIT 1
    """, (f"%{doi}%",))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row

_AFIRMACIONES_CORTAS = [
    "si", "sí", "quiero", "dale", "va", "ok", "okay", "de acuerdo",
    "porfa", "por favor", "claro", "bueno", "ya", "sale", "correcto",
]

def _es_respuesta_afirmativa_corta(texto: str) -> bool:
    """Detecta respuestas cortas tipo 'sí', 'dale', 'quiero' con las que
    un cliente acepta una invitación previa del bot (ej. '¿quieres que
    revise el estado de tu carta?'). Solo aplica a mensajes cortos para
    no capturar frases más largas que casualmente empiecen igual."""
    texto_limpio = (texto or "").strip().lower()
    if len(texto_limpio) > 25:
        return False
    return any(
        texto_limpio == a or texto_limpio.startswith(a + " ") or texto_limpio.startswith(a + ",")
        for a in _AFIRMACIONES_CORTAS
    )

def _bot_invito_a_ver_estado_carta(historial) -> bool:
    """Revisa si la última respuesta del agente (generada por Gemini,
    texto variable) mencionó carta de garantía, como señal de que una
    respuesta afirmativa del cliente ahora se refiere a eso."""
    if not historial:
        return False
    ultima_respuesta = historial[-1][1] or ""
    pl = ultima_respuesta.lower()
    return "carta" in pl and "garant" in pl

_APERTURAS_ESTADO_CARTA = [
    "Aquí tienes el estado de tu carta de garantía:",
    "Con gusto, esto es lo que veo sobre tu carta de garantía:",
    "Claro, revisé tu carta de garantía y esto es lo que encontré:",
    "Esto es lo último registrado sobre tu carta de garantía:",
]


def _responder_estado_carta(doi: str) -> str:
    fila = _consultar_estado_carta(doi)

    if not fila:
        return (
            "No encontré una carta de garantía registrada con ese DNI en el "
            "sistema. Verifica que el número esté bien escrito; si acabas de "
            "solicitarla, puede que aún no se haya actualizado. "
        )

    tipo_carta, afiliado, ipress, plan, cobertura, fecha_prog, estado_intermedio, estado_final, observacion = fila

    if estado_final:
        estado = estado_final
    elif estado_intermedio:
        estado = estado_intermedio
    else:
        estado = "EN PROCESO"

    apertura = random.choice(_APERTURAS_ESTADO_CARTA)

    partes = [apertura, f"📋 Estado: *{estado}*"]
    if afiliado:
        partes.append(f"Afiliado: {afiliado}")
    if tipo_carta:
        partes.append(f"Tipo: {tipo_carta}")
    if ipress:
        partes.append(f"Clínica/IPRESS: {ipress}")
    if cobertura:
        partes.append(f"Cobertura: {cobertura}")
    if observacion and observacion.strip() and observacion.strip().lower() != "este campo por el momento esta vacío":
        partes.append(f"Observación: {observacion.strip()}")

    #partes.append("Si necesitas más detalle, un asesor puede ayudarte al 219-1919.")

    return "\n".join(partes)

# ─────────────────────────────────────────────
# CONSULTA EN VIVO: Red de Clínicas SCTR
# ─────────────────────────────────────────────
def _es_consulta_red_clinicas(pregunta: str, historial=None) -> bool:
    """Distingue una consulta sobre DÓNDE atenderse en la red SCTR de
    una pregunta similar pero relacionada a otro tema (ej. periodo de
    latencia, EPS), que debe seguir respondiendo la base de conocimiento
    normal. Los disparadores explícitos (mencionan 'sctr' o son
    inequívocos) siempre activan la consulta; los disparadores
    ambiguos solo activan si el contexto reciente confirma que se
    habla de SCTR."""
    pl = pregunta.lower()

    disparadores_explicitos = [
        "red de clinica", "red de clínica", "clinicas afiliadas",
        "clínicas afiliadas", "clinica afiliada", "clinicas sctr",
        "clínicas sctr", "red sctr", "red medica sctr", "red médica sctr",
    ]
    if any(d in pl for d in disparadores_explicitos):
        return True

    disparadores_ambiguos = [
        "donde me atiendo", "dónde me atiendo", "donde puedo atender",
        "dónde puedo atender", "clinicas cerca", "clínicas cerca",
        "red medica", "red médica", "hospitales afiliados",
        "centros medicos", "centros médicos", "donde queda",
        "dónde queda", "clinicas en", "clínicas en", "en que clinica",
        "en qué clínica",
    ]
    if not any(d in pl for d in disparadores_ambiguos):
        return False

    # Es un disparador ambiguo: solo lo aceptamos si el propio mensaje
    # menciona "sctr", o si el contexto reciente de la conversación
    # confirma que se viene hablando de SCTR (y no de otro tema como
    # latencia o EPS).
    if "sctr" in pl:
        return True

    if historial:
        texto_reciente = " ".join(
            f"{p} {r}" for p, r in historial[-2:]
        ).lower()
        if "sctr" in texto_reciente and "latencia" not in texto_reciente:
            return True

    return False


def _obtener_ubicaciones_conocidas():
    """Devuelve distritos, provincias y departamentos únicos registrados
    en la red de clínicas, para poder detectar si el cliente ya mencionó
    una ubicación específica en su misma pregunta."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT distrito FROM sctr_red_clinicas WHERE distrito IS NOT NULL
        UNION
        SELECT DISTINCT provincia FROM sctr_red_clinicas WHERE provincia IS NOT NULL
        UNION
        SELECT DISTINCT departamento FROM sctr_red_clinicas WHERE departamento IS NOT NULL
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [r[0] for r in rows if r[0]]


def _extraer_ubicacion_de_texto(pregunta: str) -> str:
    pl = _normalizar_texto(pregunta)
    for ubicacion in _obtener_ubicaciones_conocidas():
        patron = r'\b' + re.escape(_normalizar_texto(ubicacion)) + r'\b'
        if re.search(patron, pl):
            return ubicacion
    return ""


def _contar_red_clinicas(ubicacion: str = "") -> int:
    conn = get_connection()
    cur = conn.cursor()
    if ubicacion:
        cur.execute("""
            SELECT COUNT(*) FROM sctr_red_clinicas
            WHERE distrito ILIKE %s OR provincia ILIKE %s OR departamento ILIKE %s
        """, (f"%{ubicacion}%", f"%{ubicacion}%", f"%{ubicacion}%"))
    else:
        cur.execute("SELECT COUNT(*) FROM sctr_red_clinicas")
    total = cur.fetchone()[0]
    cur.close()
    conn.close()
    return total


def _consultar_red_clinicas(ubicacion: str = "", limite: int = 8):
    conn = get_connection()
    cur = conn.cursor()
    if ubicacion:
        cur.execute("""
            SELECT nombre_comercial, telefonos, direccion, distrito, provincia, departamento
            FROM sctr_red_clinicas
            WHERE unaccent(distrito) ILIKE unaccent(%s)
               OR unaccent(provincia) ILIKE unaccent(%s)
               OR unaccent(departamento) ILIKE unaccent(%s)
            ORDER BY nombre_comercial
            LIMIT %s
        """, (f"%{ubicacion}%", f"%{ubicacion}%", f"%{ubicacion}%", limite))
    else:
        cur.execute("""
            SELECT nombre_comercial, telefonos, direccion, distrito, provincia, departamento
            FROM sctr_red_clinicas ORDER BY nombre_comercial LIMIT %s
        """, (limite,))
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


_PREGUNTA_UBICACION_CLINICA = (
    "Claro, dime en qué distrito, provincia o departamento te gustaría "
    "atenderte y te muestro las clínicas de la red SCTR más cercanas."
)


def _formatear_lista_clinicas(filas) -> str:
    bloques = []
    for nombre, telefonos, direccion, distrito, provincia, departamento in filas:
        ubicacion = ", ".join(u for u in [distrito, provincia] if u)
        partes = [f"🏥 *{nombre}*"]
        if direccion:
            partes.append(direccion)
        if ubicacion:
            partes.append(ubicacion)
        if telefonos:
            partes.append(f"☎️ {telefonos}")
        bloques.append("\n".join(partes))
    return "\n\n".join(bloques)


_CIERRES_RED_CLINICAS = [
    "¿Hay algo más en lo que pueda ayudarte? 😊",
    "Si tienes otra consulta, con gusto te ayudaré.",
    "¿Qué otra duda puedo resolver?",    
]


def _responder_red_clinicas(ubicacion: str) -> str:
    total = _contar_red_clinicas(ubicacion)
    filas = _consultar_red_clinicas(ubicacion, limite=8)

    if not filas:
        return (
            f"No encontré clínicas de la red SCTR registradas para "
            f"\"{ubicacion}\". Prueba con el nombre de otro distrito, "
            f"provincia o departamento, o escríbenos a "
            f"sanitasenlinea@sanitasperu.com para más opciones."
        )

    encabezado = f"Estas son algunas clínicas de la red SCTR en {ubicacion}:"
    cuerpo = _formatear_lista_clinicas(filas)

    pie = ""
    if total > len(filas):
        pie = (f"\n\nHay {total - len(filas)} clínicas más en esa zona; "
               f"dime un distrito más específico si quieres afinar la búsqueda.")

    cierre = random.choice(_CIERRES_RED_CLINICAS)

    return f"{encabezado}\n\n{cuerpo}{pie}\n\n{cierre}"

def _coincide_con_otro_tema_conocido(pregunta: str) -> bool:
    """Revisa si el texto coincide con cualquier otro flujo/tema conocido
    del bot (compra, latencia, estado de carta, despedida, cambio de
    tema genérico), para no forzarlo como respuesta de ubicación cuando
    en realidad el cliente cambió de tema."""
    if _parece_cambio_de_tema(pregunta):
        return True
    if _parece_despedida(pregunta):
        return True
    if _es_consulta_estado_carta(pregunta):
        return True
    if _es_consulta_periodo_latencia(pregunta):
        return True
    return False

def _bot_invito_a_ver_red_clinicas(historial) -> bool:
    """Revisa si la última respuesta del agente (generada por Gemini,
    texto variable) invitó a ver la red de clínicas SCTR, como señal de
    que una respuesta afirmativa del cliente ahora se refiere a eso."""
    if not historial:
        return False
    ultima_respuesta = historial[-1][1] or ""
    pl = ultima_respuesta.lower()
    return "red" in pl and ("clinica" in pl or "clínica" in pl)

_DISPARADORES_CAMBIO_TEMA = [
    "otra pregunta", "otra consulta", "otra duda", "una pregunta",
    "cambio de tema", "olvida eso", "olvidemos eso",
    "que es", "qué es", "cuanto cuesta", "cuánto cuesta",
    "por que", "por qué", "como funciona", "cómo funciona",
    "comprar", "cotizar", "adquirir", "afiliar", "activar",
    "reembolso", "carta de garantia", "carta de garantía",
    "latencia",
]


def _parece_cambio_de_tema(texto: str) -> bool:
    """Heurística simple para detectar que el cliente probablemente
    cambió de tema en vez de responder con una ubicación (ej. 'tengo
    otra pregunta, que es la latencia'), evitando forzar una búsqueda
    de clínicas con ese texto."""
    pl = (texto or "").lower()
    return any(d in pl for d in _DISPARADORES_CAMBIO_TEMA)

_DISPARADORES_DESPEDIDA = [
    "gracias", "listo", "eso es todo", "nada mas", "nada más",
    "ok gracias", "perfecto gracias", "muchas gracias", "ya no",
    "eso seria todo", "eso sería todo", "esta bien", "está bien",
    "de acuerdo gracias", "vale gracias",
]


def _parece_despedida(texto: str) -> bool:
    """Detecta que el cliente se está despidiendo o cerrando la
    conversación, para no forzar una búsqueda de clínicas con ese texto."""
    pl = (texto or "").strip().lower()
    return any(d in pl for d in _DISPARADORES_DESPEDIDA)

def _obtener_estado_espera(celular: str) -> str:
    """Devuelve qué está esperando el bot de este cliente ahora mismo,
    o '' si no hay nada pendiente."""
    if not celular:
        return ""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT esperando FROM sofi_agente_estado WHERE celular = %s", (celular,))
    row = cur.fetchone()
    cur.close()
    conn.close()
    return row[0] if row and row[0] else ""


def _fijar_estado_espera(celular: str, esperando: str):
    """Guarda (o limpia, si esperando='') qué está esperando el bot."""
    if not celular:
        return
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sofi_agente_estado (celular, esperando, fecha_actualizacion)
        VALUES (%s, %s, %s)
        ON CONFLICT (celular) DO UPDATE SET
            esperando = EXCLUDED.esperando,
            fecha_actualizacion = EXCLUDED.fecha_actualizacion
    """, (celular, esperando, datetime.now(LIMA_TZ)))
    conn.commit()
    cur.close()
    conn.close()

import unicodedata

def _normalizar_texto(texto: str) -> str:
    """Quita tildes/diacríticos y pasa a minúsculas, para comparar sin
    importar acentos (ej. 'jesus maria' == 'Jesús María')."""
    texto = (texto or "").lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))

# ─────────────────────────────────────────────
# CÁLCULO DETERMINÍSTICO: Periodo de Latencia
# ─────────────────────────────────────────────
_TABLA_LATENCIA = [
    (5, 9, 2),
    (10, 14, 4),
    (15, 19, 6),
    (20, 24, 8),
    (25, 29, 10),
    (30, None, 12),  # 30 a más
]


def _calcular_latencia(meses_aportados: int):
    """Calcula el periodo de latencia (en meses) según los meses
    aportados en los últimos 3 años, con la tabla oficial de Sanitas.
    Devuelve un dict con 'total', 'sanitas_peas', 'essalud', o None si
    no corresponde latencia (menos de 5 meses aportados)."""
    if meses_aportados < 5:
        return None

    total = 12  # valor por defecto para 30+
    for minimo, maximo, latencia in _TABLA_LATENCIA:
        if maximo is None:
            if meses_aportados >= minimo:
                total = latencia
                break
        elif minimo <= meses_aportados <= maximo:
            total = latencia
            break

    # La división Sanitas (PEAS) / EsSalud es siempre proporcional,
    # mitad y mitad. Si el total es impar (no debería pasar con esta
    # tabla, todos son pares), redondeamos hacia abajo para Sanitas.
    sanitas_peas = total // 2
    essalud = total - sanitas_peas

    return {"total": total, "sanitas_peas": sanitas_peas, "essalud": essalud}


def _extraer_meses_de_texto(texto: str) -> int:
    """Busca un número de meses aportados en un texto libre (ej. '12
    meses', 'aporté 8', '20'). Devuelve el número como int, o -1 si no
    se pudo interpretar."""
    coincidencias = re.findall(r'\d{1,3}', texto or "")
    if coincidencias:
        return int(coincidencias[0])
    return -1


_PREGUNTA_MESES_APORTADOS = (
    "Para calcular tu periodo de latencia, cuéntame: ¿cuántos meses "
    "aportaste en los últimos tres años antes de que terminara tu "
    "vínculo laboral?"
)


def _responder_calculo_latencia(meses_aportados: int) -> str:
    resultado = _calcular_latencia(meses_aportados)

    if resultado is None:
        return (
            f"Con {meses_aportados} meses aportados, lamentablemente no "
            f"corresponde periodo de latencia (se requiere un mínimo de 5 "
            f"meses aportados en los últimos tres años). Si tienes dudas "
            f"sobre tu caso específico, escríbenos a "
            f"sanitasenlinea@sanitasperu.com o al 219-1919."
        )

    total = resultado["total"]
    sanitas = resultado["sanitas_peas"]
    essalud = resultado["essalud"]

    return (
        f"Con {meses_aportados} meses aportados, te corresponden "
        f"*{total} meses* de periodo de latencia.\n\n"
        f"De estos {total} meses:\n"
        f"• Los primeros {sanitas} meses, Sanitas cubre los diagnósticos PEAS.\n"
        f"• Los siguientes {essalud} meses, EsSalud cubre diagnósticos que "
        f"excedan el PEAS y también la cobertura de maternidad.\n\n"
        f"¿Hay algo más en lo que pueda ayudarte? 😊"
    )


def _es_consulta_periodo_latencia(pregunta: str, historial=None) -> bool:
    """Detecta que el cliente quiere saber cuántos meses de latencia le
    corresponden (no una pregunta genérica sobre qué es la latencia)."""
    pl = pregunta.lower()
    disparadores = [
        "cuantos meses me corresponde", "cuántos meses me corresponde",
        "cuanto me corresponde", "cuánto me corresponde",
        "cuanto dura mi latencia", "cuánto dura mi latencia",
        "cuanta latencia", "cuánta latencia", "meses de latencia me",
        "cuantos meses de latencia", "cuántos meses de latencia",
    ]
    if any(d in pl for d in disparadores):
        return True
    # Si el mensaje es solo un número (o "X meses") y el contexto
    # reciente venía hablando de latencia, también lo tomamos como
    # respuesta a la pregunta de meses aportados.
    if historial:
        texto_reciente = " ".join(f"{p} {r}" for p, r in historial[-2:]).lower()
        if "latencia" in texto_reciente and re.fullmatch(r'\s*\d{1,3}\s*(meses)?\s*', pl):
            return True
    return False

def _obtener_perfil(celular: str):
    """Devuelve (resumen, fecha_actualizacion) del cliente, o ('', None)
    si no existe todavía."""
    if not celular:
        return "", None

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT resumen, fecha_actualizacion
        FROM sofi_agente_perfil_cliente
        WHERE celular = %s
    """, (celular,))
    row = cur.fetchone()
    cur.close()
    conn.close()

    if not row:
        return "", None
    return row[0] or "", row[1]


def _guardar_perfil(celular: str, resumen: str):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sofi_agente_perfil_cliente (celular, resumen, fecha_actualizacion)
        VALUES (%s, %s, %s)
        ON CONFLICT (celular) DO UPDATE SET
            resumen = EXCLUDED.resumen,
            fecha_actualizacion = EXCLUDED.fecha_actualizacion
    """, (celular, resumen, datetime.now(LIMA_TZ)))
    conn.commit()
    cur.close()
    conn.close()


def _actualizar_perfil_si_hay_hilo_cerrado(celular: str) -> str:
    """Si desde la última actualización del perfil quedaron intercambios
    de un hilo que ya se cerró (más viejos que la ventana de hilo activo,
    y aún no incorporados al resumen), los resume con Gemini y actualiza
    sofi_agente_perfil_cliente. Devuelve el resumen vigente (actualizado
    o el que ya había) para usar en el prompt."""
    if not celular:
        return ""

    resumen_actual, fecha_actualizacion = _obtener_perfil(celular)

    conn = get_connection()
    cur = conn.cursor()
    if fecha_actualizacion:
        cur.execute("""
            SELECT pregunta, respuesta
            FROM sofi_agente_historial
            WHERE celular = %s
              AND fecha > %s
              AND fecha < NOW() - INTERVAL '%s minutes'
            ORDER BY fecha
        """, (celular, fecha_actualizacion, HISTORIAL_VENTANA_MINUTOS))
    else:
        cur.execute("""
            SELECT pregunta, respuesta
            FROM sofi_agente_historial
            WHERE celular = %s
              AND fecha < NOW() - INTERVAL '%s minutes'
            ORDER BY fecha
        """, (celular, HISTORIAL_VENTANA_MINUTOS))
    pendientes = cur.fetchall()
    cur.close()
    conn.close()

    if not pendientes:
        return resumen_actual

    lineas = [f"Cliente: {p}\nAgente: {r}" for p, r in pendientes]
    prompt_resumen = f"""Actualiza el resumen de este cliente de Sanitas Perú para
un agente virtual de WhatsApp. El resumen debe ser breve (máximo 5 líneas),
enfocado en: el nombre del cliente si lo mencionó en algún momento, temas
administrativos que ha consultado (ej. cartas de garantía, SCTR, reembolsos),
tipo de plan si lo mencionó, y dudas recurrentes. Si el cliente dijo su
nombre, inclúyelo explícitamente al inicio del resumen (ej. "Nombre: Miguel.").
NO incluyas diagnósticos médicos ni datos clínicos sensibles, solo temas de
cobertura/administrativos y el nombre si se conoce.

RESUMEN ACTUAL:
{resumen_actual or "(sin resumen previo, es la primera vez que se resume a este cliente)"}

NUEVOS INTERCAMBIOS A INCORPORAR:
{chr(10).join(lineas)}

RESUMEN ACTUALIZADO:"""

    model = _get_model()
    nuevo_resumen = model.generate_content(prompt_resumen).text.strip()

    _guardar_perfil(celular, nuevo_resumen)
    return nuevo_resumen


def _obtener_historial(celular: str):
    """Trae los últimos intercambios de este celular, solo si fueron
    dentro de la ventana de tiempo definida (si no, se considera hilo
    nuevo). Devuelve una lista de (pregunta, respuesta) en orden
    cronológico (más antiguo primero)."""
    if not celular:
        return []

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT pregunta, respuesta
        FROM sofi_agente_historial
        WHERE celular = %s
          AND fecha >= NOW() - INTERVAL '%s minutes'
        ORDER BY fecha DESC
        LIMIT %s
    """, (celular, HISTORIAL_VENTANA_MINUTOS, HISTORIAL_LIMITE))
    rows = cur.fetchall()
    cur.close()
    conn.close()

    return list(reversed(rows))


def _guardar_intercambio(celular: str, pregunta: str, respuesta: str):
    if not celular:
        return
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO sofi_agente_historial (celular, pregunta, respuesta, fecha)
        VALUES (%s, %s, %s, %s)
    """, (celular, pregunta, respuesta, datetime.now(LIMA_TZ)))
    conn.commit()
    cur.close()
    conn.close()


def _get_contexto() -> str:
    """Lee de sofi_agente_kb todos los registros activos y los concatena.
    Sin caché: cada request trae el contenido más reciente, así puedes
    editar la base de conocimiento sin redeploy."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT titulo, contenido
        FROM sofi_agente_kb
        WHERE activo = TRUE
        ORDER BY id
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    partes = [f"### {titulo}\n{contenido}" for titulo, contenido in rows]
    return "\n\n---\n\n".join(partes)


def _formatear_historial(historial) -> str:
    if not historial:
        return "(sin conversación reciente)"
    lineas = []
    for pregunta_previa, respuesta_previa in historial:
        lineas.append(f"Cliente: {pregunta_previa}")
        lineas.append(f"Agente: {respuesta_previa}")
    return "\n".join(lineas)


def _construir_prompt(contexto: str, perfil: str, historial, celular: str, pregunta: str) -> str:
    bloque_historial = _formatear_historial(historial)
    ultima_respuesta = historial[-1][1] if historial else "(ninguna)"

    return f"""
# ==========================================
# IDENTIDAD
# ==========================================

Eres **Sofi**, la asistente virtual oficial de Sanitas Perú para WhatsApp.

Tu misión es brindar una atención cálida, cercana, clara y profesional. Debes hacer sentir al cliente que conversa con una asesora de Sanitas con experiencia, no con un buscador de preguntas frecuentes.

Tu personalidad es:

• Amable y empática.
• Profesional.
• Paciente.
• Explica conceptos de forma sencilla.
• Conversacional, sin perder eficiencia.
• Precisa y confiable.

Nunca menciones que eres una inteligencia artificial.

Prioriza siempre este orden:

1. Exactitud de la información.
2. Comprensión de la consulta.
3. Claridad en la explicación.
4. Cercanía en la conversación.
5. Brevedad.

# ==========================================
# PRIORIDAD: CONTINUIDAD SOBRE COINCIDENCIA DE PALABRAS
# ==========================================

Antes de aplicar cualquier otra regla, revisa el HISTORIAL RECIENTE para
identificar de qué tema específico se viene hablando.

Si el cliente usa una palabra o pronombre ambiguo que podría relacionarse
con VARIOS temas del CONTEXTO (ej. "activarlo", "cómo lo hago", "eso
cómo funciona"), interpreta esa palabra SIEMPRE en relación al tema que
se venía conversando en el HISTORIAL RECIENTE, no al tema que
coincida más por palabras sueltas en el CONTEXTO.

Ejemplo: si el HISTORIAL RECIENTE trata sobre periodo de latencia, y el
cliente pregunta "¿qué tengo que hacer para activarlo?", se refiere a
activar su cobertura de latencia — NO a comprar o activar un SCTR nuevo,
aunque el CONTEXTO tenga un artículo sobre "activar SCTR" que coincida
con la palabra "activar".

Solo cambia de tema si el cliente lo indica explícitamente (ej. menciona
otro producto por nombre, o dice "cambiando de tema", "otra pregunta").

# ==========================================
# REGLAS DE NEGOCIO
# ==========================================

Responde ÚNICAMENTE utilizando la información disponible en el CONTEXTO.

Nunca inventes:

- Coberturas
- Beneficios
- Diagnósticos
- Enfermedades
- Exclusiones
- Montos
- Porcentajes
- Plazos
- Requisitos
- Procedimientos

Si la información no existe en el CONTEXTO, responde de forma transparente:

"No cuento con información suficiente para responder esa consulta."

Luego invita cordialmente al cliente a comunicarse con:

📞 Sanitas en Línea: 219-1919
✉️ sanitasenlinea@sanitasperu.com

Nunca completes información utilizando conocimientos propios.

Si el CONTEXTO recuperado contiene una instrucción específica para realizar un trámite (por ejemplo: comprar un SCTR desde WhatsApp), utiliza únicamente esa instrucción.

No sustituyas enlaces, teléfonos, correos o canales por otros que aparezcan en artículos diferentes.

No combines información de varios artículos si hablan de procesos distintos.

Si el usuario desea realizar una acción (comprar, cotizar, renovar, excluir, consultar, reembolsar, etc.), prioriza recuperar artículos cuyo título o intención coincida con esa acción.

Nunca inventes pasos para realizar un trámite.

Si el CONTEXTO indica que el trámite se realiza por WhatsApp, responde exactamente eso.

No cambies el canal indicado en el CONTEXTO.

No sugieras páginas web si el CONTEXTO no las menciona.

Si el CONTEXTO recuperado contiene una instrucción específica para
realizar un trámite del MISMO TEMA que se viene conversando (por ejemplo:
comprar un SCTR, cuando el tema es SCTR), utiliza únicamente esa
instrucción.

# ==========================================
# MEMORIA DE LA CONVERSACIÓN
# ==========================================

Dispones de información histórica del cliente.

Úsala para comprender mejor la conversación, NO para mencionarla.

La información se divide en:

## PERFIL

Describe de forma resumida al cliente y los temas que suele consultar.
Puede incluir su nombre si lo mencionó en el pasado.

Úsalo únicamente para comprender mejor el contexto y, si tienes su
nombre, para personalizar el saludo de forma natural.

Nunca digas frases como:

"Según consultas anteriores..."

"Como preguntaste anteriormente..."

"Veo que..."

No vuelvas a mencionar conversaciones pasadas explícitamente.

## HISTORIAL RECIENTE

Corresponde a las últimas consultas realizadas.

Debe servir para:

• entender preguntas de seguimiento.
• identificar referencias como:

    "¿y ese?"
    "¿también aplica?"
    "¿cuánto cuesta?"
    "¿y si es medicamentos?"
    "¿esa cobertura?"

• evitar repetir exactamente la misma explicación.
• mantener el hilo natural de la conversación.
• responder considerando lo ya conversado.

Si el cliente cambia completamente de tema, deja de utilizar el historial anterior.

# ==========================================
# PERSONALIZACIÓN
# ==========================================

Si conoces el nombre del cliente (por el PERFIL o el HISTORIAL), úsalo ocasionalmente.

Ejemplos:

"Claro, Miguel."

"Con gusto, Miguel."

No lo uses en todas las respuestas.

Nunca vuelvas a preguntar el nombre si ya está disponible.

Si no conoces su nombre, no lo inventes ni lo pidas explícitamente salvo que sea natural para la conversación.

# ==========================================
# ESTILO DE RESPUESTA
# ==========================================

Las respuestas deben sentirse naturales.

No respondas como una base de conocimientos.

Habla como una asesora de atención al cliente.

Puedes iniciar de distintas maneras:

• Claro.
• Con gusto.
• Buena pregunta.
• Te explico.
• Entiendo tu consulta.
• Sí, depende del caso.
• Permíteme explicarte.

Varía el inicio.

No repitas siempre la misma estructura.

Primero responde de forma sencilla.

Luego agrega los detalles importantes.

Si existen varios puntos utiliza listas.

Utiliza lenguaje sencillo.

Evita copiar textos legales completos.

Si el contexto contiene lenguaje legal, explícalo con palabras simples sin alterar su significado.

Cuando debas compartir un enlace:

- Escríbelo una sola vez.
- No utilices formato Markdown.
- No repitas el mismo enlace en la respuesta.

# ==========================================
# CONTINUIDAD
# ==========================================

Si el cliente hace varias preguntas relacionadas, responde manteniendo el contexto.

Ejemplo:

Cliente:
¿Qué son las preexistencias?

Luego:

¿Las EPS las cubren?

No vuelvas a explicar qué son las preexistencias.

Responde directamente sobre la cobertura.

# ==========================================
# CIERRE Y MENÚ
# ==========================================

Solo menciona la opción de escribir *menu* cuando el cliente se está
despidiendo, agradeciendo, o da señales claras de que terminó la
conversación (ej. "gracias", "listo", "eso era todo", "ok gracias").

En ese caso, agrega al final, con un salto de línea de por medio:

Escribe *menu* para volver al menú principal.

Si la conversación sigue abierta (el cliente probablemente seguirá
preguntando), NO menciones *menu* — cierra con una de las frases de
variedad ya definidas arriba en su lugar.

# ==========================================
# VARIEDAD
# ==========================================

Evita respuestas repetitivas.

Alterna:

- saludo inicial
- forma de explicar
- cierre
- invitación a continuar

No finalices siempre con:

"¿Algo más en que te pueda ayudar?"

Alterna con frases como:

- ¿Hay algo más en lo que pueda ayudarte? 😊
- Si tienes otra consulta, con gusto te ayudaré.
- Estoy aquí para ayudarte.
- ¿Qué otra duda puedo resolver?
- Si deseas, puedo explicarte ese tema con más detalle.

# ==========================================
# SEGUIMIENTO DEL TEMA
# ==========================================

Después de responder, si el tema tiene una acción transaccional relacionada
que el agente puede resolver en vivo (por ejemplo: estado de carta de
garantía, estado de un reembolso, etc.), invita al cliente a pedirla,
SIN solicitar tú mismo el DNI o CE — eso lo pedirá el sistema en el
siguiente paso si el cliente acepta.

Ejemplo:
Si el cliente pregunta qué es una carta de garantía, tras explicarle
puedes preguntar: "¿Quieres que revise el estado de alguna carta de
garantía tuya?" — NO le pidas el DNI en este mensaje.

Si el tema no tiene una acción transaccional asociada (por ejemplo,
preexistencias, coberturas generales, requisitos), en vez de eso ofrece
profundizar en un aspecto relacionado del mismo tema, sin inventar una
acción que no existe.

No fuerces esta pregunta de seguimiento si el cliente ya está
despidiéndose o agradeciendo — en ese caso ve directo al cierre.

# ==========================================
# FORMATO
# ==========================================

• Responde normalmente entre 4 y 6 líneas.
• Si el tema es complejo puedes utilizar hasta 8 líneas.
• Usa saltos de línea.
• No escribas bloques largos.
• WhatsApp debe ser fácil de leer.

# ==========================================
# INFORMACIÓN DEL CLIENTE
# ==========================================

## CELULAR
{celular}

## PERFIL
{perfil or "(sin perfil previo)"}

## HISTORIAL RECIENTE
{bloque_historial}

## ÚLTIMA RESPUESTA ENTREGADA
{ultima_respuesta}

# ==========================================
# CONTEXTO
# ==========================================

{contexto}

# ==========================================
# CONSULTA DEL CLIENTE
# ==========================================

{pregunta}

# ==========================================
# RESPUESTA
# ==========================================
"""


def _saludo_pidiendo_consulta(perfil: str, historial) -> str:
    """El cliente escribió pero no llegó ninguna pregunta (ej. solo saludó,
    o el mensaje llegó vacío/mal formado desde el canal). Genera una línea
    breve, variada, y personalizada con el nombre SOLO si aparece en el
    perfil o el historial."""
    model = _get_model()
    bloque_historial = _formatear_historial(historial)

    prompt = f"""Eres Sofi, la asistente virtual de Sanitas Perú por WhatsApp.
El cliente escribió pero no llegó ninguna pregunta o consulta clara junto
con su mensaje. Genera SOLO una línea breve, cálida y natural pidiéndole
que cuente en qué le puedes ayudar hoy.

Reglas:
- Varía la frase; no uses siempre la misma estructura ni las mismas palabras.
- Si en el PERFIL o el HISTORIAL aparece claramente el nombre del cliente,
  úsalo de forma natural (ej. "Hola Miguel, ¿en qué puedo ayudarte hoy?").
- Si NO aparece ningún nombre, no inventes uno ni lo pidas explícitamente:
  usa una frase genérica, cálida y variada.
- No agregues explicaciones, disculpas ni texto adicional. Responde solo
  con esa línea.

PERFIL:
{perfil or "(sin datos previos)"}

HISTORIAL RECIENTE:
{bloque_historial}

RESPUESTA:"""

    return model.generate_content(prompt).text.strip()

"""Función principal reutilizable. 'celular' es obligatorio (identifica
    el hilo e historial del cliente). Si 'pregunta' viene vacía, responde
    con una línea variada pidiendo la consulta, personalizada con el
    nombre si se conoce por el perfil/historial. Úsala también desde tu
    webhook de WhatsApp existente si quieres enrutar mensajes libres
    hacia acá."""
def responder_pregunta(pregunta: str, celular: str) -> str:
    perfil = _actualizar_perfil_si_hay_hilo_cerrado(celular)
    historial = _obtener_historial(celular)

    if not pregunta:
        return _saludo_pidiendo_consulta(perfil, historial)

    esperando = _obtener_estado_espera(celular)

    # Paso 1: el bot está esperando un DNI/CE para carta de garantía
    if esperando == "dni_carta":
        doc = _es_documento_valido(pregunta)
        if doc:
            _fijar_estado_espera(celular, "")
            respuesta = _responder_estado_carta(doc)
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta
        _fijar_estado_espera(celular, "")  # ya limpia sin necesidad de chequeo extra, porque solo actúa si YA es un documento válido

    # Paso 2: el bot está esperando una ubicación para red de clínicas
    elif esperando == "ubicacion_clinica":
        if _parece_despedida(pregunta):
            _fijar_estado_espera(celular, "")
            respuesta = (
                "¡De nada! Si necesitas algo más, aquí estoy.\n\n"
                "Escribe *menu* para volver al menú principal."
            )
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta

        if _coincide_con_otro_tema_conocido(pregunta):
            # El cliente cambió de tema; limpiamos el estado y dejamos
            # que el mensaje siga el flujo normal más abajo (sin return).
            _fijar_estado_espera(celular, "")
        else:
            ubicacion = _extraer_ubicacion_de_texto(pregunta)
            if not ubicacion and 0 < len(pregunta.strip()) <= 25:
                ubicacion = pregunta.strip()
            if ubicacion:
                respuesta = _responder_red_clinicas(ubicacion)
                _guardar_intercambio(celular, pregunta, respuesta)
                return respuesta
            # No se pudo interpretar como ubicación tampoco; limpiamos
            # el estado y seguimos el flujo normal.
            _fijar_estado_espera(celular, "")
    # Paso: el bot está esperando cuántos meses aportó (periodo de latencia)
    elif esperando == "meses_latencia":
        meses = _extraer_meses_de_texto(pregunta)
        if meses >= 0:
            _fijar_estado_espera(celular, "")
            respuesta = _responder_calculo_latencia(meses)
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta
        _fijar_estado_espera(celular, "")  # no se entendió, limpiamos y seguimos

    # Paso 3: consulta directa de estado de carta de garantía
    if _es_consulta_estado_carta(pregunta):
        doc_embebido = _extraer_documento_de_texto(pregunta)
        if doc_embebido:
            respuesta = _responder_estado_carta(doc_embebido)
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta
        _fijar_estado_espera(celular, "dni_carta")
        _guardar_intercambio(celular, pregunta, _PREGUNTA_DNI_CARTA)
        return _PREGUNTA_DNI_CARTA

    # Paso 4: consulta directa de red de clínicas
    if _es_consulta_red_clinicas(pregunta, historial):
        ubicacion_embebida = _extraer_ubicacion_de_texto(pregunta)
        if ubicacion_embebida:
            respuesta = _responder_red_clinicas(ubicacion_embebida)
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta
        _fijar_estado_espera(celular, "ubicacion_clinica")
        _guardar_intercambio(celular, pregunta, _PREGUNTA_UBICACION_CLINICA)
        return _PREGUNTA_UBICACION_CLINICA

    # Consulta directa: cuántos meses de latencia le corresponden
    if _es_consulta_periodo_latencia(pregunta, historial):
        meses_embebidos = _extraer_meses_de_texto(pregunta)
        # Si el mensaje ya trae un número Y no es solo una pregunta abierta
        # (ej. "12 meses" en vez de "cuantos meses me corresponden"), calculamos directo.
        if meses_embebidos >= 0 and re.search(r'\d', pregunta):
            respuesta = _responder_calculo_latencia(meses_embebidos)
            _guardar_intercambio(celular, pregunta, respuesta)
            return respuesta
        _fijar_estado_espera(celular, "meses_latencia")
        _guardar_intercambio(celular, pregunta, _PREGUNTA_MESES_APORTADOS)
        return _PREGUNTA_MESES_APORTADOS

    # Paso 5: invitaciones de Gemini (texto libre) aceptadas con un "sí" corto
    if _bot_invito_a_ver_estado_carta(historial) and _es_respuesta_afirmativa_corta(pregunta):
        _fijar_estado_espera(celular, "dni_carta")
        _guardar_intercambio(celular, pregunta, _PREGUNTA_DNI_CARTA)
        return _PREGUNTA_DNI_CARTA

    if _bot_invito_a_ver_red_clinicas(historial) and _es_respuesta_afirmativa_corta(pregunta):
        _fijar_estado_espera(celular, "ubicacion_clinica")
        _guardar_intercambio(celular, pregunta, _PREGUNTA_UBICACION_CLINICA)
        return _PREGUNTA_UBICACION_CLINICA

    # Paso 6: flujo genérico de RAG con Gemini
    contexto = _get_contexto()
    model = _get_model()
    prompt = _construir_prompt(contexto, perfil, historial, celular, pregunta)
    respuesta = model.generate_content(prompt).text.strip()

    _guardar_intercambio(celular, pregunta, respuesta)

    return respuesta

# ─────────────────────────────────────────────
# API: /consulta_agente
# ─────────────────────────────────────────────
def handle_consulta_agente(request):
    """Espera JSON: {"pregunta": "texto del cliente", "celular": "51999999999"}
    El campo 'celular' es OBLIGATORIO. El campo 'pregunta' es opcional:
    si no llega, el agente responde pidiendo la consulta de forma variada
    (y personalizada con el nombre, si ya lo conoce).
    Devuelve JSON: {"respuesta": "texto generado"}"""
    try:
        body = request.get_json(force=True, silent=True) or {}
        pregunta = (body.get("pregunta") or "").strip()
        celular = (body.get("celular") or "").strip()

        if not celular:
            return (json.dumps({"error": "Falta el campo 'celular'"}),
                    400, {"Content-Type": "application/json"})

        respuesta = responder_pregunta(pregunta, celular)
        return (json.dumps({"respuesta": respuesta}, ensure_ascii=False),
                200, {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"error": str(e)}),
                500, {"Content-Type": "application/json"})


# ─────────────────────────────────────────────
# API: /agente_kb/perfil
# Consulta el resumen de largo plazo y el hilo reciente de un celular,
# útil para depurar durante pruebas. Requiere ?clave= válida.
# ─────────────────────────────────────────────
def handle_agente_kb_perfil(request):
    if not _clave_kb_valida(request):
        return (json.dumps({"error": "Clave inválida"}),
                401, {"Content-Type": "application/json"})
    try:
        celular = (request.args.get("celular") or "").strip()
        if not celular:
            return (json.dumps({"error": "Falta el parámetro celular"}),
                    400, {"Content-Type": "application/json"})

        resumen, fecha_actualizacion = _obtener_perfil(celular)

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT pregunta, respuesta, fecha
            FROM sofi_agente_historial
            WHERE celular = %s
            ORDER BY fecha DESC
            LIMIT 20
        """, (celular,))
        rows = cur.fetchall()
        cur.close()
        conn.close()

        historial = [
            {"pregunta": p, "respuesta": r, "fecha": str(f)}
            for p, r, f in rows
        ]

        return (json.dumps({
            "celular": celular,
            "resumen": resumen,
            "resumen_actualizado": str(fecha_actualizacion) if fecha_actualizacion else None,
            "historial_reciente": historial,
        }, ensure_ascii=False), 200, {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"error": str(e)}),
                500, {"Content-Type": "application/json"})


# ─────────────────────────────────────────────
# API: /agente_kb/upsert
# Inserta o actualiza un artículo de la base de conocimiento (UPSERT por
# título, mismo patrón que handle_sync_retencion_sms).
# ─────────────────────────────────────────────
def handle_agente_kb_upsert(request):
    """Espera JSON: {"titulo": "...", "contenido": "...", "activo": true}
    Requiere ?clave= válida en la URL."""
    if not _clave_kb_valida(request):
        return (json.dumps({"error": "Clave inválida"}),
                401, {"Content-Type": "application/json"})
    try:
        body = request.get_json(force=True, silent=True) or {}
        titulo = (body.get("titulo") or "").strip()
        contenido = (body.get("contenido") or "").strip()
        activo = body.get("activo", True)

        if not titulo or not contenido:
            return (json.dumps({"error": "titulo y contenido son obligatorios"}),
                    400, {"Content-Type": "application/json"})

        fecha = datetime.now(LIMA_TZ)

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO sofi_agente_kb (titulo, contenido, activo, fecha_actualizacion)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (titulo) DO UPDATE SET
                contenido = EXCLUDED.contenido,
                activo = EXCLUDED.activo,
                fecha_actualizacion = EXCLUDED.fecha_actualizacion
        """, (titulo, contenido, activo, fecha))
        conn.commit()
        cur.close()
        conn.close()

        return (json.dumps({"mensaje": f"Artículo '{titulo}' guardado"}, ensure_ascii=False),
                200, {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"error": str(e)}),
                500, {"Content-Type": "application/json"})


# ─────────────────────────────────────────────
# API: /agente_kb/listar
# Lista los artículos de la base de conocimiento (para revisar/editar).
# ─────────────────────────────────────────────
def handle_agente_kb_listar(request):
    """Requiere ?clave= válida en la URL."""
    if not _clave_kb_valida(request):
        return (json.dumps({"error": "Clave inválida"}),
                401, {"Content-Type": "application/json"})
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, titulo, contenido, activo, fecha_actualizacion
            FROM sofi_agente_kb
            ORDER BY id
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()

        articulos = [
            {
                "id": r[0],
                "titulo": r[1],
                "contenido": r[2],
                "activo": r[3],
                "fecha_actualizacion": str(r[4]),
            }
            for r in rows
        ]

        return (json.dumps({"articulos": articulos}, ensure_ascii=False),
                200, {"Content-Type": "application/json"})

    except Exception as e:
        return (json.dumps({"error": str(e)}),
                500, {"Content-Type": "application/json"})


# ─────────────────────────────────────────────
# API: /agente_kb/form
# Formulario web simple para listar, crear y editar artículos de la
# base de conocimiento sin necesidad de Postman ni SQL directo.
# ─────────────────────────────────────────────
def handle_agente_kb_form_web(request):
    if not _clave_kb_valida(request):
        html_denegado = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<title>Acceso no autorizado</title></head>
<body style="font-family:Arial,sans-serif;text-align:center;padding:60px;">
<h2>Acceso no autorizado</h2>
<p>Falta el parámetro <code>?clave=</code> o la clave es incorrecta.</p>
</body></html>"""
        return (html_denegado, 401, {"Content-Type": "text/html; charset=utf-8"})

    clave = request.args.get("clave", "")
    html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Base de Conocimiento - Agente Sanitas</title>
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
  header h1{margin:0;font-size:20px;font-weight:600;}
  header p{margin:4px 0 0;font-size:13px;opacity:.85;}
  .container{
    max-width:900px;
    margin:0 auto;
    padding:16px;
    display:grid;
    grid-template-columns:1fr;
    gap:16px;
  }
  @media(min-width:820px){
    .container{grid-template-columns:340px 1fr;align-items:start;}
  }
  .card{
    background:#fff;
    border-radius:10px;
    padding:18px;
    box-shadow:0 1px 4px rgba(0,0,0,.08);
  }
  .card h2{
    font-size:15px;
    color:var(--azul);
    margin:0 0 12px;
  }
  .lista{
    list-style:none;
    margin:0;
    padding:0;
    max-height:520px;
    overflow-y:auto;
  }
  .lista li{
    padding:10px 12px;
    border:1px solid var(--borde);
    border-radius:8px;
    margin-bottom:8px;
    cursor:pointer;
    font-size:13px;
    transition:.15s;
  }
  .lista li:hover{background:#eef3fa;}
  .lista li.selected{
    border-color:var(--azul-claro);
    background:#eef3fa;
  }
  .lista li .titulo{
    font-weight:700;
    color:#222;
  }
  .lista li .meta{
    font-size:11px;
    color:#888;
    margin-top:3px;
  }
  .badge{
    display:inline-block;
    font-size:10px;
    font-weight:700;
    padding:2px 6px;
    border-radius:4px;
    margin-left:6px;
  }
  .badge.activo{background:#e7f6ec;color:var(--verde);}
  .badge.inactivo{background:#fbe9e7;color:var(--rojo);}
  label{
    display:block;
    font-size:13px;
    font-weight:600;
    margin:10px 0 4px;
    color:#333;
  }
  label.req::after{content:" *";color:var(--rojo);}
  input[type=text], textarea{
    width:100%;
    padding:10px 12px;
    border:1px solid var(--borde);
    border-radius:6px;
    font-size:14px;
    background:#fff;
    font-family:inherit;
  }
  textarea{
    resize:vertical;
    min-height:320px;
    line-height:1.4;
  }
  .checkbox-row{
    display:flex;
    align-items:center;
    gap:8px;
    margin-top:12px;
  }
  .checkbox-row input{width:auto;}
  .botones{
    display:flex;
    gap:10px;
    margin-top:16px;
  }
  button{
    flex:1;
    border:none;
    padding:12px;
    font-size:14px;
    font-weight:700;
    border-radius:8px;
    cursor:pointer;
  }
  button.guardar{background:var(--azul-claro);color:#fff;}
  button.nuevo{background:#eef1f5;color:var(--azul);}
  button:disabled{background:#9bb6d6;cursor:not-allowed;}
  .msg{
    margin-top:14px;
    padding:12px 14px;
    border-radius:8px;
    font-size:14px;
    display:none;
  }
  .msg.ok{display:block;background:#e7f6ec;color:var(--verde);border:1px solid #b9e3c4;}
  .msg.error{display:block;background:#fbe9e7;color:var(--rojo);border:1px solid #f0c2bb;}
  .hint{font-size:12px;color:#888;margin-top:6px;}
</style>
</head>
<body>

<header>
  <h1>Base de Conocimiento del Agente Virtual</h1>
  <p>Sanitas Perú · Administrar preguntas frecuentes / contenido del agente</p>
</header>

<div class="container">

  <div class="card">
    <h2>Artículos</h2>
    <ul class="lista" id="lista"></ul>
  </div>

  <div class="card">
    <h2 id="formTitulo">Nuevo artículo</h2>
    <form id="kbForm">
      <label class="req">Título</label>
      <input type="text" id="titulo" required placeholder="Ej: FAQ - Reembolsos">

      <label class="req">Contenido</label>
      <textarea id="contenido" required placeholder="Escribe aquí el contenido que usará el agente para responder (formato P: / R: o texto libre)"></textarea>
      <div class="hint">Este texto se manda tal cual como contexto a Gemini. Sé claro y evita ambigüedades.</div>

      <div class="checkbox-row">
        <input type="checkbox" id="activo" checked>
        <label style="margin:0;">Artículo activo (se usa en las respuestas del agente)</label>
      </div>

      <div class="botones">
        <button type="button" class="nuevo" id="btnNuevo">Limpiar / Nuevo</button>
        <button type="submit" class="guardar" id="btnGuardar">Guardar</button>
      </div>

      <div class="msg" id="msgBox"></div>
    </form>
  </div>
</div>

<script>
  const API_BASE = window.location.origin;
  const CLAVE = "__CLAVE__";

  let articulos = [];
  let tituloOriginalSeleccionado = null;

  async function cargarLista() {
    const lista = document.getElementById("lista");
    lista.innerHTML = "<li>Cargando...</li>";
    try {
      const resp = await fetch(API_BASE + "/agente_kb/listar?clave=" + encodeURIComponent(CLAVE));
      const data = await resp.json();
      articulos = data.articulos || [];
      renderLista();
    } catch (err) {
      lista.innerHTML = "<li>Error cargando artículos.</li>";
    }
  }

  function renderLista() {
    const lista = document.getElementById("lista");
    lista.innerHTML = "";

    if (articulos.length === 0) {
      lista.innerHTML = "<li>No hay artículos todavía.</li>";
      return;
    }

    articulos.forEach(art => {
      const li = document.createElement("li");
      li.dataset.titulo = art.titulo;
      if (art.titulo === tituloOriginalSeleccionado) li.classList.add("selected");

      const badgeClass = art.activo ? "activo" : "inactivo";
      const badgeTexto = art.activo ? "ACTIVO" : "INACTIVO";

      li.innerHTML = `
        <div class="titulo">${art.titulo}<span class="badge ${badgeClass}">${badgeTexto}</span></div>
        <div class="meta">Actualizado: ${art.fecha_actualizacion}</div>
      `;
      li.addEventListener("click", () => seleccionarArticulo(art));
      lista.appendChild(li);
    });
  }

  function seleccionarArticulo(art) {
    tituloOriginalSeleccionado = art.titulo;
    document.getElementById("formTitulo").textContent = "Editando artículo";
    document.getElementById("titulo").value = art.titulo;
    document.getElementById("contenido").value = art.contenido;
    document.getElementById("activo").checked = art.activo;
    renderLista();
    ocultarMsg();
  }

  function limpiarForm() {
    tituloOriginalSeleccionado = null;
    document.getElementById("formTitulo").textContent = "Nuevo artículo";
    document.getElementById("titulo").value = "";
    document.getElementById("contenido").value = "";
    document.getElementById("activo").checked = true;
    renderLista();
    ocultarMsg();
  }

  function ocultarMsg() {
    const msgBox = document.getElementById("msgBox");
    msgBox.style.display = "none";
  }

  function mostrarMsg(texto, tipo) {
    const msgBox = document.getElementById("msgBox");
    msgBox.textContent = texto;
    msgBox.className = "msg " + tipo;
    msgBox.style.display = "block";
  }

  document.getElementById("btnNuevo").addEventListener("click", limpiarForm);

  document.getElementById("kbForm").addEventListener("submit", async (e) => {
    e.preventDefault();

    const titulo = document.getElementById("titulo").value.trim();
    const contenido = document.getElementById("contenido").value.trim();
    const activo = document.getElementById("activo").checked;

    if (!titulo || !contenido) {
      mostrarMsg("Título y contenido son obligatorios.", "error");
      return;
    }

    const btn = document.getElementById("btnGuardar");
    btn.disabled = true;
    btn.textContent = "Guardando...";

    try {
      const resp = await fetch(API_BASE + "/agente_kb/upsert?clave=" + encodeURIComponent(CLAVE), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ titulo, contenido, activo })
      });
      const data = await resp.json();

      if (resp.ok) {
        mostrarMsg("✓ " + (data.mensaje || "Guardado correctamente."), "ok");
        tituloOriginalSeleccionado = titulo;
        await cargarLista();
      } else {
        mostrarMsg("Error: " + (data.error || "No se pudo guardar."), "error");
      }
    } catch (err) {
      mostrarMsg("Error de conexión. Intenta nuevamente.", "error");
    } finally {
      btn.disabled = false;
      btn.textContent = "Guardar";
    }
  });

  cargarLista();
</script>

</body>
</html>
"""
    html = html.replace("__CLAVE__", clave)
    return (html, 200, {"Content-Type": "text/html; charset=utf-8"})