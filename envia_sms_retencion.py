# ─────────────────────────────────────────────
# API: /envia_sms_retencion
# Envía SMS de retención en lotes acotados (para no exceder el límite de
# ~200 mensajes por invocación que detectaste en la API). Se debe invocar
# repetidamente vía Cloud Scheduler (ej. cada 5-10 minutos, horario
# laboral): cada corrida procesa el siguiente lote pendiente y se detiene;
# como cada fila enviada se marca en la BD, la siguiente invocación retoma
# automáticamente donde quedó la anterior, sin duplicar envíos.
# ─────────────────────────────────────────────

LOTE_MAXIMO_MENSAJES = 180  # deja margen bajo el límite de ~200 detectado
LOTE_MAXIMO_FILAS = 200     # tope de filas a traer por consulta (cada fila puede generar hasta 3 mensajes)

MENSAJES_POR_TIPO = {
    "FAMILIA": "Hola {nombre}, aun estas a tiempo de continuar tu cobertura de salud con Sanitas. Contactanos:",
    "HIJOS":   "Hola {nombre}, aun estas a tiempo de continuar tu cobertura de salud con Sanitas. Contactanos:",
    "GENERAL": "Hola {nombre}, aun estas a tiempo de continuar tu cobertura de salud con Sanitas. Contactanos:",
}


def _enviar_sms_retencion(numero, mensaje):
    """Reutiliza el mismo canal (Danaconnect) que ya usas en
    envia_sms_leads_generico / enviar_danaconnect."""
    numero = f"+{numero}" if not numero.startswith("+") else numero
    payload = {
        "NUM":  f" {numero}",
        "TEXT": mensaje,
        "URL":  WHATSAPP_URL
    }
    headers = {
        "Content-Type":  "application/json",
        "Authorization": DANACONNECT_AUTH
    }
    response = requests.post(DANACONNECT_URL, json=payload, headers=headers, timeout=30)
    response.raise_for_status()


def handle_envia_sms_retencion(request):
    try:
        lima       = pytz.timezone("America/Lima")
        ahora      = datetime.now(lima)
        dia_semana = ahora.isoweekday()  # 1=Lunes ... 7=Domingo

        if dia_semana < 1 or dia_semana > 6:
            return (json.dumps({"mensaje": "Fuera del horario"}), 200, HEADERS)

        conn = get_connection()
        cur  = conn.cursor()

        # Por seguridad, asegura que exista la columna sms_enviado
        cur.execute("""
            ALTER TABLE retencion_sms
            ADD COLUMN IF NOT EXISTS sms_enviado TEXT DEFAULT 'N'
        """)
        conn.commit()

        cur.execute(f"""
            SELECT n_documento, primer_nombre, telefono1_ok, telefono2_ok, telefono3_ok, tipo_envio
            FROM retencion_sms
            WHERE (sms_enviado IS NULL OR sms_enviado = 'N')
              AND fecha_envio_sms::date <= CURRENT_DATE
            ORDER BY fecha_envio_sms
            LIMIT {LOTE_MAXIMO_FILAS}
        """)
        filas = cur.fetchall()

        mensajes_enviados = 0
        filas_procesadas  = 0
        errores           = 0

        for n_documento, primer_nombre, tel1, tel2, tel3, tipo_envio in filas:
            telefonos = [t.strip() for t in [tel1, tel2, tel3] if t and t.strip()]

            # Sin teléfonos válidos: marca como procesado para no reintentar por siempre
            if not telefonos:
                cur.execute("UPDATE retencion_sms SET sms_enviado = 'S' WHERE n_documento = %s", (n_documento,))
                conn.commit()
                filas_procesadas += 1
                continue

            # Si esta fila no entra completa en lo que queda del lote, se deja para la siguiente invocación
            if mensajes_enviados + len(telefonos) > LOTE_MAXIMO_MENSAJES:
                break

            plantilla = MENSAJES_POR_TIPO.get((tipo_envio or "").strip().upper(), MENSAJES_POR_TIPO["GENERAL"])
            mensaje   = plantilla.format(nombre=primer_nombre or "")

            fallo_fila = False
            for tel in telefonos:
                try:
                    _enviar_sms_retencion(tel, mensaje)
                    mensajes_enviados += 1
                    sms_aviso(f"+{tel}", "retencion")
                except Exception as e:
                    print(f"Error enviando a {tel} (doc {n_documento}): {e}")
                    errores += 1
                    fallo_fila = True

            # Se marca enviado igual, aunque haya fallado algún número puntual,
            # para no reintentar infinitamente el mismo registro problemático.
            cur.execute("UPDATE retencion_sms SET sms_enviado = 'S' WHERE n_documento = %s", (n_documento,))
            conn.commit()
            filas_procesadas += 1

        cur.close()
        conn.close()

        return (json.dumps({
            "mensaje":            "Lote procesado",
            "filas_procesadas":   filas_procesadas,
            "mensajes_enviados":  mensajes_enviados,
            "errores":            errores,
            "filas_restantes_aprox": len(filas) - filas_procesadas,
        }), 200, HEADERS)

    except Exception as e:
        return (json.dumps({"error": str(e)}), 500, HEADERS)
