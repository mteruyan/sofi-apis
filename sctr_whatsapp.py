"""
sctr_whatsapp.py  –  Handlers WhatsApp para SCTR
Integrar en main.py igual que los demás módulos existentes.

Endpoints nuevos a registrar en main.py:
  POST /sctr/whatsapp/resumen   ← cliente pide resumen de su grupo
  POST /sctr/whatsapp/excel     ← asesora pide el Excel de un grupo
"""

import json
import os, io, re
from google.cloud.sql.connector import Connector
from google.cloud import storage
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, timedelta, timezone

HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
}


# ──────────────────────────────────────────────
# DB (mismo patrón que sctr_forms.py)
# ──────────────────────────────────────────────
def get_conn():
    connector = Connector()
    return connector.connect(
        os.environ["CLOUD_SQL_CONNECTION_NAME"],
        "pg8000",
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        db="sofi_whatsapp",
    )


# ──────────────────────────────────────────────
# Chattigo: enviar mensaje de texto
# ──────────────────────────────────────────────
def chattigo_send_text(telefono: str, mensaje: str):
    """Reutiliza tu función existente de Danaconnect/Chattigo."""
    import requests as req
    url     = os.environ.get("CHATTIGO_API_URL")   # ej: https://api.chattigo.com/...
    api_key = os.environ.get("CHATTIGO_API_KEY")
    channel = os.environ.get("CHATTIGO_CHANNEL_ID")

    payload = {
        "channel_id": channel,
        "to":         telefono,
        "type":       "text",
        "text":       mensaje,
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    resp = req.post(url, json=payload, headers=headers, timeout=10)
    resp.raise_for_status()
    return resp.json()


# ──────────────────────────────────────────────
# Chattigo: enviar archivo (documento)
# ──────────────────────────────────────────────
def chattigo_send_file(telefono: str, file_url: str, filename: str, caption: str = ""):
    """Envía un documento (Excel) por WhatsApp vía Chattigo."""
    import requests as req
    url     = os.environ.get("CHATTIGO_API_URL")
    api_key = os.environ.get("CHATTIGO_API_KEY")
    channel = os.environ.get("CHATTIGO_CHANNEL_ID")

    payload = {
        "channel_id": channel,
        "to":         telefono,
        "type":       "document",
        "document": {
            "url":      file_url,
            "filename": filename,
            "caption":  caption,
        }
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    resp = req.post(url, json=payload, headers=headers, timeout=15)
    resp.raise_for_status()
    return resp.json()


# ──────────────────────────────────────────────
# Subir Excel a Cloud Storage y obtener URL firmada
# ──────────────────────────────────────────────
def subir_excel_gcs(workbook, filename: str) -> str:
    bucket_name = os.environ.get("GCS_BUCKET", "sanitas-recetas-archivos")
    client  = storage.Client()
    bucket  = client.bucket(bucket_name)
    blob    = bucket.blob(f"sctr/empleados/{filename}")

    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)

    blob.upload_from_file(buf, content_type=(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ))
    blob.make_public()

    return f"https://storage.googleapis.com/{bucket_name}/sctr/empleados/{filename}"


# ──────────────────────────────────────────────
# Generar Excel (reutilizable)
# ──────────────────────────────────────────────
def generar_excel(empleados: list, empresa: str, ruc: str) -> openpyxl.Workbook:
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Afiliados"

    # Fila título
    num_cols = 13
    ws.merge_cells(f"A1:{get_column_letter(num_cols)}1")
    ws["A1"] = f"SCTR – {empresa}"
    ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="0066CC")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    header_row = 2
    if ruc:
        ws.merge_cells(f"A2:{get_column_letter(num_cols)}2")
        ws["A2"] = f"RUC: {ruc}"
        ws["A2"].font = Font(size=10, color="555555")
        ws["A2"].alignment = Alignment(horizontal="center")
        header_row = 3

    headers = [
        "Nombres*", "ApPaterno*", "ApMaterno", "TipoTrabajador*",
        "PaisNacimiento", "TipoIdent*", "NumIdent*", "Sexo*",
        "FecNacimiento*", "Moneda*", "Remuneracion*", "EstadoCivil*", "Sede*"
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1A1A2E")
        cell.alignment = Alignment(horizontal="center")

    for i, emp in enumerate(empleados):
        row_num = header_row + 1 + i
        fill = PatternFill("solid", fgColor="F0F4F8") if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")

        # Determinar TipoIdent según número de documento
        doc = str(emp["numero_documento"]).strip()
        if len(doc) == 8 and doc.isdigit():
            tipo_ident = "DNI"
        elif len(doc) == 9 and doc.isdigit():
            tipo_ident = "CE"
        elif len(doc) == 11 and doc.isdigit():
            tipo_ident = "RUC"
        else:
            tipo_ident = "DNI"

        # Formatear fecha dd/mm/yyyy
        fnac = emp["fecha_nacimiento"]
        if hasattr(fnac, "strftime"):
            fecha_fmt = fnac.strftime("%d/%m/%Y")
        else:
            # viene como string "YYYY-MM-DD"
            try:
                from datetime import datetime as dt
                fecha_fmt = dt.strptime(str(fnac), "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                fecha_fmt = str(fnac)

        valores = [
            emp["nombres"],
            emp["apellido_paterno"],
            emp["apellido_materno"],
            "Alto Riesgo",
            emp["nacionalidad"],
            tipo_ident,
            doc,
            emp["sexo"],
            fecha_fmt,
            "S/.",
            float(emp["sueldo"]) if emp["sueldo"] else None,
            "Soltero",
            "PRINCIPAL",
        ]

        for col, val in enumerate(valores, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.fill = fill
            if col == 11 and val is not None:  # Remuneracion
                cell.number_format = '0'

    col_widths = [18, 18, 18, 14, 16, 10, 14, 8, 14, 8, 14, 12, 12]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    return wb

# ──────────────────────────────────────────────
# Devuelve el tipo de documento
# ──────────────────────────────────────────────
def get_tipo_ident(doc):
        doc = str(doc).strip()
        if len(doc) == 8 and doc.isdigit():
            return "DNI"
        elif len(doc) == 9 and doc.isdigit():
            return "CE"
        elif len(doc) == 11 and doc.isdigit():
            return "RUC"
        else:
            return "DNI"

# ──────────────────────────────────────────────
# HANDLER 1: Cliente pide resumen de su grupo
# ──────────────────────────────────────────────
def handle_sctr_resumen_whatsapp(request):
    data     = request.get_json(force=True)
    telefono = (data.get("telefono") or "").strip()
    ruc      = (data.get("ruc") or "").strip()

    if not telefono:
        return (json.dumps({"ok": False, "error": "telefono requerido"}), 400, HEADERS)
    if not ruc:
        return (json.dumps({"ok": False, "error": "ruc requerido"}), 400, HEADERS)

    conn = get_conn()
    cur  = conn.cursor()

    # ── BUSCAR GRUPO VALIDANDO RUC Y TELÉFONO ─────────────────────────────
    cur.execute("""
        SELECT id, empresa, codigo
        FROM sctr_grupos
        WHERE ruc = %s AND comprador_telefono = %s
        ORDER BY id DESC
        LIMIT 1
    """, (ruc, telefono))
    grupo = cur.fetchone()

    if not grupo:
        cur.close(); conn.close()
        return (json.dumps({
            "ok":         False,
            "encontrado": False,
            "mensaje":    "No encontramos un grupo SCTR asociado a este RUC y teléfono."
        }), 200, HEADERS)

    grupo_id, empresa, codigo = grupo

    # ── EMPLEADOS ──────────────────────────────────────────────────────────
    cur.execute("""
        SELECT nombres, apellido_paterno, apellido_materno, numero_documento,
               nacionalidad, fecha_nacimiento
        FROM sctr_empleados
        WHERE grupo_id = %s
        ORDER BY apellido_paterno, apellido_materno, nombres
    """, (grupo_id,))
    empleados = cur.fetchall()
    cur.close(); conn.close()

    lista = [{
        "nombre_completo":   f"{ap} {am}, {nombres}",
        "numero_documento":  doc,
        "nacionalidad":      nacionalidad,
        "fecha_nacimiento":  fecha_nacimiento.strftime("%d/%m/%Y") if fecha_nacimiento else "",
        "sueldo":            "****"
    } for nombres, ap, am, doc, nacionalidad, fecha_nacimiento in empleados]

    lista_whatsapp = "\n".join([
        f"{i+1}. {e['nombre_completo']} - {get_tipo_ident(e['numero_documento'])}: {e['numero_documento']} - "
        f"{e['nacionalidad']} - Nac: {e['fecha_nacimiento']}"
        for i, e in enumerate(lista)
    ])

    return (json.dumps({
        "ok":         True,
        "encontrado": True,
        "empresa":    empresa,
        "codigo":     codigo,
        "total":      str(len(empleados)),
        "lista_whatsapp":  lista_whatsapp
    }, ensure_ascii=False), 200, HEADERS)

# ──────────────────────────────────────────────
# HANDLER 2: Asesora pide Excel de un grupo
# ──────────────────────────────────────────────
def handle_sctr_excel_whatsapp(request):
    # Verificar API key
    auth = request.headers.get("Authorization", "")
    api_key = auth.split(" ", 1)[1] if auth.startswith("ApiKey ") else (
        request.headers.get("X-API-Key") or request.args.get("api_key")
    )
    if api_key != os.environ.get("API_KEY"):
        return (json.dumps({"ok": False, "error": "No autorizado"}), 401, HEADERS)

    data     = request.get_json(force=True)
    codigo   = (data.get("codigo") or "").strip().upper()

    if not codigo:
        return (json.dumps({"ok": False, "error": "codigo es requerido"}), 400, HEADERS)

    conn = get_conn()
    cur  = conn.cursor()

    cur.execute("SELECT id, empresa, ruc FROM sctr_grupos WHERE codigo = %s", (codigo,))
    grupo = cur.fetchone()

    if not grupo:
        cur.close(); conn.close()
        return (json.dumps({
            "ok": False,
            "encontrado": False,
            "mensaje": f"No existe ningún grupo con el código {codigo}."
        }), 200, HEADERS)

    grupo_id, empresa, ruc = grupo

    cur.execute("""
        SELECT nombres, apellido_paterno, apellido_materno,
               numero_documento, fecha_nacimiento, nacionalidad, sexo, sueldo
        FROM sctr_empleados
        WHERE grupo_id = %s
        ORDER BY apellido_paterno, apellido_materno, nombres
    """, (grupo_id,))
    cols = ["nombres","apellido_paterno","apellido_materno",
            "numero_documento","fecha_nacimiento","nacionalidad","sexo","sueldo"]
    empleados = [dict(zip(cols, row)) for row in cur.fetchall()]
    cur.close(); conn.close()

    if not empleados:
        return (json.dumps({
            "ok": False,
            "encontrado": True,
            "mensaje": f"El grupo {empresa} aún no tiene empleados registrados."
        }), 200, HEADERS)

    # Generar Excel y subir a GCS
    #from datetime import datetime
    #wb       = generar_excel(empleados, empresa, ruc or "")
    #fecha    = datetime.now().strftime("%Y%m%d")
    #filename = f"SCTR_{empresa.replace(' ', '_')}_{fecha}_{datetime.now().strftime('%H%M%S')}.xlsx"
    #file_url = subir_excel_gcs(wb, filename)

    import pytz

    LIMA_TZ  = pytz.timezone("America/Lima")
    ahora    = datetime.now(LIMA_TZ).replace(tzinfo=None)
    wb       = generar_excel(empleados, empresa, ruc or "")
    fecha    = ahora.strftime("%Y%m%d")
    filename = f"SCTR_{empresa.replace(' ', '_')}_{fecha}_{ahora.strftime('%H%M%S')}.xlsx"
    file_url = subir_excel_gcs(wb, filename)

    return (json.dumps({
        "ok":        True,
        "encontrado": True,
        "empresa":   empresa,
        "codigo":    codigo,
        "total":     len(empleados),
        "filename":  filename,
        "url_excel": file_url
    }, ensure_ascii=False), 200, HEADERS)


# ──────────────────────────────────────────────
# HANDLER 3: Borra un registro (empleado) de un grupo
# ──────────────────────────────────────────────
def handle_sctr_borrar_empleado(request):
    data       = request.get_json(force=True)
    telefono   = (data.get("telefono") or "").strip()
    ruc        = (data.get("ruc") or "").strip()
    numero_str = (data.get("numero") or "").strip()

    if not telefono:
        return (json.dumps({"ok": False, "error": "telefono requerido"}), 400, HEADERS)
    if not ruc:
        return (json.dumps({"ok": False, "error": "ruc requerido"}), 400, HEADERS)
    if not numero_str:
        return (json.dumps({"ok": False, "error": "numero requerido"}), 400, HEADERS)

    try:
        numero = int(numero_str)
    except ValueError:
        return (json.dumps({"ok": False, "error": "numero debe ser un valor numérico"}), 400, HEADERS)

    conn = get_conn()
    cur  = conn.cursor()

    # ── BUSCAR GRUPO VALIDANDO RUC Y TELÉFONO ─────────────────────────────
    cur.execute("""
        SELECT id, empresa, codigo
        FROM sctr_grupos
        WHERE ruc = %s AND comprador_telefono = %s
        ORDER BY id DESC
        LIMIT 1
    """, (ruc, telefono))
    grupo = cur.fetchone()

    if not grupo:
        cur.close(); conn.close()
        return (json.dumps({
            "ok":         False,
            "encontrado": False,
            "mensaje":    "No encontramos un grupo SCTR asociado a este RUC y teléfono."
        }), 200, HEADERS)

    grupo_id, empresa, codigo = grupo

    # ── TRAER EMPLEADOS EN EL MISMO ORDEN QUE EL RESUMEN ───────────────────
    cur.execute("""
        SELECT id, nombres, apellido_paterno, apellido_materno
        FROM sctr_empleados
        WHERE grupo_id = %s
        ORDER BY apellido_paterno, apellido_materno, nombres
    """, (grupo_id,))
    empleados = cur.fetchall()

    if numero < 1 or numero > len(empleados):
        cur.close(); conn.close()
        return (json.dumps({
            "ok":      False,
            "mensaje": f"El número {numero} no es válido. La lista tiene {len(empleados)} registros."
        }), 200, HEADERS)

    empleado_id, nombres, ap, am = empleados[numero - 1]
    nombre_completo = f"{ap} {am}, {nombres}"

    cur.execute("DELETE FROM sctr_empleados WHERE id = %s", (empleado_id,))
    conn.commit()
    cur.close(); conn.close()

    return (json.dumps({
        "ok":               True,
        "empresa":          empresa,
        "codigo":           codigo,
        "eliminado":        nombre_completo,
        "numero_eliminado": numero,
        "total_restante":   len(empleados) - 1
    }, ensure_ascii=False), 200, HEADERS)