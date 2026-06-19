from flask import Blueprint, request, jsonify
from google.cloud.sql.connector import Connector
import sqlalchemy
import os

sctr_bp = Blueprint('sctr', __name__)

def get_db_connection():
    connector = Connector()
    conn = connector.connect(
        os.environ.get("CLOUD_SQL_CONNECTION_NAME"),
        "pg8000",
        user=os.environ.get("DB_USER"),
        password=os.environ.get("DB_PASSWORD"),
        db="sofi_whatsapp",
    )
    return conn

# ─── GET /sctr/{codigo} → sirve el formulario HTML ───
def handle_sctr_form(request, codigo):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT empresa FROM sctr_grupos WHERE codigo = %s", (codigo,))
    row = cur.fetchone()
    cur.close()
    conn.close()

    if not row:
        return "Enlace no válido o expirado.", 404

    empresa = row[0]
    html = generar_html(codigo, empresa)
    return html, 200, {'Content-Type': 'text/html; charset=utf-8'}


# ─── POST /sctr/{codigo}/submit → guarda datos del empleado ───
def handle_sctr_submit(request, codigo):
    data = request.get_json(force=True)

    # Validar campos requeridos
    for campo in ['nombres', 'apellido_paterno', 'apellido_materno', 'numero_documento',
                  'fecha_nacimiento', 'sexo', 'sueldo']:
        if not data.get(campo):
            return jsonify({"ok": False, "error": f"Campo requerido: {campo}"}), 400

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT id FROM sctr_grupos WHERE codigo = %s", (codigo,))
    grupo = cur.fetchone()
    if not grupo:
        cur.close()
        conn.close()
        return jsonify({"ok": False, "error": "Código inválido"}), 404

    grupo_id = grupo[0]

    # Evitar documento duplicado en el mismo grupo
    cur.execute(
        "SELECT id FROM sctr_empleados WHERE grupo_id = %s AND numero_documento = %s",
        (grupo_id, data['numero_documento'])
    )
    if cur.fetchone():
        cur.close()
        conn.close()
        return jsonify({"ok": False, "error": "Este número de documento ya fue registrado en este grupo."}), 409

    cur.execute("""
        INSERT INTO sctr_empleados
            (grupo_id, nombres, apellido_paterno, apellido_materno,
             numero_documento, fecha_nacimiento, nacionalidad, sexo, sueldo)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        grupo_id,
        data['nombres'].strip().upper(),
        data['apellido_paterno'].strip().upper(),
        data['apellido_materno'].strip().upper(),
        data['numero_documento'].strip(),
        data['fecha_nacimiento'],
        data.get('nacionalidad', 'Peruana'),
        data['sexo'],
        data.get('sueldo') or None
    ))
    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"ok": True, "mensaje": "Datos registrados correctamente."})


def generar_html(codigo, empresa):
  BASE_URL = "https://sntspe.com"
  OG_IMAGE = "https://storage.googleapis.com/sanitas-recetas-archivos/static/registro_sctr.png"

  return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta property="og:type"         content="website">
    <meta property="og:url"          content="{BASE_URL}/sctr/{codigo}">
    <meta property="og:site_name"    content="Sanitas Perú">
    <meta property="og:title"        content="Registro SCTR – {empresa}">
    <meta property="og:description"  content="Completa los datos de tus trabajadores fácilmente. Puedes compartir el link para que el llenado sea más rápido.">
    <meta property="og:image"        content="{OG_IMAGE}">
    <meta property="og:image:width"  content="1200">
    <meta property="og:image:height" content="630">
    <meta name="twitter:card"        content="summary_large_image">
    <meta name="twitter:title"       content="Registro SCTR – {empresa}">
    <meta name="twitter:description" content="Completa los datos de tus trabajadores fácilmente. Puedes compartir el link para que el llenado sea más rápido.">
    <meta name="twitter:image"       content="{OG_IMAGE}">
<title>Registro SCTR – {empresa}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: #f0f4f8;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }}
  .card {{
    background: white;
    border-radius: 16px;
    padding: 32px 28px;
    max-width: 480px;
    width: 100%;
    box-shadow: 0 4px 24px rgba(0,0,0,0.08);
  }}
  .logo {{ text-align: center; margin-bottom: 8px; }}
  .logo img {{ height: 48px; }}
  h1 {{ font-size: 1.2rem; color: #1a1a2e; text-align: center; margin-bottom: 4px; }}
  .empresa {{ text-align: center; color: #0066cc; font-weight: 600; margin-bottom: 24px; font-size: 0.95rem; }}
  label {{ display: block; font-size: 0.85rem; color: #555; margin-bottom: 4px; margin-top: 16px; font-weight: 500; }}
  input, select {{
    width: 100%; padding: 12px 14px;
    border: 1.5px solid #e0e0e0; border-radius: 10px;
    font-size: 1rem; color: #1a1a2e;
    transition: border 0.2s;
    background: #fafafa;
  }}
  input:focus, select:focus {{
    outline: none; border-color: #0066cc; background: white;
  }}
  .required {{ color: #e53e3e; }}
  button {{
    width: 100%; margin-top: 28px;
    background: #0066cc; color: white;
    border: none; border-radius: 10px;
    padding: 14px; font-size: 1.05rem;
    font-weight: 600; cursor: pointer;
    transition: background 0.2s;
  }}
  button:hover {{ background: #0052a3; }}
  button:disabled {{ background: #aaa; cursor: not-allowed; }}
  .msg {{
    margin-top: 20px; padding: 14px;
    border-radius: 10px; text-align: center;
    font-weight: 500; display: none;
  }}
  .msg.ok {{ background: #e6f4ea; color: #2d6a4f; }}
  .msg.err {{ background: #fdecea; color: #c0392b; }}
  .footer {{ text-align: center; margin-top: 24px; font-size: 0.75rem; color: #aaa; }}
  .section-title {{
    font-size: 0.75rem; font-weight: 700; color: #0066cc;
    text-transform: uppercase; letter-spacing: 0.05em;
    margin-top: 24px; margin-bottom: 2px;
    border-bottom: 1px solid #e8f0fe; padding-bottom: 6px;
  }}
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <img src="https://www.sanitasperu.com/wp-content/uploads/2021/03/logo-sanitas.png"
         onerror="this.style.display='none'" alt="Sanitas">
  </div>
  <h1>Registro de datos SCTR</h1>
  <div class="empresa">{empresa}</div>
 
  <div class="section-title">Datos personales</div>
 
  <label>Nombres <span class="required">*</span></label>
  <input id="nombres" type="text" placeholder="Ej: Juan Carlos" autocomplete="given-name">
 
  <label>Apellido Paterno <span class="required">*</span></label>
  <input id="apellido_paterno" type="text" placeholder="Ej: Pérez" autocomplete="family-name">
 
  <label>Apellido Materno <span class="required">*</span></label>
  <input id="apellido_materno" type="text" placeholder="Ej: García">
 
  <label>País de nacimiento <span class="required">*</span></label>
  <select id="nacionalidad">
    <option value="Perú" selected>Perú</option>
    <option value="Afganistán">Afganistán</option>
    <option value="Albania">Albania</option>
    <option value="Alemania">Alemania</option>
    <option value="Andorra">Andorra</option>
    <option value="Angola">Angola</option>
    <option value="Anguila">Anguila</option>
    <option value="Antártida">Antártida</option>
    <option value="Antigua y Barbuda">Antigua y Barbuda</option>
    <option value="Antillas Holandesas">Antillas Holandesas</option>
    <option value="Arabia Saudita">Arabia Saudita</option>
    <option value="Argelia">Argelia</option>
    <option value="Argentina">Argentina</option>
    <option value="Armenia">Armenia</option>
    <option value="Aruba">Aruba</option>
    <option value="Australia">Australia</option>
    <option value="Austria">Austria</option>
    <option value="Azerbaiyán">Azerbaiyán</option>
    <option value="Bahamas">Bahamas</option>
    <option value="Bahrein">Bahrein</option>
    <option value="Bangladesh">Bangladesh</option>
    <option value="Barbados">Barbados</option>
    <option value="Belarús">Belarús</option>
    <option value="Bélgica">Bélgica</option>
    <option value="Belice">Belice</option>
    <option value="Benin">Benin</option>
    <option value="Bermuda">Bermuda</option>
    <option value="Bolivia">Bolivia</option>
    <option value="Bonaire, San Eustaquio y Saba">Bonaire, San Eustaquio y Saba</option>
    <option value="Bosnia y Herzegovina">Bosnia y Herzegovina</option>
    <option value="Botswana">Botswana</option>
    <option value="Brasil">Brasil</option>
    <option value="Brunei Darussalam">Brunei Darussalam</option>
    <option value="Bulgaria">Bulgaria</option>
    <option value="Burkina Faso">Burkina Faso</option>
    <option value="Burundi">Burundi</option>
    <option value="Bután">Bután</option>
    <option value="Cabo Verde">Cabo Verde</option>
    <option value="Camboya">Camboya</option>
    <option value="Camerún">Camerún</option>
    <option value="Canadá">Canadá</option>
    <option value="Chad">Chad</option>
    <option value="Chile">Chile</option>
    <option value="China">China</option>
    <option value="Chipre">Chipre</option>
    <option value="Ciudad del Vaticano">Ciudad del Vaticano</option>
    <option value="Colombia">Colombia</option>
    <option value="Congo (Brazzaville)">Congo (Brazzaville)</option>
    <option value="Congo RD">Congo RD</option>
    <option value="Corea del Norte">Corea del Norte</option>
    <option value="Corea del Sur">Corea del Sur</option>
    <option value="Costa de Marfil">Costa de Marfil</option>
    <option value="Costa Rica">Costa Rica</option>
    <option value="Croacia">Croacia</option>
    <option value="Cuba">Cuba</option>
    <option value="Curaçao">Curaçao</option>
    <option value="Dinamarca">Dinamarca</option>
    <option value="Dominica">Dominica</option>
    <option value="Ecuador">Ecuador</option>
    <option value="Egipto">Egipto</option>
    <option value="El Salvador">El Salvador</option>
    <option value="Emiratos Árabes Unidos">Emiratos Árabes Unidos</option>
    <option value="Eritrea">Eritrea</option>
    <option value="Eslovaquia">Eslovaquia</option>
    <option value="Eslovenia">Eslovenia</option>
    <option value="España">España</option>
    <option value="Estados Unidos">Estados Unidos</option>
    <option value="Estados Unidos (Islas Menores Alejadas)">Estados Unidos (Islas Menores Alejadas)</option>
    <option value="Estonia">Estonia</option>
    <option value="Etiopía">Etiopía</option>
    <option value="Filipinas">Filipinas</option>
    <option value="Finlandia">Finlandia</option>
    <option value="Fiyi">Fiyi</option>
    <option value="Francia">Francia</option>
    <option value="Gabón">Gabón</option>
    <option value="Gambia">Gambia</option>
    <option value="Georgia">Georgia</option>
    <option value="Georgia del Sur e Islas Sandwich">Georgia del Sur e Islas Sandwich</option>
    <option value="Ghana">Ghana</option>
    <option value="Gibraltar">Gibraltar</option>
    <option value="Granada">Granada</option>
    <option value="Grecia">Grecia</option>
    <option value="Groenlandia">Groenlandia</option>
    <option value="Guadalupe">Guadalupe</option>
    <option value="Guam">Guam</option>
    <option value="Guatemala">Guatemala</option>
    <option value="Guayana">Guayana</option>
    <option value="Guayana Francesa">Guayana Francesa</option>
    <option value="Guernesey">Guernesey</option>
    <option value="Guinea">Guinea</option>
    <option value="Guinea Ecuatorial">Guinea Ecuatorial</option>
    <option value="Guinea-Bissau">Guinea-Bissau</option>
    <option value="Haití">Haití</option>
    <option value="Honduras">Honduras</option>
    <option value="Hong Kong">Hong Kong</option>
    <option value="Hungría">Hungría</option>
    <option value="India">India</option>
    <option value="Indonesia">Indonesia</option>
    <option value="Irak">Irak</option>
    <option value="Irán">Irán</option>
    <option value="Irlanda">Irlanda</option>
    <option value="Isla Bouvet">Isla Bouvet</option>
    <option value="Isla de Man">Isla de Man</option>
    <option value="Isla de Navidad">Isla de Navidad</option>
    <option value="Isla Norfolk">Isla Norfolk</option>
    <option value="Islandia">Islandia</option>
    <option value="Islas Åland">Islas Åland</option>
    <option value="Islas Caimán">Islas Caimán</option>
    <option value="Islas Cocos (Keeling)">Islas Cocos (Keeling)</option>
    <option value="Islas Comoros">Islas Comoros</option>
    <option value="Islas Cook">Islas Cook</option>
    <option value="Islas Feroe">Islas Feroe</option>
    <option value="Islas Heard y McDonald">Islas Heard y McDonald</option>
    <option value="Islas Malvinas (Falkland)">Islas Malvinas (Falkland)</option>
    <option value="Islas Marianas del Norte">Islas Marianas del Norte</option>
    <option value="Islas Marshall">Islas Marshall</option>
    <option value="Islas Pitcairn">Islas Pitcairn</option>
    <option value="Islas Salomón">Islas Salomón</option>
    <option value="Islas Turcas y Caicos">Islas Turcas y Caicos</option>
    <option value="Islas Vírgenes Británicas">Islas Vírgenes Británicas</option>
    <option value="Islas Vírgenes de EE.UU.">Islas Vírgenes de EE.UU.</option>
    <option value="Israel">Israel</option>
    <option value="Italia">Italia</option>
    <option value="Jamaica">Jamaica</option>
    <option value="Japón">Japón</option>
    <option value="Jersey">Jersey</option>
    <option value="Jordania">Jordania</option>
    <option value="Kazajstán">Kazajstán</option>
    <option value="Kenia">Kenia</option>
    <option value="Kirguistán">Kirguistán</option>
    <option value="Kiribati">Kiribati</option>
    <option value="Kosovo">Kosovo</option>
    <option value="Kuwait">Kuwait</option>
    <option value="Laos">Laos</option>
    <option value="Lesoto">Lesoto</option>
    <option value="Letonia">Letonia</option>
    <option value="Líbano">Líbano</option>
    <option value="Liberia">Liberia</option>
    <option value="Libia">Libia</option>
    <option value="Liechtenstein">Liechtenstein</option>
    <option value="Lituania">Lituania</option>
    <option value="Luxemburgo">Luxemburgo</option>
    <option value="Macao">Macao</option>
    <option value="Macedonia">Macedonia</option>
    <option value="Madagascar">Madagascar</option>
    <option value="Malasia">Malasia</option>
    <option value="Malawi">Malawi</option>
    <option value="Maldivas">Maldivas</option>
    <option value="Malí">Malí</option>
    <option value="Malta">Malta</option>
    <option value="Marruecos">Marruecos</option>
    <option value="Martinica">Martinica</option>
    <option value="Mauricio">Mauricio</option>
    <option value="Mauritania">Mauritania</option>
    <option value="Mayotte">Mayotte</option>
    <option value="México">México</option>
    <option value="Micronesia">Micronesia</option>
    <option value="Moldova">Moldova</option>
    <option value="Mónaco">Mónaco</option>
    <option value="Mongolia">Mongolia</option>
    <option value="Montenegro">Montenegro</option>
    <option value="Montserrat">Montserrat</option>
    <option value="Mozambique">Mozambique</option>
    <option value="Myanmar">Myanmar</option>
    <option value="Namibia">Namibia</option>
    <option value="Nauru">Nauru</option>
    <option value="Nepal">Nepal</option>
    <option value="Nicaragua">Nicaragua</option>
    <option value="Níger">Níger</option>
    <option value="Nigeria">Nigeria</option>
    <option value="Niue">Niue</option>
    <option value="Noruega">Noruega</option>
    <option value="Nueva Caledonia">Nueva Caledonia</option>
    <option value="Nueva Zelanda">Nueva Zelanda</option>
    <option value="Omán">Omán</option>
    <option value="Países Bajos">Países Bajos</option>
    <option value="Pakistán">Pakistán</option>
    <option value="Palau">Palau</option>
    <option value="Palestina (Gaza)">Palestina (Gaza)</option>
    <option value="Panamá">Panamá</option>
    <option value="Papua Nueva Guinea">Papua Nueva Guinea</option>
    <option value="Paraguay">Paraguay</option>
    <option value="Polinesia Francés">Polinesia Francés</option>
    <option value="Polonia">Polonia</option>
    <option value="Portugal">Portugal</option>
    <option value="Puerto Rico">Puerto Rico</option>
    <option value="Qatar">Qatar</option>
    <option value="Reino Unido">Reino Unido</option>
    <option value="República Centroafricana">República Centroafricana</option>
    <option value="República Checa">República Checa</option>
    <option value="República Dominicana">República Dominicana</option>
    <option value="Reunión">Reunión</option>
    <option value="Ruanda">Ruanda</option>
    <option value="Rumania">Rumania</option>
    <option value="Rusia">Rusia</option>
    <option value="Sáhara Occidental">Sáhara Occidental</option>
    <option value="Saint Martin">Saint Martin</option>
    <option value="Saint-Barthélemy">Saint-Barthélemy</option>
    <option value="Samoa">Samoa</option>
    <option value="Samoa Americana">Samoa Americana</option>
    <option value="San Cristóbal y Nieves">San Cristóbal y Nieves</option>
    <option value="San Marino">San Marino</option>
    <option value="San Pedro y Miquelón">San Pedro y Miquelón</option>
    <option value="San Vicente y las Granadinas">San Vicente y las Granadinas</option>
    <option value="Santa Helena - Ascensión y Tristán da Cunha">Santa Helena - Ascensión y Tristán da Cunha</option>
    <option value="Santa Lucía">Santa Lucía</option>
    <option value="Santo Tomé y Príncipe">Santo Tomé y Príncipe</option>
    <option value="Senegal">Senegal</option>
    <option value="Serbia">Serbia</option>
    <option value="Seychelles">Seychelles</option>
    <option value="Sierra Leona">Sierra Leona</option>
    <option value="Singapur">Singapur</option>
    <option value="Siria">Siria</option>
    <option value="Somalia">Somalia</option>
    <option value="Sri Lanka">Sri Lanka</option>
    <option value="Suazilandia">Suazilandia</option>
    <option value="Sudáfrica">Sudáfrica</option>
    <option value="Sudán">Sudán</option>
    <option value="Sudan Sur">Sudan Sur</option>
    <option value="Suecia">Suecia</option>
    <option value="Suiza">Suiza</option>
    <option value="Suriname">Suriname</option>
    <option value="Svalbard y Jan Mayen">Svalbard y Jan Mayen</option>
    <option value="Tailandia">Tailandia</option>
    <option value="Taiwan">Taiwan</option>
    <option value="Tanzania">Tanzania</option>
    <option value="Tayikistán">Tayikistán</option>
    <option value="Territorio Brtiánico del Océano lndico">Territorio Brtiánico del Océano lndico</option>
    <option value="Territorios Franceses del Sur">Territorios Franceses del Sur</option>
    <option value="Timor Oriental">Timor Oriental</option>
    <option value="Togo">Togo</option>
    <option value="Tokelau">Tokelau</option>
    <option value="Tonga">Tonga</option>
    <option value="Trinidad y Tobago">Trinidad y Tobago</option>
    <option value="Túnez">Túnez</option>
    <option value="Turkmenistán">Turkmenistán</option>
    <option value="Turquía">Turquía</option>
    <option value="Tuvalu">Tuvalu</option>
    <option value="Ucrania">Ucrania</option>
    <option value="Uganda">Uganda</option>
    <option value="Uruguay">Uruguay</option>
    <option value="Uzbekistán">Uzbekistán</option>
    <option value="Vanuatu">Vanuatu</option>
    <option value="Venezuela">Venezuela</option>
    <option value="Viet Nam">Viet Nam</option>
    <option value="Wallis y Futuna">Wallis y Futuna</option>
    <option value="Yemen">Yemen</option>
    <option value="Yibuti">Yibuti</option>
    <option value="Zambia">Zambia</option>
    <option value="Zimbabue">Zimbabue</option>
  </select>
 
  <label>DNI / CE <span class="required">*</span></label>
  <input id="numero_documento" type="text" maxlength="12" placeholder="DNI o CE" inputmode="numeric">
 
  <label>Sexo <span class="required">*</span></label>
  <select id="sexo">
    <option value="">— Seleccionar —</option>
    <option value="M" selected>Masculino</option>
    <option value="F">Femenino</option>
  </select>
 
  <label>Fecha de nacimiento <span class="required">*</span></label>
  <input id="fecha" type="date">
 
  <div class="section-title">Datos laborales</div>
 
  <label>Sueldo mensual (S/.) <span class="required">*</span></label>
  <input id="sueldo" type="number" min="0" step="0.01" placeholder="Ej: 1500.00" inputmode="decimal">
 
  
  <button id="btn" onclick="enviar()">Registrar mis datos</button>
  <button id="btn-nuevo" onclick="nuevo()" style="display:none; width:100%; margin-top:12px;
    background:white; color:#0066cc; border:2px solid #0066cc; border-radius:10px;
    padding:14px; font-size:1.05rem; font-weight:600; cursor:pointer;">
    + Registrar otro trabajador
  </button>
  <div class="msg" id="msg"></div>
  <div class="footer">Sanitas Perú · Tus datos están protegidos</div>
</div>
 
<script>
async function enviar() {{
  const btn = document.getElementById('btn');

  const nombres          = document.getElementById('nombres').value.trim();
  const apellido_paterno = document.getElementById('apellido_paterno').value.trim();
  const apellido_materno = document.getElementById('apellido_materno').value.trim();
  const numero_documento = document.getElementById('numero_documento').value.trim();
  const sexo             = document.getElementById('sexo').value;
  const fecha            = document.getElementById('fecha').value;
  const sueldo           = document.getElementById('sueldo').value;

  if (!nombres || !apellido_paterno || !apellido_materno || !numero_documento || !sexo || !fecha || !sueldo) {{
    mostrar('Por favor completa todos los campos obligatorios.', 'err');
    return;
  }}
  if (!/^\\d{{8,12}}$/.test(numero_documento)) {{
    mostrar('El número de documento debe tener entre 8 y 12 dígitos.', 'err');
    return;
  }}

  btn.disabled = true;
  btn.textContent = 'Enviando...';

  try {{
    const res = await fetch('/sctr/{codigo}/submit', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{
        nombres:          nombres,
        apellido_paterno: apellido_paterno,
        apellido_materno: apellido_materno,
        numero_documento: numero_documento,
        sexo:             sexo,
        fecha_nacimiento: fecha,
        nacionalidad:     document.getElementById('nacionalidad').value,
        sueldo:           parseFloat(sueldo)
      }})
    }});
    const data = await res.json();
    if (data.ok) {{
      mostrar('✅ ¡Gracias! Tus datos fueron registrados correctamente.', 'ok');
      btn.style.display = 'none';
      document.querySelectorAll('input, select').forEach(el => el.disabled = true);
      document.getElementById('btn-nuevo').style.display = 'block';
    }} else {{
      mostrar(data.error || 'Ocurrió un error. Intenta nuevamente.', 'err');
      btn.disabled = false;
      btn.textContent = 'Registrar mis datos';
    }}
  }} catch(e) {{
    mostrar('Error de conexión. Intenta nuevamente.', 'err');
    btn.disabled = false;
    btn.textContent = 'Registrar mis datos';
  }}
}}

function nuevo() {{
  document.querySelectorAll('input').forEach(el => {{ el.disabled = false; el.value = ''; }});
  document.querySelectorAll('select').forEach(el => {{ el.disabled = false; el.selectedIndex = 0; }});
  document.getElementById('btn').style.display = 'block';
  document.getElementById('btn').disabled = false;
  document.getElementById('btn').textContent = 'Registrar mis datos';
  document.getElementById('btn-nuevo').style.display = 'none';
  document.getElementById('msg').style.display = 'none';
  document.getElementById('nombres').focus();
}}

function mostrar(texto, tipo) {{
  const msg = document.getElementById('msg');
  msg.textContent = texto;
  msg.className = 'msg ' + tipo;
  msg.style.display = 'block';
}}
</script>
</body>
</html>"""



# ─── POST /sctr/crear_grupo → asesora crea el link para una empresa ───
def handle_sctr_crear_grupo(request):
    auth = request.headers.get("Authorization", "")
    if auth.startswith("ApiKey "):
        api_key = auth.split(" ", 1)[1]
    else:
        api_key = request.headers.get("X-API-Key") or request.args.get("api_key")
    if api_key != os.environ.get("API_KEY"):
        return jsonify({"ok": False, "error": "No autorizado"}), 401

    data = request.get_json(force=True)
    if not data.get('empresa'):
        return jsonify({"ok": False, "error": "Campo requerido: empresa"}), 400

    conn = get_db_connection()
    cur = conn.cursor()

    # Verificar si ya existe un grupo con el mismo RUC
    ruc = data.get('ruc', '').strip()
    if ruc:
        cur.execute("SELECT codigo FROM sctr_grupos WHERE ruc = %s ORDER BY id DESC LIMIT 1", (ruc,))
        existente = cur.fetchone()
        if existente:
            cur.close()
            conn.close()
            codigo   = existente[0]
            base_url = os.environ.get("BASE_URL", "https://sntspe.com")
            return jsonify({
                "ok":     True,
                "existe": True,
                "codigo": codigo,
                "link":   f"{base_url}/sctr/{codigo}"
            })

    import random, string
    codigo = ''.join(random.choices(string.ascii_uppercase + string.digits, k=7))

    cur.execute("""
        INSERT INTO sctr_grupos (codigo, empresa, ruc, comprador_nombre, comprador_telefono, asesora_telefono)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (
        codigo,
        data['empresa'].strip().upper(),
        ruc,
        data.get('comprador_nombre', ''),
        data.get('comprador_telefono', ''),
        data.get('asesora_telefono', '')
    ))
    conn.commit()
    cur.close()
    conn.close()

    base_url = os.environ.get("BASE_URL", "https://sntspe.com")
    return jsonify({
        "ok":     True,
        "existe": False,
        "codigo": codigo,
        "link":   f"{base_url}/sctr/{codigo}"
    })