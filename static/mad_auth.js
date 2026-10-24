(function () {
  const KEY = 'mad_token';
  const almacen = {
    get() { try { return localStorage.getItem(KEY); } catch (e) { return null; } },
    set(t) { try { localStorage.setItem(KEY, t); } catch (e) {} },
    del() { try { localStorage.removeItem(KEY); } catch (e) {} },
  };

  async function llamar(path, body, tokenParcial) {
    const headers = { 'Content-Type': 'application/json' };
    const t = tokenParcial || almacen.get();
    if (t) headers['Authorization'] = 'Bearer ' + t;
    const res = await fetch(path, {
      method: body === undefined ? 'GET' : 'POST',
      headers,
      body: body === undefined ? undefined : JSON.stringify(body),
      cache: 'no-store',
    });
    let data = {};
    try { data = await res.json(); } catch (e) {}
    return { ok: res.ok, status: res.status, data };
  }

  // fetch con el token; si la sesión venció, vuelve al login
  async function authFetch(url, opts = {}) {
    const headers = Object.assign({}, opts.headers || {});
    const t = almacen.get();
    if (t) headers['Authorization'] = 'Bearer ' + t;
    const res = await fetch(url, Object.assign({}, opts, { headers }));
    if (res.status === 401) {
      almacen.del();
      if (MadAuth.paginaLogin) location.href = MadAuth.paginaLogin + '?volver=' + encodeURIComponent(location.pathname);
      else location.reload();
      throw new Error('Sesión vencida');
    }
    return res;
  }

  const CSS = `
  #madAuth{position:fixed;inset:0;z-index:5000;background:#f4f6f9;display:flex;align-items:center;justify-content:center;padding:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif}
  #madAuth .card{background:#fff;border-radius:14px;padding:26px 22px;width:100%;max-width:380px;box-shadow:0 2px 14px rgba(0,56,118,.12);text-align:center;color:#1f2937}
  #madAuth .logo{width:50px;height:50px;border-radius:14px;margin:0 auto 12px;background:linear-gradient(135deg,#003876,#007A87);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:20px}
  #madAuth h1{font-size:18px;margin:0 0 4px;color:#003876}
  #madAuth p{font-size:13px;color:#6b7a82;margin:0 0 16px;line-height:1.45}
  #madAuth input{width:100%;padding:12px 14px;border:1px solid #d7dde5;border-radius:8px;font-size:16px;margin-bottom:10px;box-sizing:border-box}
  #madAuth input.codigo{text-align:center;letter-spacing:6px;font-size:22px}
  #madAuth button{width:100%;padding:13px;background:#0057b8;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer}
  #madAuth button:disabled{opacity:.6}
  #madAuth .error{color:#c0392b;font-size:12.5px;min-height:18px;margin-top:10px}
  #madAuth .qr{display:flex;justify-content:center;margin:6px 0 10px}
  #madAuth .secreto{font-family:monospace;font-size:12px;background:#f4f6f9;padding:8px;border-radius:6px;word-break:break-all;margin-bottom:12px}
  #madAuth .pasos{font-size:12.5px;color:#6b7a82;text-align:left;margin:0 0 12px;padding-left:18px}`;

  function crearUI(titulo) {
    const style = document.createElement('style'); style.textContent = CSS;
    const root = document.createElement('div'); root.id = 'madAuth';
    root.innerHTML = '<div class="card"><div class="logo">S</div><h1></h1><div class="cuerpo"></div></div>';
    root.querySelector('h1').textContent = titulo;
    document.head.appendChild(style); document.body.appendChild(root);
    return {
      render(html) { root.querySelector('.cuerpo').innerHTML = html; },
      q(sel) { return root.querySelector(sel); },
      quitar() { root.remove(); style.remove(); },
    };
  }

  function iniciar({ rol, titulo }) {
    return new Promise(async (resolve) => {
      // ¿Ya hay una sesión válida para este rol?
      if (almacen.get()) {
        const r = await llamar('/mad/auth/yo');
        if (r.ok && r.data.rol === rol) return resolve(r.data);
        almacen.del();
      }

      const ui = crearUI(titulo || 'Acceso');
      let parcial = null;

      async function enviar(boton, accion) {
        const err = ui.q('.error'); err.textContent = '';
        boton.disabled = true;
        try { await accion(err); }
        catch (e) { const el = ui.q('.error'); if (el) el.textContent = 'Error de conexión, intenta nuevamente'; }
        finally { boton.disabled = false; }
      }

      function expirada(r) {
        if (r.status === 401) { pantallaLogin('Tu sesión expiró, vuelve a ingresar'); return true; }
        return false;
      }

      function siguiente(paso) {
        if (paso === 'cambiar_password') pantallaPassword();
        else if (paso === 'enrolar_totp') pantallaEnrolar();
        else pantallaCodigo();
      }

      function pantallaLogin(mensaje) {
        ui.render(`<p>Ingresa con tu usuario y clave</p>
          <form><input id="u" autocomplete="username" autocapitalize="none" placeholder="Usuario">
          <input id="p" type="password" autocomplete="current-password" placeholder="Clave">
          <button type="submit">Ingresar</button></form><div class="error"></div>`);
        if (mensaje) ui.q('.error').textContent = mensaje;
        ui.q('form').onsubmit = (e) => { e.preventDefault(); enviar(ui.q('button'), async (err) => {
          const r = await llamar('/mad/auth/login', { usuario: ui.q('#u').value, password: ui.q('#p').value });
          if (!r.ok) { err.textContent = r.data.error || 'No se pudo ingresar'; return; }
          parcial = r.data.token_parcial;
          siguiente(r.data.siguiente);
        }); };
      }

      function pantallaPassword() {
        ui.render(`<p>Por seguridad, crea una clave nueva: mínimo 10 caracteres, con letras y números.</p>
          <form><input id="n1" type="password" autocomplete="new-password" placeholder="Nueva clave">
          <input id="n2" type="password" autocomplete="new-password" placeholder="Repite la nueva clave">
          <button type="submit">Guardar clave</button></form><div class="error"></div>`);
        ui.q('form').onsubmit = (e) => { e.preventDefault(); enviar(ui.q('button'), async (err) => {
          const n1 = ui.q('#n1').value, n2 = ui.q('#n2').value;
          if (n1 !== n2) { err.textContent = 'Las claves no coinciden'; return; }
          const r = await llamar('/mad/auth/cambiar-password', { nueva_password: n1 }, parcial);
          if (expirada(r)) return;
          if (!r.ok) { err.textContent = r.data.error || 'No se pudo guardar'; return; }
          siguiente(r.data.siguiente);
        }); };
      }

      async function pantallaEnrolar() {
        const r = await llamar('/mad/auth/totp/iniciar', {}, parcial);
        if (expirada(r)) return;
        if (!r.ok) { pantallaLogin(r.data.error || 'No se pudo configurar Google Authenticator'); return; }
        ui.render(`<p>Configura <b>Google Authenticator</b> (obligatorio):</p>
          <ol class="pasos"><li>Instala Google Authenticator en tu celular.</li>
          <li>Toca <b>+</b> y escanea este código QR.</li>
          <li>Escribe el código de 6 dígitos que aparece.</li></ol>
          <div class="qr" id="qr"></div>
          <div class="secreto">¿No puedes escanear? Ingresa esta clave: <span id="sec"></span></div>
          <form><input id="c" class="codigo" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="000000">
          <button type="submit">Activar y entrar</button></form><div class="error"></div>`);
        ui.q('#sec').textContent = r.data.secreto;
        if (window.QRCode) new QRCode(ui.q('#qr'), { text: r.data.otpauth_uri, width: 180, height: 180 });
        bindCodigo();
      }

      function pantallaCodigo() {
        ui.render(`<p>Ingresa el código de 6 dígitos de <b>Google Authenticator</b></p>
          <form><input id="c" class="codigo" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="000000">
          <button type="submit">Verificar</button></form><div class="error"></div>`);
        ui.q('#c').focus();
        bindCodigo();
      }

      function bindCodigo() {
        ui.q('form').onsubmit = (e) => { e.preventDefault(); enviar(ui.q('button'), async (err) => {
          const r = await llamar('/mad/auth/totp/verificar', { codigo: ui.q('#c').value }, parcial);
          if (expirada(r)) return;
          if (!r.ok) { err.textContent = r.data.error || 'Código incorrecto'; ui.q('#c').value = ''; return; }
          if (r.data.usuario.rol !== rol) {
            await llamar('/mad/auth/logout', {}, r.data.token);
            pantallaLogin('Este usuario no tiene acceso a esta pantalla');
            return;
          }
          almacen.set(r.data.token);
          ui.quitar();
          resolve(r.data.usuario);
        }); };
      }

      pantallaLogin();
    });
  }

  async function cerrarSesion() {
    try { await llamar('/mad/auth/logout', {}); } catch (e) {}
    almacen.del();
    location.reload();
  }

  const MadAuth = { iniciar, fetch: authFetch, cerrarSesion, paginaLogin: null };
  window.MadAuth = MadAuth;
})();

window.fmtProgramada = function (iso) {
  const TZ = 'America/Lima', d = new Date(iso);
  const dia = x => x.toLocaleDateString('en-CA', { timeZone: TZ });
  const hora = d.toLocaleTimeString('es-PE', { timeZone: TZ, hour: '2-digit', minute: '2-digit' });
  if (dia(d) === dia(new Date())) return `Hoy ${hora}`;
  if (dia(d) === dia(new Date(Date.now() + 86400000))) return `Mañana ${hora}`;
  return d.toLocaleDateString('es-PE', { timeZone: TZ, weekday: 'short', day: '2-digit', month: '2-digit' }) + ` ${hora}`;
};
