/**
 * Invitación "Cumple Nico": confirmaciones y pagos.
 *
 * Va entero en Código.gs, reemplazando todo lo que había.
 * La primera vez, elegí la función "configurar" arriba y tocá Ejecutar.
 *
 * Lo que cambia seguido (fecha, lugar, precio, alias, titular) se edita
 * en la pestaña "Invitación" de la planilla, no acá.
 */

const PLANILLA_ID = '1IXz1OIY2vBBwrMLt_LnP3XVAYnqfZOMvmAvJqPYM3p4';
const PESTANA_DATOS = 'Invitación';

const DATOS_INICIALES = [
  ['Fecha y hora', 'Viernes 4 de diciembre · 23 hs'],
  ['Lugar', 'A confirmar'],
  ['Precio por persona', ''],
  ['Alias', ''],
  ['Titular de la cuenta', ''],
];

const ENCABEZADOS = [
  'Fecha', 'ID', 'Nombre', 'Respuesta', 'Lleva', 'Clave (no tocar)',
  'Pagó', 'Monto', 'Fecha pago', 'Clave pago (no tocar)',
];
const ETIQUETAS = { si: 'Va', quizas: 'Quizás', no: 'No va' };

// ---------- página ----------

function doGet() {
  const datos = JSON.stringify(leerDatos_()).replace(/</g, '\\u003c');
  const html = PAGINA.replace('__DATOS__', function () { return datos; });
  return HtmlService.createHtmlOutput(html)
    .setTitle('Cumple Nico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- funciones que llama la invitación ----------

function responder(nombre, respuesta) {
  nombre = limpiarNombre_(nombre);
  if (!ETIQUETAS[respuesta]) throw new Error('Elegí si venís o no.');
  const id = idPara_(nombre);
  const t = Date.now();
  const clave = ['RSVP', id, respuesta, t, '', b64_(nombre)].join('~');
  conLock_(function () {
    const sh = hojaRespuestas_();
    asegurarEncabezados_(sh);
    const fila = buscarFila_(sh, id);
    if (fila) {
      sh.getRange(fila, 1, 1, 4).setValues([[new Date(t), id, nombre, ETIQUETAS[respuesta]]]);
      sh.getRange(fila, 6).setValue(clave);
    } else {
      sh.appendRow([new Date(t), id, nombre, ETIQUETAS[respuesta], '', clave]);
    }
  });
  return estado(nombre);
}

function registrarPago(nombre, monto) {
  nombre = limpiarNombre_(nombre);
  const id = idPara_(nombre);
  monto = parseInt(String(monto == null ? '' : monto).replace(/[^\d]/g, ''), 10);
  if (!(monto > 0 && monto <= 10000000)) throw new Error('Poné cuánto transferiste, solo el número.');
  const t = Date.now();
  const clave = ['PAGO', id, t, monto, b64_(nombre)].join('~');
  conLock_(function () {
    const sh = hojaRespuestas_();
    asegurarEncabezados_(sh);
    let fila = buscarFila_(sh, id);
    if (!fila) {
      sh.appendRow([new Date(t), id, nombre, ETIQUETAS.si, '', ['RSVP', id, 'si', t, '', b64_(nombre)].join('~')]);
      fila = sh.getLastRow();
    }
    sh.getRange(fila, 7, 1, 4).setValues([['Avisó que transfirió', monto, new Date(t), clave]]);
    sh.getRange(fila, 8).setNumberFormat('"$"#,##0');
  });
  return estado(nombre);
}

function estado(nombre) {
  nombre = limpiarNombre_(nombre);
  const id = idPara_(nombre);
  const sh = hojaRespuestas_();
  const fila = buscarFila_(sh, id);
  if (!fila) return { id: id, nombre: nombre, respuesta: '', monto: 0 };
  const v = sh.getRange(fila, 1, 1, ENCABEZADOS.length).getValues()[0];
  const m = String(v[5]).match(/^RSVP~[^~]*~(si|quizas|no)~/);
  return {
    id: id,
    nombre: String(v[2] || nombre),
    respuesta: m ? m[1] : '',
    monto: Number(v[7]) || 0,
  };
}

// ---------- configuración (se corre una vez desde el editor) ----------

function configurar() {
  const ss = planilla_();
  asegurarEncabezados_(ss.getSheets()[0]);
  let hoja = ss.getSheetByName(PESTANA_DATOS);
  if (!hoja) {
    hoja = ss.insertSheet(PESTANA_DATOS, ss.getNumSheets());
    hoja.getRange(1, 1, DATOS_INICIALES.length, 2).setValues(DATOS_INICIALES);
    hoja.getRange(1, 1, DATOS_INICIALES.length, 1).setFontWeight('bold');
    hoja.setColumnWidth(1, 180);
    hoja.setColumnWidth(2, 320);
  }
  return 'Listo. Completá precio, alias y titular en la pestaña "' + PESTANA_DATOS + '".';
}

// ---------- ayudantes ----------

function planilla_() {
  return SpreadsheetApp.openById(PLANILLA_ID);
}

// Las respuestas van siempre en la primera pestaña: la página del cumple lee esa.
function hojaRespuestas_() {
  return planilla_().getSheets()[0];
}

function leerDatos_() {
  const hoja = planilla_().getSheetByName(PESTANA_DATOS);
  let textos = DATOS_INICIALES;
  let crudos = DATOS_INICIALES;
  if (hoja && hoja.getLastRow() > 0) {
    const rango = hoja.getRange(1, 1, hoja.getLastRow(), 2);
    textos = rango.getDisplayValues();
    crudos = rango.getValues();
  }
  const fila = function (etiqueta) {
    const e = normalizar_(etiqueta);
    for (let i = 0; i < textos.length; i++) if (normalizar_(textos[i][0]) === e) return i;
    return -1;
  };
  const texto = function (etiqueta) {
    const i = fila(etiqueta);
    return i < 0 ? '' : String(textos[i][1]).trim();
  };
  const iPrecio = fila('Precio por persona');
  let precio = 0;
  if (iPrecio >= 0) {
    const crudo = crudos[iPrecio][1];
    precio = typeof crudo === 'number' ? Math.round(crudo) : parseInt(String(crudo).replace(/[^\d]/g, ''), 10) || 0;
  }
  return {
    fecha: texto('Fecha y hora') || DATOS_INICIALES[0][1],
    lugar: texto('Lugar'),
    precio: precio > 0 ? precio : 0,
    alias: texto('Alias'),
    titular: texto('Titular de la cuenta'),
  };
}

function asegurarEncabezados_(sh) {
  const actuales = sh.getRange(1, 1, 1, ENCABEZADOS.length).getDisplayValues()[0];
  for (let i = 0; i < ENCABEZADOS.length; i++) {
    if (!String(actuales[i]).trim()) sh.getRange(1, i + 1).setValue(ENCABEZADOS[i]);
  }
}

function buscarFila_(sh, id) {
  const ultima = sh.getLastRow();
  if (ultima < 2) return 0;
  const ids = sh.getRange(2, 2, ultima - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function limpiarNombre_(nombre) {
  const n = String(nombre == null ? '' : nombre)
    .replace(/\s+/g, ' ')
    .replace(/^[\s=+\-@]+/, '')
    .trim()
    .slice(0, 60);
  if (!n) throw new Error('Escribí tu nombre.');
  return n;
}

function sinAcentos_(s) {
  const mapa = { á: 'a', à: 'a', ä: 'a', â: 'a', ã: 'a', é: 'e', è: 'e', ë: 'e', ê: 'e', í: 'i', ì: 'i', ï: 'i', î: 'i',
    ó: 'o', ò: 'o', ö: 'o', ô: 'o', õ: 'o', ú: 'u', ù: 'u', ü: 'u', û: 'u', ñ: 'n', ç: 'c' };
  return String(s).toLowerCase().replace(/[áàäâãéèëêíìïîóòöôõúùüûñç]/g, function (c) { return mapa[c]; });
}

function normalizar_(s) {
  return sinAcentos_(s).replace(/\s+/g, ' ').trim();
}

// Mismo formato de ID que ya usa la planilla: "Felipe Alborch" -> "n-felipe-alborch".
function idPara_(nombre) {
  const slug = sinAcentos_(nombre).replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 50);
  if (!slug) throw new Error('Escribí tu nombre con letras.');
  return 'n-' + slug;
}

function b64_(s) {
  return Utilities.base64EncodeWebSafe(s, Utilities.Charset.UTF_8).replace(/=+$/, '');
}

// ---------- la invitación ----------

const PAGINA = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bagel+Fat+One&family=DM+Mono:wght@500&family=Figtree:wght@400;500;600;700&display=swap">
<style>
/* Una sola columna: título grande, cuándo y dónde, confirmar, después pagar */
:root {
  --bg: #FAF6FB; --surface: #FFFFFF; --surface-2: #F3ECF5; --fg: #23162B; --muted: #75667D; --line: #E6DCEA;
  --accent: #D11F69; --accent-ink: #FFFFFF; --accent-soft: #FBE1EC; --hl: #FFD45C;
  --ok: #17885A; --ok-bg: #DAF2E5; --err: #B3261E;
  --font-display: "Bagel Fat One", "Arial Rounded MT Bold", "Trebuchet MS", system-ui, sans-serif;
  --font-body: "Figtree", system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  --font-mono: "DM Mono", ui-monospace, Menlo, Consolas, monospace;
  color-scheme: light;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #17101C; --surface: #211827; --surface-2: #2C2133; --fg: #F5EEF7; --muted: #AC9CB3; --line: #3A2D42;
    --accent: #FF5A9A; --accent-ink: #1C0B14; --accent-soft: #46182E; --hl: #6E4C10;
    --ok: #52D29C; --ok-bg: #163829; --err: #FF8A80;
    color-scheme: dark;
  }
}
* { box-sizing: border-box; }
html { -webkit-text-size-adjust: 100%; }
body {
  margin: 0; background: var(--bg); color: var(--fg);
  font: 16px/1.5 var(--font-body);
  padding: max(28px, env(safe-area-inset-top, 0px)) max(16px, env(safe-area-inset-right, 0px)) max(40px, env(safe-area-inset-bottom, 0px)) max(16px, env(safe-area-inset-left, 0px));
}
main { max-width: 460px; margin: 0 auto; display: grid; gap: 18px; }
button, input { font: inherit; color: inherit; }
button { cursor: pointer; }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
[hidden] { display: none !important; }

.kicker { margin: 0; font-weight: 600; color: var(--accent); letter-spacing: .02em; }
h1 { font-family: var(--font-display); font-weight: 400; font-size: clamp(56px, 18vw, 88px); line-height: .9; margin: 0; }
h1 .hl { color: var(--accent); background: linear-gradient(transparent 58%, var(--hl) 58% 86%, transparent 86%); padding-inline: 4px; }
.info { display: grid; gap: 6px; margin: 0; }
.info div { display: grid; grid-template-columns: 72px minmax(0, 1fr); gap: 8px; align-items: baseline; }
.info dt { font-size: 12px; text-transform: uppercase; letter-spacing: .08em; color: var(--muted); font-weight: 600; }
.info dd { margin: 0; font-weight: 600; font-size: 18px; }
.nota { margin: 0; color: var(--muted); }

.card { background: var(--surface); border: 1px solid var(--line); border-radius: 18px; padding: 20px; display: grid; gap: 14px; }
.card h2 { margin: 0; font-size: 20px; }
form { display: grid; gap: 14px; }
.lbl { font-size: 14px; font-weight: 600; color: var(--muted); display: block; margin-bottom: 6px; }
.txt { width: 100%; border: 1px solid var(--line); background: var(--bg); border-radius: 12px; padding: 12px 14px; font-size: 17px; }
.txt:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
.opts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.opt { border: 1px solid var(--line); background: var(--bg); border-radius: 12px; padding: 12px 6px; font-weight: 700; }
.opt[aria-checked="true"] { background: var(--fg); color: var(--bg); border-color: var(--fg); }
.btn { border: 0; background: var(--accent); color: var(--accent-ink); font-weight: 700; font-size: 17px; border-radius: 12px; padding: 14px 18px; }
.btn:disabled { opacity: .6; cursor: progress; }
.btn.sec { background: var(--surface-2); color: var(--fg); font-size: 15px; padding: 12px 14px; }
.link { border: 0; background: none; padding: 0; color: var(--accent); font-weight: 600; text-decoration: underline; justify-self: start; }
.err { margin: 0; color: var(--err); font-weight: 600; font-size: 15px; }
.hecho { margin: 0; font-size: 18px; font-weight: 600; }

.monto { margin: 0; font-family: var(--font-display); font-size: 44px; line-height: 1; color: var(--accent); font-variant-numeric: tabular-nums; }
.monto small { font-family: var(--font-body); font-size: 15px; font-weight: 600; color: var(--muted); }
.alias { display: flex; gap: 8px; }
.alias .txt { font-family: var(--font-mono); font-size: 17px; min-width: 0; flex: 1 1 auto; }
.titular { margin: 0; color: var(--muted); }
.titular b { color: var(--fg); }
.peso { display: flex; align-items: center; gap: 8px; }
.peso span { font-weight: 700; font-size: 20px; color: var(--muted); }
.gracias { background: var(--ok-bg); color: var(--ok); border-radius: 12px; padding: 12px 14px; margin: 0; font-weight: 600; }

.toast { position: fixed; left: 50%; bottom: max(18px, env(safe-area-inset-bottom, 0px)); transform: translateX(-50%);
  background: var(--fg); color: var(--bg); padding: 10px 16px; border-radius: 999px; font-weight: 600; font-size: 15px; }
</style>
</head>
<body>
<main>
  <header style="display:grid;gap:14px">
    <p class="kicker">Te invito a mi cumple</p>
    <h1>Cumple <span class="hl">Nico</span></h1>
    <dl class="info">
      <div><dt>Cuándo</dt><dd id="fecha"></dd></div>
      <div><dt>Dónde</dt><dd id="lugar"></dd></div>
    </dl>
    <p class="nota" id="nota"></p>
  </header>

  <section class="card" id="paso-rsvp" aria-labelledby="t-rsvp">
    <h2 id="t-rsvp">¿Venís?</h2>
    <form id="f-rsvp" novalidate>
      <div>
        <label class="lbl" for="nombre">Tu nombre y apellido</label>
        <input class="txt" id="nombre" autocomplete="name" maxlength="60" placeholder="Ej: Felipe Alborch">
      </div>
      <div class="opts" role="radiogroup" aria-labelledby="t-rsvp">
        <button type="button" class="opt" role="radio" aria-checked="false" data-r="si">Voy</button>
        <button type="button" class="opt" role="radio" aria-checked="false" data-r="quizas">Quizás</button>
        <button type="button" class="opt" role="radio" aria-checked="false" data-r="no">No puedo</button>
      </div>
      <button type="submit" class="btn" id="b-rsvp">Confirmar</button>
      <p class="err" id="e-rsvp" role="alert" hidden></p>
    </form>
  </section>

  <section class="card" id="paso-listo" hidden>
    <p class="hecho" id="m-listo"></p>
    <button type="button" class="link" id="cambiar">Cambiar mi respuesta</button>
  </section>

  <section class="card" id="pago" aria-labelledby="t-pago" hidden>
    <h2 id="t-pago">Tu parte del alquiler</h2>
    <p id="pago-pendiente" class="nota" hidden>Todavía estoy cerrando el alquiler. En unos días te paso cuánto sale; volvé a este mismo link para pagar.</p>
    <div id="pago-listo" style="display:grid;gap:14px" hidden>
      <p class="monto"><span id="p-precio"></span> <small>por persona</small></p>
      <div>
        <label class="lbl" for="p-alias">Transferí a este alias</label>
        <div class="alias">
          <input class="txt" id="p-alias" readonly>
          <button type="button" class="btn sec" id="copiar">Copiar</button>
        </div>
      </div>
      <p class="titular" id="p-titular" hidden></p>
      <p class="gracias" id="m-pago" hidden></p>
      <button type="button" class="link" id="corregir" hidden>Me equivoqué con el monto</button>
      <form id="f-pago" novalidate>
        <div>
          <label class="lbl" for="monto">Cuando transfieras, avisame cuánto mandaste</label>
          <div class="peso"><span aria-hidden="true">$</span><input class="txt" id="monto" inputmode="numeric" autocomplete="off"></div>
        </div>
        <button type="submit" class="btn" id="b-pago">Ya transferí</button>
        <p class="err" id="e-pago" role="alert" hidden></p>
      </form>
    </div>
  </section>
</main>
<div class="toast" id="toast" role="status" aria-live="polite" hidden></div>

<script>
(function () {
  var D = __DATOS__;
  var $ = function (id) { return document.getElementById(id); };
  var yo = { nombre: '', respuesta: '', monto: 0 };
  var elegida = '';

  function plata(n) { return '$' + Number(n).toLocaleString('es-AR'); }
  function guardar(nombre) { try { localStorage.setItem('cumple-nico-nombre', nombre); } catch (e) {} }
  function recordado() { try { return localStorage.getItem('cumple-nico-nombre') || ''; } catch (e) { return ''; } }
  function llamar(fn, args) {
    return new Promise(function (ok, mal) {
      var r = google.script.run.withSuccessHandler(ok).withFailureHandler(mal);
      r[fn].apply(r, args);
    });
  }
  function mensajeError(e) {
    var m = (e && e.message ? e.message : String(e || '')).replace(/^(Exception|Error):\\s*/, '');
    return m || 'No se pudo guardar. Probá de nuevo en un rato.';
  }
  var tt;
  function toast(msg) {
    var t = $('toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(tt); tt = setTimeout(function () { t.hidden = true; }, 2400);
  }

  $('fecha').textContent = D.fecha;
  $('lugar').textContent = D.lugar || 'A confirmar';
  var cobrar = D.precio > 0 && D.alias;
  $('nota').textContent = cobrar
    ? 'Alquilo un lugar y lo pagamos entre todos: sale ' + plata(D.precio) + ' por persona.'
    : 'Alquilo un lugar y lo pagamos entre todos. En unos días paso cuánto sale por persona.';

  function elegir(r) {
    elegida = r;
    document.querySelectorAll('.opt').forEach(function (b) { b.setAttribute('aria-checked', String(b.dataset.r === r)); });
  }
  document.querySelector('.opts').addEventListener('click', function (e) {
    var b = e.target.closest('.opt'); if (b) elegir(b.dataset.r);
  });

  function mostrarListo() {
    var n = yo.nombre.split(' ')[0];
    var msg = { si: '¡Genial, ' + n + '! Te espero.',
      quizas: 'Dale, ' + n + '. Cuando sepas, volvé a este link y confirmá.',
      no: 'Qué lástima, ' + n + '. ¡Gracias por avisar!' }[yo.respuesta];
    $('m-listo').textContent = msg;
    $('paso-rsvp').hidden = true;
    $('paso-listo').hidden = false;
    var paga = yo.respuesta === 'si';
    $('pago').hidden = !paga;
    if (!paga) return;
    $('pago-pendiente').hidden = !!cobrar;
    $('pago-listo').hidden = !cobrar;
    if (!cobrar) return;
    $('p-precio').textContent = plata(D.precio);
    $('p-alias').value = D.alias;
    $('p-titular').hidden = !D.titular;
    $('p-titular').innerHTML = '';
    if (D.titular) { $('p-titular').append('A nombre de '); var b = document.createElement('b'); b.textContent = D.titular; $('p-titular').append(b); }
    mostrarPago(false);
  }
  function mostrarPago(corrigiendo) {
    var avisado = yo.monto > 0 && !corrigiendo;
    $('m-pago').hidden = !avisado;
    $('corregir').hidden = !avisado;
    $('f-pago').hidden = avisado;
    if (avisado) $('m-pago').textContent = 'Me avisaste que transferiste ' + plata(yo.monto) + '. ¡Gracias!';
    else $('monto').value = String(yo.monto || D.precio).replace(/\\B(?=(\\d{3})+(?!\\d))/g, '.');
  }

  $('f-rsvp').addEventListener('submit', function (e) {
    e.preventDefault();
    var err = $('e-rsvp'); err.hidden = true;
    var nombre = $('nombre').value.replace(/\\s+/g, ' ').trim();
    if (!nombre) { err.textContent = 'Escribí tu nombre.'; err.hidden = false; $('nombre').focus(); return; }
    if (!elegida) { err.textContent = 'Elegí si venís, quizás o no.'; err.hidden = false; return; }
    var b = $('b-rsvp'); b.disabled = true; b.textContent = 'Guardando…';
    llamar('responder', [nombre, elegida]).then(function (r) {
      yo = r; guardar(nombre); mostrarListo();
    }).catch(function (x) {
      err.textContent = mensajeError(x); err.hidden = false;
    }).then(function () { b.disabled = false; b.textContent = 'Confirmar'; });
  });

  $('cambiar').addEventListener('click', function () {
    $('nombre').value = yo.nombre; elegir(yo.respuesta);
    $('paso-listo').hidden = true; $('pago').hidden = true; $('paso-rsvp').hidden = false;
  });
  $('corregir').addEventListener('click', function () { mostrarPago(true); $('monto').focus(); });

  $('copiar').addEventListener('click', function () {
    var inp = $('p-alias');
    var listo = function () { toast('Alias copiado'); };
    var aMano = function () {
      inp.focus(); inp.select();
      try { if (document.execCommand('copy')) { listo(); return; } } catch (e) {}
      toast('Mantené apretado el alias para copiarlo');
    };
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(D.alias).then(listo, aMano);
      else aMano();
    } catch (e) { aMano(); }
  });

  $('f-pago').addEventListener('submit', function (e) {
    e.preventDefault();
    var err = $('e-pago'); err.hidden = true;
    var monto = $('monto').value.replace(/[^\\d]/g, '');
    if (!monto || Number(monto) <= 0) { err.textContent = 'Poné cuánto transferiste, solo el número.'; err.hidden = false; return; }
    var b = $('b-pago'); b.disabled = true; b.textContent = 'Avisando…';
    llamar('registrarPago', [yo.nombre, monto]).then(function (r) {
      yo = r; mostrarPago(false);
    }).catch(function (x) {
      err.textContent = mensajeError(x); err.hidden = false;
    }).then(function () { b.disabled = false; b.textContent = 'Ya transferí'; });
  });

  var nombreGuardado = recordado();
  if (nombreGuardado) {
    $('nombre').value = nombreGuardado;
    llamar('estado', [nombreGuardado]).then(function (r) {
      if (r && r.respuesta) { yo = r; mostrarListo(); }
    }).catch(function () {});
  }
})();
</script>
</body>
</html>`;
