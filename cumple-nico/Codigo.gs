/**
 * Invitación "Cumple Nico": confirmaciones y pagos.
 *
 * Va entero en Código.gs, reemplazando todo lo que había.
 * La primera vez, elegí la función "configurar" arriba y tocá Ejecutar.
 *
 * Lo que cambia seguido (fecha, lugar, precios, CVU o alias, titular) se edita
 * en la pestaña "Invitación" de la planilla, no acá.
 *
 * Mercado Pago: el Access Token va en Configuración del proyecto →
 * Propiedades del script, con el nombre MP_ACCESS_TOKEN. Nunca en este código.
 */

const PLANILLA_ID = '1IXz1OIY2vBBwrMLt_LnP3XVAYnqfZOMvmAvJqPYM3p4';
const PESTANA_DATOS = 'Invitación';
const PRECIO = 'Precio por persona';
const PRECIO_MP = 'Precio con Mercado Pago';
const CUENTA = 'Alias o CBU/CVU';

// $400.000 de alquiler entre 80 personas = $5.000. Con Mercado Pago se suma la
// comisión (6,29% + IVA si la plata se acredita al instante) para que igual te lleguen $5.000.
const DATOS_INICIALES = [
  ['Fecha y hora', 'Viernes 4 de diciembre · 23 hs'],
  ['Lugar', 'A confirmar'],
  [PRECIO, 5000],
  [PRECIO_MP, 5450],
  [CUENTA, ''],
  ['Titular de la cuenta', ''],
];

const ENCABEZADOS = [
  'Fecha', 'ID', 'Nombre', 'Respuesta', 'Lleva', 'Clave (no tocar)',
  'Pagó', 'Monto', 'Fecha pago', 'Clave pago (no tocar)', 'Código (no tocar)',
];
const ETIQUETAS = { si: 'Va', quizas: 'Quizás', no: 'No va' };
const AVISO = 'Avisó que transfirió';
const PAGO_MP = 'Pagó por Mercado Pago';

const MP_API = 'https://api.mercadopago.com';
const MP_REF = 'cumple-nico:';

// ---------- página ----------

function doGet() {
  const datos = leerDatos_();
  const json = JSON.stringify(datos).replace(/</g, '\\u003c');
  const html = PAGINA.replace('__DATOS__', function () { return json; });
  return HtmlService.createHtmlOutput(html)
    .setTitle('Cumple Nico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- funciones que llama la invitación ----------

// Cada invitado recibe un código propio al confirmar. Sin ese código nadie
// puede ver la respuesta de otro: lo que contestó cada uno lo ves solo vos.
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
      if (!String(sh.getRange(fila, 11).getValue())) sh.getRange(fila, 11).setValue(nuevoCodigo_());
    } else {
      sh.appendRow([new Date(t), id, nombre, ETIQUETAS[respuesta], '', clave, '', '', '', '', nuevoCodigo_()]);
    }
  });
  return estadoPorId_(id, nombre);
}

function estado(nombre, codigo) {
  const id = idPara_(limpiarNombre_(nombre));
  return delInvitado_(estadoPorId_(id, nombre), codigo);
}

function registrarPago(nombre, codigo, monto) {
  nombre = limpiarNombre_(nombre);
  const id = idPara_(nombre);
  delInvitado_(estadoPorId_(id, nombre), codigo, true);
  monto = parseInt(String(monto == null ? '' : monto).replace(/[^\d]/g, ''), 10);
  if (!(monto > 0 && monto <= 10000000)) throw new Error('Poné cuánto transferiste, solo el número.');
  const t = Date.now();
  const clave = ['PAGO', id, t, monto, b64_(nombre)].join('~');
  conLock_(function () {
    const sh = hojaRespuestas_();
    const fila = buscarFila_(sh, id);
    if (String(sh.getRange(fila, 7).getValue()) === PAGO_MP) throw new Error('Tu pago por Mercado Pago ya está registrado.');
    sh.getRange(fila, 7, 1, 4).setValues([[AVISO, monto, new Date(t), clave]]);
    sh.getRange(fila, 8).setNumberFormat('"$"#,##0');
  });
  return estadoPorId_(id, nombre);
}

function linkDePago(nombre, codigo) {
  nombre = limpiarNombre_(nombre);
  const id = idPara_(nombre);
  delInvitado_(estadoPorId_(id, nombre), codigo, true);
  const d = leerDatos_();
  if (!d.mp) throw new Error('El pago con Mercado Pago no está activado. Transferí al CVU.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('pend:' + id, String(Date.now()));
  const guardado = props.getProperty('link:' + id + ':' + d.precioMp);
  if (guardado) return guardado;
  const pref = mp_('post', '/checkout/preferences', {
    items: [{ id: 'cumple-nico', title: 'Cumple Nico: tu parte del alquiler', quantity: 1, currency_id: 'ARS', unit_price: d.precioMp }],
    external_reference: MP_REF + id,
    metadata: { invitado: nombre },
    binary_mode: true, // aprobado o rechazado al momento, sin pagos en efectivo que quedan pendientes
    statement_descriptor: 'CUMPLE NICO',
  });
  if (!pref.init_point) throw new Error('Mercado Pago no devolvió el link de pago. Probá de nuevo.');
  props.setProperty('link:' + id + ':' + d.precioMp, pref.init_point);
  return pref.init_point;
}

// La invitación la llama cuando el invitado vuelve de Mercado Pago a la pestaña.
function comprobarPago(nombre, codigo) {
  nombre = limpiarNombre_(nombre);
  const id = idPara_(nombre);
  delInvitado_(estadoPorId_(id, nombre), codigo, true);
  if (tokenMp_() && actualizarPagoMp_(id)) PropertiesService.getScriptProperties().deleteProperty('pend:' + id);
  return estadoPorId_(id, nombre);
}

// ---------- Mercado Pago ----------

// Corre sola cada 10 minutos (la programa "configurar"): busca los pagos aprobados
// de quienes abrieron el link de pago y los marca en la planilla.
function revisarPagos() {
  if (!tokenMp_()) return;
  const props = PropertiesService.getScriptProperties();
  const todas = props.getProperties();
  Object.keys(todas).forEach(function (k) {
    if (k.indexOf('pend:') !== 0) return;
    const id = k.slice(5);
    try {
      if (actualizarPagoMp_(id) || Date.now() - Number(todas[k]) > 45 * 864e5) props.deleteProperty(k);
    } catch (err) {
      console.warn('No pude revisar el pago de ' + id + ': ' + err.message);
    }
  });
}

function actualizarPagoMp_(id) {
  const ref = MP_REF + id;
  const res = mp_('get', '/v1/payments/search?sort=date_created&criteria=desc&limit=50&external_reference=' + encodeURIComponent(ref));
  const ok = (res.results || []).filter(function (p) { return p.status === 'approved' && p.external_reference === ref; });
  if (!ok.length) return false;
  let bruto = 0;
  let neto = 0;
  ok.forEach(function (p) {
    const monto = Number(p.transaction_amount) || 0;
    const llega = p.transaction_details ? Number(p.transaction_details.net_received_amount) : 0;
    bruto += monto;
    neto += llega > 0 ? llega : monto;
  });
  const nombre = (ok[0].metadata && ok[0].metadata.invitado) || '';
  marcarPagoMp_(id, nombre, Math.round(bruto), Math.round(neto), new Date(ok[0].date_approved || ok[0].date_created || Date.now()));
  return true;
}

function marcarPagoMp_(id, nombre, bruto, neto, fecha) {
  conLock_(function () {
    const sh = hojaRespuestas_();
    asegurarEncabezados_(sh);
    let fila = buscarFila_(sh, id);
    if (!fila) {
      nombre = String(nombre || '').replace(/^[\s=+\-@]+/, '').slice(0, 60);
      if (!nombre) return;
      sh.appendRow([fecha, id, nombre, ETIQUETAS.si, '', ['RSVP', id, 'si', fecha.getTime(), '', b64_(nombre)].join('~'), '', '', '', '', nuevoCodigo_()]);
      fila = sh.getLastRow();
    }
    const actual = sh.getRange(fila, 7, 1, 2).getValues()[0];
    if (String(actual[0]) === PAGO_MP && Number(actual[1]) === bruto) return;
    const quien = String(sh.getRange(fila, 3).getValue() || nombre);
    const clave = ['PAGOMP', id, fecha.getTime(), bruto, b64_(quien), neto].join('~');
    sh.getRange(fila, 7, 1, 4).setValues([[PAGO_MP, bruto, fecha, clave]]);
    sh.getRange(fila, 8).setNumberFormat('"$"#,##0');
  });
}

function mp_(metodo, ruta, cuerpo) {
  const token = tokenMp_();
  if (!token) throw new Error('Falta el Access Token de Mercado Pago.');
  const opciones = { method: metodo, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  if (cuerpo) {
    opciones.contentType = 'application/json';
    opciones.payload = JSON.stringify(cuerpo);
  }
  const r = UrlFetchApp.fetch(MP_API + ruta, opciones);
  const texto = r.getContentText();
  let datos = {};
  try { datos = texto ? JSON.parse(texto) : {}; } catch (err) { datos = {}; }
  if (r.getResponseCode() >= 300) {
    throw new Error('Mercado Pago respondió ' + r.getResponseCode() + (datos.message ? ': ' + datos.message : ''));
  }
  return datos;
}

function tokenMp_() {
  return String(PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN') || '').trim();
}

// ---------- configuración (se corre desde el editor) ----------

function configurar() {
  const ss = planilla_();
  asegurarEncabezados_(ss.getSheets()[0]);

  let hoja = ss.getSheetByName(PESTANA_DATOS);
  if (!hoja) hoja = ss.insertSheet(PESTANA_DATOS, ss.getNumSheets());
  // Agrega las filas que falten, sin tocar lo que ya completaste
  const etiquetas = hoja.getLastRow() ? hoja.getRange(1, 1, hoja.getLastRow(), 1).getDisplayValues().map(function (r) { return normalizar_(r[0]); }) : [];
  DATOS_INICIALES.forEach(function (fila) {
    if (etiquetas.indexOf(normalizar_(fila[0])) >= 0) return;
    const n = hoja.getLastRow() + 1;
    // El CBU va como texto: si no, Sheets lo toma como número y le saca los ceros.
    if (fila[0] === CUENTA) hoja.getRange(n, 2).setNumberFormat('@');
    hoja.getRange(n, 1, 1, 2).setValues([fila]);
    etiquetas.push(normalizar_(fila[0]));
  });
  if (etiquetas.indexOf(normalizar_('Plata que avisaron o pagaron')) < 0) {
    const resp = "'" + ss.getSheets()[0].getName().replace(/'/g, "''") + "'";
    const n = hoja.getLastRow() + 2;
    hoja.getRange(n, 1, 2, 2).setValues([
      ['Plata que avisaron o pagaron', '=SUM(' + resp + '!H2:H)'],
      ['Personas que avisaron o pagaron', '=COUNT(' + resp + '!H2:H)'],
    ]);
    hoja.getRange(n, 2).setNumberFormat('"$"#,##0');
  }
  hoja.getRange(1, 1, hoja.getLastRow(), 1).setFontWeight('bold');
  hoja.setColumnWidth(1, 220);
  hoja.setColumnWidth(2, 320);

  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'revisarPagos'; })) {
    ScriptApp.newTrigger('revisarPagos').timeBased().everyMinutes(10).create();
  }

  let mp = 'Mercado Pago: falta el Access Token (Configuración del proyecto → Propiedades del script → MP_ACCESS_TOKEN).';
  if (tokenMp_()) {
    try {
      const yo = mp_('get', '/users/me');
      mp = 'Mercado Pago conectado a la cuenta ' + (yo.nickname || yo.email || yo.id) + '.';
    } catch (err) {
      mp = 'Mercado Pago no aceptó el Access Token: ' + err.message;
    }
  }
  const msg = 'Planilla lista. ' + mp;
  console.log(msg);
  return msg;
}

// ---------- ayudantes ----------

function planilla_() {
  return SpreadsheetApp.openById(PLANILLA_ID);
}

// Las respuestas van siempre en la primera pestaña: tu página del cumple lee esa.
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
  const texto = function () {
    for (let k = 0; k < arguments.length; k++) {
      const i = fila(arguments[k]);
      if (i >= 0) return String(textos[i][1]).trim();
    }
    return '';
  };
  const numero = function (etiqueta) {
    const i = fila(etiqueta);
    if (i < 0) return 0;
    const v = crudos[i][1];
    const n = typeof v === 'number' ? Math.round(v) : parseInt(String(v).replace(/[^\d]/g, ''), 10);
    return n > 0 ? n : 0;
  };
  const precio = numero(PRECIO);
  const precioMp = numero(PRECIO_MP) || precio;
  return {
    fecha: texto('Fecha y hora') || DATOS_INICIALES[0][1],
    lugar: texto('Lugar'),
    precio: precio,
    precioMp: precioMp,
    mp: !!tokenMp_() && precioMp > 0,
    alias: texto(CUENTA, 'Alias', 'CBU', 'CVU'),
    titular: texto('Titular de la cuenta'),
  };
}

function estadoPorId_(id, nombre) {
  const sh = hojaRespuestas_();
  const fila = buscarFila_(sh, id);
  if (!fila) return { id: id, nombre: String(nombre || ''), respuesta: '', monto: 0, mp: false, codigo: '' };
  const v = sh.getRange(fila, 1, 1, ENCABEZADOS.length).getValues()[0];
  const m = String(v[5]).match(/^RSVP~[^~]*~(si|quizas|no)~/);
  return {
    id: id,
    nombre: String(v[2] || nombre || ''),
    respuesta: m ? m[1] : '',
    monto: Number(v[7]) || 0,
    mp: String(v[6]) === PAGO_MP,
    codigo: String(v[10] || ''),
  };
}

// Sin el código de esa persona no se muestra nada de su respuesta.
function delInvitado_(est, codigo, obligatorio) {
  if (est.codigo && String(codigo || '') === est.codigo) return est;
  if (obligatorio) throw new Error('Primero confirmá tu asistencia arriba.');
  return { id: est.id, nombre: est.nombre, respuesta: '', monto: 0, mp: false, codigo: '' };
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

function nuevoCodigo_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
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
.chico { margin: 0; color: var(--muted); font-size: 14px; }

.card { background: var(--surface); border: 1px solid var(--line); border-radius: 18px; padding: 20px; display: grid; gap: 14px; }
.card h2 { margin: 0; font-size: 20px; }
.grupo { display: grid; gap: 12px; }
form { display: grid; gap: 14px; }
.lbl { font-size: 14px; font-weight: 600; color: var(--muted); display: block; margin-bottom: 6px; }
.txt { width: 100%; border: 1px solid var(--line); background: var(--bg); border-radius: 12px; padding: 12px 14px; font-size: 17px; }
.txt:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
.opts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.opt { border: 1px solid var(--line); background: var(--bg); border-radius: 12px; padding: 12px 6px; font-weight: 700; }
.opt[aria-checked="true"] { background: var(--fg); color: var(--bg); border-color: var(--fg); }
.btn { border: 0; background: var(--accent); color: var(--accent-ink); font-weight: 700; font-size: 17px; border-radius: 12px; padding: 14px 18px; display: block; text-align: center; text-decoration: none; }
.btn:disabled, .btn[aria-disabled="true"] { opacity: .6; cursor: progress; }
.btn.sec { background: var(--surface-2); color: var(--fg); font-size: 15px; padding: 12px 14px; }
.link { border: 0; background: none; padding: 0; color: var(--accent); font-weight: 600; text-decoration: underline; justify-self: start; text-align: left; }
.err { margin: 0; color: var(--err); font-weight: 600; font-size: 15px; }
.hecho { margin: 0; font-size: 18px; font-weight: 600; }

.monto { margin: 0; font-family: var(--font-display); font-size: 44px; line-height: 1; color: var(--accent); font-variant-numeric: tabular-nums; }
.monto small { font-family: var(--font-body); font-size: 15px; font-weight: 600; color: var(--muted); }
.alias { display: flex; flex-wrap: wrap; gap: 8px; }
.alias .txt { font-family: var(--font-mono); font-size: 17px; min-width: 0; flex: 1 1 240px; }
.alias .btn { flex: 0 0 auto; }
.titular { margin: 0; color: var(--muted); }
.titular b { color: var(--fg); }
.peso { display: flex; align-items: center; gap: 8px; }
.peso span { font-weight: 700; font-size: 20px; color: var(--muted); }
.gracias { background: var(--ok-bg); color: var(--ok); border-radius: 12px; padding: 12px 14px; margin: 0; font-weight: 600; }
.estado-mp { background: var(--surface-2); border-radius: 12px; padding: 12px 14px; margin: 0; font-weight: 600; }
.estado-mp.mal { color: var(--err); }

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
    <div id="pago-listo" class="grupo" hidden>
      <p class="monto"><span id="p-precio"></span> <small>por persona</small></p>
      <p class="estado-mp" id="estado-mp" role="status" hidden></p>
      <p class="gracias" id="m-pago" hidden></p>

      <div id="mp" class="grupo" hidden>
        <a class="btn" id="b-mp" href="#" target="_blank" rel="noopener" aria-disabled="true">Preparando el pago…</a>
        <button type="button" class="btn sec" id="b-revisar" hidden>Ya pagué, revisar</button>
        <p class="chico" id="mp-nota"></p>
        <button type="button" class="link" id="ver-cvu" hidden></button>
      </div>

      <div id="transfer" class="grupo" hidden>
        <div>
          <label class="lbl" for="p-alias" id="l-alias">Transferí a este alias</label>
          <div class="alias">
            <input class="txt" id="p-alias" readonly>
            <button type="button" class="btn sec" id="copiar">Copiar</button>
          </div>
        </div>
        <p class="titular" id="p-titular" hidden></p>
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
    </div>
  </section>
</main>
<div class="toast" id="toast" role="status" aria-live="polite" hidden></div>

<script>
(function () {
  var D = __DATOS__;
  var $ = function (id) { return document.getElementById(id); };
  var yo = { nombre: '', respuesta: '', monto: 0, mp: false, codigo: '' };
  var elegida = '';
  var verCvu = false;
  var linkMp = '';
  var pagando = false;
  var ultimaRevision = 0;

  function plata(n) { return '$' + Number(n).toLocaleString('es-AR'); }
  function guardar(p) { try { localStorage.setItem('cumple-nico', JSON.stringify({ nombre: p.nombre, codigo: p.codigo })); } catch (e) {} }
  function recordado() { try { return JSON.parse(localStorage.getItem('cumple-nico') || 'null'); } catch (e) { return null; } }
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

  // Un CBU o CVU son 22 números; los CVU (Mercado Pago, Ualá...) empiezan con 000000
  var cuenta = /^[\\d\\s-]+$/.test(D.alias) ? D.alias.replace(/[\\s-]/g, '') : D.alias;
  var tipo = /^\\d{22}$/.test(cuenta) ? (cuenta.indexOf('000000') === 0 ? 'CVU' : 'CBU') : 'alias';
  var hayCuenta = !!cuenta;
  var cobrar = D.precio > 0 && (D.mp || hayCuenta);

  $('fecha').textContent = D.fecha;
  $('lugar').textContent = D.lugar || 'A confirmar';
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
    $('l-alias').textContent = 'Transferí ' + plata(D.precio) + ' a este ' + tipo;
    $('p-alias').value = cuenta;
    $('p-titular').hidden = !D.titular;
    $('p-titular').innerHTML = '';
    if (D.titular) { $('p-titular').append('A nombre de '); var b = document.createElement('b'); b.textContent = D.titular; $('p-titular').append(b); }
    var extra = D.precioMp - D.precio;
    $('mp-nota').textContent = (extra > 0 ? 'Incluye ' + plata(extra) + ' de comisión de Mercado Pago. ' : '') +
      'Se confirma solo cuando Mercado Pago aprueba el pago.';
    $('ver-cvu').textContent = 'Prefiero transferir ' + (extra > 0 ? plata(D.precio) + ' ' : '') + 'al ' + tipo + (extra > 0 ? ', sin comisión' : '');
    $('ver-cvu').hidden = !hayCuenta;
    mostrarPago(false);
  }

  function mostrarPago(corrigiendo) {
    var pagoMp = yo.mp && yo.monto > 0;
    var avisado = !pagoMp && yo.monto > 0 && !corrigiendo;
    $('m-pago').hidden = !(pagoMp || avisado);
    if (pagoMp) $('m-pago').textContent = 'Tu pago de ' + plata(yo.monto) + ' entró por Mercado Pago. ¡Gracias!';
    else if (avisado) $('m-pago').textContent = 'Me avisaste que transferiste ' + plata(yo.monto) + '. ¡Gracias!';
    var conMp = D.mp && !pagoMp && !avisado && !verCvu && !corrigiendo;
    $('mp').hidden = !conMp;
    $('transfer').hidden = pagoMp || !hayCuenta || (conMp && !verCvu);
    $('corregir').hidden = !avisado;
    $('f-pago').hidden = avisado;
    if (!avisado) $('monto').value = String((!yo.mp && yo.monto) || D.precio).replace(/\\B(?=(\\d{3})+(?!\\d))/g, '.');
    if (conMp) prepararMp();
  }

  function prepararMp() {
    var a = $('b-mp');
    if (linkMp) return;
    a.setAttribute('aria-disabled', 'true'); a.textContent = 'Preparando el pago…';
    llamar('linkDePago', [yo.nombre, yo.codigo]).then(function (url) {
      linkMp = url; a.href = url; a.setAttribute('aria-disabled', 'false');
      a.textContent = 'Pagar ' + plata(D.precioMp) + ' con Mercado Pago';
    }).catch(function (x) {
      $('mp').hidden = true;
      $('transfer').hidden = !hayCuenta;
      avisoMp(mensajeError(x), true);
    });
  }
  $('b-mp').addEventListener('click', function (e) {
    if (!linkMp) { e.preventDefault(); return; }
    pagando = true;
    $('b-revisar').hidden = false;
    avisoMp('Cuando termines en Mercado Pago, volvé a esta pestaña: el pago se confirma solo.', false);
  });
  $('b-revisar').addEventListener('click', function () { revisar(true); });

  // Le pregunta a Mercado Pago si el pago de esta persona ya está aprobado
  function revisar(aMano) {
    if (!yo.codigo || yo.mp || Date.now() - ultimaRevision < 8000) return;
    ultimaRevision = Date.now();
    var b = $('b-revisar'); b.disabled = true; b.textContent = 'Revisando…';
    llamar('comprobarPago', [yo.nombre, yo.codigo]).then(function (r) {
      if (r && r.mp) { yo = r; pagando = false; avisoMp('', false); mostrarPago(false); }
      else if (aMano) avisoMp('Todavía no veo el pago. Si ya pagaste, puede tardar unos minutos: se confirma solo.', false);
    }).catch(function (x) {
      if (aMano) avisoMp(mensajeError(x), true);
    }).then(function () { b.disabled = false; b.textContent = 'Ya pagué, revisar'; });
  }
  $('ver-cvu').addEventListener('click', function () { verCvu = true; mostrarPago(false); $('p-alias').focus(); });

  function avisoMp(msg, mal) {
    var el = $('estado-mp');
    el.hidden = !msg; el.textContent = msg || ''; el.classList.toggle('mal', !!mal);
  }

  // Al volver a esta pestaña después de pagar, se fija si el pago ya entró
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible' && pagando) revisar(false);
  });

  $('f-rsvp').addEventListener('submit', function (e) {
    e.preventDefault();
    var err = $('e-rsvp'); err.hidden = true;
    var nombre = $('nombre').value.replace(/\\s+/g, ' ').trim();
    if (!nombre) { err.textContent = 'Escribí tu nombre.'; err.hidden = false; $('nombre').focus(); return; }
    if (!elegida) { err.textContent = 'Elegí si venís, quizás o no.'; err.hidden = false; return; }
    var b = $('b-rsvp'); b.disabled = true; b.textContent = 'Guardando…';
    llamar('responder', [nombre, elegida]).then(function (r) {
      yo = r; linkMp = ''; verCvu = false; guardar(r); mostrarListo();
    }).catch(function (x) {
      err.textContent = mensajeError(x); err.hidden = false;
    }).then(function () { b.disabled = false; b.textContent = 'Confirmar'; });
  });

  $('cambiar').addEventListener('click', function () {
    $('nombre').value = yo.nombre; elegir(yo.respuesta); avisoMp('', false);
    $('paso-listo').hidden = true; $('pago').hidden = true; $('paso-rsvp').hidden = false;
  });
  $('corregir').addEventListener('click', function () { mostrarPago(true); $('monto').focus(); });

  $('copiar').addEventListener('click', function () {
    var inp = $('p-alias');
    var nombreDato = tipo === 'alias' ? 'Alias' : tipo;
    var listo = function () { toast(nombreDato + ' copiado'); };
    var aMano = function () {
      inp.focus(); inp.select();
      try { if (document.execCommand('copy')) { listo(); return; } } catch (e) {}
      toast('Mantené apretado el ' + tipo + ' para copiarlo');
    };
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(cuenta).then(listo, aMano);
      else aMano();
    } catch (e) { aMano(); }
  });

  $('f-pago').addEventListener('submit', function (e) {
    e.preventDefault();
    var err = $('e-pago'); err.hidden = true;
    var monto = $('monto').value.replace(/[^\\d]/g, '');
    if (!monto || Number(monto) <= 0) { err.textContent = 'Poné cuánto transferiste, solo el número.'; err.hidden = false; return; }
    var b = $('b-pago'); b.disabled = true; b.textContent = 'Avisando…';
    llamar('registrarPago', [yo.nombre, yo.codigo, monto]).then(function (r) {
      yo = r; verCvu = false; mostrarPago(false);
    }).catch(function (x) {
      err.textContent = mensajeError(x); err.hidden = false;
    }).then(function () { b.disabled = false; b.textContent = 'Ya transferí'; });
  });

  // Alguien que ya confirmó desde este celu: se muestra su respuesta
  var guardado = recordado();
  if (guardado && guardado.nombre) {
    $('nombre').value = guardado.nombre;
    llamar('estado', [guardado.nombre, guardado.codigo]).then(function (r) {
      if (r && r.respuesta) { yo = r; mostrarListo(); }
    }).catch(function () {});
  }
})();
</script>
</body>
</html>`;
