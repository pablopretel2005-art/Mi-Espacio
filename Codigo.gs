/**
 * Mi espacio · servidor (Google Apps Script)
 *
 * Va dentro de la hoja de Google: Extensiones > Apps Script.
 * Es lo único que lee y escribe la hoja. La app (index.html en GitHub Pages) le habla por internet.
 *
 * Todo se hace desde el menú «Mi espacio» de la hoja (no con el botón Ejecutar del editor):
 *   1. Preparar la hoja
 *   2. Poner o cambiar la contraseña
 *
 * La contraseña no se guarda tal cual: solo una «huella» con sal en las propiedades del script.
 *
 * @OnlyCurrentDoc
 */

const HOJA = 'datos';
const CABECERA = ['clave', 'rev', 'versión', 'editado', 'borrado', 'carpeta', 'título (solo para leer)', 'datos'];
const COL_DATOS = 8;            // los datos (JSON) empiezan en la columna H
const TROZO = 40000;            // caracteres por celda (Google admite 50.000)
const MAX_JSON = 400000;        // tamaño máximo de un elemento
const MAX_CAMBIOS = 300;        // cambios por petición
const SESION_DIAS = 60;         // una sesión caduca tras 60 días sin usarse
const MAX_FALLOS = 5;           // intentos fallidos antes del bloqueo
const BLOQUEO_MIN = 15;         // minutos de bloqueo
const ITERACIONES = 2000;       // vueltas del cifrado de la contraseña
const CLAVE_RE = /^(app|e:[A-Za-z0-9_~:@+\-]{1,80})$/;

/* ================= menú de la hoja ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Mi espacio')
    .addItem('1. Preparar la hoja', 'prepararHoja')
    .addItem('2. Poner o cambiar la contraseña', 'ponerContrasena')
    .addSeparator()
    .addItem('Ver dispositivos conectados', 'verSesiones')
    .addItem('Cerrar todas las sesiones', 'cerrarSesiones')
    .addItem('Limpiar borrados antiguos', 'limpiarBorrados')
    .addToUi();
}

function prepararHoja() {
  const sh = hoja_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('REV')) props.setProperty('REV', '0');
  const tiene = !!props.getProperty('PW');
  SpreadsheetApp.getUi().alert('Hoja preparada',
    'La pestaña «' + HOJA + '» está lista. No la edites a mano.\n\n' +
    (tiene ? 'Ya hay una contraseña puesta.' : 'Siguiente paso: menú Mi espacio > 2. Poner o cambiar la contraseña.'),
    SpreadsheetApp.getUi().ButtonSet.OK);
  return sh.getName();
}

function ponerContrasena() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Contraseña de Mi espacio',
    'Escribe la contraseña (mínimo 8 caracteres). La usarás al entrar en cada dispositivo.\nOjo: se ve mientras la escribes.',
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const pw = r1.getResponseText();
  if (pw.length < 8) { ui.alert('La contraseña tiene que tener al menos 8 caracteres. No se ha cambiado nada.'); return; }
  if (pw.length > 200) { ui.alert('Demasiado larga. No se ha cambiado nada.'); return; }
  const r2 = ui.prompt('Repite la contraseña', 'Escríbela otra vez para comprobar.', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (r2.getResponseText() !== pw) { ui.alert('No coinciden. No se ha cambiado nada.'); return; }

  const sal = Utilities.getUuid() + Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PW', JSON.stringify({ s: sal, h: derivar_(pw, sal, ITERACIONES), i: ITERACIONES }));
  if (!props.getProperty('REV')) props.setProperty('REV', '0');
  const n = borrarSesiones_();
  CacheService.getScriptCache().remove('fallos');
  ui.alert('Contraseña guardada',
    'Listo. ' + (n ? 'Se han cerrado ' + n + ' sesiones abiertas: tendrás que volver a entrar en tus dispositivos.' : 'Ya puedes entrar desde la app.'),
    ui.ButtonSet.OK);
}

function verSesiones() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const lineas = [];
  for (const k in props) {
    if (k.indexOf('S_') !== 0) continue;
    try {
      const s = JSON.parse(props[k]);
      lineas.push('• ' + s.d + ' · último uso ' + Utilities.formatDate(new Date(s.u), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
    } catch (e) {}
  }
  SpreadsheetApp.getUi().alert('Dispositivos conectados', lineas.length ? lineas.join('\n') : 'Ninguno.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function cerrarSesiones() {
  const n = borrarSesiones_();
  SpreadsheetApp.getUi().alert('Se han cerrado ' + n + ' sesiones. Cada dispositivo tendrá que volver a poner la contraseña.');
}

// Quita las filas de elementos borrados hace más de 30 días.
function limpiarBorrados() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { SpreadsheetApp.getUi().alert('La hoja está ocupada. Prueba en un momento.'); return; }
  let n = 0;
  try {
    const sh = hoja_();
    const vals = sh.getDataRange().getValues();
    const limite = Date.now() - 30 * 864e5;
    for (let r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][4]) === '1' && Number(vals[r][3]) < limite) { sh.deleteRow(r + 1); n++; }
    }
    if (n) {
      // Los dispositivos que llevan mucho sin conectarse harán una descarga completa.
      const props = PropertiesService.getScriptProperties();
      const rev = Number(props.getProperty('REV') || 0) + 1;
      props.setProperties({ REV: String(rev), COMPLETO_DESDE: String(rev) });
    }
  } finally { lock.releaseLock(); }
  SpreadsheetApp.getUi().alert(n ? 'Se han quitado ' + n + ' filas antiguas.' : 'No había nada que limpiar.');
}

/* ================= aplicación web ================= */

function doGet() {
  return ContentService.createTextOutput('Mi espacio: el servidor funciona. Esta dirección se usa desde la app.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false, error: 'formato' }); }
  if (!req || typeof req !== 'object') return salida_({ ok: false, error: 'formato' });
  try {
    if (req.op === 'ping') return salida_({ ok: true, now: Date.now(), configurado: !!PropertiesService.getScriptProperties().getProperty('PW') });
    if (req.op === 'login') return salida_(login_(req));
    const ses = sesion_(req.token);
    if (!ses) return salida_({ ok: false, error: 'auth' });
    if (req.op === 'pull') return salida_(pull_(req));
    if (req.op === 'push') return salida_(push_(req));
    if (req.op === 'logout') { PropertiesService.getScriptProperties().deleteProperty(ses.clave); return salida_({ ok: true }); }
    return salida_({ ok: false, error: 'op' });
  } catch (err) {
    return salida_({ ok: false, error: 'servidor', msg: String(err && err.message || err).slice(0, 200) });
  }
}

/* ================= contraseña y sesiones ================= */

function login_(req) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PW');
  if (!raw) return { ok: false, error: 'sin_contrasena' };
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('fallos') || 0);
  if (fallos >= MAX_FALLOS) return { ok: false, error: 'bloqueado', min: BLOQUEO_MIN };
  const pw = typeof req.password === 'string' ? req.password.slice(0, 200) : '';
  const guardada = JSON.parse(raw);
  if (!pw || !iguales_(derivar_(pw, guardada.s, guardada.i), guardada.h)) {
    const n = fallos + 1;
    cache.put('fallos', String(n), BLOQUEO_MIN * 60);
    Utilities.sleep(1200);
    return n >= MAX_FALLOS ? { ok: false, error: 'bloqueado', min: BLOQUEO_MIN } : { ok: false, error: 'contrasena', quedan: MAX_FALLOS - n };
  }
  cache.remove('fallos');
  limpiarSesiones_();
  const token = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now()));
  const ahora = Date.now();
  const disp = (typeof req.device === 'string' ? req.device : '').replace(/[^\wÀ-ÿ ·.,()\-]/g, '').slice(0, 40) || 'Dispositivo';
  props.setProperty('S_' + huella_(token), JSON.stringify({ d: disp, c: ahora, u: ahora }));
  return { ok: true, token: token, now: ahora };
}

function sesion_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) return null;
  const props = PropertiesService.getScriptProperties();
  const clave = 'S_' + huella_(token);
  const raw = props.getProperty(clave);
  if (!raw) return null;
  let s;
  try { s = JSON.parse(raw); } catch (e) { return null; }
  const ahora = Date.now();
  if (ahora - s.u > SESION_DIAS * 864e5) { props.deleteProperty(clave); return null; }
  if (ahora - s.u > 12 * 3600e3) { s.u = ahora; props.setProperty(clave, JSON.stringify(s)); }
  s.clave = clave;
  return s;
}

function limpiarSesiones_() {
  const props = PropertiesService.getScriptProperties();
  const todas = props.getProperties();
  const ahora = Date.now();
  for (const k in todas) {
    if (k.indexOf('S_') !== 0) continue;
    try { if (ahora - JSON.parse(todas[k]).u > SESION_DIAS * 864e5) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  }
}

function borrarSesiones_() {
  const props = PropertiesService.getScriptProperties();
  const todas = props.getProperties();
  let n = 0;
  for (const k in todas) if (k.indexOf('S_') === 0) { props.deleteProperty(k); n++; }
  return n;
}

function derivar_(pw, sal, iter) {
  const pwB = Utilities.newBlob(pw).getBytes();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + pw, Utilities.Charset.UTF_8);
  for (let i = 0; i < iter; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(pwB));
  return hex_(h);
}
function huella_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }
function hex_(bytes) { return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join(''); }
function iguales_(a, b) {
  a = String(a); b = String(b);
  let d = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}

/* ================= datos ================= */

// Devuelve lo que ha cambiado desde la revisión «since» (o todo, si hace falta).
function pull_(req) {
  const props = PropertiesService.getScriptProperties();
  const rev = Number(props.getProperty('REV') || 0);
  const desde = Number(props.getProperty('COMPLETO_DESDE') || 0);
  const since = Math.max(0, Math.floor(Number(req.since) || 0));
  if (since > 0 && since === rev) return { ok: true, rev: rev, same: true, now: Date.now() };
  const completo = !(since > 0 && since < rev && since >= desde);
  const items = [];
  for (const f of leer_(hoja_())) {
    if (completo ? f.del : f.rev <= since) continue;
    const it = { k: f.k, v: f.ver, d: f.del ? 1 : 0 };
    if (!f.del) { try { it.j = JSON.parse(f.json); } catch (e) { continue; } }
    items.push(it);
  }
  return { ok: true, rev: rev, full: completo, items: items, now: Date.now() };
}

// Guarda un lote de cambios. Cada cambio lleva la versión que vio la app («b»):
// si otro dispositivo lo cambió antes, se rechaza y se devuelve lo que hay ahora.
function push_(req) {
  const cambios = Array.isArray(req.changes) ? req.changes : [];
  if (cambios.length > MAX_CAMBIOS) return { ok: false, error: 'demasiados' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'ocupado' };
  try {
    const sh = hoja_();
    const props = PropertiesService.getScriptProperties();
    let rev = Number(props.getProperty('REV') || 0);
    const filas = leer_(sh);
    const mapa = {};
    for (const f of filas) mapa[f.k] = f;
    let ultima = Math.max(1, sh.getLastRow());
    const nuevas = [], cambiadas = [], res = [];
    const ahora = Date.now();

    for (const c of cambios) {
      const err = validar_(c);
      if (err) { res.push({ k: c && typeof c.k === 'string' ? c.k.slice(0, 90) : '', ok: false, error: err }); continue; }
      const actual = mapa[c.k];
      const json = c.del ? '' : JSON.stringify(c.j);
      if (actual) {
        if (c.del && actual.del) { res.push({ k: c.k, ok: true, v: actual.ver }); continue; }
        if (!c.del && !actual.del && actual.json === json) { res.push({ k: c.k, ok: true, v: actual.ver }); continue; }
        if ((Number(c.b) || 0) !== actual.ver) {
          const r = { k: c.k, ok: false, error: 'conflicto', v: actual.ver, d: actual.del ? 1 : 0 };
          if (!actual.del) { try { r.j = JSON.parse(actual.json); } catch (e) {} }
          res.push(r);
          continue;
        }
      } else if (c.del) { res.push({ k: c.k, ok: true, v: 0 }); continue; }

      rev++;
      const ver = actual ? actual.ver + 1 : 1;
      const editado = Math.min(Number(c.t) || ahora, ahora + 60000);
      const fila = construir_(c.k, rev, ver, editado, !!c.del, c.del ? (actual && actual.leer) : c.j, json);
      if (actual) {
        cambiadas.push({ row: actual.row, fila: fila });
        actual.ver = ver; actual.del = !!c.del; actual.json = json;
      } else {
        ultima++;
        nuevas.push(fila);
        mapa[c.k] = { k: c.k, row: ultima, ver: ver, del: !!c.del, json: json, rev: rev };
      }
      res.push({ k: c.k, ok: true, v: ver });
    }

    const ancho = Math.max(sh.getLastColumn(), COL_DATOS, ...nuevas.map(f => f.length), ...cambiadas.map(c => c.fila.length));
    if (ancho > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), ancho - sh.getMaxColumns());
    const rellenar = f => { while (f.length < ancho) f.push(''); return f; };
    for (const c of cambiadas) sh.getRange(c.row, 1, 1, ancho).setValues([rellenar(c.fila)]);
    if (nuevas.length) {
      const primera = ultima - nuevas.length + 1;
      if (ultima > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), ultima - sh.getMaxRows());
      sh.getRange(primera, 1, nuevas.length, ancho).setNumberFormat('@').setValues(nuevas.map(rellenar));
    }
    props.setProperty('REV', String(rev));
    SpreadsheetApp.flush();
    return { ok: true, rev: rev, res: res, now: Date.now() };
  } finally {
    lock.releaseLock();
  }
}

function validar_(c) {
  if (!c || typeof c !== 'object') return 'formato';
  if (typeof c.k !== 'string' || !CLAVE_RE.test(c.k)) return 'clave';
  if (!Number.isFinite(Number(c.b || 0)) || Number(c.b || 0) < 0) return 'version';
  if (c.del) return null;
  const j = c.j;
  if (!j || typeof j !== 'object' || Array.isArray(j)) return 'datos';
  if (JSON.stringify(j).length > MAX_JSON) return 'grande';
  if (c.k === 'app') {
    if (!j.items || typeof j.items !== 'object' || !Array.isArray(j.order)) return 'datos';
    if (j.order.length > 500) return 'grande';
    return null;
  }
  if (typeof j.folder !== 'string' || j.folder.length > 120) return 'datos';
  if (j.title !== undefined && (typeof j.title !== 'string' || j.title.length > 300)) return 'datos';
  if (j.body !== undefined && (typeof j.body !== 'string' || j.body.length > 300000)) return 'datos';
  for (const n of ['pos', 'created', 'updated', 'done', 'amt']) if (j[n] !== undefined && !Number.isFinite(j[n])) return 'datos';
  for (const s of ['cat', 'day', 'src', 'upto']) if (j[s] !== undefined && (typeof j[s] !== 'string' || j[s].length > 120)) return 'datos';
  return null;
}

function construir_(k, rev, ver, editado, del, j, json) {
  let carpeta = '', titulo = '';
  if (k === 'app') titulo = 'Pantalla de inicio';
  else if (j) {
    carpeta = String(j.folder || '');
    titulo = String(j.title || '').trim() || String(j.body || '').trim().split('\n')[0];
    if (j.amt) titulo = (titulo || 'Movimiento') + ' · ' + (j.amt / 100).toFixed(2) + ' €';
  }
  const fila = ['~' + k, rev, ver, editado, del ? '1' : '', legible_(carpeta), legible_(titulo)];
  for (let i = 0; i < json.length; i += TROZO) fila.push('~' + json.slice(i, i + TROZO));
  if (fila.length < COL_DATOS) fila.push('');
  return fila;
}

function leer_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  const vals = sh.getRange(2, 1, n - 1, Math.max(sh.getLastColumn(), COL_DATOS)).getValues();
  const out = [];
  for (let r = 0; r < vals.length; r++) {
    const fila = vals[r];
    const k = String(fila[0] || '').replace(/^~/, '');
    if (!k) continue;
    let json = '';
    for (let c = COL_DATOS - 1; c < fila.length; c++) {
      const v = String(fila[c] || '');
      if (!v) break;
      json += v.charAt(0) === '~' ? v.slice(1) : v;
    }
    out.push({
      k: k, row: r + 2, rev: Number(fila[1]) || 0, ver: Number(fila[2]) || 0,
      del: String(fila[4]) === '1', json: json,
      leer: { folder: String(fila[5] || '').replace(/^'/, ''), title: String(fila[6] || '').replace(/^'/, '') },
    });
  }
  return out;
}

function hoja_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, CABECERA.length).setValues([CABECERA]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.setColumnWidth(COL_DATOS, 300);
  }
  return sh;
}

// Columnas solo para leer: evita que Sheets tome el texto como una fórmula.
function legible_(s) {
  s = String(s || '').replace(/\s+/g, ' ').slice(0, 120);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
