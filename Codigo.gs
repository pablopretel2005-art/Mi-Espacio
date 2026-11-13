/**
 * Mi espacio · servidor (Google Apps Script)
 *
 * Va dentro de la hoja de Google: Extensiones > Apps Script.
 * Es lo único que lee y escribe la hoja. La app (index.html en GitHub Pages) le habla por internet.
 *
 * Todo se hace desde el menú «Mi espacio» de la hoja (no con el botón Ejecutar del editor):
 *   1. Preparar la hoja
 *   2. Poner o cambiar la contraseña
 *   3. Conectar Google Calendar (da permiso para LEER tu calendario: widget de hoy y exámenes)
 *
 * Atajos de Siri: la app crea una clave aparte (Configuración → Atajos de Siri) que solo sirve
 * para añadir gastos, hábitos, recordatorios y notas, y para preguntar por tu agenda (hoy, mañana,
 * próxima clase, próximo examen) y por lo que llevas gastado este mes.
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
    .addItem('3. Conectar Google Calendar', 'conectarCalendario')
    .addSeparator()
    .addItem('Ver dispositivos conectados', 'verSesiones')
    .addItem('Cerrar todas las sesiones', 'cerrarSesiones')
    .addItem('Limpiar borrados antiguos', 'limpiarBorrados')
    .addItem('Desactivar atajos de Siri', 'desactivarAtajos')
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

// Pide el permiso de Google Calendar (la primera vez sale la ventana de autorización)
// y comprueba qué calendarios ve la app.
function conectarCalendario() {
  const ui = SpreadsheetApp.getUi();
  const cals = calendarios_();
  const uni = buscarCalendario_(cals, 'Universidad');
  ui.alert('Google Calendar conectado',
    'La app puede leer ' + cals.length + ' calendarios:\n' + cals.map(c => '• ' + c.getName()).join('\n') + '\n\n' +
    (uni ? 'Para Universidad usaré «' + uni.getName() + '».' : 'No hay ninguno llamado «Universidad»: podrás elegir otro desde la app.') +
    '\n\nÚltimo paso: Implementar > Gestionar implementaciones > lápiz > Versión: Nueva versión > Implementar.',
    ui.ButtonSet.OK);
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
    if (req.op === 'atajo') { try { return texto_(atajo_(req)); } catch (err) { return texto_('Ha fallado el servidor de Mi espacio. Inténtalo en un momento.'); } }
    const ses = sesion_(req.token);
    if (!ses) return salida_({ ok: false, error: 'auth' });
    if (req.op === 'pull') return salida_(pull_(req));
    if (req.op === 'push') return salida_(push_(req));
    if (req.op === 'cal') return salida_(calendario_(req));
    if (req.op === 'atajos') return salida_(atajosClave_(req));
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
  for (const n of ['pos', 'created', 'updated', 'done', 'amt', 'dow', 'goal']) if (j[n] !== undefined && !Number.isFinite(j[n])) return 'datos';
  for (const s of ['cat', 'day', 'src', 'upto', 't', 'subj', 'room', 'hs', 'he', 'col']) if (j[s] !== undefined && (typeof j[s] !== 'string' || j[s].length > 120)) return 'datos';
  for (const s of ['mail', 'ph', 'hours', 'url']) if (j[s] !== undefined && (typeof j[s] !== 'string' || j[s].length > 2000)) return 'datos';
  if (j.hd !== undefined && (typeof j.hd !== 'string' || j.hd.length > 40000 || (j.hd && !/^\d{8}(,\d{8})*$/.test(j.hd)))) return 'datos';   // días de un hábito
  return null;
}

/* ================= Google Calendar (solo lectura) ================= */

// Eventos de hoy y los próximos días de todos tus calendarios visibles, y lo próximo de los
// calendarios de la universidad. Se guarda 2 minutos en caché para no leer el calendario en cada consulta.
function calendario_(req) {
  const dias = Math.max(1, Math.min(31, Math.floor(Number(req.dias) || 7)));
  const atras = Math.max(0, Math.min(7, Math.floor(Number(req.atras) || 0)));   // días hacia atrás (para ver la semana entera en el horario)
  const unis = (Array.isArray(req.unis) ? req.unis : []).filter(x => typeof x === 'string' && x.trim()).map(x => x.trim().slice(0, 100)).slice(0, 5);
  const tz = Session.getScriptTimeZone();
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const cache = CacheService.getScriptCache();
  const clave = 'cal_' + huella_(JSON.stringify([hoy, dias, atras, unis])).slice(0, 40);
  const guardado = cache.get(clave);
  if (guardado) { try { return JSON.parse(guardado); } catch (e) {} }

  let cals;
  try { cals = calendarios_(); }
  catch (e) { return { ok: false, error: 'calendario', msg: 'Falta el permiso de Google Calendar. En la hoja: menú Mi espacio > 3. Conectar Google Calendar.' }; }
  const inicio = Utilities.parseDate(hoy + ' 00:00', tz, 'yyyy-MM-dd HH:mm');
  const fin = new Date(inicio.getTime() + dias * 864e5 + 3600e3);
  const desde = new Date(inicio.getTime() - atras * 864e5);
  const ev = [];
  for (const c of cals) {
    try { for (const e of c.getEvents(desde, fin)) { const x = evento_(e, c, false); if (x) ev.push(x); } } catch (e) {}
    if (ev.length > 600) break;
  }
  ev.sort(ordenEventos_);
  const uni = {};
  const finUni = new Date(inicio.getTime() + 366 * 864e5);
  for (const nombre of unis) {
    const c = buscarCalendario_(cals, nombre);
    if (!c) { uni[nombre] = { found: false }; continue; }
    const lista = [];
    try { for (const e of c.getEvents(inicio, finUni)) { const x = evento_(e, c, true); if (x) lista.push(x); if (lista.length >= 80) break; } } catch (e) {}
    lista.sort(ordenEventos_);
    uni[nombre] = { found: true, name: c.getName(), col: c.getColor(), ev: lista };
  }
  const res = { ok: true, tz: tz, hoy: hoy, dias: dias, atras: atras, ev: ev, uni: uni, cals: cals.map(c => c.getName()), now: Date.now() };
  const txt = JSON.stringify(res);
  if (txt.length < 90000) cache.put(clave, txt, 120);
  return res;
}

function calendarios_() {
  return CalendarApp.getAllCalendars().filter(c => { try { return !c.isHidden(); } catch (e) { return true; } });
}

function sinAcentos_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); }
function buscarCalendario_(cals, nombre) {
  const n = sinAcentos_(nombre);
  return cals.find(c => sinAcentos_(c.getName()) === n) || cals.find(c => sinAcentos_(c.getName()).indexOf(n) >= 0) || null;
}

// all: día completo (s y e son 'AAAA-MM-DD', e no incluido); si no, s y e son milisegundos.
function evento_(e, c, conDetalle) {
  try {
    if (e.getMyStatus && e.getMyStatus() === CalendarApp.GuestStatus.NO) return null;
  } catch (x) {}
  const tz = c.getTimeZone() || Session.getScriptTimeZone();
  const o = { t: String(e.getTitle() || '').slice(0, 200), cal: c.getName(), col: c.getColor() };
  if (e.isAllDayEvent()) {
    o.all = 1;
    o.s = Utilities.formatDate(e.getAllDayStartDate(), tz, 'yyyy-MM-dd');
    o.e = Utilities.formatDate(e.getAllDayEndDate(), tz, 'yyyy-MM-dd');
  } else {
    o.s = e.getStartTime().getTime();
    o.e = e.getEndTime().getTime();
  }
  // clave para ordenar (no se envía a la app)
  Object.defineProperty(o, '_k', { value: o.all ? e.getAllDayStartDate().getTime() : o.s, enumerable: false });
  const loc = String(e.getLocation() || '').trim();
  if (loc) o.loc = loc.slice(0, 160);
  if (conDetalle) {
    const d = String(e.getDescription() || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').trim();
    if (d) o.desc = d.slice(0, 400);
  }
  return o;
}
function ordenEventos_(a, b) {
  return a._k - b._k || (b.all ? 1 : 0) - (a.all ? 1 : 0);
}

/* ================= atajos de Siri ================= */
// Los atajos de la app Atajos del iPhone llaman aquí con { op: 'atajo', clave, accion, … }.
// Usan una clave propia (no la contraseña): solo permite AÑADIR cosas y leer la agenda de hoy.
// La clave la crea la app (Configuración → Atajos de Siri); aquí se guarda solo su huella.
// La respuesta es texto normal, para que el atajo la enseñe o Siri la diga en voz alta.

function atajosClave_(req) {
  const props = PropertiesService.getScriptProperties();
  if (req.accion === 'crear') {
    const clave = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now())).slice(0, 32);
    props.setProperty('ATAJO', huella_('atajo:' + clave));
    return { ok: true, activo: true, clave: clave };
  }
  if (req.accion === 'quitar') { props.deleteProperty('ATAJO'); return { ok: true, activo: false }; }
  return { ok: true, activo: !!props.getProperty('ATAJO') };
}

function desactivarAtajos() {
  PropertiesService.getScriptProperties().deleteProperty('ATAJO');
  SpreadsheetApp.getUi().alert('Atajos de Siri desactivados. Para volver a usarlos, crea una clave nueva en la app.');
}

function texto_(t) { return ContentService.createTextOutput(String(t)).setMimeType(ContentService.MimeType.TEXT); }

function atajo_(crudo) {
  // Los nombres de los campos se aceptan con tilde, mayúsculas o espacios («Acción», «Importe »…),
  // que es como a veces los escribe el teclado del iPhone.
  const req = {};
  for (const k in crudo) { const n = sinAcentos_(k).replace(/[^a-z]/g, ''); if (n && !(n in req)) req[n] = crudo[k]; }
  const guardada = PropertiesService.getScriptProperties().getProperty('ATAJO');
  if (!guardada) return 'Los atajos están desactivados. Actívalos en la app: Configuración, Atajos de Siri.';
  const clave = typeof req.clave === 'string' ? req.clave.trim().toLowerCase() : '';
  if (!/^[0-9a-f]{32}$/.test(clave) || !iguales_(huella_('atajo:' + clave), guardada)) {
    Utilities.sleep(800);
    return 'La clave del atajo no es correcta. Cópiala otra vez desde la app: Configuración, Atajos de Siri.';
  }
  const accion = sinAcentos_(req.accion);
  const d = estado_();
  if (accion === 'gasto') return atajoGasto_(req, d);
  if (accion === 'habito') return atajoHabito_(req, d);
  if (accion === 'recordatorio') return atajoRecordatorio_(req, d);
  if (accion === 'nota') return atajoNota_(req, d);
  if (accion === 'hoy') return atajoHoy_(d);
  if (accion === 'manana') return atajoManana_(d);
  if (accion === 'clase') return atajoClase_(d);
  if (accion === 'examen') return atajoExamen_(d);
  if (accion === 'mes' || accion === 'gastos') return atajoMes_(req, d);
  if (!accion) return 'Al atajo le falta el campo «accion» (sin tilde). Revisa los pasos en la app: Configuración, Atajos de Siri.';
  return 'No conozco la acción «' + String(req.accion || '').slice(0, 30) + '». Usa gasto, habito, recordatorio, nota, hoy, manana, clase, examen o mes.';
}

// Inicio y entradas tal como están ahora en la hoja.
function estado_() {
  let home = { items: {}, order: [] };
  const entradas = [];
  for (const f of leer_(hoja_())) {
    if (f.del) continue;
    let j; try { j = JSON.parse(f.json); } catch (e) { continue; }
    if (f.k === 'app') home = j;
    else entradas.push({ id: f.k.slice(2), ver: f.ver, j: j });
  }
  // si la app aún no ha guardado el inicio, se usan las secciones que trae de serie
  if (!home.items || !Object.keys(home.items).length) home = { order: [], items: { notas: { kind: 'notes', name: 'Notas' }, finanzas: { kind: 'money', name: 'Finanzas' }, recordatorios: { kind: 'list', name: 'Recordatorios' }, universidad: { kind: 'uni', name: 'Universidad' } } };
  return { home: home, items: Object.keys(home.items || {}).map(id => Object.assign({ id: id }, home.items[id])), entradas: entradas };
}
// Carpeta de un tipo: la que se llame como «carpeta» (si se dice), si no la preferida o la primera.
function carpeta_(d, kind, nombre, preferida) {
  const deTipo = d.items.filter(i => i.kind === kind);
  if (nombre) { const n = sinAcentos_(nombre); const x = deTipo.find(i => sinAcentos_(i.name) === n) || deTipo.find(i => sinAcentos_(i.name).indexOf(n) >= 0); if (x) return x; }
  return deTipo.find(i => i.id === preferida) || deTipo[0] || null;
}
function nuevoId_() { return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function guardar_(cambios) {
  const r = push_({ changes: cambios });
  if (!r.ok) return false;
  return r.res.every(x => x.ok);
}
function hoy_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function euros_(c) { return (c / 100).toFixed(2).replace('.', ',').replace(/,00$/, '') + ' €'; }
function limpio_(s, max) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, max); }

const CATEGORIAS_ = [
  ['super', /mercadona|lidl|carrefour|dia\b|aldi|super|compra/],
  ['comida', /cafe|bar\b|restaurante|comida|cena|desayuno|almuerzo|tapas|burger|pizza|kebab|mcdonald|telepizza/],
  ['transporte', /bus|metro|tranvia|tren|renfe|taxi|uber|cabify|gasolina|parking|bici/],
  ['suscripciones', /netflix|spotify|hbo|disney|amazon prime|icloud|youtube|chatgpt|claude/],
  ['salud', /farmacia|medico|dentista|fisio/],
  ['estudios', /libro|apuntes|fotocopia|universidad|matricula|curso|academia/],
  ['ocio', /cine|concierto|entrada|fiesta|discoteca|teatro|museo|juego/],
  ['regalos', /regalo|cumple/],
  ['compras', /ropa|zara|zapatillas|amazon|tienda/],
  ['viajes', /hotel|vuelo|avion|ryanair|vueling|airbnb|viaje/],
];
function atajoGasto_(req, d) {
  const f = carpeta_(d, 'money', req.carpeta, 'finanzas');
  if (!f) return 'No tienes ninguna sección de Finanzas. Créala en la app con el botón +.';
  let amt;
  if (typeof req.importe === 'number') amt = Math.round(req.importe * 100);
  else {
    let s = String(req.importe || '').replace(/[^\d.,]/g, '');
    if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    amt = Math.round(parseFloat(s) * 100);
  }
  if (!(amt > 0 && amt <= 1e9)) return 'No he entendido el importe. Dímelo con números, por ejemplo 4,50.';
  let concepto = limpio_(req.concepto, 60);
  concepto = concepto.charAt(0).toUpperCase() + concepto.slice(1);
  const n = sinAcentos_(concepto), hoy = hoy_();
  // categoría: la de la última vez que apuntaste lo mismo; si no, por palabras; si no, «otros»
  let cat = '';
  const previos = d.entradas.filter(e => e.j.folder === f.id && !e.j.rep && !e.j.inc && n && sinAcentos_(e.j.title) === n).sort((a, b) => (b.j.created || 0) - (a.j.created || 0));
  if (previos.length && previos[0].j.cat) cat = previos[0].j.cat;
  if (!cat) for (const c of CATEGORIAS_) if (c[1].test(n)) { cat = c[0]; break; }
  if (!cat) cat = 'otros';
  const ahora = Date.now();
  const j = { folder: f.id, title: concepto, body: '', pos: 0, created: ahora, updated: ahora, amt: amt, cat: cat, day: hoy };
  if (!guardar_([{ k: 'e:' + nuevoId_(), b: 0, t: ahora, j: j }])) return 'No se ha podido guardar. Inténtalo otra vez.';
  const mes = hoy.slice(0, 7);
  let total = amt;
  for (const e of d.entradas) if (e.j.folder === f.id && !e.j.rep && !e.j.inc && e.j.day && e.j.day.slice(0, 7) === mes && e.j.day <= hoy) total += e.j.amt || 0;
  let t = 'Apuntado: ' + euros_(amt) + (concepto ? ' en ' + concepto : '') + '. Este mes llevas ' + euros_(total) + '.';
  if (f.budget) { const q = f.budget - total; t += q >= 0 ? ' Te quedan ' + euros_(q) + ' del presupuesto.' : ' Te has pasado del presupuesto en ' + euros_(-q) + '.'; }
  return t;
}

function diaClave_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd'); }
function atajoHabito_(req, d) {
  const carpetas = d.items.filter(i => i.kind === 'habits').map(i => i.id);
  const habitos = d.entradas.filter(e => carpetas.indexOf(e.j.folder) >= 0);
  if (!habitos.length) return 'No tienes hábitos todavía. Créalos en la app: botón +, Hábitos.';
  const n = sinAcentos_(req.nombre);
  if (!n) return 'Dime qué hábito quieres marcar.';
  const h = habitos.find(e => sinAcentos_(e.j.title) === n) || habitos.find(e => sinAcentos_(e.j.title).indexOf(n) >= 0) || habitos.find(e => n.indexOf(sinAcentos_(e.j.title)) >= 0);
  if (!h) return 'No encuentro el hábito «' + limpio_(req.nombre, 40) + '». Tus hábitos son: ' + habitos.map(e => e.j.title).join(', ') + '.';
  const hoy = diaClave_(new Date());
  const dias = h.j.hd ? h.j.hd.split(',') : [];
  const ya = dias.indexOf(hoy) >= 0;
  if (!ya) {
    dias.push(hoy); dias.sort();
    const j = Object.assign({}, h.j, { hd: dias.join(','), updated: Date.now() });
    if (!guardar_([{ k: 'e:' + h.id, b: h.ver, t: Date.now(), j: j }])) return 'No se ha podido guardar. Inténtalo otra vez.';
  }
  const set = {}; dias.forEach(x => { set[x] = 1; });
  const goal = h.j.goal || 7;
  let t = (ya ? h.j.title + ' ya estaba marcado hoy.' : h.j.title + ' marcado.');
  if (goal >= 7) {
    let racha = 0; const dd = new Date();
    while (set[diaClave_(dd)]) { racha++; dd.setDate(dd.getDate() - 1); }
    t += racha > 1 ? ' Llevas ' + racha + ' días seguidos.' : ' Primer día de la racha.';
  } else {
    const dd = new Date(); const dow = (Number(Utilities.formatDate(dd, Session.getScriptTimeZone(), 'u')) + 6) % 7;   // 0 = lunes
    let semana = 0; for (let i = 0; i <= dow; i++) { const x = new Date(dd.getTime() - i * 864e5); if (set[diaClave_(x)]) semana++; }
    t += ' Esta semana: ' + semana + ' de ' + goal + '.' + (semana >= goal ? ' Objetivo cumplido.' : '');
  }
  return t;
}

function atajoRecordatorio_(req, d) {
  const f = carpeta_(d, 'list', req.carpeta, 'recordatorios');
  if (!f) return 'No tienes ninguna lista. Créala en la app con el botón +.';
  const lineas = String(req.texto || '').split(/\n|,\s*y\s+|\s+y\s+luego\s+/).map(s => limpio_(s, 300)).filter(Boolean).slice(0, 20);
  if (!lineas.length) return 'Dime qué quieres añadir.';
  const deLista = d.entradas.filter(e => e.j.folder === f.id);
  let min = deLista.reduce((m, e) => Math.min(m, Number(e.j.pos) || 0), Infinity);
  if (!isFinite(min)) min = 1000;
  const ahora = Date.now();
  const cambios = lineas.map((x, i) => ({ k: 'e:' + nuevoId_() + i, b: 0, t: ahora, j: { folder: f.id, title: x, body: '', pos: min - 1000 * (lineas.length - i), created: ahora, updated: ahora } }));
  if (!guardar_(cambios)) return 'No se ha podido guardar. Inténtalo otra vez.';
  const pendientes = deLista.filter(e => !e.j.done).length + lineas.length;
  return (lineas.length === 1 ? 'Añadido a ' + f.name + '.' : 'Añadidas ' + lineas.length + ' cosas a ' + f.name + '.') + ' Tienes ' + pendientes + (pendientes === 1 ? ' pendiente.' : ' pendientes.');
}

function atajoNota_(req, d) {
  const f = carpeta_(d, 'notes', req.carpeta, 'notas');
  if (!f) return 'No tienes ninguna carpeta de notas. Créala en la app con el botón +.';
  const texto = String(req.texto || '').trim().slice(0, 20000);
  if (!texto) return 'La nota está vacía.';
  const ahora = Date.now();
  const j = { folder: f.id, title: limpio_(req.titulo, 120), body: texto, pos: 0, created: ahora, updated: ahora };
  if (!guardar_([{ k: 'e:' + nuevoId_(), b: 0, t: ahora, j: j }])) return 'No se ha podido guardar. Inténtalo otra vez.';
  return 'Nota guardada en ' + f.name + '.';
}

// Lo que queda hoy: eventos de Google Calendar y clases del horario, por orden.
function atajoHoy_(d) {
  const tz = Session.getScriptTimeZone(), ahora = Date.now(), hoy = hoy_();
  const hora = ms => Utilities.formatDate(new Date(ms), tz, 'H:mm');
  const cosas = [];
  let sinCal = false;
  const cal = calendario_({ dias: 1, atras: 0, unis: [] });
  if (cal.ok) {
    for (const e of cal.ev) {
      if (e.all) { if (e.s <= hoy && hoy < e.e) cosas.push({ s: 0, t: e.t + ' (todo el día)' }); continue; }
      if (e.e > ahora && Utilities.formatDate(new Date(e.s), tz, 'yyyy-MM-dd') <= hoy) cosas.push({ s: e.s, t: (e.s > ahora ? 'a las ' + hora(e.s) + ', ' : 'ahora, hasta las ' + hora(e.e) + ', ') + e.t + (e.loc ? ' en ' + e.loc : '') });
    }
  } else sinCal = true;
  const dow = Number(Utilities.formatDate(new Date(), tz, 'u'));   // 1 = lunes
  const unis = d.items.filter(i => i.kind === 'uni').map(i => i.id);
  for (const e of d.entradas) {
    const j = e.j;
    if (j.t !== 'class' || j.dow !== dow || unis.indexOf(j.folder) < 0 || !/^\d\d:\d\d$/.test(j.hs || '') || !/^\d\d:\d\d$/.test(j.he || '')) continue;
    const s = Utilities.parseDate(hoy + ' ' + j.hs, tz, 'yyyy-MM-dd HH:mm').getTime(), fin = Utilities.parseDate(hoy + ' ' + j.he, tz, 'yyyy-MM-dd HH:mm').getTime();
    if (fin > ahora) cosas.push({ s: s, t: (s > ahora ? 'a las ' + hora(s) + ', ' : 'ahora, hasta las ' + hora(fin) + ', ') + 'clase de ' + j.title + (j.room ? ' en ' + j.room : '') });
  }
  cosas.sort((a, b) => a.s - b.s);
  let t;
  if (!cosas.length) t = 'No te queda nada más hoy.';
  else if (cosas.length === 1) t = 'Hoy te queda una cosa: ' + cosas[0].t + '.';
  else t = 'Hoy te quedan ' + cosas.length + ' cosas: ' + cosas.slice(0, 8).map(c => c.t).join('; ') + '.';
  if (sinCal) t += ' (No he podido leer tu Google Calendar: falta darle permiso desde la hoja.)';
  return t;
}

const DIAS_ = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];
const MESES_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
// Fecha 'AAAA-MM-DD' de dentro de n días, día de la semana (1 = lunes) y días entre dos fechas.
function diaMas_(n) {
  const tz = Session.getScriptTimeZone();
  const medio = Utilities.parseDate(hoy_() + ' 12:00', tz, 'yyyy-MM-dd HH:mm');
  return Utilities.formatDate(new Date(medio.getTime() + n * 864e5), tz, 'yyyy-MM-dd');
}
function dowDe_(dia) { const p = dia.split('-').map(Number); return ((new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay() + 6) % 7) + 1; }
function diasEntre_(a, b) { const x = a.split('-').map(Number), y = b.split('-').map(Number); return Math.round((Date.UTC(y[0], y[1] - 1, y[2]) - Date.UTC(x[0], x[1] - 1, x[2])) / 864e5); }
function fechaLarga_(dia) { const p = dia.split('-').map(Number); return 'el ' + DIAS_[dowDe_(dia) - 1] + ' ' + p[2] + ' de ' + MESES_[p[1] - 1]; }
function clases_(d) {
  const unis = d.items.filter(i => i.kind === 'uni').map(i => i.id);
  return d.entradas.map(e => e.j).filter(j => j.t === 'class' && unis.indexOf(j.folder) >= 0 && /^\d\d:\d\d$/.test(j.hs || '') && /^\d\d:\d\d$/.test(j.he || ''));
}

// Todo lo de mañana: eventos de Google Calendar y clases del horario, por orden.
function atajoManana_(d) {
  const tz = Session.getScriptTimeZone(), dia = diaMas_(1);
  const hora = ms => Utilities.formatDate(new Date(ms), tz, 'H:mm');
  const cosas = [];
  let sinCal = false;
  const cal = calendario_({ dias: 2, atras: 0, unis: [] });
  if (cal.ok) {
    for (const e of cal.ev) {
      if (e.all) { if (e.s <= dia && dia < e.e) cosas.push({ s: 0, t: e.t + ' (todo el día)' }); continue; }
      if (Utilities.formatDate(new Date(e.s), tz, 'yyyy-MM-dd') === dia) cosas.push({ s: e.s, t: 'a las ' + hora(e.s) + ', ' + e.t + (e.loc ? ' en ' + e.loc : '') });
    }
  } else sinCal = true;
  const dow = dowDe_(dia);
  for (const j of clases_(d)) {
    if (j.dow !== dow) continue;
    const s = Utilities.parseDate(dia + ' ' + j.hs, tz, 'yyyy-MM-dd HH:mm').getTime();
    cosas.push({ s: s, t: 'a las ' + hora(s) + ', clase de ' + j.title + (j.room ? ' en ' + j.room : '') });
  }
  cosas.sort((a, b) => a.s - b.s);
  let t;
  if (!cosas.length) t = 'Mañana no tienes nada apuntado.';
  else if (cosas.length === 1) t = 'Mañana tienes una cosa: ' + cosas[0].t + '.';
  else t = 'Mañana tienes ' + cosas.length + ' cosas: ' + cosas.slice(0, 8).map(c => c.t).join('; ') + '.';
  if (sinCal) t += ' (No he podido leer tu Google Calendar: falta darle permiso desde la hoja.)';
  return t;
}

// La clase que tienes ahora y la siguiente del horario (mira hasta una semana).
function atajoClase_(d) {
  const lista = clases_(d);
  if (!lista.length) return 'No tienes clases en tu horario. Añádelas en la app: Universidad, Horario.';
  const tz = Session.getScriptTimeZone(), ahora = Date.now();
  const hora = ms => Utilities.formatDate(new Date(ms), tz, 'H:mm');
  let actual = null, sig = null;
  for (let n = 0; n <= 7 && !sig; n++) {
    const dia = diaMas_(n), dow = dowDe_(dia);
    const delDia = lista.filter(j => j.dow === dow).map(j => ({ j: j, s: Utilities.parseDate(dia + ' ' + j.hs, tz, 'yyyy-MM-dd HH:mm').getTime(), e: Utilities.parseDate(dia + ' ' + j.he, tz, 'yyyy-MM-dd HH:mm').getTime(), n: n, dia: dia })).sort((a, b) => a.s - b.s);
    for (const c of delDia) {
      if (c.s <= ahora && c.e > ahora) { if (!actual) actual = c; continue; }
      if (c.s > ahora) { sig = c; break; }
    }
  }
  const cuando = c => (c.n === 0 ? 'hoy' : c.n === 1 ? 'mañana' : 'el ' + DIAS_[dowDe_(c.dia) - 1]) + ' a las ' + hora(c.s);
  const sitio = c => (c.j.room ? ', en ' + c.j.room : '');
  let t = '';
  if (actual) t = 'Ahora tienes ' + actual.j.title + sitio(actual) + ', hasta las ' + hora(actual.e) + '. ';
  if (sig) t += (actual ? 'Después, ' : 'Tu próxima clase es ') + sig.j.title + ', ' + cuando(sig) + sitio(sig) + '.';
  else if (!actual) t = 'No tienes más clases esta semana.';
  return t.trim();
}

// Lo siguiente del calendario de la universidad (exámenes y entregas).
function atajoExamen_(d) {
  const uni = d.items.filter(i => i.kind === 'uni')[0];
  if (!uni) return 'No tienes la sección Universidad. Créala en la app con el botón +.';
  const nombre = uni.calName || 'Universidad';
  const cal = calendario_({ dias: 1, atras: 0, unis: [nombre] });
  if (!cal.ok) return 'No he podido leer tu Google Calendar: falta darle permiso desde la hoja.';
  const u = cal.uni[nombre];
  if (!u || !u.found) return 'No encuentro el calendario «' + nombre + '» en tu Google Calendar.';
  const tz = Session.getScriptTimeZone(), hoy = hoy_();
  const diaDe = e => (e.all ? e.s : Utilities.formatDate(new Date(e.s), tz, 'yyyy-MM-dd'));
  const prox = u.ev.filter(e => diaDe(e) >= hoy);
  if (!prox.length) return 'No tienes nada próximo en el calendario «' + u.name + '».';
  const falta = e => { const n = diasEntre_(hoy, diaDe(e)); return n <= 0 ? 'hoy' : n === 1 ? 'mañana' : fechaLarga_(diaDe(e)) + ', dentro de ' + n + ' días'; };
  const hora = e => (e.all ? '' : ' a las ' + Utilities.formatDate(new Date(e.s), tz, 'H:mm'));
  let t = 'Lo siguiente es ' + prox[0].t + ', ' + falta(prox[0]) + hora(prox[0]) + '.';
  if (prox[1]) t += ' Después, ' + prox[1].t + ', ' + falta(prox[1]) + '.';
  return t;
}

// Lo que llevas gastado este mes y lo que queda del presupuesto.
function atajoMes_(req, d) {
  const f = carpeta_(d, 'money', req.carpeta, 'finanzas');
  if (!f) return 'No tienes ninguna sección de Finanzas. Créala en la app con el botón +.';
  const hoy = hoy_(), mes = hoy.slice(0, 7);
  let total = 0, n = 0;
  const porCat = {};
  for (const e of d.entradas) {
    const j = e.j;
    if (j.folder !== f.id || j.rep || j.inc || !j.day || j.day.slice(0, 7) !== mes || j.day > hoy) continue;
    total += j.amt || 0; n++;
    porCat[j.cat || 'otros'] = (porCat[j.cat || 'otros'] || 0) + (j.amt || 0);
  }
  if (!n) return 'Este mes todavía no has apuntado ningún gasto.';
  let t = 'Este mes llevas ' + euros_(total) + ' en ' + n + (n === 1 ? ' gasto.' : ' gastos.');
  const mayor = Object.keys(porCat).sort((a, b) => porCat[b] - porCat[a])[0];
  if (n > 1 && mayor) t += ' Donde más, en ' + mayor.replace('super', 'súper') + ': ' + euros_(porCat[mayor]) + '.';
  if (f.budget) { const q = f.budget - total; t += q >= 0 ? ' Te quedan ' + euros_(q) + ' del presupuesto.' : ' Te has pasado del presupuesto en ' + euros_(-q) + '.'; }
  return t;
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
