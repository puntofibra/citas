/**
 * CITAS PRO — backend multi-negocio (Google Apps Script + Google Sheets)
 * API: POST JSON (text/plain) -> { action, token, ...datos }
 * Despliegue: Implementar > Nueva implementación > App web
 *   Ejecutar como: yo  |  Acceso: cualquier usuario
 * Zona horaria del proyecto: Europe/Madrid (appsscript.json)
 */

/* ============ CONFIG ============ */
const ADMIN_EMAIL = 'futurmovil.com@gmail.com';
const MYPOS_TAG = 'https://mypos.com/@futurmovil';
const PLANES = [                 // planes de suscripción (precio total en €)
  { meses: 1, dias: 30, precio: 15 },
  { meses: 3, dias: 90, precio: 36 },
  { meses: 6, dias: 180, precio: 60 },
  { meses: 12, dias: 365, precio: 100 }
];
const PRECIO_MES = PLANES[0].precio;
const DIAS_PRUEBA = 14;         // prueba gratuita al registrarse
const APP_NOMBRE = 'Citas Pro';
const SESION_HORAS = 24 * 14;   // duración de la sesión

const HOJAS = {
  Negocios: ['id','slug','nombre','email','salt','hash','estado','venceEn','color','logo','portada','telefono','direccion','whatsapp','descripcion','paso','antelacion','creado','resetCode','resetExp'],
  Servicios: ['negocioId','id','categoria','nombre','duracion','precio','activo'],
  Profesionales: ['negocioId','id','nombre','servicios','horario','activo'],
  Citas: ['id','negocioId','profId','servId','fecha','hora','dur','cliente','tel','email','notas','estado','creado','codigo'],
  Sesiones: ['token','negocioId','expira']
};
const COLS_TEXTO = ['fecha','hora','venceEn','creado','tel','whatsapp','telefono','codigo','resetCode'];

/* ============ SETUP (ejecutar una vez) ============ */
function setup() {
  const ss = SS();
  Object.keys(HOJAS).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    const h = HOJAS[n];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    h.forEach((k, i) => { if (COLS_TEXTO.indexOf(k) >= 0) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@'); });
  });
  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return 'Hojas listas';
}

/* ============ ENTRADA HTTP ============ */
function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.m) return manifiesto_(q.m, q.b);
  return salida({ ok: true, app: APP_NOMBRE });
}
/* Manifiesto PWA de cada negocio: así el cliente instala "la app de su peluquería" con su nombre */
function manifiesto_(slug, base) {
  base = String(base || '');
  if (!/^https:\/\/[a-z0-9.-]+\/[^?#]*\/$/i.test(base)) base = 'https://puntofibra.github.io/citas/';
  let n = null; try { n = negPorSlug(slug); } catch (err) { }
  const nombre = n ? String(n.nombre) : APP_NOMBRE;
  const start = base + (n ? '?n=' + encodeURIComponent(n.slug) : '');
  const color = n && /^#[0-9a-f]{6}$/i.test(n.color || '') ? n.color : '#2563eb';
  const m = { id: start, name: nombre, short_name: nombre.slice(0, 14), description: 'Reserva tu cita en ' + nombre,
    start_url: start, scope: base, display: 'standalone', orientation: 'portrait', background_color: '#070d1a', theme_color: color,
    icons: [{ src: base + 'icon-192.png', sizes: '192x192', type: 'image/png' }, { src: base + 'icon-512.png', sizes: '512x512', type: 'image/png' },
      { src: base + 'maskable-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' }] };
  return ContentService.createTextOutput(JSON.stringify(m)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  try {
    const p = JSON.parse((e.postData && e.postData.contents) || '{}');
    out = ruta(p);
    if (out && out.ok === undefined) out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return salida(out);
}
function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function ruta(p) {
  const a = String(p.action || '');
  const publicas = { negocio: negocioPublico, disp: dispPublica, dispRango: dispRango, reservar: reservar, cancelar: cancelarPublica,
    registro: registro, login: login, recuperar: recuperar, reset: resetPass, llaveSesion: llaveSesion, llaveEstado: llaveEstado };
  if (publicas[a]) return publicas[a](p);

  const s = sesion(p.token);
  if (!s) throw new Error('Sesión caducada. Entra de nuevo.');
  if (s.admin) {
    const adm = { yo: yoAdmin, adminNegocios: adminNegocios, adminDias: adminDias, adminEstado: adminEstado,
      totpStatus: totpStatus, totpStart: totpStart, totpConfirm: totpConfirm, totpDisable: totpDisable, logout: logout };
    if (!adm[a]) throw new Error('Acción no permitida');
    return adm[a](p, s);
  }
  const neg = { yo: yoNegocio, guardarNegocio: guardarNegocio, servGuardar: servGuardar, servBorrar: servBorrar,
    profGuardar: profGuardar, profBorrar: profBorrar, subirImagen: subirImagen, quitarImagen: quitarImagen, agenda: agenda, citaEstado: citaEstado, citaManual: citaManual,
    cambiarPass: cambiarPass, logout: logout };
  if (!neg[a]) throw new Error('Acción no permitida');
  return neg[a](p, s);
}

/* ============ UTILIDADES DE HOJA ============ */
const SHEET_ID = '1CIXvLIm-4OqxEOrj5VCow3yyWLLxYJwSzfMLprPdAmk';
let _ss = null; const _sh = {}, _tab = {};
function SS() { return _ss || (_ss = SpreadsheetApp.openById(SHEET_ID)); }
function hoja(n) { return _sh[n] || (_sh[n] = SS().getSheetByName(n)); }
/* Caché: cada tabla se guarda en CacheService y se invalida en cada escritura (mucho más rápido que leer la hoja). */
function cacheTabla_(n, v) { try { const s = JSON.stringify(v); if (s.length < 95000) CacheService.getScriptCache().put('t2_' + n, s, 600); else CacheService.getScriptCache().remove('t2_' + n); } catch (e) { } }
function tocar_(n) { delete _tab[n]; try { CacheService.getScriptCache().remove('t2_' + n); } catch (e) { } }
function norm(k, v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz(), k === 'hora' ? 'HH:mm' : 'yyyy-MM-dd');
  return v;
}
function tz() { return Session.getScriptTimeZone() || 'Europe/Madrid'; }
function tabla(n) {
  if (_tab[n]) return _tab[n].map(o => Object.assign({}, o));
  if (!_enLock) try { const c = CacheService.getScriptCache().get('t2_' + n); if (c) { _tab[n] = JSON.parse(c); return _tab[n].map(o => Object.assign({}, o)); } } catch (e) { }
  const out = leerHoja_(n);
  _tab[n] = out; cacheTabla_(n, out);
  return out.map(o => Object.assign({}, o));
}
function leerHoja_(n) {
  const sh = hoja(n);
  const v = sh.getDataRange().getValues();
  const h = v.shift();
  const out = [];
  v.forEach((r, i) => {
    if (r[0] === '' || r[0] === null) return;
    const o = { _r: i + 2 };
    h.forEach((k, j) => { o[k] = norm(k, r[j]); });
    out.push(o);
  });
  return out;
}
function fila(n, o) { return HOJAS[n].map(k => (o[k] === undefined ? '' : o[k])); }
function insertar(n, o) {
  const sh = hoja(n);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, HOJAS[n].length).setValues([fila(n, o)]);
  o._r = r; tocar_(n);
  return o;
}
function guardarFila(n, o) { hoja(n).getRange(o._r, 1, 1, HOJAS[n].length).setValues([fila(n, o)]); tocar_(n); }
function borrarFila(n, o) { hoja(n).deleteRow(o._r); tocar_(n); }
function uid() { return Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function hoyStr() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
function ahoraMin() { const d = new Date(); return +Utilities.formatDate(d, tz(), 'H') * 60 + +Utilities.formatDate(d, tz(), 'm'); }
function minutos(h) { const x = String(h).split(':'); return +x[0] * 60 + +x[1]; }
function hhmm(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }
function sumaDias(fecha, n) { const x = fecha.split('-'); const d = new Date(+x[0], +x[1] - 1, +x[2] + n); return Utilities.formatDate(d, tz(), 'yyyy-MM-dd'); }
function diaSemana(fecha) { const x = fecha.split('-'); return new Date(+x[0], +x[1] - 1, +x[2]).getDay(); }
function conLock(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { Object.keys(_tab).forEach(k => delete _tab[k]); _enLock = true; return fn(); } finally { _enLock = false; l.releaseLock(); } }
let _enLock = false;
function txt(v, max) { return String(v === undefined || v === null ? '' : v).trim().slice(0, max || 200); }
function esFecha(f) { return /^\d{4}-\d{2}-\d{2}$/.test(f); }
function esHora(h) { return /^\d{2}:\d{2}$/.test(h); }
function esUrl(u) { return u === '' || /^https:\/\/[^\s]+$/i.test(u); }
function limpiaTel(t) { return String(t || '').replace(/[^\d+]/g, '').slice(0, 16); }
function mayus(s) { return String(s || '').trim().toLowerCase().replace(/(^|\s)\S/g, m => m.toUpperCase()).slice(0, 80); }

/* ============ SEGURIDAD ============ */
function sha(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function props() { return PropertiesService.getScriptProperties(); }
function limite(clave, max, seg) {
  const c = CacheService.getScriptCache();
  const n = +(c.get(clave) || 0);
  if (n >= max) throw new Error('Demasiados intentos. Espera unos minutos.');
  c.put(clave, String(n + 1), seg);
}
function nuevaSesion(negId) {
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  const exp = Date.now() + SESION_HORAS * 3600000;
  insertar('Sesiones', { token: token, negocioId: negId, expira: exp });
  CacheService.getScriptCache().put('s_' + token, negId, 21600);
  return token;
}
function sesion(token) {
  token = txt(token, 100);
  if (!token) return null;
  const c = CacheService.getScriptCache();
  let id = c.get('s_' + token);
  if (!id) {
    const f = tabla('Sesiones').filter(x => x.token === token && +x.expira > Date.now())[0];
    if (!f) return null;
    id = f.negocioId;
    c.put('s_' + token, id, 21600);
  }
  return id === 'ADMIN' ? { admin: true, token: token } : { negocioId: id, token: token };
}
function logout(p) {
  const f = tabla('Sesiones').filter(x => x.token === p.token)[0];
  if (f) borrarFila('Sesiones', f);
  CacheService.getScriptCache().remove('s_' + p.token);
  return {};
}
function hashPass(salt, clienteHash) { return sha(salt + ':' + clienteHash); }
// El navegador envía ya SHA-256(email + contraseña); nunca llega la clave en claro.
function validaClienteHash(h) { if (!/^[a-f0-9]{64}$/.test(String(h))) throw new Error('Contraseña no válida'); return h; }

/* ============ REGISTRO / LOGIN / RECUPERAR ============ */
function slugDe(nombre, excluirId) {
  let base = String(nombre).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'negocio';
  const usados = tabla('Negocios').filter(n => n.id !== excluirId).map(n => n.slug);
  let s = base, i = 2;
  while (usados.indexOf(s) >= 0) s = base + '-' + i++;
  return s;
}
function horarioDefecto() {
  const m = [['09:00', '14:00'], ['16:00', '20:00']];
  return { '0': [], '1': m, '2': m, '3': m, '4': m, '5': m, '6': [['10:00', '14:00']] };
}
function registro(p) {
  const email = txt(p.email, 120).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('Correo no válido');
  if (email === ADMIN_EMAIL) throw new Error('Ese correo está reservado');
  const nombre = txt(p.nombre, 60);
  if (nombre.length < 2) throw new Error('Escribe el nombre de tu negocio');
  const ch = validaClienteHash(p.clave);
  limite('reg_' + email, 5, 3600);
  return conLock(() => {
    if (tabla('Negocios').some(n => String(n.email).toLowerCase() === email)) throw new Error('Ese correo ya tiene una cuenta');
    const salt = uid() + uid();
    const id = uid();
    const neg = insertar('Negocios', {
      id: id, slug: slugDe(nombre), nombre: nombre, email: email, salt: salt, hash: hashPass(salt, ch), estado: 'activo',
      venceEn: sumaDias(hoyStr(), DIAS_PRUEBA), color: '#c026d3', logo: '', portada: '', telefono: '', direccion: '', whatsapp: '',
      descripcion: '', paso: 30, antelacion: 60, creado: hoyStr(), resetCode: '', resetExp: ''
    });
    insertar('Profesionales', { negocioId: id, id: uid(), nombre: 'Equipo', servicios: '', horario: JSON.stringify(horarioDefecto()), activo: 1 });
    insertar('Servicios', { negocioId: id, id: uid(), categoria: 'General', nombre: 'Servicio de ejemplo', duracion: 30, precio: 20, activo: 1 });
    return { token: nuevaSesion(id), slug: neg.slug };
  });
}
function login(p) {
  const email = txt(p.email, 120).toLowerCase();
  const clave = String(p.clave || '');
  const codigo = String(p.codigo || '').replace(/\s/g, '');
  // Admin con Google Authenticator: correo vacío + código de 6 cifras
  if (!email && /^\d{6}$/.test(codigo)) {
    limite('totp_login', 8, 600);
    if (!verificaTotp(codigo)) throw new Error('Código incorrecto');
    return { token: nuevaSesion('ADMIN'), admin: true };
  }
  if (!email) throw new Error('Escribe tu correo');
  limite('login_' + email, 8, 600);
  const ch = validaClienteHash(clave);
  if (email === ADMIN_EMAIL) {
    const pr = props();
    const salt = pr.getProperty('ADMIN_SALT'), h = pr.getProperty('ADMIN_HASH');
    if (!h) throw new Error('Define tu contraseña con "He olvidado mi contraseña"');
    if (hashPass(salt, ch) !== h) throw new Error('Correo o contraseña incorrectos');
    return { token: nuevaSesion('ADMIN'), admin: true };
  }
  const n = tabla('Negocios').filter(x => String(x.email).toLowerCase() === email)[0];
  if (!n || hashPass(n.salt, ch) !== n.hash) throw new Error('Correo o contraseña incorrectos');
  if (n.estado === 'bloqueado') throw new Error('Cuenta suspendida. Contacta con soporte.');
  return { token: nuevaSesion(n.id) };
}
function recuperar(p) {
  const email = txt(p.email, 120).toLowerCase();
  limite('rec_' + email, 3, 3600);
  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  const exp = Date.now() + 15 * 60000;
  if (email === ADMIN_EMAIL) {
    props().setProperties({ ADMIN_RESET: sha(codigo), ADMIN_RESET_EXP: String(exp) });
  } else {
    const n = tabla('Negocios').filter(x => String(x.email).toLowerCase() === email)[0];
    if (!n) return {}; // no revelar si existe
    n.resetCode = sha(codigo); n.resetExp = exp;
    guardarFila('Negocios', n);
  }
  MailApp.sendEmail(email, APP_NOMBRE + ': código para cambiar tu contraseña',
    'Tu código es ' + codigo + '. Caduca en 15 minutos. Si no lo has pedido tú, ignora este correo.');
  return {};
}
function resetPass(p) {
  const email = txt(p.email, 120).toLowerCase();
  const ch = validaClienteHash(p.clave);
  limite('rst_' + email, 6, 900);
  const cod = sha(String(p.codigo || '').trim());
  if (email === ADMIN_EMAIL) {
    const pr = props();
    if (!pr.getProperty('ADMIN_RESET') || Date.now() > +pr.getProperty('ADMIN_RESET_EXP') || pr.getProperty('ADMIN_RESET') !== cod) throw new Error('Código incorrecto o caducado');
    const salt = uid() + uid();
    pr.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: hashPass(salt, ch) });
    pr.deleteProperty('ADMIN_RESET');
    return {};
  }
  const n = tabla('Negocios').filter(x => String(x.email).toLowerCase() === email)[0];
  if (!n || !n.resetCode || Date.now() > +n.resetExp || n.resetCode !== cod) throw new Error('Código incorrecto o caducado');
  n.salt = uid() + uid(); n.hash = hashPass(n.salt, ch); n.resetCode = ''; n.resetExp = '';
  guardarFila('Negocios', n);
  return {};
}
function cambiarPass(p, s) {
  const ch = validaClienteHash(p.clave), nuevo = validaClienteHash(p.nueva);
  const n = tabla('Negocios').filter(x => x.id === s.negocioId)[0];
  if (hashPass(n.salt, ch) !== n.hash) throw new Error('La contraseña actual no es correcta');
  n.salt = uid() + uid(); n.hash = hashPass(n.salt, nuevo);
  guardarFila('Negocios', n);
  return {};
}

/* ============ GOOGLE AUTHENTICATOR (TOTP) ============ */
function b32(s) {
  const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  let bits = '';
  s.replace(/=+$/, '').toUpperCase().split('').forEach(c => { bits += ('00000' + A.indexOf(c).toString(2)).slice(-5); });
  const out = [];
  for (let i = 0; i + 8 <= bits.length; i += 8) out.push(parseInt(bits.substr(i, 8), 2));
  return out;
}
function signed(b) { return b > 127 ? b - 256 : b; }
function totpDe(secret, t) {
  const key = b32(secret).map(signed);
  let cnt = Math.floor(t / 30000);
  const msg = [0, 0, 0, 0, 0, 0, 0, 0];
  for (let i = 7; i >= 0; i--) { msg[i] = signed(cnt & 255); cnt = Math.floor(cnt / 256); }
  const sig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, msg, key);
  const o = sig[sig.length - 1] & 15;
  const n = ((sig[o] & 127) << 24) | ((sig[o + 1] & 255) << 16) | ((sig[o + 2] & 255) << 8) | (sig[o + 3] & 255);
  return ('000000' + (n % 1000000)).slice(-6);
}
function codigoValido(secret, codigo) {
  const now = Date.now();
  return [-30000, 0, 30000].some(d => totpDe(secret, now + d) === codigo);
}
function verificaTotp(codigo) {
  const pr = props();
  if (pr.getProperty('TOTP_ON') !== '1') return false;
  return codigoValido(pr.getProperty('TOTP_SECRET'), codigo);
}
function totpStatus() { return { activo: props().getProperty('TOTP_ON') === '1' }; }
function totpStart() {
  const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
  let sec = '';
  for (let i = 0; i < 32; i++) sec += A[Math.floor(Math.random() * 32)];
  props().setProperty('TOTP_PENDIENTE', sec);
  const uri = 'otpauth://totp/' + encodeURIComponent(APP_NOMBRE + ':' + ADMIN_EMAIL) + '?secret=' + sec + '&issuer=' + encodeURIComponent(APP_NOMBRE);
  return { uri: uri, secret: sec };
}
function totpConfirm(p) {
  const pr = props();
  const sec = pr.getProperty('TOTP_PENDIENTE');
  if (!sec) throw new Error('Empieza de nuevo la configuración');
  if (!codigoValido(sec, String(p.code || '').replace(/\s/g, ''))) throw new Error('Código incorrecto');
  pr.setProperties({ TOTP_SECRET: sec, TOTP_ON: '1' });
  pr.deleteProperty('TOTP_PENDIENTE');
  return {};
}
function totpDisable() { props().setProperty('TOTP_ON', '0'); return {}; }

/* ============ LLAVE NFC (Central de Llaves) ============
 * El cliente usa el mismo protocolo que Bonificaciones:
 *   llaveSesion -> { url, codigo }     (la página abre/enseña la llave)
 *   llaveEstado {codigo} -> { estado: 'pendiente'|'ok'|'denegado'|'caducado', token?, motivo? }
 * PEGAR AQUÍ las funciones llaveSesion y llaveEstado del Codigo.gs de Bonificaciones
 * (y su constante con la URL de la Central de Llaves), cambiando solo el nombre de la app
 * y, cuando la llave sea válida, devolviendo { estado:'ok', token: nuevaSesion('ADMIN') }.
 */
const LLAVE_CENTRAL_URL = 'https://script.google.com/macros/s/AKfycbxqb6AB1Tg5xgygsU2XzDSw8fqfk9wHPOJFTEJ3LHBNebhOL72FSfM_EvVcQs6_HIFW/exec';
const LLAVE_APP_ID = 'citas';
const LLAVE_APP_SECRETO = '06871618935d068476804b5d';
const LLAVE_ADMINS = ['manu'];
function llaveLlamar_(datos) {
  const r = UrlFetchApp.fetch(LLAVE_CENTRAL_URL, {
    method: 'post', contentType: 'application/json', followRedirects: true, muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({ app: LLAVE_APP_ID, secreto: LLAVE_APP_SECRETO }, datos))
  });
  let d;
  try { d = JSON.parse(r.getContentText()); } catch (e) { throw new Error('La central de llaves no responde'); }
  if (!d.ok) throw new Error(d.error || 'Error en la central de llaves');
  return d;
}
function llaveSesion() {
  const d = llaveLlamar_({ accion: 'sesion' });
  CacheService.getScriptCache().put('llv_' + d.codigo, '1', 600);
  return { codigo: d.codigo, url: d.url };
}
function llaveEstado(p) {
  const codigo = String(p.codigo || '');
  const cache = CacheService.getScriptCache();
  if (!codigo || !cache.get('llv_' + codigo)) return { estado: 'caducado' };
  const e = llaveLlamar_({ accion: 'estado', codigo: codigo });
  if (e.estado === 'esperando') return { estado: 'pendiente' };
  if (e.estado === 'caducado') { cache.remove('llv_' + codigo); return { estado: 'caducado' }; }
  if (e.estado === 'denegado') { cache.remove('llv_' + codigo); return { estado: 'denegado', motivo: e.motivo || 'Llave no autorizada' }; }
  const id = llaveLlamar_({ accion: 'canjear', codigo: codigo });
  cache.remove('llv_' + codigo);
  if (LLAVE_ADMINS.indexOf(String(id.usuario || '').toLowerCase()) < 0) return { estado: 'denegado', motivo: 'Esta llave no es de administrador' };
  return { estado: 'ok', token: nuevaSesion('ADMIN'), admin: true };
}

/* ============ DATOS PÚBLICOS DEL NEGOCIO ============ */
function activo(n) { return n.estado !== 'bloqueado' && String(n.venceEn) >= hoyStr(); }
function datosPublicos(n) {
  return { id: n.id, slug: n.slug, nombre: n.nombre, color: n.color, logo: n.logo, portada: n.portada, telefono: n.telefono,
    direccion: n.direccion, whatsapp: n.whatsapp, descripcion: n.descripcion, paso: +n.paso, antelacion: +n.antelacion };
}
function serviciosDe(id, todos) { return tabla('Servicios').filter(s => s.negocioId === id && (todos || +s.activo === 1)).map(s => ({ id: s.id, categoria: s.categoria, nombre: s.nombre, duracion: +s.duracion, precio: +s.precio, activo: +s.activo })); }
function profesDe(id, todos) {
  return tabla('Profesionales').filter(x => x.negocioId === id && (todos || +x.activo === 1)).map(x => ({
    id: x.id, nombre: x.nombre, servicios: x.servicios ? String(x.servicios).split(',') : [], horario: JSON.parse(x.horario || '{}'), activo: +x.activo }));
}
function negPorSlug(slug) {
  const n = tabla('Negocios').filter(x => x.slug === txt(slug, 60))[0];
  if (!n) throw new Error('Negocio no encontrado');
  return n;
}
function negocioPublico(p) {
  const n = negPorSlug(p.slug);
  const out = { negocio: datosPublicos(n), activo: activo(n), hoy: hoyStr() };
  if (out.activo) { out.servicios = serviciosDe(n.id); out.profesionales = profesDe(n.id).map(x => ({ id: x.id, nombre: x.nombre, servicios: x.servicios, horario: x.horario })); }
  return out;
}

/* ============ DISPONIBILIDAD ============ */
function citasDia(negId, fecha) { return tabla('Citas').filter(c => c.negocioId === negId && c.fecha === fecha && c.estado !== 'cancelada'); }
function slotsProf(neg, prof, dur, fecha, citas) {
  const h = (prof.horario || {})[String(diaSemana(fecha))] || [];
  const paso = Math.max(5, +neg.paso || 30);
  const hoy = hoyStr();
  const minimo = fecha === hoy ? ahoraMin() + (+neg.antelacion || 0) : (fecha < hoy ? 99999 : -1);
  const ocup = citas.filter(c => c.profId === prof.id).map(c => [minutos(c.hora), minutos(c.hora) + (+c.dur || 30)]);
  const rangos = h.map(r => [minutos(r[0]), minutos(r[1])]);
  return rejilla_(rangos, ocup, paso).filter(t => t >= minimo && rangos.some(r => t >= r[0] && t + dur <= r[1]) && !ocup.some(o => t < o[1] && t + dur > o[0]));
}
/* Horas candidatas: cada "paso" minutos desde que abre, y cuando termina una cita la cuenta
   vuelve a empezar justo en ese minuto (así no se pierden huecos entre citas). */
function rejilla_(rangos, ocup, paso) {
  const out = [];
  const oc = ocup.slice().sort((a, b) => a[0] - b[0]);
  rangos.forEach(r => {
    let cur = r[0];
    const dentro = oc.filter(o => o[1] > r[0] && o[0] < r[1]);
    dentro.forEach(o => {
      for (let t = cur; t < Math.min(o[0], r[1]); t += paso) out.push(t);
      cur = Math.max(cur, o[1]);
    });
    for (let t = cur; t < r[1]; t += paso) out.push(t);
  });
  return [...new Set(out)].sort((a, b) => a - b);
}
function calcDisp(neg, servId, profId, fecha) {
  if (!esFecha(fecha)) throw new Error('Fecha no válida');
  const ids = String(servId || '').split(',').map(s => s.trim()).filter(Boolean).slice(0, 10);
  const todos = serviciosDe(neg.id);
  const lista = ids.map(id => todos.filter(s => s.id === id)[0]);
  if (!ids.length || lista.some(s => !s)) throw new Error('Servicio no disponible');
  const serv = { id: ids.join(','), nombre: lista.map(s => s.nombre).join(' + '),
    duracion: lista.reduce((a, s) => a + Number(s.duracion || 0), 0), precio: lista.reduce((a, s) => a + Number(s.precio || 0), 0) };
  let profs = profesDe(neg.id).filter(x => !x.servicios.length || ids.every(id => x.servicios.indexOf(id) >= 0));
  if (profId && profId !== 'any') profs = profs.filter(x => x.id === profId);
  const citas = citasDia(neg.id, fecha);
  const mapa = {};
  profs.forEach(pr => slotsProf(neg, pr, serv.duracion, fecha, citas).forEach(t => { if (!mapa[t]) mapa[t] = pr.id; }));
  return { serv: serv, slots: Object.keys(mapa).map(Number).sort((a, b) => a - b).map(t => ({ h: hhmm(t), p: mapa[t] })) };
}
/* ============ RESERVAS DE GRUPO (varias personas) ============ */
function personasDe_(p) {
  let ps = p.personas;
  if (typeof ps === 'string') { try { ps = JSON.parse(ps); } catch (e) { ps = null; } }
  if (!Array.isArray(ps) || ps.length < 2) return null;
  if (ps.length > 6) throw new Error('Máximo 6 personas por reserva');
  return ps.map((x, i) => ({ nombre: mayus(txt(x && x.nombre, 60)), serv: txt(x && x.serv, 300) }));
}
function calcDispGrupo(neg, personas, fecha) {
  if (!esFecha(fecha)) throw new Error('Fecha no válida');
  const todos = serviciosDe(neg.id), profs = profesDe(neg.id), citas = citasDia(neg.id, fecha);
  const dw = String(diaSemana(fecha)), paso = Math.max(5, +neg.paso || 30), hoy = hoyStr();
  const minimo = fecha === hoy ? ahoraMin() + (+neg.antelacion || 0) : (fecha < hoy ? 99999 : -1);
  const pers = personas.map((x, i) => {
    const ids = String(x.serv || '').split(',').map(s => s.trim()).filter(Boolean).slice(0, 10);
    const lista = ids.map(id => todos.filter(s => s.id === id)[0]);
    if (!ids.length || lista.some(s => !s)) throw new Error('Elige los servicios de la persona ' + (i + 1));
    const cap = profs.filter(pr => !pr.servicios.length || ids.every(id => pr.servicios.indexOf(id) >= 0));
    if (!cap.length) throw new Error('Ningún profesional hace todos los servicios de la persona ' + (i + 1));
    return { nombre: x.nombre, ids: ids, serv: ids.join(','), nombreServ: lista.map(s => s.nombre).join(' + '),
      dur: lista.reduce((a, s) => a + Number(s.duracion || 0), 0), precio: lista.reduce((a, s) => a + Number(s.precio || 0), 0), cap: cap };
  });
  const rangos = {}; profs.forEach(pr => { rangos[pr.id] = ((pr.horario || {})[dw] || []).map(r => [minutos(r[0]), minutos(r[1])]); });
  const ocBase = {}; profs.forEach(pr => { ocBase[pr.id] = citas.filter(c => c.profId === pr.id).map(c => [minutos(c.hora), minutos(c.hora) + (+c.dur || 30)]); });
  const libre = (pid, s, d, oc) => rangos[pid].some(r => s >= r[0] && s + d <= r[1]) && !oc[pid].some(o => s < o[1] && s + d > o[0]);
  const capIds = {}; pers.forEach(pe => pe.cap.forEach(c => { capIds[c.id] = 1; }));
  const inicios = [...new Set([].concat.apply([], Object.keys(capIds).map(pid => rejilla_(rangos[pid], ocBase[pid], paso))))].sort((a, b) => a - b);
  const slots = [];
  for (const t0 of inicios) {
    if (t0 < minimo) continue;
    const oc = {}; Object.keys(ocBase).forEach(k => { oc[k] = ocBase[k].slice(); });
    const asig = []; let acum = 0, ok = true;
    for (let i = 0; i < pers.length && ok; i++) {
      const pe = pers[i]; let hecho = false;
      const cand = [...new Set([].concat.apply([t0], pe.cap.map(c => rejilla_(rangos[c.id], oc[c.id], paso))))].filter(s => s >= t0 && s <= t0 + acum && s >= minimo).sort((a, b) => a - b);
      for (const s of cand) {
        const pr = pe.cap.filter(c => libre(c.id, s, pe.dur, oc))[0];
        if (pr) { oc[pr.id].push([s, s + pe.dur]); asig.push({ p: pr.id, h: hhmm(s), dur: pe.dur }); hecho = true; break; }
      }
      if (!hecho) ok = false; else acum += pe.dur;
    }
    if (ok) slots.push({ h: hhmm(t0), p: '', a: asig });
  }
  return { pers: pers, slots: slots };
}
function dispRango(p) {
  const neg = negPorSlug(p.slug);
  if (!activo(neg)) throw new Error('Este negocio no admite reservas ahora mismo');
  const desde = txt(p.desde, 10), hasta = txt(p.hasta, 10);
  if (!esFecha(desde) || !esFecha(hasta) || hasta < desde) throw new Error('Fechas no válidas');
  const dias = {};
  const ps = personasDe_(p);
  for (let f = desde, i = 0; f <= hasta && i < 42; f = sumaDias(f, 1), i++) dias[f] = ps ? calcDispGrupo(neg, ps, f).slots.map(s => ({ h: s.h, p: '' })) : calcDisp(neg, txt(p.serv, 300), txt(p.prof, 20), f).slots;
  return { dias: dias };
}
function dispPublica(p) {
  const neg = negPorSlug(p.slug);
  if (!activo(neg)) throw new Error('Este negocio no admite reservas ahora mismo');
  const ps = personasDe_(p);
  if (ps) return { slots: calcDispGrupo(neg, ps, txt(p.fecha, 10)).slots.map(s => ({ h: s.h, p: '' })) };
  return { slots: calcDisp(neg, txt(p.serv, 300), txt(p.prof, 20), txt(p.fecha, 10)).slots };
}

/* ============ RESERVAS ============ */
function reservar(p) {
  const neg = negPorSlug(p.slug);
  if (!activo(neg)) throw new Error('Este negocio no admite reservas ahora mismo');
  const cliente = mayus(p.nombre), tel = limpiaTel(p.tel);
  if (cliente.length < 2) throw new Error('Escribe tu nombre');
  if (tel.replace(/\D/g, '').length < 9) throw new Error('Escribe un teléfono válido');
  const email = txt(p.email, 120).toLowerCase();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('Correo no válido');
  limite('res_' + neg.id + '_' + tel, 5, 3600);
  const fecha = txt(p.fecha, 10), hora = txt(p.hora, 5);
  if (!esHora(hora)) throw new Error('Hora no válida');
  const ps = personasDe_(p);
  if (ps) return conLock(() => {
    const g = calcDispGrupo(neg, ps, fecha);
    const slot = g.slots.filter(s => s.h === hora)[0];
    if (!slot) throw new Error('Esa hora acaba de ocuparse. Elige otra.');
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    const profs = profesDe(neg.id, true), notas = txt(p.notas, 260);
    const ids = []; let finMax = 0;
    g.pers.forEach((pe, i) => {
      const a = slot.a[i];
      const quien = i === 0 ? cliente : (pe.nombre || 'Acompañante ' + (i + 1)) + ' · con ' + cliente;
      const c = insertar('Citas', { id: uid(), negocioId: neg.id, profId: a.p, servId: pe.serv, fecha: fecha, hora: a.h, dur: pe.dur,
        cliente: quien, tel: tel, email: email, notas: (notas ? notas + ' · ' : '') + 'Grupo de ' + g.pers.length, estado: 'confirmada', creado: hoyStr(), codigo: codigo });
      ids.push(c.id); finMax = Math.max(finMax, minutos(a.h) + pe.dur);
    });
    const nomProf = id => { const x = profs.filter(y => y.id === id)[0]; return x ? x.nombre : ''; };
    const detalle = g.pers.map((pe, i) => ({ nombre: i === 0 ? cliente : (pe.nombre || 'Acompañante ' + (i + 1)), servicio: pe.nombreServ, hora: slot.a[i].h, profesional: nomProf(slot.a[i].p), precio: pe.precio }));
    return { id: ids[0], ids: ids, codigo: codigo, fecha: fecha, hora: hora, dur: finMax - minutos(hora), personas: g.pers.length, detalle: detalle,
      servicio: g.pers.length + ' personas: ' + detalle.map(x => x.nombre.split(' ')[0] + ' (' + x.servicio + ')').join(', '),
      precio: g.pers.reduce((a, x) => a + x.precio, 0), profesional: [...new Set(detalle.map(x => x.profesional).filter(Boolean))].join(', ') };
  });
  return conLock(() => {
    const d = calcDisp(neg, txt(p.serv, 300), txt(p.prof, 20), fecha);
    const slot = d.slots.filter(s => s.h === hora)[0];
    if (!slot) throw new Error('Esa hora acaba de ocuparse. Elige otra.');
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    const c = insertar('Citas', { id: uid(), negocioId: neg.id, profId: slot.p, servId: d.serv.id, fecha: fecha, hora: hora, dur: d.serv.duracion,
      cliente: cliente, tel: tel, email: email, notas: txt(p.notas, 300), estado: 'confirmada', creado: hoyStr(), codigo: codigo });
    const pr = profesDe(neg.id, true).filter(x => x.id === slot.p)[0];
    return { id: c.id, codigo: codigo, fecha: fecha, hora: hora, dur: d.serv.duracion, servicio: d.serv.nombre, precio: d.serv.precio, profesional: pr ? pr.nombre : '' };
  });
}
function cancelarPublica(p) {
  return conLock(() => {
    const c = tabla('Citas').filter(x => x.id === txt(p.id, 20))[0];
    if (!c || String(c.codigo) !== txt(p.codigo, 10)) throw new Error('Cita no encontrada');
    tabla('Citas').filter(x => x.negocioId === c.negocioId && String(x.codigo) === String(c.codigo) && x.fecha === c.fecha && x.tel === c.tel && x.estado !== 'cancelada')
      .forEach(x => { x.estado = 'cancelada'; guardarFila('Citas', x); });
    return {};
  });
}

/* ============ PANEL DEL NEGOCIO ============ */
function miNegocio(s) { return tabla('Negocios').filter(x => x.id === s.negocioId)[0]; }
function yoNegocio(p, s) {
  const n = miNegocio(s);
  if (!n) throw new Error('Cuenta no encontrada');
  return { negocio: datosPublicos(n), email: n.email, venceEn: n.venceEn, activo: activo(n), hoy: hoyStr(), estado: n.estado,
    servicios: serviciosDe(n.id, true), profesionales: profesDe(n.id, true),
    pago: { tag: MYPOS_TAG, precio: PRECIO_MES, dias: 30, planes: PLANES } };
}
function guardarNegocio(p, s) {
  return conLock(() => {
    const n = miNegocio(s), d = p.datos || {};
    const nombre = txt(d.nombre, 60);
    if (nombre.length < 2) throw new Error('El nombre es obligatorio');
    if (!/^#[0-9a-fA-F]{6}$/.test(d.color || '')) throw new Error('Color no válido');
    ['logo', 'portada'].forEach(k => { if (!esUrl(txt(d[k], 400))) throw new Error('El enlace de ' + k + ' debe empezar por https://'); });
    let slug = String(d.slug || '').toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 40);
    if (slug.length < 3) slug = n.slug;
    if (tabla('Negocios').some(x => x.id !== n.id && x.slug === slug)) throw new Error('Ese enlace ya está en uso, prueba otro');
    let wa = String(d.whatsapp || '').replace(/\D/g, '');
    if (wa.length === 9) wa = '34' + wa;
    n.nombre = nombre; n.slug = slug; n.color = d.color; n.logo = txt(d.logo, 400); n.portada = txt(d.portada, 400);
    n.telefono = limpiaTel(d.telefono); n.whatsapp = wa; n.direccion = txt(d.direccion, 160); n.descripcion = txt(d.descripcion, 400);
    n.paso = [10, 15, 20, 30, 45, 60].indexOf(+d.paso) >= 0 ? +d.paso : 30;
    n.antelacion = Math.max(0, Math.min(10080, +d.antelacion || 0));
    guardarFila('Negocios', n);
    return { negocio: datosPublicos(n) };
  });
}
/* ============ IMÁGENES (logo / portada) EN DRIVE ============ */
// Se guardan en la carpeta "logos", junto al Sheet. Ejecuta autorizarDrive() una vez desde el editor para dar permiso.
function autorizarDrive() { return carpetaLogos_().getName() + ' lista'; }
function carpetaLogos_() {
  const id = props().getProperty('LOGOS_DIR');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { } }
  const ps = DriveApp.getFileById(SHEET_ID).getParents();
  const padre = ps.hasNext() ? ps.next() : DriveApp.getRootFolder();
  const it = padre.getFoldersByName('logos');
  const f = it.hasNext() ? it.next() : padre.createFolder('logos');
  props().setProperty('LOGOS_DIR', f.getId());
  return f;
}
function idDrive_(url) { const m = String(url || '').match(/googleusercontent\.com\/d\/([\w-]{20,})/); return m ? m[1] : ''; }
function subirImagen(p, s) {
  const tipo = p.tipo === 'portada' ? 'portada' : 'logo';
  const mime = String(p.mime || '');
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) < 0) throw new Error('Formato no válido (usa JPG, PNG o WEBP)');
  const bytes = Utilities.base64Decode(String(p.datos || ''));
  if (!bytes.length) throw new Error('Imagen vacía');
  if (bytes.length > 3 * 1024 * 1024) throw new Error('La imagen es demasiado grande (máx. 3 MB)');
  return conLock(() => {
    const n = miNegocio(s);
    const ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' }[mime];
    const file = carpetaLogos_().createFile(Utilities.newBlob(bytes, mime, n.slug + '-' + tipo + '-' + Date.now() + '.' + ext));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://lh3.googleusercontent.com/d/' + file.getId();
    const viejo = idDrive_(n[tipo]);
    if (viejo) { try { DriveApp.getFileById(viejo).setTrashed(true); } catch (e) { } }
    n[tipo] = url;
    guardarFila('Negocios', n);
    return { url: url, negocio: datosPublicos(n) };
  });
}
function quitarImagen(p, s) {
  const tipo = p.tipo === 'portada' ? 'portada' : 'logo';
  return conLock(() => {
    const n = miNegocio(s);
    const viejo = idDrive_(n[tipo]);
    if (viejo) { try { DriveApp.getFileById(viejo).setTrashed(true); } catch (e) { } }
    n[tipo] = ''; guardarFila('Negocios', n);
    return { negocio: datosPublicos(n) };
  });
}
function servGuardar(p, s) {
  return conLock(() => {
    const d = p.datos || {};
    const nombre = txt(d.nombre, 80);
    const dur = Math.round(+d.duracion);
    const precio = Math.round((+d.precio || 0) * 100) / 100;
    if (!nombre) throw new Error('Escribe el nombre del servicio');
    if (!(dur >= 5 && dur <= 600)) throw new Error('Duración entre 5 y 600 minutos');
    if (precio < 0 || precio > 9999) throw new Error('Precio no válido');
    const cat = txt(d.categoria, 40) || 'General';
    let f = d.id ? tabla('Servicios').filter(x => x.id === d.id && x.negocioId === s.negocioId)[0] : null;
    if (f) { f.nombre = nombre; f.duracion = dur; f.precio = precio; f.categoria = cat; f.activo = d.activo === 0 ? 0 : 1; guardarFila('Servicios', f); }
    else insertar('Servicios', { negocioId: s.negocioId, id: uid(), categoria: cat, nombre: nombre, duracion: dur, precio: precio, activo: 1 });
    return { servicios: serviciosDe(s.negocioId, true) };
  });
}
function servBorrar(p, s) {
  return conLock(() => {
    const f = tabla('Servicios').filter(x => x.id === p.id && x.negocioId === s.negocioId)[0];
    if (f) { f.activo = 0; guardarFila('Servicios', f); }
    return { servicios: serviciosDe(s.negocioId, true) };
  });
}
function validaHorario(h) {
  const out = {};
  for (let d = 0; d < 7; d++) {
    out[d] = ((h || {})[d] || []).map(r => {
      if (!esHora(r[0]) || !esHora(r[1]) || minutos(r[0]) >= minutos(r[1])) throw new Error('Revisa las horas del horario');
      return [r[0], r[1]];
    }).slice(0, 4);
  }
  return out;
}
function profGuardar(p, s) {
  return conLock(() => {
    const d = p.datos || {};
    const nombre = txt(d.nombre, 60);
    if (!nombre) throw new Error('Escribe el nombre del profesional');
    const ids = serviciosDe(s.negocioId, true).map(x => x.id);
    const sv = (d.servicios || []).filter(x => ids.indexOf(x) >= 0);
    const hor = JSON.stringify(validaHorario(d.horario));
    let f = d.id ? tabla('Profesionales').filter(x => x.id === d.id && x.negocioId === s.negocioId)[0] : null;
    if (f) { f.nombre = nombre; f.servicios = sv.join(','); f.horario = hor; f.activo = d.activo === 0 ? 0 : 1; guardarFila('Profesionales', f); }
    else insertar('Profesionales', { negocioId: s.negocioId, id: uid(), nombre: nombre, servicios: sv.join(','), horario: hor, activo: 1 });
    return { profesionales: profesDe(s.negocioId, true) };
  });
}
function profBorrar(p, s) {
  return conLock(() => {
    const act = profesDe(s.negocioId).length;
    const f = tabla('Profesionales').filter(x => x.id === p.id && x.negocioId === s.negocioId)[0];
    if (f && act <= 1 && +f.activo === 1) throw new Error('Debe quedar al menos un profesional activo');
    if (f) { f.activo = 0; guardarFila('Profesionales', f); }
    return { profesionales: profesDe(s.negocioId, true) };
  });
}
function agenda(p, s) {
  const desde = txt(p.desde, 10), hasta = txt(p.hasta, 10);
  if (!esFecha(desde) || !esFecha(hasta)) throw new Error('Fechas no válidas');
  const citas = tabla('Citas').filter(c => c.negocioId === s.negocioId && c.fecha >= desde && c.fecha <= hasta)
    .map(c => ({ id: c.id, profId: c.profId, servId: c.servId, fecha: c.fecha, hora: c.hora, dur: +c.dur, cliente: c.cliente, tel: c.tel, email: c.email, notas: c.notas, estado: c.estado }));
  citas.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
  return { citas: citas };
}
function citaEstado(p, s) {
  return conLock(() => {
    const c = tabla('Citas').filter(x => x.id === p.id && x.negocioId === s.negocioId)[0];
    if (!c) throw new Error('Cita no encontrada');
    if (['confirmada', 'completada', 'cancelada', 'no vino'].indexOf(p.estado) < 0) throw new Error('Estado no válido');
    c.estado = p.estado;
    guardarFila('Citas', c);
    return {};
  });
}
function citaManual(p, s) {
  return conLock(() => {
    const d = p.datos || {};
    const fecha = txt(d.fecha, 10), hora = txt(d.hora, 5);
    if (!esFecha(fecha) || !esHora(hora)) throw new Error('Fecha u hora no válidas');
    const serv = serviciosDe(s.negocioId, true).filter(x => x.id === d.serv)[0];
    const prof = profesDe(s.negocioId, true).filter(x => x.id === d.prof)[0];
    if (!serv || !prof) throw new Error('Elige servicio y profesional');
    const dur = Math.round(+d.dur) || serv.duracion;
    const ini = minutos(hora);
    const choca = citasDia(s.negocioId, fecha).some(c => c.profId === prof.id && ini < minutos(c.hora) + +c.dur && ini + dur > minutos(c.hora));
    if (choca) throw new Error('Esa franja ya está ocupada para ese profesional');
    const c = insertar('Citas', { id: uid(), negocioId: s.negocioId, profId: prof.id, servId: serv.id, fecha: fecha, hora: hora, dur: dur,
      cliente: mayus(d.cliente) || 'Bloqueado', tel: limpiaTel(d.tel), email: '', notas: txt(d.notas, 300), estado: 'confirmada', creado: hoyStr(), codigo: String(Math.floor(100000 + Math.random() * 900000)) });
    return { id: c.id };
  });
}

/* ============ PANEL ADMIN ============ */
function yoAdmin() { return { admin: true, precio: PRECIO_MES, planes: PLANES, hoy: hoyStr(), email: ADMIN_EMAIL }; }
function adminNegocios() {
  const citas = tabla('Citas');
  return {
    negocios: tabla('Negocios').map(n => ({ id: n.id, slug: n.slug, nombre: n.nombre, email: n.email, estado: n.estado, venceEn: n.venceEn,
      creado: n.creado, citas: citas.filter(c => c.negocioId === n.id).length, activo: activo(n) }))
  };
}
function adminDias(p) {
  return conLock(() => {
    const n = tabla('Negocios').filter(x => x.id === p.id)[0];
    if (!n) throw new Error('Negocio no encontrado');
    const dias = Math.round(+p.dias);
    if (!dias || Math.abs(dias) > 3650) throw new Error('Días no válidos');
    const base = String(n.venceEn) > hoyStr() ? n.venceEn : hoyStr();
    n.venceEn = sumaDias(base, dias);
    guardarFila('Negocios', n);
    return { venceEn: n.venceEn };
  });
}
function adminEstado(p) {
  return conLock(() => {
    const n = tabla('Negocios').filter(x => x.id === p.id)[0];
    if (!n) throw new Error('Negocio no encontrado');
    n.estado = p.estado === 'bloqueado' ? 'bloqueado' : 'activo';
    guardarFila('Negocios', n);
    return {};
  });
}

/* Limpieza opcional: programa un activador diario para esta función */
function limpiarSesiones() {
  const sh = SS().getSheetByName('Sesiones');
  leerHoja_('Sesiones').filter(x => +x.expira < Date.now()).sort((a, b) => b._r - a._r).forEach(x => sh.deleteRow(x._r));
  tocar_('Sesiones');
}
