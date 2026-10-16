// ============================================================
// Hallazgos de Planta · API para Google Sheets + Drive · v2
// ------------------------------------------------------------
// Qué hace:
//  - Login solo con el nombre de usuario, con sesión (token) de 7 días
//  - Permisos revisados en el servidor, según Niveles y privilegios
//  - Guarda hallazgos en la pestaña "hallazgos" (una columna por campo)
//  - Bitácora de solo-agregar (nadie puede editarla desde la app)
//  - Catálogos, privilegios y opciones de cierre en la pestaña "Config"
//  - Fotos en Drive
//  - Resumen diario por correo (ver sección CORREO al final)
// ============================================================

// ---------- AJUSTES (puedes cambiarlos) ----------
const SHEET_NAME    = 'hallazgos';   // Pestaña de datos
const USERS_SHEET   = 'Usuarios';    // Pestaña de usuarios
const AREAS_SHEET   = 'Areas';       // Pestaña antigua de áreas (SIN tilde)
const CATS_SHEET    = 'Categorias';  // Pestaña antigua de categorías (SIN tilde)
const CONFIG_SHEET  = 'Config';      // Se crea sola
const LOG_SHEET     = 'Bitacora';    // Se crea sola
const PHOTO_FOLDER  = 'Hallazgos Planta - Fotos';

const SESSION_DAYS  = 7;             // Duración de la sesión
const FOLIO_PREFIX  = 'HP-';         // Consecutivo de hallazgos: HP-0001, HP-0002, …
const FOLIO_DIGITS  = 4;             // Cantidad mínima de dígitos
const LEGACY_GET_WRITES = false;     // true = deja funcionar la app ANTERIOR (sin sesión). Déjalo en false.

// --- Correo diario ---
const APP_URL                  = '';  // Dirección de la app (GitHub Pages) para el botón del correo
const CORREO_DESTINATARIOS     = '';  // Correos separados por coma. Vacío = el dueño de este script
const CORREO_HORA              = 7;   // Hora a la que se envía (0 a 23)
const CORREO_SOLO_DIAS_HABILES = true;
const DIAS_AVISO               = 3;   // "Próximos a vencer" = vencen en los siguientes N días

// ---------- CAMPOS ----------
const ROW_KEYS = ['id','nuevo','archivo','tipo','origen','fr','cat','desc','crit','area','ubic','equipo',
  'req','freq','ot','contratista','ccosto','cmon','tc','fentrega','refs','ref','reporto','porId','resp',
  'accion','segcap','inv','comp','est','av','seg','segId','fseg','cerro','cerroId','fcierre','obs','pkg',
  'wbs','hrs','hrr','bs','bf','foto','fotoDb','despues','despuesDb','ts','upd','borrado','orig','esp','borr','autor','folio'];
const NUM_FIELDS  = ['av','hrs','hrr','ccosto','tc','ts','upd'];
const BOOL_FIELDS = ['nuevo','fotoDb','despuesDb','borrado','borr'];
const DATE_FIELDS = ['fr','comp','freq','fentrega','fseg','fcierre','bs','bf'];
const COMPRAS_KEYS = ['req','freq','ot','contratista','ccosto','cmon','tc','fentrega','refs'];
const CFG_KEYS = { catalogos:['catalogos','baja'], privilegios:['niveles'], cierre:['cierreOpc'], plan:['editar'] };

const PRIV_DEFAULT = {
  supervisor: ['crear','editar','estatus','cerrar','reabrir','fotoDespues','compras','asignar','catalogos','baja','excel','respaldo'],
  usuario:    ['crear','editar','estatus','fotoDespues']
};

// ============================================================
// ROUTER
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || '';
  try {
    let r;
    switch (action) {
      case 'login':  r = login(p.name); break;
      case 'ver':    auth_(p.token); r = { ok: true, ver: getVer_() }; break;
      case 'getAll': r = getAll_(auth_(p.token)); break;
      // --- compatibilidad con la app anterior (apagado por defecto) ---
      case 'updateRow': if (!LEGACY_GET_WRITES) throw err_('auth', 'Actualice la app: esta versión ya no acepta escrituras sin sesión.'); r = legacyUpdateRow_(JSON.parse(p.row)); break;
      case 'deleteRow': if (!LEGACY_GET_WRITES) throw err_('auth', 'Actualice la app: esta versión ya no acepta escrituras sin sesión.'); r = legacyDeleteRow_(p.id, p.user); break;
      default: r = { ok: false, error: 'Acción desconocida: ' + action };
    }
    return json(r);
  } catch (err) {
    return json({ ok: false, error: err.message, code: err.code || '' });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let r;
    switch (body.action) {
      case 'batch':       r = batch_(auth_(body.token), body.ops || []); break;
      case 'uploadPhoto': r = uploadPhoto_(auth_(body.token), body.base64, body.name); break;
      case 'assignFolio': r = assignFolio_(auth_(body.token), body.id); break;
      case 'getPhotos':   r = getPhotos_(auth_(body.token), body.ids || []); break;
      case 'saveUsers':   r = saveUsers_(auth_(body.token), body.users || []); break;
      case 'deleteUser':  r = deleteUser_(auth_(body.token), body.name); break;
      default: r = { ok: false, error: 'Acción POST desconocida: ' + body.action };
    }
    return json(r);
  } catch (err) {
    return json({ ok: false, error: err.message, code: err.code || '' });
  }
}

// ============================================================
// UTILIDADES
// ============================================================
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function err_(code, msg) { const e = new Error(msg); e.code = code; return e; }

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
}
// admin / usuario / supervisor / otros niveles (en minúsculas, sin acentos)
function normRol_(r) {
  r = norm_(r);
  if (!r || r === 'user' || r === 'usuario') return 'usuario';
  if (/^admin/.test(r)) return 'admin';
  return r.replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 20) || 'usuario';
}
function todayStr_() { return Utilities.formatDate(new Date(), ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }

function dataSheet_() {
  const s = ss().getSheetByName(SHEET_NAME);
  if (!s) throw new Error('No existe la pestaña ' + SHEET_NAME);
  return s;
}
function usersSheet_() {
  const s = ss().getSheetByName(USERS_SHEET);
  if (!s) throw new Error('No existe la pestaña ' + USERS_SHEET);
  return s;
}
function textSheet_(name, headers) {
  let s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), headers.length).setNumberFormat('@'); // texto plano: Sheets no convierte fechas ni fórmulas
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return s;
}
function cfgSheet_() { return textSheet_(CONFIG_SHEET, ['clave', 'valor', 'actualizado', 'por']); }
function logSheet_() { return textSheet_(LOG_SHEET, ['ts', 'fecha', 'hora', 'uid', 'usuario', 'id', 'movimiento', 'detalle']); }

// ============================================================
// USUARIOS Y SESIÓN
// ============================================================
// Encuentra las columnas por su encabezado (Nombre / Rol / Activo / PIN).
function userInfo_() {
  const sh = usersSheet_();
  const v = sh.getDataRange().getValues();
  const h = (v[0] || []).map(norm_);
  const find = function () { for (let i = 0; i < arguments.length; i++) { const x = h.indexOf(arguments[i]); if (x >= 0) return x; } return -1; };
  const rolIdx = find('rol', 'role', 'nivel', 'perfil');
  return {
    sh: sh, v: v, h: h,
    n:   Math.max(0, find('nombre', 'name', 'usuario', 'user')),
    rol: rolIdx >= 0 ? rolIdx : 1,            // sin encabezado "Rol": usa la columna B como antes
    act: find('activo', 'estado', 'status')
  };
}
function readUsers_() {
  const info = userInfo_(), out = [];
  for (let i = 1; i < info.v.length; i++) {
    const r = info.v[i], n = String(r[info.n] || '').trim();
    if (!n) continue;
    let act = true;
    if (info.act >= 0) { const a = norm_(r[info.act]); act = !(a === 'no' || a === '0' || a === 'false' || a === 'baja' || a === 'inactivo'); }
    out.push({ n: n, rol: normRol_(r[info.rol]), act: act, row: i + 1 });
  }
  return out;
}
function usersCached_() {
  const c = CacheService.getScriptCache(), raw = c.get('USERS');
  if (raw) { try { return JSON.parse(raw); } catch (_) {} }
  const u = readUsers_();
  try { c.put('USERS', JSON.stringify(u), 60); } catch (_) {}
  return u;
}

function login(name) {
  if (!name) return { ok: false, error: 'Nombre vacío' };
  const key = norm_(name);
  const u = readUsers_().filter(function (x) { return norm_(x.n) === key; })[0];
  if (!u) return { ok: false, error: 'Usuario no encontrado' };
  if (!u.act) return { ok: false, error: 'Usuario dado de baja. Contacte al administrador.' };
  return { ok: true, nombre: u.n, rol: u.rol, token: newToken_(u.n), dias: SESSION_DAYS };
}

function newToken_(name) {
  const P = PropertiesService.getScriptProperties(), now = Date.now(), all = P.getProperties();
  Object.keys(all).forEach(function (k) {            // limpia sesiones vencidas
    if (k.indexOf('T_') === 0) { try { if (JSON.parse(all[k]).exp < now) P.deleteProperty(k); } catch (_) { P.deleteProperty(k); } }
  });
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  P.setProperty('T_' + t, JSON.stringify({ n: name, exp: now + SESSION_DAYS * 864e5 }));
  return t;
}
function auth_(token) {
  if (!token) throw err_('auth', 'Sesión no iniciada');
  const raw = PropertiesService.getScriptProperties().getProperty('T_' + token);
  if (!raw) throw err_('auth', 'Sesión expirada');
  let t; try { t = JSON.parse(raw); } catch (_) { throw err_('auth', 'Sesión inválida'); }
  if (t.exp < Date.now()) throw err_('auth', 'Sesión expirada');
  const u = usersCached_().filter(function (x) { return norm_(x.n) === norm_(t.n); })[0];
  if (!u || !u.act) throw err_('auth', 'Usuario no autorizado');
  return { n: u.n, rol: u.rol };
}

// ---------- privilegios (se leen de Config > privilegios) ----------
function privMatrix_() {
  const c = CacheService.getScriptCache(), raw = c.get('PRIV');
  if (raw) { try { return JSON.parse(raw); } catch (_) {} }
  let m = {};
  const cfg = readConfig_().privilegios;
  if (cfg && cfg.m) m = cfg.m;
  Object.keys(PRIV_DEFAULT).forEach(function (k) {
    if (!m[k]) { m[k] = {}; PRIV_DEFAULT[k].forEach(function (p) { m[k][p] = true; }); }
  });
  try { c.put('PRIV', JSON.stringify(m), 60); } catch (_) {}
  return m;
}
function can_(user, priv) {
  if (user.rol === 'admin') return true;
  if (priv === 'borrar') return false;              // borrar solo lo hace el Administrador
  const m = privMatrix_(), set = m[user.rol] || m.usuario || {};
  return !!set[priv];
}
function need_(user, priv, msg) { if (!can_(user, priv)) throw err_('perm', msg || 'Su nivel no tiene permiso para esto.'); }

// ============================================================
// CONFIG (catálogos, privilegios, cierre, plan)
// ============================================================
function readConfig_() {
  const v = cfgSheet_().getDataRange().getValues(), parts = {}, out = {};
  for (let i = 1; i < v.length; i++) {
    const k = String(v[i][0] || ''), m = k.match(/^(.+)#(\d+)$/);
    if (!m) continue;
    (parts[m[1]] = parts[m[1]] || [])[Number(m[2])] = String(v[i][1] == null ? '' : v[i][1]);
  }
  Object.keys(parts).forEach(function (k) { try { out[k] = JSON.parse(parts[k].join('')); } catch (_) {} });
  return out;
}
function setConfig_(user, key, value) {
  const allowed = CFG_KEYS[key];
  if (!allowed) throw new Error('Configuración no permitida: ' + key);
  if (!allowed.some(function (p) { return can_(user, p); })) throw err_('perm', 'Su nivel no puede cambiar esta configuración.');
  const s = JSON.stringify(value);
  if (s.length > 190000) throw new Error('La configuración es demasiado grande.');
  const sh = cfgSheet_(), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).replace(/#\d+$/, '') === key) sh.deleteRow(i + 1);
  }
  const now = new Date().toISOString();
  for (let j = 0, n = 0; j < s.length || n === 0; j += 40000, n++) {
    sh.appendRow([key + '#' + n, s.slice(j, j + 40000), now, user.n]);
  }
  CacheService.getScriptCache().remove('PRIV');
}

// ============================================================
// HALLAZGOS
// ============================================================
function truthy_(v) { return v === true || /^(true|si|sí|1)$/i.test(String(v == null ? '' : v).trim()); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
function cell_(key, val) {
  if (val == null) return '';
  if (BOOL_FIELDS.indexOf(key) >= 0) return (val === true || val === 'TRUE' || val === 'true' || val === 1) ? 'TRUE' : '';
  if (NUM_FIELDS.indexOf(key) >= 0) { const n = Number(val); return (val === '' || !isFinite(n)) ? '' : n; }
  let s = String(val);
  if (s.charAt(0) === '=') s = ' ' + s;               // evita que un texto se ejecute como fórmula
  return s.length > 45000 ? s.slice(0, 45000) : s;
}
function read_(key, val) {
  if (isDate_(val)) return Utilities.formatDate(val, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (BOOL_FIELDS.indexOf(key) >= 0) return val === true || /^(true|si|sí|1)$/i.test(String(val).trim());
  if (NUM_FIELDS.indexOf(key) >= 0) { if (val === '' || val == null) return ''; const n = Number(val); return isFinite(n) ? n : ''; }
  if (val == null) return '';
  if (typeof val === 'number') return String(val);
  return val;
}
function readAllRows_() {
  const v = dataSheet_().getDataRange().getValues();
  const h = (v[0] || []).map(String), rows = [];
  for (let i = 1; i < v.length; i++) {
    const o = {};
    h.forEach(function (k, j) { if (k && k.charAt(0) !== '_') o[k] = read_(k, v[i][j]); });
    if (!String(o.id || '').trim()) continue;
    o.id = String(o.id);
    rows.push(o);
  }
  return rows;
}
function readList_(sheetName) {
  try {
    const sh = ss().getSheetByName(sheetName);
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).map(function (r) { return String(r[0] || '').trim(); }).filter(Boolean);
  } catch (_) { return []; }
}
function readLog_(limit) {
  const sh = logSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - limit + 1);
  return sh.getRange(start, 1, last - start + 1, 8).getValues().map(function (r) {
    return { ts: Number(r[0]) || 0, f: String(isDate_(r[1]) ? read_('fr', r[1]) : r[1]), h: String(r[2]), uid: null,
             nm: String(r[4]), id: String(r[5]), a: String(r[6]), d: String(r[7]) };
  });
}
function usersForClient_() {
  return readUsers_().map(function (u) { return { n: u.n, rol: u.rol, b: u.act ? 0 : 1 }; });
}
function getVer_() { return Number(PropertiesService.getScriptProperties().getProperty('VER') || 0); }
function bump_() { const P = PropertiesService.getScriptProperties(), v = getVer_() + 1; P.setProperty('VER', String(v)); return v; }

function getAll_(user) {
  const rows = readAllRows_().filter(function (r) {
    return !r.borr || user.rol === 'admin' || norm_(r.autor) === norm_(user.n);
  });
  const out = {
    ok: true, ver: getVer_(), me: { n: user.n, rol: user.rol },
    rows: rows, log: readLog_(1000), config: readConfig_(),
    areas: readList_(AREAS_SHEET), categorias: readList_(CATS_SHEET)
  };
  if (can_(user, 'usuarios') || can_(user, 'niveles')) out.users = usersForClient_();
  return out;
}

// ---------- escrituras en lote ----------
function batch_(user, ops) {
  if (!ops || !ops.length) return { ok: true, ver: getVer_(), prev: getVer_(), errors: [] };
  if (ops.length > 200) throw new Error('Demasiadas operaciones en una sola solicitud');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const prev = getVer_(), errors = [];
    const sh = dataSheet_();
    const S = { sh: sh, data: sh.getDataRange().getValues(), rowOf: {}, last: 0 };
    S.h = (S.data[0] || []).map(String);
    if (S.h.indexOf('id') < 0) throw new Error('No hay columna id');
    S.last = S.data.length;
    const idc = S.h.indexOf('id');
    for (let i = 1; i < S.data.length; i++) S.rowOf[String(S.data[i][idc])] = i + 1;

    ops.forEach(function (op, i) {
      try {
        if (op.t === 'row') applyRow_(S, user, op);
        else if (op.t === 'log') addLog_(user, op.entry || {});
        else if (op.t === 'cfg') setConfig_(user, String(op.key), op.value);
        else throw new Error('Operación desconocida');
      } catch (e) { errors.push({ i: i, error: e.message, code: e.code || '' }); }
    });
    return { ok: true, ver: bump_(), prev: prev, errors: errors };
  } finally { lock.releaseLock(); }
}

function applyRow_(S, user, op) {
  const id = String(op.id || '').trim();
  if (!id || id.length > 60) throw new Error('Fila sin id');
  const patch = op.patch || {};
  const keys = Object.keys(patch).filter(function (k) {
    return k !== 'id' && /^[A-Za-z0-9_]{1,30}$/.test(k) && k.charAt(0) !== '_' && (ROW_KEYS.indexOf(k) >= 0 || S.h.indexOf(k) >= 0);
  });
  const idx = S.rowOf[id];
  const idc = S.h.indexOf('id');

  // --- borradores y permisos ---
  const colBorr = S.h.indexOf('borr'), colAutor = S.h.indexOf('autor');
  const curRow = idx ? S.data[idx - 1] : null;
  const isDraft = !!(curRow && colBorr >= 0 && truthy_(curRow[colBorr]));
  const mine = !!(curRow && colAutor >= 0 && norm_(curRow[colAutor]) === norm_(user.n));
  const ownDraft = isDraft && (mine || user.rol === 'admin');
  if (isDraft && !ownDraft) throw err_('perm', 'Este borrador es de otra persona.');

  if (idx) {                                          // el consecutivo, el autor y el estado de borrador no se editan desde el navegador
    ['folio', 'autor', 'borr'].forEach(function (k) { const x = keys.indexOf(k); if (x >= 0) keys.splice(x, 1); });
  } else {
    ['folio', 'autor'].forEach(function (k) { const x = keys.indexOf(k); if (x >= 0) keys.splice(x, 1); });
    if (user.rol !== 'admin') { patch.borr = true; if (keys.indexOf('borr') < 0) keys.push('borr'); }   // todo hallazgo nuevo nace como borrador
    patch.autor = user.n; keys.push('autor');
  }

  if (patch.borrado && !ownDraft) need_(user, 'borrar', 'Solo el Administrador puede borrar registros.');
  if (keys.some(function (k) { return COMPRAS_KEYS.indexOf(k) >= 0 && (idx || String(patch[k]) !== ''); })) need_(user, 'compras', 'Su nivel no puede capturar requisición ni compras.');
  if (keys.indexOf('despues') >= 0 || keys.indexOf('despuesDb') >= 0) {
    if (patch.despues || patch.despuesDb || idx) need_(user, 'fotoDespues', 'Su nivel no puede subir la foto del después.');
  }
  if (!idx) {
    need_(user, 'crear', 'Su nivel no puede crear hallazgos.');
    if (patch.est === 'Cerrado') need_(user, 'cerrar', 'Su nivel no puede cerrar hallazgos.');
  } else if (isDraft) {
    need_(user, 'crear', 'Su nivel no puede crear hallazgos.');
  } else {
    const estCol = S.h.indexOf('est'), curEst = estCol >= 0 ? String(S.data[idx - 1][estCol] || '') : '';
    if (patch.est === 'Cerrado' && curEst !== 'Cerrado') need_(user, 'cerrar', 'Su nivel no puede cerrar hallazgos.');
    if (curEst === 'Cerrado' && patch.est && patch.est !== 'Cerrado') need_(user, 'reabrir', 'Su nivel no puede reabrir hallazgos.');
    const general = keys.some(function (k) { return COMPRAS_KEYS.indexOf(k) < 0 && ['despues', 'despuesDb', 'borrado', 'upd', 'ts'].indexOf(k) < 0; });
    if (general && !(can_(user, 'editar') || can_(user, 'estatus') || can_(user, 'asignar') || can_(user, 'catalogos'))) {
      throw err_('perm', 'Su nivel no puede editar hallazgos.');
    }
  }

  // --- columnas nuevas (se crean solas) ---
  keys.forEach(function (k) {
    if (S.h.indexOf(k) < 0) {
      S.h.push(k);
      S.sh.getRange(1, S.h.length).setValue(k);
      if (NUM_FIELDS.indexOf(k) < 0 && DATE_FIELDS.indexOf(k) < 0) {
        S.sh.getRange(2, S.h.length, Math.max(1, S.sh.getMaxRows() - 1), 1).setNumberFormat('@');
      }
    }
  });

  const cur = idx ? S.data[idx - 1].slice() : [];
  while (cur.length < S.h.length) cur.push('');
  cur[idc] = id;
  keys.forEach(function (k) { cur[S.h.indexOf(k)] = cell_(k, patch[k]); });

  // --- fotos: máximo 5 por lado ---
  ['foto', 'despues'].forEach(function (k) {
    const c = S.h.indexOf(k);
    if (c >= 0 && String(cur[c] || '').split('|').filter(function (x) { return /^https:\/\//.test(x.trim()); }).length > 5) {
      throw new Error('Máximo 5 fotos por hallazgo (' + (k === 'foto' ? 'antes' : 'después') + ').');
    }
  });
  // --- no se puede cerrar sin al menos una foto del después ---
  const g = function (k) { const c = S.h.indexOf(k); return c >= 0 ? cur[c] : ''; };
  if (String(g('est')) === 'Cerrado' && (patch.est !== undefined || patch.av !== undefined)) {
    const legacy = String(g('archivo') || '').split('|').map(function (x) { return x.trim(); });
    const fotos = String(g('despues') || '').split('|').map(function (x) { return x.trim(); })
      .filter(function (x) { return /^https:\/\//.test(x) && legacy.indexOf(x) < 0; });
    if (!fotos.length && String(g('despuesDb')) !== 'TRUE') throw err_('perm', 'Para cerrar un hallazgo se necesita al menos una foto del después.');
  }

  if (idx) {
    S.data[idx - 1] = cur;
    S.sh.getRange(idx, 1, 1, S.h.length).setValues([cur]);
  } else {
    S.last++;
    S.data.push(cur);
    S.rowOf[id] = S.last;
    S.sh.getRange(S.last, 1, 1, S.h.length).setValues([cur]);
  }
}

function addLog_(user, e) {
  const now = new Date();
  logSheet_().appendRow([
    String(Number(e.ts) || now.getTime()),
    String(e.f || todayStr_()).slice(0, 10),
    String(e.h || '').slice(0, 8),
    '', user.n,                                       // el usuario lo pone el servidor, no el navegador
    String(e.id || '').slice(0, 60),
    String(e.a || '').slice(0, 60),
    String(e.d || '').slice(0, 500)
  ]);
}

// ---------- usuarios ----------
function saveUsers_(user, list) {
  need_(user, 'usuarios', 'Su nivel no puede administrar usuarios.');
  const info = userInfo_(), sh = info.sh;
  let h = info.v[0].map(String), col = { n: info.n, rol: info.rol, act: info.act };
  if (col.act < 0) { h.push('Activo'); col.act = h.length - 1; sh.getRange(1, col.act + 1).setValue('Activo'); }

  const byName = {};
  for (let i = 1; i < info.v.length; i++) { const n = norm_(info.v[i][info.n]); if (n) byName[n] = i + 1; }

  // al menos un administrador activo
  const futuros = list.filter(function (u) { return u && !u.b && normRol_(u.rol) === 'admin'; }).length;
  const hayOtros = readUsers_().some(function (u) { return u.act && u.rol === 'admin' && !list.some(function (x) { return norm_(x.n) === norm_(u.n); }); });
  if (!futuros && !hayOtros) throw new Error('Debe quedar al menos un Administrador activo.');

  const errors = [];
  list.forEach(function (u, i) {
    try {
      const nombre = String((u && u.n) || '').replace(/\s+/g, ' ').trim().slice(0, 80);
      if (!nombre) return;
      const rol = normRol_(u.rol), activo = u.b ? 'No' : 'Sí', mismo = norm_(nombre) === norm_(user.n);
      if (mismo && (u.b || (user.rol === 'admin' && rol !== 'admin'))) throw new Error('No puede darse de baja ni quitarse el nivel de Administrador a sí mismo.');
      const row = byName[norm_(nombre)];
      if (row) {
        sh.getRange(row, col.rol + 1).setValue(rol);
        sh.getRange(row, col.act + 1).setValue(activo);
      } else {
        const nr = new Array(Math.max(h.length, col.act + 1)).fill('');
        nr[col.n] = nombre; nr[col.rol] = rol; nr[col.act] = activo;
        sh.appendRow(nr);
      }
    } catch (e) { errors.push({ i: i, error: e.message }); }
  });
  CacheService.getScriptCache().remove('USERS');
  return { ok: true, ver: bump_(), errors: errors, users: usersForClient_() };
}
function deleteUser_(user, name) {
  need_(user, 'borrar', 'Solo el Administrador puede eliminar usuarios.');
  if (norm_(name) === norm_(user.n)) throw new Error('No puede eliminarse a sí mismo.');
  const info = userInfo_();
  for (let i = 1; i < info.v.length; i++) {
    if (norm_(info.v[i][info.n]) === norm_(name)) {
      info.sh.deleteRow(i + 1);
      CacheService.getScriptCache().remove('USERS');
      return { ok: true, ver: bump_(), users: usersForClient_() };
    }
  }
  return { ok: false, error: 'No encontrado' };
}

// ============================================================
// FOTOS (Drive)
// ============================================================
function uploadPhoto_(user, base64, name) {
  if (!(can_(user, 'crear') || can_(user, 'fotoDespues'))) throw err_('perm', 'Su nivel no puede subir fotos.');
  if (!base64) throw new Error('Falta la imagen');
  if (String(base64).length > 8 * 1024 * 1024) throw new Error('La imagen es demasiado grande');
  const folder = getOrCreateFolder(PHOTO_FOLDER);
  const clean = String(base64).replace(/^data:image\/\w+;base64,/, '');
  const safe = String(name || ('foto_' + Date.now() + '.jpg')).replace(/[^\w.\-]+/g, '_').slice(0, 80);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(clean), 'image/jpeg', safe));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return { ok: true, url: 'https://lh3.googleusercontent.com/d/' + id, id: id };
}
function getOrCreateFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// ============================================================
// CONSECUTIVO (FOLIO)
// ============================================================
// Todo hallazgo nace como borrador con un identificador temporal (B-...). Al pulsar
// "Guardar y asignar folio" el servidor le pone el siguiente número de la serie
// (HP-0001, HP-0002, …). El número se reserva dentro de un candado, así que dos personas
// nunca reciben el mismo, y nunca se reutiliza aunque después se borre el hallazgo.
function nextFolio_(rowOf) {
  const P = PropertiesService.getScriptProperties();
  let n = Number(P.getProperty('FOLIO') || 0);
  if (!n) {                                           // primera vez: continúa después del mayor HP-#### que ya exista
    const re = new RegExp('^' + FOLIO_PREFIX.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\d+)$');
    Object.keys(rowOf).forEach(function (id) { const m = id.match(re); if (m) n = Math.max(n, Number(m[1])); });
  }
  let id;
  do { n++; id = FOLIO_PREFIX + String(n).padStart(FOLIO_DIGITS, '0'); } while (rowOf[id]);
  P.setProperty('FOLIO', String(n));
  return id;
}
function assignFolio_(user, oldId) {
  need_(user, 'crear', 'Su nivel no puede crear hallazgos.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const prev = getVer_(), sh = dataSheet_(), data = sh.getDataRange().getValues(), h = data[0].map(String);
    const col = function (k) { return h.indexOf(k); };
    const idc = col('id'), rowOf = {};
    for (let i = 1; i < data.length; i++) rowOf[String(data[i][idc])] = i + 1;
    const idx = rowOf[String(oldId)];
    if (!idx) throw new Error('No se encontró el borrador.');
    const cur = data[idx - 1].slice();
    if (col('borr') < 0 || !truthy_(cur[col('borr')])) throw new Error('Este hallazgo ya tiene folio.');
    if (user.rol !== 'admin' && norm_(cur[col('autor')]) !== norm_(user.n)) throw err_('perm', 'Este borrador es de otra persona.');
    if (truthy_(cur[col('borrado')])) throw new Error('El borrador fue eliminado.');
    const folio = nextFolio_(rowOf);
    ['folio', 'upd'].forEach(function (k) {
      if (col(k) < 0) { h.push(k); sh.getRange(1, h.length).setValue(k); }
    });
    while (cur.length < h.length) cur.push('');
    cur[idc] = folio; cur[col('folio')] = folio; cur[col('borr')] = ''; cur[col('upd')] = Date.now();
    sh.getRange(idx, 1, 1, h.length).setValues([cur]);
    const row = {};
    h.forEach(function (k, j) { if (k && k.charAt(0) !== '_') row[k] = read_(k, cur[j]); });
    row.id = folio;
    return { ok: true, folio: folio, oldId: String(oldId), row: row, prev: prev, ver: bump_() };
  } finally { lock.releaseLock(); }
}

// ============================================================
// FOTOS PARA EL EXCEL
// ============================================================
// El navegador intenta bajar las fotos directo; las que Google no deje (por permisos del navegador)
// se piden aquí. Solo entrega archivos que estén dentro de la carpeta de fotos de la app.
function getPhotos_(user, ids) {
  if (!ids.length) return { ok: true, photos: {} };
  if (ids.length > 12) throw new Error('Pida máximo 12 fotos por vez.');
  const folderId = getOrCreateFolder(PHOTO_FOLDER).getId(), out = {};
  ids.forEach(function (id) {
    try {
      if (!/^[A-Za-z0-9_-]{10,80}$/.test(String(id))) return;
      const f = DriveApp.getFileById(String(id)), par = f.getParents();
      if (!par.hasNext() || par.next().getId() !== folderId) return;
      let b = null;
      try { b = f.getThumbnail(); } catch (_) {}
      if (!b) { b = f.getBlob(); if (b.getBytes().length > 1500000) return; }
      out[id] = { mime: b.getContentType() || 'image/jpeg', b64: Utilities.base64Encode(b.getBytes()) };
    } catch (_) {}
  });
  return { ok: true, photos: out };
}

// ============================================================
// COMPATIBILIDAD CON LA APP ANTERIOR (solo si LEGACY_GET_WRITES = true)
// ============================================================
function legacyUpdateRow_(row) {
  if (!row || !row.id) throw new Error('Fila sin id');
  const sh = dataSheet_(), values = sh.getDataRange().getValues(), headers = values[0], idCol = headers.indexOf('id');
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) if (String(values[i][idCol]) === String(row.id)) { rowIndex = i + 1; break; }
  const newRow = headers.map(function (h) { return (h in row) ? row[h] : ''; });
  if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]); else sh.appendRow(newRow);
  bump_();
  return { ok: true };
}
function legacyDeleteRow_(id, userName) {
  const u = readUsers_().filter(function (x) { return norm_(x.n) === norm_(userName); })[0];
  if (!u || !u.act || u.rol !== 'admin') throw new Error('Solo Admin puede borrar');
  const sh = dataSheet_(), values = sh.getDataRange().getValues(), idCol = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) if (String(values[i][idCol]) === String(id)) { sh.deleteRow(i + 1); bump_(); return { ok: true }; }
  return { ok: false, error: 'No encontrado' };
}

// ============================================================
// PREPARACIÓN (ejecutar UNA VEZ a mano desde el editor)
// ============================================================
// Crea las pestañas Config y Bitacora, la carpeta de fotos y pide los permisos.
function prepararHoja() {
  cfgSheet_(); logSheet_(); dataSheet_(); usersSheet_();
  const folder = getOrCreateFolder(PHOTO_FOLDER);
  const h = dataSheet_().getRange(1, 1, 1, Math.max(1, dataSheet_().getLastColumn())).getValues()[0].map(String);
  Logger.log('Pestañas listas. Carpeta de fotos: ' + folder.getName() + ' · Columnas de hallazgos: ' + h.length);
  const u = userInfo_();
  Logger.log('Usuarios: columna de nombre = ' + (u.n + 1) + ', columna de nivel = ' + (u.rol + 1) + ' (' + u.v[0][u.rol] + ')');
  Logger.log('Usuarios leídos: ' + readUsers_().map(function (x) { return x.n + ' (' + x.rol + (x.act ? '' : ', baja') + ')'; }).join(', '));
}
function autorizarDrive() { prepararHoja(); }   // nombre anterior, por compatibilidad

// ============================================================
// CORREO · RESUMEN DIARIO
// ============================================================
// 1) Ajusta CORREO_DESTINATARIOS y CORREO_HORA arriba.
// 2) Ejecuta probarCorreo() una vez (te pedirá permiso para enviar correos).
// 3) Ejecuta crearDisparadorDiario() una vez. Para apagarlo: quitarDisparadorDiario().
function enviarResumenDiario() { enviarResumen_(false); }
function probarCorreo() { enviarResumen_(true); }

function enviarResumen_(forzar) {
  if (!forzar && CORREO_SOLO_DIAS_HABILES) {
    const d = new Date().getDay();
    if (d === 0 || d === 6) return;
  }
  const hoy = todayStr_();
  const en = new Date(); en.setDate(en.getDate() + DIAS_AVISO);
  const limite = Utilities.formatDate(en, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const ay = new Date(); ay.setDate(ay.getDate() - 1);
  const ayer = Utilities.formatDate(ay, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');

  const rows = readAllRows_().filter(function (r) { return !r.borrado && !r.borr && r.desc !== ''; });
  const activo = function (r) { return r.est !== 'Cerrado' && r.est !== 'Cancelado'; };
  const act = rows.filter(activo);
  const dias = function (a, b) { return Math.round((Date.parse(b + 'T12:00:00Z') - Date.parse(a + 'T12:00:00Z')) / 864e5); };
  const orden = { Alta: 0, Media: 1, Baja: 2 };

  const vencidos = act.filter(function (r) { return r.comp && r.comp < hoy; })
    .sort(function (a, b) { return (orden[a.crit] - orden[b.crit]) || (a.comp < b.comp ? -1 : 1); });
  const proximos = act.filter(function (r) { return r.comp && r.comp >= hoy && r.comp <= limite; })
    .sort(function (a, b) { return a.comp < b.comp ? -1 : 1; });
  const entVenc = act.filter(function (r) { return r.fentrega && r.fentrega < hoy; });
  const entProx = act.filter(function (r) { return r.fentrega && r.fentrega >= hoy && r.fentrega <= limite; });
  const cerrAyer = rows.filter(function (r) { return r.est === 'Cerrado' && (r.fcierre === ayer || r.fcierre === hoy); });

  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const th = 'style="text-align:left;padding:5px 8px;background:#173a63;color:#fff;font-size:12px"';
  const td = 'style="padding:5px 8px;border-bottom:1px solid #ddd;font-size:12px;vertical-align:top"';
  const tabla = function (lista, cols, max) {
    if (!lista.length) return '<p style="color:#666;font-size:12px">Ninguno.</p>';
    const filas = lista.slice(0, max).map(function (r) {
      return '<tr>' + cols.map(function (c) { return '<td ' + td + '>' + c.f(r) + '</td>'; }).join('') + '</tr>';
    }).join('');
    return '<table style="border-collapse:collapse;width:100%"><tr>' + cols.map(function (c) { return '<th ' + th + '>' + c.t + '</th>'; }).join('') + '</tr>' + filas + '</table>' +
      (lista.length > max ? '<p style="color:#666;font-size:12px">… y ' + (lista.length - max) + ' más. Véalos en la app.</p>' : '');
  };
  const cId = { t: 'ID', f: function (r) { return esc(r.id); } };
  const cDesc = { t: 'Hallazgo', f: function (r) { return esc(String(r.desc).slice(0, 90)); } };
  const cArea = { t: 'Área', f: function (r) { return esc(r.area || '—'); } };
  const cResp = { t: 'Responsable', f: function (r) { return esc(r.resp || '—'); } };
  const cCrit = { t: 'Criticidad', f: function (r) { return esc(r.crit || ''); } };
  const cComp = { t: 'Compromiso', f: function (r) { return esc(r.comp); } };
  const cAtr  = { t: 'Atraso', f: function (r) { return dias(r.comp, hoy) + ' d'; } };
  const cEntr = { t: 'Entrega contratista', f: function (r) { return esc(r.contratista || '—') + ' · ' + esc(r.fentrega); } };
  const cReq  = { t: 'Requisición', f: function (r) { return esc(r.req || '—'); } };
  const kpi = function (n, t, rojo) {
    return '<td style="padding:10px 14px;text-align:center;border:1px solid #ddd"><div style="font-size:22px;font-weight:700;color:' + (rojo && n ? '#c0392b' : '#173a63') + '">' + n + '</div><div style="font-size:11px;color:#555">' + t + '</div></td>';
  };

  const html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:760px">' +
    '<h2 style="color:#173a63;margin:0 0 4px">Hallazgos de planta · resumen del ' + esc(hoy) + '</h2>' +
    '<p style="color:#555;margin:0 0 12px;font-size:12px">Belticos · Mantenimiento</p>' +
    '<table style="border-collapse:collapse;margin-bottom:14px"><tr>' +
    kpi(act.length, 'activos') + kpi(vencidos.length, 'vencidos', true) + kpi(proximos.length, 'vencen en ' + DIAS_AVISO + ' días') +
    kpi(entVenc.length, 'entregas de contratista vencidas', true) + kpi(cerrAyer.length, 'cerrados ayer/hoy') + '</tr></table>' +
    '<h3 style="color:#c0392b;margin:16px 0 4px">Vencidos (' + vencidos.length + ')</h3>' + tabla(vencidos, [cId, cDesc, cArea, cResp, cCrit, cComp, cAtr], 40) +
    '<h3 style="color:#b86e0c;margin:16px 0 4px">Próximos a vencer (' + proximos.length + ')</h3>' + tabla(proximos, [cId, cDesc, cArea, cResp, cCrit, cComp], 40) +
    '<h3 style="color:#c0392b;margin:16px 0 4px">Entregas de contratista vencidas (' + entVenc.length + ')</h3>' + tabla(entVenc, [cId, cDesc, cReq, cEntr, cResp], 30) +
    '<h3 style="color:#b86e0c;margin:16px 0 4px">Entregas de contratista próximas (' + entProx.length + ')</h3>' + tabla(entProx, [cId, cDesc, cReq, cEntr], 30) +
    (APP_URL ? '<p style="margin-top:18px"><a href="' + esc(APP_URL) + '" style="background:#f2c200;color:#1b2330;padding:9px 16px;text-decoration:none;font-weight:700;border-radius:5px">Abrir la app</a></p>' : '') +
    '<p style="color:#888;font-size:10px;margin-top:18px">Correo automático · Powered by Marcos Carrillo</p></div>';

  const para = String(CORREO_DESTINATARIOS || '').trim() || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: para,
    subject: 'Hallazgos de planta · ' + hoy + ' · ' + vencidos.length + ' vencido(s)' + (entVenc.length ? ' · ' + entVenc.length + ' entrega(s) tarde' : ''),
    htmlBody: html,
    name: 'Hallazgos de Planta'
  });
}

function crearDisparadorDiario() {
  quitarDisparadorDiario();
  ScriptApp.newTrigger('enviarResumenDiario').timeBased().everyDays(1).atHour(CORREO_HORA).create();
  Logger.log('Listo: el resumen se enviará todos los días cerca de las ' + CORREO_HORA + ':00.');
}
function quitarDisparadorDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarResumenDiario') ScriptApp.deleteTrigger(t);
  });
}
