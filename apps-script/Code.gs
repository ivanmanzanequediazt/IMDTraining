/**
 * IMDTraining · sincronización con Google Sheets
 *
 * Este programa hace de "portero" entre las apps y la hoja "IMDTraining datos":
 * recibe los entrenos de cada móvil, los guarda en la hoja y devuelve a cada uno
 * solo lo que le corresponde. Sin un código válido no da nada.
 *
 * Funciones que se pueden ejecutar a mano desde el editor (desplegable + Ejecutar):
 *   - setup                     Prepara la hoja. Se puede repetir: no borra nada.
 *   - regenerarTokenEntrenador  Móvil perdido: crea un código de entrenador nuevo
 *                               (el anterior deja de valer).
 *
 * CÓMO PUBLICAR UN CAMBIO DE ESTE PROGRAMA SIN ROMPER LAS APPS:
 *   Implementar > Gestionar implementaciones > lápiz (Editar) > Versión: "Nueva versión" > Implementar.
 *   NUNCA "Implementar > Nueva implementación": cambia la dirección y las apps dejan de sincronizar.
 *
 * Especificación: PLAN_SINCRONIZACION_GOOGLE_SHEETS.md, sección 4.
 */

const STORES = ['profiles', 'sessions', 'strength', 'tests', 'plans', 'wellness']; // 'tracks' va aparte
const ALL_STORES = STORES.concat(['tracks']);
const TABS = ['config', 'athletes', 'profiles', 'sessions', 'tracks', 'strength', 'tests', 'plans', 'wellness'];
const PROFILE_FIELDS = ['hrMax', 'hrMaxEstimated', 'hrRest', 'weight', 'method', 'percents', 'manual'];

const PART = 45000;         // caracteres de data por celda (límite de Sheets: 50 000)
const PAGE_RECORDS = 500;   // registros por página de pull
const PAGE_CHARS = 4e6;     // caracteres de data por página de pull
const MAX_PUSH = 500;       // registros por push (el cliente manda como mucho 100)
const LOCK_MS = 20000;      // espera máxima del bloqueo
const READ_BATCH = 100;     // registros cuya data se lee de una vez en pull
const READ_GAP = 10;        // filas de hueco que se leen de más para juntar lecturas

const RECORD_HEADERS = ['id', 'userId', 'author', 'updatedAt', 'serverAt', 'deleted', 'data'];
const RECORD_FORMATS = ['@', '@', '@', '0', '0', 'General', '@'];
const HEADERS = {
  config: ['key', 'value'],
  athletes: ['id_unico', 'token', 'createdAt', 'revoked'],
  tracks: ['id', 'userId', 'part', 'parts', 'updatedAt', 'serverAt', 'deleted', 'data'],
};
// '@' = texto plano, para que Sheets no convierta ids como 2024-05-01 o 1e5 en fechas o números
const FORMATS = {
  config: ['@', '@'],
  athletes: ['@', '@', '0', 'General'],
  tracks: ['@', '@', '0', '0', '0', '0', 'General', '@'],
};
STORES.forEach(function (s) { HEADERS[s] = RECORD_HEADERS; FORMATS[s] = RECORD_FORMATS; });

/* ================= Funciones para ejecutar a mano ================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  if (props.getProperty('lastServerAt') === null) props.setProperty('lastServerAt', '0');

  TABS.forEach(function (name) {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    const headers = HEADERS[name];
    if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    FORMATS[name].forEach(function (f, i) {
      sheet.getRange(1, i + 1, sheet.getMaxRows(), 1).setNumberFormat(f);
    });
  });

  SS = ss;
  const config = ss.getSheetByName('config');
  let rows = readConfig_();
  if (!rows.coachToken) config.appendRow(['coachToken', '']);
  if (!rows.version) config.appendRow(['version', '1']);
  rows = readConfig_();
  if (!rows.coachToken.value) setConfig_('coachToken', Utilities.getUuid());
  SpreadsheetApp.flush();
  Logger.log('Código de entrenador: ' + readConfig_().coachToken.value);
}

function regenerarTokenEntrenador() {
  const res = withLock_(function () {
    const token = Utilities.getUuid();
    setConfig_('coachToken', token);
    return { ok: true, coachToken: token };
  });
  if (!res.ok) throw new Error('La hoja está ocupada. Vuelve a pulsar Ejecutar en unos segundos.');
  Logger.log('Código de entrenador nuevo: ' + res.coachToken + ' (el anterior ya no vale)');
}

/* ================= Puerta de entrada ================= */

function doGet(e) {
  return json_({ ok: true });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!body || typeof body !== 'object') return json_({ ok: false, error: 'bad_request' });
  try {
    const who = identify_(body.token);
    if (!who) return json_({ ok: false, error: 'token' });
    switch (body.action) {
      case 'ping': return json_(ping_(who));
      case 'push': return json_(withLock_(function () { return push_(who, body.records); }));
      case 'pull': return json_(withLock_(function () { return pull_(who, Number(body.since) || 0); }));
      case 'invite':
      case 'revoke':
      case 'rotateCoachToken':
        if (who.role !== 'coach') return json_({ ok: false, error: 'forbidden' });
        return json_(withLock_(function () { return COACH_ACTIONS[body.action](body); }));
      default: return json_({ ok: false, error: 'bad_request' });
    }
  } catch (err) {
    console.error(err); // visible en "Ejecuciones" del editor
    return json_({ ok: false, error: 'server', detail: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_MS)) return { ok: false, error: 'busy' };
  try {
    const res = fn();
    SpreadsheetApp.flush(); // que lo escrito quede visible antes de soltar el bloqueo
    return res;
  } finally {
    lock.releaseLock();
  }
}

/* ================= Quién llama y qué puede tocar ================= */

// {role:'coach'} | {role:'athlete', id} | null. El rol sale siempre del token, nunca de la petición.
function identify_(token) {
  if (typeof token !== 'string' || !token) return null;
  const coach = readConfig_().coachToken;
  if (coach && coach.value && token === coach.value) return { role: 'coach' };
  const a = readAthletes_().byToken[token];
  if (!a || a.revoked) return null;
  return { role: 'athlete', id: a.id };
}

function canAccess_(who, userId, athletes) {
  if (who.role === 'athlete') return userId === who.id;
  const a = athletes.byId[userId];
  return !!a && !a.revoked;
}

function ping_(who) {
  const res = { ok: true, role: who.role };
  if (who.role === 'athlete') res.id_unico = who.id;
  const version = readConfig_().version;
  res.serverTime = Date.now();
  res.version = Number(version && version.value) || 1;
  return res;
}

/* ================= push ================= */

function push_(who, records) {
  if (!Array.isArray(records) || records.length > MAX_PUSH) return { ok: false, error: 'bad_request' };
  const athletes = readAthletes_();
  const meta = readAllMeta_();
  const writer = newWriter_();
  const clock = newClock_();
  const accepted = [];
  const rejected = [];

  records.forEach(function (rec) {
    const key = keyOf_(rec);
    if (!valid_(rec)) { rejected.push(withReason_(key, 'invalid')); return; }
    const old = meta[rec.store][rec.id];
    if (!canAccess_(who, rec.userId, athletes) ||
        (rec.store === 'profiles' && rec.id !== rec.userId) ||
        (old && old.userId !== rec.userId)) {           // no se puede "adueñar" de un registro de otro
      rejected.push(withReason_(key, 'forbidden'));
      return;
    }
    if (old && old.updatedAt >= rec.updatedAt) {
      const r = withReason_(key, 'older');
      r.current = readRecord_(old, meta);
      rejected.push(r);
      return;
    }
    let data = '';
    if (!rec.deleted) {
      if (rec.store === 'profiles') data = JSON.stringify(pick_(rec.data, PROFILE_FIELDS));
      else if (rec.store === 'tracks') data = String(rec.data);
      else data = JSON.stringify(rec.data === undefined ? null : rec.data);
    }
    if (rec.store !== 'tracks' && data.length > PART) { rejected.push(withReason_(key, 'too_big')); return; }

    const serverAt = clock.next();
    const author = old ? old.author : who.role;      // quién lo creó; no cambia al actualizar
    const m = rec.store === 'tracks'
      ? writeTrack_(writer, old, rec, serverAt, data)
      : writeRecord_(writer, old, rec, author, serverAt, data);
    m.cached = toRecord_(m, data, meta);              // por si el mismo lote lo vuelve a tocar
    meta[rec.store][rec.id] = m;
    accepted.push({ store: rec.store, id: rec.id, serverAt: serverAt });
  });

  clock.save();
  writer.flush();
  return { ok: true, accepted: accepted, rejected: rejected };
}

function valid_(r) {
  if (!r || typeof r !== 'object') return false;
  if (ALL_STORES.indexOf(r.store) < 0 || !okId_(r.id) || !okId_(r.userId)) return false;
  if (typeof r.updatedAt !== 'number' || !isFinite(r.updatedAt)) return false;
  if (r.store === 'tracks' && !r.deleted && typeof r.data !== 'string') return false;
  return true;
}

function okId_(s) {
  return typeof s === 'string' && s.length > 0 && s.length <= 64;
}

function keyOf_(r) {
  return {
    store: r && typeof r.store === 'string' ? r.store : null,
    id: r && typeof r.id === 'string' ? r.id : null,
  };
}

function withReason_(key, reason) {
  return { store: key.store, id: key.id, reason: reason };
}

function pick_(obj, fields) {
  const out = {};
  if (!obj || typeof obj !== 'object') return out;
  fields.forEach(function (f) { if (Object.prototype.hasOwnProperty.call(obj, f)) out[f] = obj[f]; });
  return out;
}

function writeRecord_(writer, old, rec, author, serverAt, data) {
  const row = old ? old.row : writer.newRow(rec.store);
  const deleted = !!rec.deleted;
  writer.set(rec.store, row, [rec.id, rec.userId, author, rec.updatedAt, serverAt, deleted, data]);
  return { store: rec.store, id: rec.id, row: row, userId: rec.userId, author: author,
           updatedAt: rec.updatedAt, serverAt: serverAt, deleted: deleted };
}

// El track se trocea en partes de PART caracteres; todas comparten serverAt.
// Un tombstone es una sola fila con parts=1, deleted=TRUE y data vacía.
function writeTrack_(writer, old, rec, serverAt, data) {
  const deleted = !!rec.deleted;
  const chunks = [];
  for (let i = 0; i < data.length; i += PART) chunks.push(data.slice(i, i + PART));
  if (!chunks.length) chunks.push('');
  const oldRows = old ? old.rows : [];
  const rows = chunks.map(function (c, i) { return i < oldRows.length ? oldRows[i] : writer.newRow('tracks'); });
  oldRows.slice(chunks.length).forEach(function (r) { writer.remove('tracks', r); });
  rows.forEach(function (r, i) {
    writer.set('tracks', r, [rec.id, rec.userId, i, chunks.length, rec.updatedAt, serverAt, deleted, chunks[i]]);
  });
  return { store: 'tracks', id: rec.id, rows: rows, userId: rec.userId,
           updatedAt: rec.updatedAt, serverAt: serverAt, deleted: deleted };
}

/* ================= pull ================= */

function pull_(who, since) {
  const athletes = readAthletes_();
  const meta = readAllMeta_();
  const candidates = [];
  ALL_STORES.forEach(function (s) {
    Object.keys(meta[s]).forEach(function (id) {
      const m = meta[s][id];
      if (m.serverAt > since && canAccess_(who, m.userId, athletes)) candidates.push(m);
    });
  });
  candidates.sort(function (a, b) { return a.serverAt - b.serverAt; });

  const out = [];
  let chars = 0;
  let i = 0;
  while (i < candidates.length && out.length < PAGE_RECORDS && chars < PAGE_CHARS) {
    const batch = candidates.slice(i, i + READ_BATCH);
    const raws = readRaws_(batch);                    // solo la data de estas filas
    for (let j = 0; j < batch.length && out.length < PAGE_RECORDS && chars < PAGE_CHARS; j++) {
      out.push(toRecord_(batch[j], raws[j], meta));
      chars += raws[j].length;
      i++;
    }
  }
  return {
    ok: true,
    cursor: out.length ? out[out.length - 1].serverAt : since,
    more: out.length < candidates.length,
    records: out,
  };
}

/* ================= invite / revoke / rotateCoachToken ================= */

const COACH_ACTIONS = {
  invite: function (body) {
    const athletes = readAthletes_();
    if (body.id_unico !== undefined) {                // reenviar: devuelve el token que ya tiene
      const a = athletes.byId[String(body.id_unico)];
      if (!a || a.revoked) return { ok: false, error: 'bad_request', detail: 'id_unico desconocido o desvinculado' };
      return { ok: true, id_unico: a.id, token: a.token };
    }
    let id;
    do { id = 'd' + Utilities.getUuid().replace(/-/g, '').slice(0, 8); } while (athletes.byId[id]);
    const token = Utilities.getUuid();
    const writer = newWriter_();
    writer.set('athletes', writer.newRow('athletes'), [id, token, Date.now(), false]);
    writer.flush();
    return { ok: true, id_unico: id, token: token };
  },

  revoke: function (body) {
    const a = readAthletes_().byId[String(body.id_unico)];
    if (!a) return { ok: false, error: 'bad_request', detail: 'id_unico desconocido' };
    sheet_('athletes').getRange(a.row, 4).setValue(true);
    return { ok: true };
  },

  rotateCoachToken: function () {
    const token = Utilities.getUuid();
    setConfig_('coachToken', token);
    return { ok: true, coachToken: token };
  },
};

/* ================= Lectura de la hoja ================= */

let SS = null;

function sheet_(name) {
  if (!SS) {
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) throw new Error('Falta ejecutar setup');
    SS = SpreadsheetApp.openById(id);
  }
  const sheet = SS.getSheetByName(name);
  if (!sheet) throw new Error('Falta la pestaña ' + name + ': ejecuta setup');
  return sheet;
}

function bool_(v) {
  return v === true || v === 'TRUE' || v === 'true';
}

// {key: {row, value}}
function readConfig_() {
  const sheet = sheet_('config');
  const out = {};
  const n = sheet.getLastRow() - 1;
  if (n < 1) return out;
  sheet.getRange(2, 1, n, 2).getValues().forEach(function (v, i) {
    if (v[0] !== '') out[String(v[0])] = { row: i + 2, value: String(v[1]) };
  });
  return out;
}

function setConfig_(key, value) {
  const sheet = sheet_('config');
  const row = readConfig_()[key];
  if (row) sheet.getRange(row.row, 2).setValue(value);
  else sheet.appendRow([key, value]);
}

function readAthletes_() {
  const sheet = sheet_('athletes');
  const out = { byId: {}, byToken: {} };
  const n = sheet.getLastRow() - 1;
  if (n < 1) return out;
  sheet.getRange(2, 1, n, 4).getValues().forEach(function (v, i) {
    if (v[0] === '') return;
    const a = { row: i + 2, id: String(v[0]), token: String(v[1]), revoked: bool_(v[3]) };
    out.byId[a.id] = a;
    if (a.token) out.byToken[a.token] = a;
  });
  return out;
}

function readAllMeta_() {
  const meta = {};
  ALL_STORES.forEach(function (s) { meta[s] = readMeta_(s); });
  return meta;
}

// Metadatos sin la columna data: id -> {store, id, row | rows, userId, author, updatedAt, serverAt, deleted}
function readMeta_(store) {
  const sheet = sheet_(store);
  const out = {};
  const n = sheet.getLastRow() - 1;
  if (n < 1) return out;
  if (store === 'tracks') {
    const parts = {};
    sheet.getRange(2, 1, n, 7).getValues().forEach(function (v, i) {
      const id = String(v[0]);
      if (!id) return;
      const part = Number(v[2]) || 0;
      if (!out[id] || part === 0) {
        out[id] = { store: store, id: id, userId: String(v[1]), updatedAt: Number(v[4]),
                    serverAt: Number(v[5]), deleted: bool_(v[6]) };
      }
      (parts[id] = parts[id] || []).push({ row: i + 2, part: part });
    });
    Object.keys(out).forEach(function (id) {
      out[id].rows = parts[id].sort(function (a, b) { return a.part - b.part; }).map(function (p) { return p.row; });
    });
    return out;
  }
  sheet.getRange(2, 1, n, 6).getValues().forEach(function (v, i) {
    const id = String(v[0]);
    if (!id) return;
    out[id] = { store: store, id: id, row: i + 2, userId: String(v[1]), author: String(v[2]),
                updatedAt: Number(v[3]), serverAt: Number(v[4]), deleted: bool_(v[5]) };
  });
  return out;
}

// Cadena de data de cada metadato (tracks: partes ya unidas). '' si es tombstone.
function readRaws_(metas) {
  const rowsBy = {};
  metas.forEach(function (m) {
    if (m.cached || m.deleted) return;
    const list = rowsBy[m.store] = rowsBy[m.store] || [];
    if (m.store === 'tracks') m.rows.forEach(function (r) { list.push(r); });
    else list.push(m.row);
  });
  const cells = {};
  Object.keys(rowsBy).forEach(function (s) {
    cells[s] = readRows_(sheet_(s), s === 'tracks' ? 8 : 7, rowsBy[s]);
  });
  return metas.map(function (m) {
    if (m.deleted) return '';
    if (m.cached) return typeof m.cached.data === 'string' ? m.cached.data : JSON.stringify(m.cached.data);
    if (m.store === 'tracks') return m.rows.map(function (r) { return String(cells.tracks[r]); }).join('');
    return String(cells[m.store][m.row]);
  });
}

// Lee la columna col de las filas pedidas, juntando filas cercanas en una sola lectura
function readRows_(sheet, col, rows) {
  const out = {};
  if (!rows.length) return out;
  const sorted = rows.slice().sort(function (a, b) { return a - b; });
  let start = sorted[0];
  let prev = sorted[0];
  const readRun = function (a, b) {
    sheet.getRange(a, col, b - a + 1, 1).getValues().forEach(function (v, k) { out[a + k] = v[0]; });
  };
  for (let k = 1; k < sorted.length; k++) {
    if (sorted[k] - prev > READ_GAP) { readRun(start, prev); start = sorted[k]; }
    prev = sorted[k];
  }
  readRun(start, prev);
  return out;
}

function readRecord_(m, meta) {
  if (m.cached) return m.cached;
  return toRecord_(m, readRaws_([m])[0], meta);
}

function toRecord_(m, raw, meta) {
  const isTrack = m.store === 'tracks';
  let data = null;
  if (!m.deleted) {
    if (isTrack) data = raw;
    else { try { data = JSON.parse(raw); } catch (err) { data = null; } }
  }
  // La pestaña tracks no tiene columna author: es el de su sesión (mismo id)
  const session = isTrack ? meta.sessions[m.id] : null;
  return {
    store: m.store, id: m.id, userId: m.userId,
    author: isTrack ? (session ? session.author : null) : m.author,
    updatedAt: m.updatedAt, serverAt: m.serverAt, deleted: m.deleted, data: data,
  };
}

/* ================= Escritura en bloque ================= */

// Acumula filas y las escribe al final en tramos contiguos (una llamada por tramo).
// Los borrados de filas sobrantes de tracks van al final y de abajo arriba.
function newWriter_() {
  const sheets = {};
  const get = function (name) {
    if (!sheets[name]) {
      const sheet = sheet_(name);
      sheets[name] = { sheet: sheet, next: sheet.getLastRow() + 1, rows: {}, del: {} };
    }
    return sheets[name];
  };
  return {
    newRow: function (name) { return get(name).next++; },
    set: function (name, row, values) { get(name).rows[row] = values; },
    remove: function (name, row) { const s = get(name); delete s.rows[row]; s.del[row] = true; },
    flush: function () {
      Object.keys(sheets).forEach(function (name) {
        const s = sheets[name];
        const formats = FORMATS[name];
        const rows = Object.keys(s.rows).map(Number).sort(function (a, b) { return a - b; });
        if (rows.length) {
          const last = rows[rows.length - 1];
          const max = s.sheet.getMaxRows();
          if (last > max) s.sheet.insertRowsAfter(max, last - max);
        }
        let k = 0;
        while (k < rows.length) {
          let end = k;
          while (end + 1 < rows.length && rows[end + 1] === rows[end] + 1) end++;
          const values = rows.slice(k, end + 1).map(function (r) { return s.rows[r]; });
          const range = s.sheet.getRange(rows[k], 1, values.length, formats.length);
          range.setNumberFormats(values.map(function () { return formats; }));
          range.setValues(values);
          k = end + 1;
        }
        Object.keys(s.del).map(Number).sort(function (a, b) { return b - a; })
          .forEach(function (r) { s.sheet.deleteRow(r); });
      });
    },
  };
}

// serverAt estrictamente creciente y único. Solo se usa dentro del bloqueo.
function newClock_() {
  const props = PropertiesService.getScriptProperties();
  let last = Number(props.getProperty('lastServerAt') || 0);
  return {
    next: function () { last = Math.max(Date.now(), last + 1); return last; },
    save: function () { props.setProperty('lastServerAt', String(last)); },
  };
}
