/**
 * IMDTraining · diagnóstico de la sincronización (solo lee, no cambia nada)
 *
 * Ejecutar a mano: desplegable de funciones > diagnosticoSync > Ejecutar.
 * El resultado sale abajo, en "Registro de ejecución". No muestra contraseñas ni nombres.
 */
function diagnosticoSync() {
  const athletes = readAthletes_();
  const apodos = {};
  try {
    const dep = sheet_(DEP_TAB);
    const n = dep.getLastRow() - 1;
    if (n > 0) dep.getRange(2, 1, n, 2).getValues().forEach(function (v) { if (v[0]) apodos[String(v[0])] = String(v[1] || ''); });
  } catch (e) { /* sin pestaña deportistas */ }

  const stats = {};
  const get = function (id) {
    return stats[id] || (stats[id] = { sessions: 0, strength: 0, tests: 0, plans: 0, wellness: 0, profiles: 0, last: 0, lastSrc: '' });
  };
  ['profiles', 'sessions', 'strength', 'tests', 'plans', 'wellness'].forEach(function (s) {
    const meta = readMeta_(s);
    Object.keys(meta).forEach(function (id) {
      const m = meta[id];
      if (m.deleted) return;
      const st = get(m.userId);
      st[s]++;
      if ((s === 'sessions' || s === 'strength') && m.updatedAt > st.last) { st.last = m.updatedAt; st.lastSrc = s === 'strength' ? 'fuerza' : 'aeróbico'; }
    });
  });

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const fmt = function (t) { return t ? Utilities.formatDate(new Date(t), tz, 'dd/MM/yyyy HH:mm') : '—'; };
  const ids = Object.keys(athletes.byId).sort(function (a, b) { return athletes.byId[a].row - athletes.byId[b].row; });
  Object.keys(stats).forEach(function (id) { if (ids.indexOf(id) < 0) ids.push(id); });
  const lines = ['CÓDIGO       APODO        ESTADO          FUERZA AERÓB. TEST PLANES BIENEST.  ÚLTIMO ENTRENO SUBIDO'];
  ids.forEach(function (id) {
    const a = athletes.byId[id];
    const st = stats[id] || get(id);
    const estado = !a ? 'sin invitación' : a.revoked ? 'desvinculado' : 'activo';
    const pad = function (s, n) { s = String(s); return s + Array(Math.max(1, n - s.length + 1)).join(' '); };
    lines.push(pad(id, 13) + pad(apodos[id] || '—', 13) + pad(estado, 16) + pad(st.strength, 7) + pad(st.sessions, 7) +
      pad(st.tests, 5) + pad(st.plans, 7) + pad(st.wellness, 10) + fmt(st.last) + (st.lastSrc ? ' (' + st.lastSrc + ')' : ''));
  });
  Logger.log('\n' + lines.join('\n'));
}

/**
 * Pasa los entrenos de un código antiguo al código nuevo del mismo deportista (no borra nada).
 * Solo cambia de quién son (userId); el contenido y la fecha de cada registro no se tocan.
 * Si un registro de bienestar ya existe con el código nuevo para ese día, se deja el del código nuevo.
 */
function moverDatos_(from, to) {
  const athletes = readAthletes_();
  if (!athletes.byId[from]) throw new Error('No existe el código antiguo ' + from);
  if (!athletes.byId[to] || athletes.byId[to].revoked) throw new Error('El código nuevo ' + to + ' no existe o está desvinculado');
  const res = withLock_(function () {
    const clock = newClock_();
    const out = {};
    ['sessions', 'strength', 'tests', 'plans', 'wellness', 'tracks'].forEach(function (s) {
      const sh = sheet_(s);
      const n = sh.getLastRow() - 1;
      out[s] = { moved: 0, skipped: 0 };
      if (n < 1) return;
      const cols = s === 'tracks' ? 8 : 7;
      const vals = sh.getRange(2, 1, n, cols).getValues();
      const ids = {};
      vals.forEach(function (v) { ids[String(v[0])] = true; });
      const trackAt = {};
      vals.forEach(function (v, i) {
        if (String(v[1]) !== from) return;
        if (s === 'wellness') {
          const nid = to + String(v[0]).slice(from.length);
          if (ids[nid]) { out[s].skipped++; return; }
          v[0] = nid;
          ids[nid] = true;
        }
        v[1] = to;
        if (s === 'tracks') { const id = String(v[0]); v[5] = trackAt[id] || (trackAt[id] = clock.next()); }
        else v[4] = clock.next();
        sh.getRange(i + 2, 1, 1, cols).setValues([v]);
        out[s].moved++;
      });
    });
    clock.save();
    return { ok: true, out: out };
  });
  if (!res.ok) throw new Error('La hoja está ocupada. Vuelve a pulsar Ejecutar en unos segundos.');
  Logger.log('Movido de ' + from + ' a ' + to + ':\n' + Object.keys(res.out).map(function (s) {
    return '  ' + s + ': ' + res.out[s].moved + ' movidos' + (res.out[s].skipped ? ', ' + res.out[s].skipped + ' ya estaban en el nuevo' : '');
  }).join('\n'));
}

// Para usarla, crea una función como esta con los dos códigos (antiguo, nuevo) y ejecútala una vez:
// function moverDatosDeportista() { moverDatos_('dCODIGOVIEJO', 'dCODIGONUEVO'); }

/* ================= Limpieza: empezar de cero salvo lo que se elija ================= */
//
// Paso 1 · limpiezaPaso1Preparar: crea la pestaña «limpieza» con un código por fila y lo que tiene cada uno.
//          En la columna «Qué hacer» eliges para cada código: Mantener todo · Solo aeróbico · Vaciar.
// Paso 2 · limpiezaPaso2Ejecutar: hace lo elegido. Antes, haz una copia de la hoja (Archivo > Hacer una copia).
//   - Vaciar: se borran sus entrenos de fuerza, aeróbicos, test, programados y bienestar. Sigue vinculado.
//   - Solo aeróbico: se borra todo menos las sesiones aeróbicas (con su registro de pulsómetro).
//   - Mantener todo: no se toca.
// El borrado llega a las apps con la sincronización (también a los móviles de los deportistas).
// Además quita de «programacion» las filas de quien no sea «Mantener todo» (si no, se volverían a publicar)
// y desmarca «Importar» en «fuentes» para sus Excel.

const LIMPIEZA_TAB = 'limpieza';
const LIMPIEZA_OPC = ['Mantener todo', 'Solo aeróbico', 'Vaciar'];

function limpiezaPaso1Preparar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const athletes = readAthletes_();
  const apodos = {};
  try {
    const dep = sheet_(DEP_TAB);
    const n = dep.getLastRow() - 1;
    if (n > 0) dep.getRange(2, 1, n, 2).getValues().forEach(function (v) { if (v[0]) apodos[String(v[0]).trim()] = String(v[1] || ''); });
  } catch (e) { /* sin pestaña deportistas */ }
  const stats = {};
  const get = function (id) { return stats[id] || (stats[id] = { strength: 0, sessions: 0, tests: 0, plans: 0, wellness: 0 }); };
  ['sessions', 'strength', 'tests', 'plans', 'wellness'].forEach(function (s) {
    const meta = readMeta_(s);
    Object.keys(meta).forEach(function (id) { if (!meta[id].deleted) get(meta[id].userId)[s]++; });
  });
  const ids = Object.keys(athletes.byId).sort(function (a, b) { return athletes.byId[a].row - athletes.byId[b].row; });
  Object.keys(stats).forEach(function (id) { if (ids.indexOf(id) < 0) ids.push(id); });

  let sh = ss.getSheetByName(LIMPIEZA_TAB);
  const prev = {};
  if (sh) {
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 1, n, 9).getValues().forEach(function (v) { if (v[0]) prev[String(v[0])] = String(v[8] || ''); });
    sh.clear();
    sh.getRange('I:I').clearDataValidations();
  } else sh = ss.insertSheet(LIMPIEZA_TAB);
  const rows = ids.map(function (id) {
    const a = athletes.byId[id];
    const st = stats[id] || get(id);
    return [id, apodos[id] || '', !a ? 'sin invitación' : a.revoked ? 'desvinculado' : 'activo',
      st.strength, st.sessions, st.tests, st.plans, st.wellness, prev[id] || ''];
  });
  sh.getRange(1, 1, 1, 9).setValues([['Código', 'Apodo', 'Estado', 'Fuerza', 'Aeróbico', 'Test', 'Programados', 'Bienestar', 'Qué hacer']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, 9).setValues(rows);
    sh.getRange(2, 9, rows.length, 1).setBackground('#fff8e1').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(LIMPIEZA_OPC, true).setAllowInvalid(false).build());
  }
  sh.getRange('I1').setNote('Elige para cada código: Mantener todo · Solo aeróbico · Vaciar. Luego ejecuta limpiezaPaso2Ejecutar.');
  sh.autoResizeColumns(1, 9);
  sh.activate();
  Logger.log('Pestaña «limpieza» lista con ' + rows.length + ' códigos. Rellena la columna «Qué hacer» en todas las filas.');
}

function limpiezaPaso2Ejecutar() {
  const sh = sheet_(LIMPIEZA_TAB);
  const n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('La pestaña «limpieza» está vacía. Ejecuta antes limpiezaPaso1Preparar.');
  const accion = {};
  const faltan = [];
  sh.getRange(2, 1, n, 9).getValues().forEach(function (v) {
    const id = String(v[0]).trim();
    if (!id) return;
    const a = String(v[8]).trim();
    if (LIMPIEZA_OPC.indexOf(a) < 0) faltan.push(id + (v[1] ? ' (' + v[1] + ')' : ''));
    accion[id] = a;
  });
  if (faltan.length) throw new Error('Falta elegir «Qué hacer» en: ' + faltan.join(', ') + '. No se ha borrado nada.');

  const imp = LockService.getDocumentLock();          // que no coincida con la importación automática
  if (!imp.tryLock(60000)) throw new Error('Se está importando un Excel. Vuelve a ejecutar en un minuto. No se ha borrado nada.');
  try {
    const res = withLock_(function () { return limpiar_(accion); });
    if (!res || res.ok === false) throw new Error('La hoja está ocupada sincronizando. Vuelve a ejecutar en unos segundos. No se ha borrado nada.');
    const p = limpiarProgramacion_(accion);
    const f = limpiarFuentes_(accion);
    const c = res.count;
    Logger.log('Limpieza hecha:\n' +
      '  Borrados: ' + c.strength + ' de fuerza, ' + c.sessions + ' aeróbicos, ' + c.tests + ' test, ' + c.plans + ' programados, ' + c.wellness + ' de bienestar\n' +
      (c.unlinked ? '  ' + c.unlinked + ' aeróbicos que eran parte de una sesión de fuerza ahora salen como sesión normal\n' : '') +
      '  Pestaña programacion: ' + p + ' filas quitadas\n' +
      '  Pestaña fuentes: ' + f + ' Excel desmarcados\n' +
      'Las apps lo reciben en la próxima sincronización (abre la app o pulsa Actualizar).');
  } finally { imp.releaseLock(); }
}

// Borra en la hoja (con tombstones para que el borrado llegue a las apps) lo elegido en cada código
function limpiar_(accion) {
  const meta = readAllMeta_();
  const writer = newWriter_();
  const clock = newClock_();
  const now = Date.now();
  const count = { strength: 0, sessions: 0, tests: 0, plans: 0, wellness: 0, unlinked: 0 };
  const borra = function (s, m) {
    const a = accion[m.userId];
    if (a === 'Vaciar') return true;
    if (a === 'Solo aeróbico') return s !== 'sessions' && s !== 'tracks';
    return false;
  };
  const ts = function (m) { return Math.max(now, m.updatedAt + 1); };
  ['sessions', 'strength', 'tests', 'plans', 'wellness', 'tracks'].forEach(function (s) {
    Object.keys(meta[s]).forEach(function (id) {
      const m = meta[s][id];
      if (m.deleted || !borra(s, m)) return;
      const rec = { store: s, id: id, userId: m.userId, updatedAt: ts(m), deleted: true };
      if (s === 'tracks') writeTrack_(writer, m, rec, clock.next(), '');
      else { writeRecord_(writer, m, rec, m.author, clock.next(), ''); count[s]++; }
    });
  });
  // Aeróbicos que se quedan pero colgaban de una sesión de fuerza borrada: pasan a ser sesiones normales
  const hijos = Object.keys(meta.sessions).map(function (id) { return meta.sessions[id]; })
    .filter(function (m) { return !m.deleted && accion[m.userId] === 'Solo aeróbico'; });
  const raws = hijos.length ? readRaws_(hijos) : [];
  hijos.forEach(function (m, i) {
    let d = null;
    try { d = JSON.parse(raws[i]); } catch (e) { d = null; }
    if (!d || !d.parentId) return;
    delete d.parentId;
    writeRecord_(writer, m, { store: 'sessions', id: m.id, userId: m.userId, updatedAt: ts(m), deleted: false },
      m.author, clock.next(), JSON.stringify(d));
    count.unlinked++;
  });
  clock.save();
  writer.flush();
  return { ok: true, count: count };
}

// Quita de «programacion» las filas de los códigos que no son «Mantener todo».
// Deportista, Fecha y Sesión vacíos heredan de la fila de arriba: en las filas que se quedan se escriben si hace falta.
function limpiarProgramacion_(accion) {
  let prog;
  try { prog = sheet_(PROG_TAB); } catch (e) { return 0; }
  const last = prog.getLastRow();
  if (last < 2) return 0;
  const dep = readDeportistas_();
  const vals = prog.getRange(2, 1, last - 1, PROG_COLS).getValues();
  const vacia = function (x) { return x === '' || x === null; };
  const ef = { apodo: '', fecha: '', sesion: '' };       // valores heredados (tal cual están en la celda)
  const escrito = { apodo: null, fecha: null, sesion: null };
  const keep = [];
  let quitadas = 0;
  vals.forEach(function (v) {
    const has = v.slice(0, PC.accion + 1).some(function (x) { return x !== '' && x !== false && x !== null; });
    if (!vacia(v[PC.apodo])) ef.apodo = v[PC.apodo];
    if (!vacia(v[PC.fecha])) ef.fecha = v[PC.fecha];
    if (!vacia(v[PC.sesion])) ef.sesion = v[PC.sesion];
    if (!has) return;
    const a = dep.byApodo[norm_(ef.apodo)];
    if (a && accion[a.id] && accion[a.id] !== 'Mantener todo') { quitadas++; return; }
    const row = v.slice();
    ['apodo', 'fecha', 'sesion'].forEach(function (k) {
      if (vacia(row[PC[k]]) && String(escrito[k]) !== String(ef[k])) row[PC[k]] = ef[k];
      escrito[k] = row[PC[k]] === '' ? escrito[k] : row[PC[k]];
    });
    keep.push(row);
  });
  if (!quitadas) return 0;
  prog.getRange(2, 1, last - 1, PROG_COLS).clearContent();
  if (keep.length) prog.getRange(2, 1, keep.length, PROG_COLS).setValues(keep);
  return quitadas;
}

// Desmarca «Importar» en los Excel de quien no es «Mantener todo»
function limpiarFuentes_(accion) {
  let f;
  try { f = sheet_(FUENTES_TAB); } catch (e) { return 0; }
  const n = f.getLastRow() - 1;
  if (n < 1) return 0;
  const dep = readDeportistas_();
  let k = 0;
  f.getRange(2, 1, n, 4).getValues().forEach(function (v, i) {
    if (!bool_(v[3])) return;
    const quienes = parseQuienes_(str_(v[0])) || [];
    const fuera = quienes.some(function (q) {
      const a = dep.byApodo[norm_(q.apodo)];
      return a && accion[a.id] && accion[a.id] !== 'Mantener todo';
    });
    if (fuera) { f.getRange(i + 2, 4).setValue(false); k++; }
  });
  return k;
}
