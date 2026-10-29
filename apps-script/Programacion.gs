/**
 * IMDTraining · programar entrenamientos desde la hoja
 *
 * Iván escribe los entrenamientos en la pestaña «programacion» (una fila por ejercicio) y, al pulsar
 * IMDTraining > Publicar programación, se guardan en la pestaña «plans». Desde ahí las apps los
 * descargan solas con la sincronización de siempre (Código.gs). No hace falta volver a implementar:
 * esto se ejecuta desde la hoja, no desde las apps.
 *
 * Menú IMDTraining (aparece al abrir la hoja):
 *   - Publicar programación   Pasa a la app lo nuevo o cambiado. Se puede pulsar las veces que quieras.
 *   - Actualizar estados      Solo mira si los deportistas han hecho los entrenamientos.
 *   - Preparar pestañas       Crea «deportistas» y «programacion» si faltan. No borra nada.
 *
 * Especificación: PLAN_PROGRAMACION_DESDE_LA_HOJA.md
 */

const PROG_TAB = 'programacion';
const DEP_TAB = 'deportistas';
const PROG_HEADERS = ['Deportista', 'Fecha', 'Sesión', 'Tipo', 'Ejercicio', 'Variante', 'Unilateral', 'Pliometría',
  'Series', 'Reps', 'Kg', '%1RM', 'RPE', 'RIR', 'Intensidad', 'Minutos', 'Zona', 'Intervalos', 'Indicaciones',
  'Acción', 'Estado', 'id', 'Origen'];
const PC = {}; // índice (0..21) de cada columna por nombre corto
['apodo', 'fecha', 'sesion', 'tipo', 'ejercicio', 'variante', 'unilateral', 'plyo', 'series', 'reps', 'kg', 'pct',
  'rpe', 'rir', 'int', 'minutos', 'zona', 'intervalos', 'notas', 'accion', 'estado', 'id', 'origen']
  .forEach(function (k, i) { PC[k] = i; });
const PROG_COLS = PROG_HEADERS.length;          // 23 (Origen: filas importadas de Excel, oculta)
const INT_BY_NAME = { baja: 1, media: 2, alta: 3, b: 1, m: 2, a: 3 };
const STATUS_TXT = { pending: 'Pendiente', inprogress: 'En curso', done: 'Hecho' };

/* ================= Menú ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('IMDTraining')
    .addItem('Publicar programación', 'publicarProgramacion')
    .addItem('Actualizar estados', 'actualizarEstados')
    .addSeparator()
    .addItem('Importar de Excel', 'importarExcel')
    .addItem('Importar y publicar ahora', 'importarYPublicar')
    .addItem('Automático: activar (cada 30 min)', 'activarAutomatico')
    .addItem('Automático: desactivar', 'desactivarAutomatico')
    .addSeparator()
    .addItem('Preparar pestañas de programación', 'prepararProgramacion')
    .addToUi();
}

/* ================= Preparar las pestañas ================= */

function prepararProgramacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // deportistas: código (lo rellena el programa) + apodo (Iván)
  let dep = ss.getSheetByName(DEP_TAB);
  if (!dep) {
    dep = ss.insertSheet(DEP_TAB);
    dep.getRange(1, 1, 1, 2).setValues([['codigo', 'apodo']]).setFontWeight('bold').setBackground('#e8eaed');
    dep.setFrozenRows(1);
    dep.getRange('A:A').setNumberFormat('@').setFontColor('#5f6368');
    dep.getRange('B:B').setNumberFormat('@');
    dep.setColumnWidth(1, 130);
    dep.setColumnWidth(2, 160);
    dep.getRange('A1').setNote('El programa añade aquí el código de cada deportista vinculado. No lo cambies.\n' +
      'En la app, el código sale en el perfil de cada deportista (pestaña Datos).');
    dep.getRange('B1').setNote('Escribe un apodo o iniciales para cada código (sin nombres completos). ' +
      'Es lo que eliges en la columna Deportista de «programacion». No repitas apodos.');
  }

  // programacion
  let prog = ss.getSheetByName(PROG_TAB);
  const isNew = !prog;
  if (!prog) prog = ss.insertSheet(PROG_TAB);
  if (prog.getLastRow() === 0) prog.getRange(1, 1, 1, PROG_COLS).setValues([PROG_HEADERS]);
  if (prog.getMaxColumns() < PROG_COLS) prog.insertColumnsAfter(prog.getMaxColumns(), PROG_COLS - prog.getMaxColumns());
  if (prog.getRange(1, PC.origen + 1).getValue() === '') prog.getRange(1, PC.origen + 1).setValue('Origen');
  const head = prog.getRange(1, 1, 1, PROG_COLS);
  head.setFontWeight('bold').setBackground('#e8eaed').setWrap(true).setVerticalAlignment('middle');
  prog.setFrozenRows(1);
  prog.setFrozenColumns(3);

  const n = prog.getMaxRows() - 1;
  const col = function (k) { return prog.getRange(2, PC[k] + 1, n, 1); };
  const list = function (values) {
    return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  };
  col('apodo').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(DEP_TAB).getRange('B2:B'), true).setAllowInvalid(true).build());
  col('fecha').setNumberFormat('dd/mm/yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  col('tipo').setDataValidation(list(['Fuerza', 'Activación', 'Aeróbico']));
  col('unilateral').insertCheckboxes();
  col('plyo').insertCheckboxes();
  col('int').setDataValidation(list(['Baja', 'Media', 'Alta']));
  col('zona').setDataValidation(list(['Z1', 'Z2', 'Z3', 'Z4', 'Z5']));
  col('accion').setDataValidation(list(['Borrar']));
  ['sesion', 'ejercicio', 'variante', 'intervalos', 'notas'].forEach(function (k) { col(k).setNumberFormat('@'); });
  col('estado').setFontColor('#5f6368').setFontStyle('italic');
  col('id').setNumberFormat('@');
  col('origen').setNumberFormat('@');
  prog.hideColumns(PC.id + 1, 2);

  const widths = { apodo: 110, fecha: 95, sesion: 130, tipo: 85, ejercicio: 150, variante: 110, unilateral: 75,
    plyo: 75, series: 55, reps: 55, kg: 55, pct: 55, rpe: 50, rir: 50, int: 85, minutos: 65, zona: 60,
    intervalos: 170, notas: 220, accion: 80, estado: 190 };
  Object.keys(widths).forEach(function (k) { prog.setColumnWidth(PC[k] + 1, widths[k]); });

  // Colores de la columna Estado
  const est = prog.getRange(2, PC.estado + 1, n, 1);
  const rules = prog.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (g) { return g.getColumn() === PC.estado + 1; });
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Error')
    .setBackground('#fce8e6').setFontColor('#a50e0e').setRanges([est]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Hecho')
    .setBackground('#e6f4ea').setFontColor('#137333').setRanges([est]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✓')
    .setBackground('#e8f0fe').setFontColor('#174ea6').setRanges([est]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Borrado')
    .setFontColor('#80868b').setStrikethrough(true).setRanges([est]).build());
  prog.setConditionalFormatRules(rules);

  prog.getRange(1, PC.apodo + 1).setNote(
    'CÓMO SE PROGRAMA\n' +
    '• Una fila por ejercicio. Una sesión = filas con el mismo Deportista + Fecha + Sesión. ' +
    'El orden de las filas es el orden de los ejercicios.\n' +
    '• Deportista, Fecha y Sesión vacíos = los de la fila de arriba: escríbelos solo cuando cambien.\n' +
    '• Series distintas (pirámides): varias filas seguidas del mismo ejercicio.\n' +
    '• Kg o %1RM (si pones %1RM, la app calcula los kilos).\n' +
    '• Pliometría: marca la casilla; Reps = contactos; elige Intensidad.\n' +
    '• Aeróbico: Tipo = Aeróbico; Ejercicio = Carrera, Bicicleta…; Minutos + Zona, o Intervalos como ' +
    '«10 Z2, 3 Z4, 10 Z2» (0,5 = 30 s). Una sola fila por sesión aeróbica.\n' +
    '• Indicaciones: se usa la primera que escribas en la sesión.\n' +
    '• Al acabar: menú IMDTraining > Publicar programación.\n' +
    '• Para quitar algo ya publicado: Acción = Borrar y Publicar. Borrar la fila sin más NO lo quita de la app.\n' +
    '• Puedes copiar y pegar filas o semanas enteras.');
  prog.getRange(1, PC.estado + 1).setNote('Lo escribe el programa al publicar. No hace falta tocarlo.');

  if (isNew) {
    const d = function (days) { const x = new Date(); x.setHours(0, 0, 0, 0); x.setDate(x.getDate() + days); return x; };
    const e = function (o) { const r = PROG_HEADERS.map(function () { return ''; });
      Object.keys(o).forEach(function (k) { r[PC[k]] = o[k]; }); r[PC.unilateral] = !!o.unilateral; r[PC.plyo] = !!o.plyo; return r; };
    const rows = [
      e({ apodo: 'Ejemplo', fecha: d(3), sesion: 'Pierna A', tipo: 'Fuerza', ejercicio: 'Sentadilla', series: 4, reps: 6, pct: 80, rir: 2, notas: 'Calienta 10 min' }),
      e({ ejercicio: 'Zancada', unilateral: true, series: 3, reps: 8, kg: 12, rpe: 7 }),
      e({ ejercicio: 'Drop jump', plyo: true, series: 3, reps: 6, int: 'Alta' }),
      e({ fecha: d(5), sesion: 'Series', tipo: 'Aeróbico', ejercicio: 'Carrera', intervalos: '10 Z2, 3 Z4, 2 Z1, 3 Z4, 10 Z2' }),
    ];
    prog.getRange(2, 1, rows.length, PROG_COLS).setValues(rows);
  }

  SS = ss;
  syncDeportistas_();
  SpreadsheetApp.flush();
  ss.setActiveSheet(prog);
  toast_('Pestañas listas. Pon un apodo a cada código en «deportistas».');
}

/* ================= Publicar ================= */

function publicarProgramacion() {
  const prog = sheet_(PROG_TAB);
  const tz = prog.getParent().getSpreadsheetTimeZone();
  const res = withLock_(function () { return publish_(prog, tz, Date.now()); });
  if (!res || res.ok === false) {
    alert_('La hoja está ocupada sincronizando con las apps. Vuelve a pulsar Publicar en unos segundos.');
    return;
  }
  writeStatus_(prog, res.n, res.estados, res.ids);
  const s = res.stats;
  toast_('Sesiones: ' + s.published + ' publicada' + (s.published === 1 ? '' : 's') + ' · ' + s.same + ' sin cambios' +
    (s.deleted ? ' · ' + s.deleted + ' borrada' + (s.deleted === 1 ? '' : 's') : '') +
    (s.errors ? ' · ' + s.errors + ' con error (mira la columna Estado)' : ''));
}

// Lee la pestaña, valida y escribe en plans. Devuelve {ok, n, estados:{fila:texto}, ids:{fila:id}, stats}.
// Se ejecuta dentro del bloqueo, igual que push_.
function publish_(prog, tz, now) {
  syncDeportistas_();
  const dep = readDeportistas_();
  const athletes = readAthletes_();
  const meta = { plans: readMeta_('plans'), sessions: {} };
  const last = prog.getLastRow();
  const values = last < 1 ? [] : prog.getRange(1, 1, last, PROG_COLS).getValues();
  const fmt = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  const rows = parseRows_(values, fmt);
  const groups = groupSessions_(rows);
  const writer = newWriter_();
  const clock = newClock_();
  const estados = {};
  const ids = {};
  const stats = { published: 0, same: 0, deleted: 0, errors: 0 };
  const setAll = function (g, txt, id) {
    g.rows.forEach(function (r) { estados[r.row] = txt; ids[r.row] = id === undefined ? r.id : id; });
  };
  const live = function (id) { const m = id && meta.plans[id]; return m && !m.deleted ? m : null; };
  const tombstone = function (m) {
    meta.plans[m.id] = writeRecord_(writer, m, { store: 'plans', id: m.id, userId: m.userId, updatedAt: now, deleted: true },
      m.author, clock.next(), '');
  };
  const existingOf = function (m) { return m ? readRecord_(m, meta).data : null; };

  // 1) A qué grupo pertenece cada id (por si se han copiado filas con su id oculto)
  assignIds_(groups, function (id) {
    const m = live(id);
    if (!m) return null;
    const d = existingOf(m) || {};
    return { userId: m.userId, date: d.date, title: d.title };
  }, function (apodo) { const a = dep.byApodo[norm_(apodo)]; return a ? a.id : null; });

  groups.forEach(function (g) {
    // Borrar
    if (g.borrar) {
      const m = live(g.id);
      if (m) { tombstone(m); stats.deleted++; }
      setAll(g, 'Borrado', g.id || '');
      return;
    }
    // Validar
    let err = validateGroup_(g);
    const a = !err && dep.byApodo[norm_(g.apodo)];
    if (!err && !a) err = 'deportista desconocido (pon su apodo en «deportistas»)';
    if (!err && a.dup) err = 'el apodo «' + g.apodo + '» está repetido en «deportistas»';
    const ath = !err && athletes.byId[a.id];
    if (!err && (!ath || ath.revoked)) err = 'deportista desvinculado';
    if (err) { setAll(g, 'Error: ' + err); stats.errors++; return; }

    const cid = a.id;
    let id = g.id;
    let old = live(id);
    if (old && old.userId !== cid) {               // se ha cambiado el deportista: los planes no cambian de dueño
      tombstone(old);
      id = null;
      old = null;
    }
    if (!id || (meta.plans[id] && !old)) id = newPlanId_(meta.plans);   // id nuevo, o el antiguo estaba borrado
    const existing = existingOf(old);
    const built = buildPlan_(g, id, cid, existing, now);
    if (!built.ok) { setAll(g, 'Error: ' + built.error); stats.errors++; return; }
    if (!built.changed) {
      setAll(g, 'Sin cambios · ' + (STATUS_TXT[existing.status] || 'Pendiente'), id);
      stats.same++;
      return;
    }
    const data = JSON.stringify(built.data);
    if (data.length > PART) { setAll(g, 'Error: sesión demasiado larga'); stats.errors++; return; }
    meta.plans[id] = writeRecord_(writer, old, { store: 'plans', id: id, userId: cid, updatedAt: now, deleted: false },
      old ? old.author : 'coach', clock.next(), data);
    meta.plans[id].cached = { store: 'plans', id: id, userId: cid, author: meta.plans[id].author, updatedAt: now,
      serverAt: meta.plans[id].serverAt, deleted: false, data: built.data };
    setAll(g, '✓ Publicado · ' + (STATUS_TXT[built.data.status] || 'Pendiente'), id);
    stats.published++;
  });

  // 2) Ids que estaban en la hoja y ya no usa ninguna sesión (p. ej. se juntaron dos sesiones): se quitan de la app
  const used = {};
  groups.forEach(function (g) { g.rows.forEach(function (r) { if (ids[r.row]) used[ids[r.row]] = true; }); });
  rows.forEach(function (r) {
    if (r.id && !used[r.id]) { const m = live(r.id); if (m) { tombstone(m); stats.deleted++; } used[r.id] = true; }
  });

  clock.save();
  writer.flush();
  return { ok: true, n: Math.max(0, values.length - 1), estados: estados, ids: ids, stats: stats };
}

/* ================= Actualizar estados ================= */

function actualizarEstados() {
  const prog = sheet_(PROG_TAB);
  const n = prog.getLastRow() - 1;
  if (n < 1) return;
  const vals = prog.getRange(2, 1, n, PROG_COLS).getValues();
  const plans = readMeta_('plans');
  const metas = [];
  vals.forEach(function (v) { const m = plans[String(v[PC.id])]; if (m && metas.indexOf(m) < 0) metas.push(m); });
  const raws = readRaws_(metas);
  const status = {};
  metas.forEach(function (m, i) {
    if (m.deleted) { status[m.id] = 'Borrado'; return; }
    let d = null;
    try { d = JSON.parse(raws[i]); } catch (e) { d = null; }
    status[m.id] = d ? (STATUS_TXT[d.status] || 'Pendiente') : null;
  });
  const out = vals.map(function (v) {
    const id = String(v[PC.id]);
    const cur = String(v[PC.estado]);
    const st = status[id];
    if (!id || !st) return [cur];
    if (st === 'Borrado') return ['Borrado'];
    if (/^Error/.test(cur)) return [cur];
    return [(/^Sin cambios/.test(cur) ? 'Sin cambios · ' : '✓ Publicado · ') + st];
  });
  prog.getRange(2, PC.estado + 1, n, 1).setValues(out);
  const done = out.filter(function (r) { return /Hecho$/.test(r[0]); }).length;
  toast_('Estados actualizados' + (done ? ' · ' + done + ' fila' + (done === 1 ? '' : 's') + ' de sesiones hechas' : ''));
}

/* ================= Pestaña deportistas ================= */

// Añade los códigos vinculados que aún no están. No borra ni cambia nada.
function syncDeportistas_() {
  const dep = sheet_(DEP_TAB);
  const athletes = readAthletes_();
  const n = dep.getLastRow() - 1;
  const have = {};
  if (n > 0) dep.getRange(2, 1, n, 1).getValues().forEach(function (v) { have[String(v[0])] = true; });
  const add = Object.keys(athletes.byId)
    .filter(function (id) { return !athletes.byId[id].revoked && !have[id]; })
    .sort(function (a, b) { return athletes.byId[a].row - athletes.byId[b].row; })
    .map(function (id) { return [id, '']; });
  if (add.length) {
    const r = dep.getLastRow() + 1;
    dep.getRange(r, 1, add.length, 2).setNumberFormat('@').setValues(add);
  }
}

// {byApodo: {apodo normalizado: {id, dup}}}
function readDeportistas_() {
  const dep = sheet_(DEP_TAB);
  const out = { byApodo: {} };
  const n = dep.getLastRow() - 1;
  if (n < 1) return out;
  dep.getRange(2, 1, n, 2).getValues().forEach(function (v) {
    const id = String(v[0]).trim();
    const k = norm_(v[1]);
    if (!id || !k) return;
    if (out.byApodo[k] && out.byApodo[k].id !== id) out.byApodo[k].dup = true;
    else out.byApodo[k] = { id: id, dup: false };
  });
  return out;
}

/* ================= Escritura de Estado e id ================= */

function writeStatus_(prog, n, estados, ids) {
  if (n < 1) return;
  const cur = prog.getRange(2, PC.estado + 1, n, 2).getValues();
  const out = cur.map(function (v, i) {
    const row = i + 2;
    if (!(row in estados)) return ['', ''];   // fila vacía: sin estado ni id (si se escribe algo nuevo ahí, será otra sesión)
    return [estados[row], ids[row] || ''];
  });
  prog.getRange(2, PC.estado + 1, n, 2).setNumberFormat('@').setValues(out);
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'IMDTraining', 8); } catch (e) { Logger.log(msg); }
}

function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function newPlanId_(plans) {
  let id;
  do { id = 'h' + Utilities.getUuid().replace(/-/g, '').slice(0, 11); } while (plans[id]);
  return id;
}

/* ================= Núcleo: de filas a planes (sin SpreadsheetApp; se prueba con Node) ================= */

function norm_(v) {
  return String(v === null || v === undefined ? '' : v).trim().replace(/\s+/g, ' ').toLowerCase();
}

function str_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

// null si vacío; NaN si no es un número
function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/\s/g, '').replace(',', '.').replace(/%$/, '');
  if (s === '') return null;
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
}

function zone_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const m = String(v).trim().match(/^z?\s*([1-5])$/i);
  return m ? Number(m[1]) : NaN;
}

// Celda de fecha -> 'YYYY-MM-DD' | '' (vacía) | null (no válida)
function date_(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : fmt(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return ymd_(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/);
  if (m) return ymd_(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  return null;
}

function ymd_(y, mo, d) {
  const t = new Date(Date.UTC(y, mo - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return null;
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

// Deportista, Fecha y Sesión vacíos = los de la fila de arriba (así solo se escriben cuando cambian)
function parseRows_(values, fmt) {
  const out = [];
  const prev = { apodo: '', fecha: '', sesion: '' };
  for (let i = 1; i < values.length; i++) {
    const v = values[i];
    const has = v.slice(0, PC.accion + 1).some(function (x) { return x !== '' && x !== false && x !== null; });
    if (!has) continue;
    const tipo = norm_(v[PC.tipo]);
    const intTxt = norm_(v[PC.int]);
    const apodo = str_(v[PC.apodo]) || prev.apodo;
    const fecha = v[PC.fecha] === '' || v[PC.fecha] === null ? prev.fecha : date_(v[PC.fecha], fmt);
    const sesion = str_(v[PC.sesion]) || prev.sesion;
    prev.apodo = apodo; prev.fecha = fecha; prev.sesion = sesion;
    out.push({
      row: i + 1,
      apodo: apodo,
      fecha: fecha,
      sesion: sesion,
      kind: tipo === '' || tipo.indexOf('fue') === 0 ? 'strength'
        : tipo.indexOf('act') === 0 ? 'act'
        : tipo.indexOf('aer') === 0 || tipo.indexOf('aeró') === 0 ? 'hr' : 'bad',
      ejercicio: str_(v[PC.ejercicio]),
      variante: str_(v[PC.variante]),
      unilateral: bool_(v[PC.unilateral]),
      plyo: bool_(v[PC.plyo]),
      series: num_(v[PC.series]),
      reps: num_(v[PC.reps]),
      kg: num_(v[PC.kg]),
      pct: num_(v[PC.pct]),
      rpe: num_(v[PC.rpe]),
      rir: num_(v[PC.rir]),
      int: intTxt === '' ? null : (INT_BY_NAME[intTxt] || ([1, 2, 3].indexOf(num_(intTxt)) >= 0 ? num_(intTxt) : NaN)),
      minutos: num_(v[PC.minutos]),
      zona: zone_(v[PC.zona]),
      intervalos: str_(v[PC.intervalos]),
      notas: str_(v[PC.notas]),
      borrar: norm_(v[PC.accion]) === 'borrar',
      id: str_(v[PC.id]),
    });
  }
  return out;
}

// Sesiones = filas con el mismo deportista + fecha + sesión, en el orden en que aparecen
function groupSessions_(rows) {
  const by = {};
  const out = [];
  rows.forEach(function (r) {
    const key = norm_(r.apodo) + '|' + (r.fecha === null ? 'x' : r.fecha) + '|' + norm_(r.sesion);
    let g = by[key];
    if (!g) {
      g = by[key] = { key: key, apodo: r.apodo, fecha: r.fecha, sesion: r.sesion, rows: [], ids: [], borrar: false, id: '' };
      out.push(g);
    }
    g.rows.push(r);
    if (r.borrar) g.borrar = true;
    if (r.id && g.ids.indexOf(r.id) < 0) g.ids.push(r.id);
  });
  return out;
}

// Cada id solo puede ser de una sesión. Si varias lo reclaman (filas copiadas con su id oculto), se lo queda
// la que coincide con lo publicado (mismo deportista, fecha y título); si ninguna, la primera. Las demás, id nuevo.
// infoOf(id) -> {userId, date, title} | null ; cidOf(apodo) -> id_unico | null
function assignIds_(groups, infoOf, cidOf) {
  const claims = {};
  groups.forEach(function (g) { g.ids.forEach(function (id) { (claims[id] = claims[id] || []).push(g); }); });
  const owner = {};
  Object.keys(claims).forEach(function (id) {
    const list = claims[id];
    const info = infoOf(id);
    let win = list[0];
    if (info) {
      const match = list.filter(function (g) {
        return cidOf(g.apodo) === info.userId && g.fecha === info.date && norm_(g.sesion) === norm_(info.title);
      })[0];
      if (match) win = match;
    }
    owner[id] = win;
  });
  groups.forEach(function (g) {
    g.id = g.ids.filter(function (id) { return owner[id] === g; })[0] || '';
  });
}

function validateGroup_(g) {
  if (!g.apodo) return 'falta el deportista';
  if (g.fecha === null) return 'fecha no válida';
  if (!g.fecha) return 'falta la fecha';
  if (!g.sesion) return 'falta el nombre de la sesión';
  return '';
}

// Interpreta "10 Z2, 3 Z4, 0,5 Z5", "10'Z2 + 3'Z4" o "10 Z2, 4x(3 Z4, 2 Z1), 10 Z2". null si algo no encaja.
function parseIntervals_(text) {
  let s = String(text || '');
  for (let k = 0; k < 5 && /\(/.test(s); k++) {
    s = s.replace(/(\d+)\s*[x×*]\s*\(([^()]*)\)/gi, function (all, n, inner) {
      n = Number(n);
      return n >= 1 && n <= 50 ? Array(n + 1).join(inner + ', ') : '#';
    });
  }
  const re = /(\d+(?:[.,]\d+)?)\s*(?:['’′]|min\.?|minutos?)?\s*(?:en\s+)?z\s*([1-5])(?![\d])/gi;
  const out = [];
  let m;
  while ((m = re.exec(s))) {
    const min = Number(m[1].replace(',', '.'));
    if (!(min > 0) || min > 600) return null;
    out.push({ min: min, zone: Number(m[2]) });
  }
  if (!out.length || out.length > 100) return null;
  const rest = s.replace(re, '');
  return /^[\s,;+y]*$/i.test(rest) ? out : null;
}

const DEFAULT_HR = { type: 'Carrera', minutes: 40, zone: 2 };

// Lo que compara Publicar para saber si algo ha cambiado
function contentOf_(d) {
  return JSON.stringify({ title: d.title || '', kind: d.kind || 'strength', date: d.date || '', notes: d.notes || '',
    exercises: d.kind === 'hr' ? null : (d.exercises || []), hr: d.kind === 'hr' ? (d.hr || null) : null,
    cardio: d.kind === 'hr' ? null : (d.cardio || null) });
}

// -> {ok:true, data, changed} | {ok:false, error}
function buildPlan_(g, id, cid, existing, now) {
  const kinds = {};
  g.rows.forEach(function (r) { kinds[r.kind] = true; });
  if (kinds.bad) return { ok: false, error: 'el Tipo tiene que ser Fuerza, Activación o Aeróbico' };
  const aer = g.rows.filter(function (r) { return r.kind === 'hr'; });
  if (aer.length > 1) return { ok: false, error: 'solo una fila de Aeróbico por sesión' };
  const fuerza = g.rows.filter(function (r) { return r.kind !== 'hr'; });
  const notes = (g.rows.filter(function (r) { return r.notas; })[0] || {}).notas || '';
  const content = { title: g.sesion, kind: fuerza.length ? 'strength' : 'hr', date: g.fecha, notes: notes };
  let cardio = null;
  if (aer.length) {
    const r = aer[0];
    const type = r.ejercicio || 'Carrera';
    if (r.intervalos) {
      const blocks = parseIntervals_(r.intervalos);
      if (!blocks) return { ok: false, error: 'intervalos no válidos: escribe por ejemplo 10 Z2, 3 Z4, 10 Z2' };
      cardio = { type: type, minutes: Math.round(blocks.reduce(function (a, b) { return a + b.min; }, 0)), zone: DEFAULT_HR.zone,
        mode: 'int', blocks: blocks };
    } else {
      if (!(r.minutos > 0) || !(r.zona >= 1)) return { ok: false, error: 'pon Minutos y Zona, o Intervalos' };
      cardio = { type: type, minutes: r.minutos, zone: r.zona };
    }
  }

  if (!fuerza.length) {
    content.hr = cardio;
    content.exercises = [];
  } else {
    const exs = [];
    let last = null;
    for (let i = 0; i < fuerza.length; i++) {
      const r = fuerza[i];
      const n = 'fila ' + r.row + ': ';
      if (!r.ejercicio) return { ok: false, error: n + 'falta el ejercicio' };
      if (r.series !== null && (!(r.series >= 1) || r.series > 20 || r.series % 1)) return { ok: false, error: n + 'Series tiene que ser un número entre 1 y 20' };
      if (r.reps !== null && (!(r.reps >= 1) || r.reps % 1)) return { ok: false, error: n + 'Reps tiene que ser un número entero' };
      const bad = [['Kg', r.kg, 0, 1000], ['%1RM', r.pct, 0, 120], ['RPE', r.rpe, 0, 10], ['RIR', r.rir, 0, 10]]
        .filter(function (x) { return x[1] !== null && !(x[1] >= x[2] && x[1] <= x[3]); })[0];
      if (bad) return { ok: false, error: n + bad[0] + ' no válido' };
      if (r.int !== null && isNaN(r.int)) return { ok: false, error: n + 'Intensidad tiene que ser Baja, Media o Alta' };
      const pct = r.plyo || r.pct === null ? null : (r.pct > 0 && r.pct <= 1.2 ? Math.round(r.pct * 1000) / 10 : r.pct);
      const key = [norm_(r.ejercicio), norm_(r.variante), r.unilateral, r.plyo, r.kind].join('|');
      if (!last || last.key !== key) {
        last = { key: key, ex: { name: r.ejercicio, variant: r.variante, unilateral: r.unilateral, sets: [] } };
        if (r.plyo) last.ex.plyo = true;
        if (r.kind === 'act') last.ex.act = true;    // activación: se marca como hecha, no suma carga
        exs.push(last);
      }
      const ex = last.ex;
      if (!r.plyo) {
        const mode = pct !== null ? 'pct' : r.kg !== null ? 'kg' : null;
        if (mode) {
          if (last.mode && last.mode !== mode) return { ok: false, error: n + 'en ' + r.ejercicio + ' usa kg o %1RM, no los dos' };
          last.mode = mode;
          if (mode === 'pct') ex.load = 'pct';
        }
      }
      for (let k = 0; k < (r.series === null ? 1 : r.series); k++) {
        const s = { reps: r.reps };
        if (r.plyo) s.int = r.int;
        else { s.kg = pct !== null ? null : r.kg; if (pct !== null) s.pct = pct; }
        s.rpe = r.rpe;
        s.rir = r.rir;
        ex.sets.push(s);
      }
    }
    content.exercises = exs.map(function (x) { return x.ex; });
    content.hr = existing && existing.hr && existing.kind !== 'hr' ? existing.hr : { type: DEFAULT_HR.type, minutes: DEFAULT_HR.minutes, zone: DEFAULT_HR.zone };
    content.cardio = cardio;                        // aeróbico al final de la sesión de fuerza (concurrente)
  }

  const changed = !existing || contentOf_(existing) !== contentOf_(content);
  const data = {};
  if (existing) Object.keys(existing).forEach(function (k) { data[k] = existing[k]; });
  Object.keys(content).forEach(function (k) { data[k] = content[k]; });
  data.id = id;
  data.userId = cid;
  data.status = existing && existing.status ? existing.status : 'pending';
  data.createdAt = existing && existing.createdAt ? existing.createdAt : now;
  data.src = 'hoja';
  return { ok: true, data: data, changed: changed };
}

/* Para las pruebas con Node (en Apps Script no existe module) */
if (typeof module !== 'undefined') module.exports = { parseRows_: parseRows_, groupSessions_: groupSessions_, assignIds_: assignIds_,
  buildPlan_: buildPlan_, parseIntervals_: parseIntervals_, contentOf_: contentOf_ };
