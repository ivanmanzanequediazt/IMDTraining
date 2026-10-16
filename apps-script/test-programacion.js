// Pruebas de Programacion.gs sin Google: simula la hoja en memoria y ejecuta Código.gs + Programacion.gs.
// Uso:  node apps-script/test-programacion.js
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

/* ---------- Hoja simulada ---------- */
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.data[this.r - 1 + i] || [];
      const o = [];
      for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v === undefined ? '' : v); }
      out.push(o);
    }
    return out;
  }
  setValues(v) {
    v.forEach((row, i) => {
      const R = this.sh.data[this.r - 1 + i] || (this.sh.data[this.r - 1 + i] = []);
      row.forEach((x, j) => { R[this.c - 1 + j] = x; });
    });
    for (let i = 0; i < this.sh.data.length; i++) if (!this.sh.data[i]) this.sh.data[i] = [];
    return this;
  }
  setValue(x) { return this.setValues([[x]]); }
  getColumn() { return this.c; }
}
['setNumberFormat', 'setNumberFormats', 'setFontWeight', 'setBackground', 'setWrap', 'setVerticalAlignment', 'setNote',
  'setFontColor', 'setFontStyle', 'setDataValidation', 'insertCheckboxes']
  .forEach(m => { Range.prototype[m] = function () { return this; }; });

class Sheet {
  constructor(ss, name, data) { this.ss = ss; this.name = name; this.data = data || []; }
  getLastRow() { for (let i = this.data.length - 1; i >= 0; i--) if ((this.data[i] || []).some(x => x !== '' && x !== undefined && x !== false)) return i + 1; return 0; }
  getMaxRows() { return Math.max(this.data.length, 1000); }
  getMaxColumns() { return 26; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
  getParent() { return this.ss; }
  insertRowsAfter() {}
  deleteRow(r) { this.data.splice(r - 1, 1); }
  appendRow(v) { this.data[this.getLastRow()] = v.slice(); }
  setFrozenRows() {}
}
class SS {
  constructor() { this.sheets = {}; this.toasts = []; }
  add(name, data) { return (this.sheets[name] = new Sheet(this, name, data)); }
  getSheetByName(n) { return this.sheets[n] || null; }
  getSpreadsheetTimeZone() { return 'UTC'; }
  getId() { return 'ss1'; }
  toast(m) { this.toasts.push(m); }
}

function load() {
  const ss = new SS();
  const props = { SHEET_ID: 'ss1', lastServerAt: '0' };
  let nowOffset = 0;
  const ctx = {
    console, Logger: { log: () => {} },
    SpreadsheetApp: { openById: () => ss, getActiveSpreadsheet: () => ss, getActive: () => ss, flush() {},
      getUi: () => ({ alert: m => ss.toasts.push('ALERT ' + m) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => !ctx.__busy, releaseLock() {} }) },
    Utilities: { getUuid: () => crypto.randomUUID(), formatDate: d => d.toISOString().slice(0, 10) },
    ContentService: { createTextOutput: s => ({ setMimeType: () => s }), MimeType: {} },
    __busy: false,
  };
  vm.createContext(ctx);
  const dir = __dirname;
  const src = fs.readFileSync(path.join(dir, 'Code.gs'), 'utf8') + '\n' + fs.readFileSync(path.join(dir, 'Programacion.gs'), 'utf8') +
    '\n;this.__api = {publicarProgramacion, actualizarEstados, pull_, push_, doPost, readMeta_};';
  vm.runInContext(src, ctx, { filename: 'gs.js' });
  // Pestañas de Código.gs
  ss.add('config', [['key', 'value'], ['coachToken', 'COACH'], ['version', '1']]);
  ss.add('athletes', [['id_unico', 'token', 'createdAt', 'revoked'], ['d1', 'T1', 1, false], ['d2', 'T2', 2, false], ['d3', 'T3', 3, true]]);
  ['profiles', 'sessions', 'strength', 'tests', 'plans', 'wellness'].forEach(n => ss.add(n, [['id', 'userId', 'author', 'updatedAt', 'serverAt', 'deleted', 'data']]));
  ss.add('tracks', [['id', 'userId', 'part', 'parts', 'updatedAt', 'serverAt', 'deleted', 'data']]);
  ss.add('deportistas', [['codigo', 'apodo']]);
  ss.add('programacion', [['h']]);
  return { ss, ctx, api: ctx.__api };
}

/* ---------- Utilidades de prueba ---------- */
let fails = 0, passes = 0;
function ok(cond, msg) { if (cond) passes++; else { fails++; console.log('  ✗ ' + msg); } }
function eq(a, b, msg) { ok(JSON.stringify(a) === JSON.stringify(b), msg + '\n     esperado ' + JSON.stringify(b) + '\n     obtenido ' + JSON.stringify(a)); }
const COLS = ['apodo', 'fecha', 'sesion', 'tipo', 'ejercicio', 'variante', 'unilateral', 'plyo', 'series', 'reps', 'kg', 'pct',
  'rpe', 'rir', 'int', 'minutos', 'zona', 'intervalos', 'notas', 'accion', 'estado', 'id'];
function row(o) { return COLS.map(k => (k in o ? o[k] : (k === 'unilateral' || k === 'plyo' ? false : ''))); }
function setProg(ss, rows) { ss.sheets.programacion.data = [COLS.slice()].concat(rows.map(row)); }
function col(ss, k) { return ss.sheets.programacion.data.slice(1).map(r => r[COLS.indexOf(k)]); }
function plans(api) {
  const r = JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ token: 'COACH', action: 'pull', since: 0 }) } }));
  return r.records.filter(x => x.store === 'plans');
}
function live(api) { return plans(api).filter(p => !p.deleted); }
const D = s => new Date(s + 'T00:00:00Z');
function test(name, fn) { console.log('· ' + name); fn(); }

/* ---------- Núcleo ---------- */
const core = require('./Programacion.gs');
test('Intervalos', () => {
  eq(core.parseIntervals_('10 Z2, 3 Z4, 0,5 Z5'), [{ min: 10, zone: 2 }, { min: 3, zone: 4 }, { min: 0.5, zone: 5 }], 'coma decimal y separador');
  eq(core.parseIntervals_("10'Z2+3'Z4"), [{ min: 10, zone: 2 }, { min: 3, zone: 4 }], 'apóstrofo y +');
  eq(core.parseIntervals_('10Z2,3Z4'), [{ min: 10, zone: 2 }, { min: 3, zone: 4 }], 'sin espacios');
  eq(core.parseIntervals_('10 min Z1, 2x(3 Z4, 1,5 Z1), 5 Z1').length, 6, 'repeticiones 2x(...)');
  eq(core.parseIntervals_('10 Z7'), null, 'zona inexistente');
  eq(core.parseIntervals_('10 Z2, correr suave'), null, 'texto que no encaja');
  eq(core.parseIntervals_(''), null, 'vacío');
});

/* ---------- Publicar de punta a punta ---------- */
test('Primera publicación: deportistas, fuerza, pirámide, %1RM, pliometría, intervalos', () => {
  const { ss, api } = load();
  setProg(ss, []);
  api.publicarProgramacion();                              // crea la lista de deportistas
  eq(ss.sheets.deportistas.data.slice(1).map(r => r[0]), ['d1', 'd2'], 'deportistas: solo los no desvinculados');
  ss.sheets.deportistas.data[1][1] = 'DG';
  ss.sheets.deportistas.data[2][1] = 'Ana';
  setProg(ss, [
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna A', tipo: 'Fuerza', ejercicio: 'Sentadilla', series: 3, reps: 8, kg: 70, rpe: 7, notas: 'Calienta 10 min' },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna A', tipo: 'Fuerza', ejercicio: 'Sentadilla', series: 2, reps: 5, kg: 80, rpe: 8 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna A', ejercicio: 'Peso muerto rumano', series: 3, reps: 10, pct: '70%', rir: 2 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna A', ejercicio: 'Zancada', unilateral: true, series: 3, reps: 8, kg: '12,5' },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna A', ejercicio: 'Drop jump', plyo: true, series: 3, reps: 6, int: 'Alta' },
    {},
    { apodo: 'dg ', fecha: '22/10/2026', sesion: 'Series', tipo: 'Aeróbico', ejercicio: 'Carrera', intervalos: '10 Z2, 3 Z4, 10 Z2' },
    { apodo: 'Ana', fecha: D('2026-10-21'), sesion: 'Rodaje', tipo: 'Aerobico', minutos: 40, zona: 'Z2' },
  ]);
  api.publicarProgramacion();
  const st = col(ss, 'estado');
  eq(st, ['✓ Publicado · Pendiente', '✓ Publicado · Pendiente', '✓ Publicado · Pendiente', '✓ Publicado · Pendiente', '✓ Publicado · Pendiente', '',
    '✓ Publicado · Pendiente', '✓ Publicado · Pendiente'], 'estados');
  const ids = col(ss, 'id');
  ok(ids[0] && ids.slice(0, 5).every(x => x === ids[0]), 'mismo id en las filas de la sesión');
  ok(ids[6] && ids[6] !== ids[0] && ids[7] && ids[7] !== ids[6], 'ids distintos por sesión');
  eq(ids[5], '', 'fila vacía sin id');
  const P = Object.fromEntries(live(api).map(p => [p.data.title, p]));
  const f = P['Pierna A'];
  eq([f.userId, f.author, f.data.kind, f.data.date, f.data.notes, f.data.status, f.data.src], ['d1', 'coach', 'strength', '2026-10-20', 'Calienta 10 min', 'pending', 'hoja'], 'cabecera del plan de fuerza');
  eq(f.data.exercises.map(e => e.name), ['Sentadilla', 'Peso muerto rumano', 'Zancada', 'Drop jump'], 'ejercicios en orden; pirámide juntada');
  eq(f.data.exercises[0].sets.map(s => s.reps + 'x' + s.kg + '@' + s.rpe), ['8x70@7', '8x70@7', '8x70@7', '5x80@8', '5x80@8'], 'pirámide');
  eq([f.data.exercises[1].load, f.data.exercises[1].sets[0]], ['pct', { reps: 10, kg: null, pct: 70, rpe: null, rir: 2 }], '%1RM');
  eq([f.data.exercises[2].unilateral, f.data.exercises[2].sets[0].kg], [true, 12.5], 'unilateral y coma decimal');
  eq([f.data.exercises[3].plyo, f.data.exercises[3].sets[0]], [true, { reps: 6, int: 3, rpe: null, rir: null }], 'pliometría');
  eq(P.Series.data.hr, { type: 'Carrera', minutes: 23, zone: 2, mode: 'int', blocks: [{ min: 10, zone: 2 }, { min: 3, zone: 4 }, { min: 10, zone: 2 }] }, 'intervalos (apodo con mayúsculas/espacio y fecha en texto)');
  eq([P.Series.data.date, P.Series.userId], ['2026-10-22', 'd1'], 'fecha en texto');
  eq([P.Rodaje.userId, P.Rodaje.data.kind, P.Rodaje.data.hr], ['d2', 'hr', { type: 'Carrera', minutes: 40, zone: 2 }], 'aeróbico continuo; tipo vacío = Carrera');

  // El deportista solo ve lo suyo
  const ana = JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ token: 'T2', action: 'pull', since: 0 }) } }));
  eq(ana.records.map(r => r.data.title), ['Rodaje'], 'Ana solo recibe su plan');

  // H3: sin cambios
  const before = plans(api).map(p => p.serverAt);
  api.publicarProgramacion();
  eq(col(ss, 'estado')[0], 'Sin cambios · Pendiente', 'sin cambios');
  eq(plans(api).map(p => p.serverAt), before, 'nada reescrito');

  // H5/H6: el deportista lo hace (push desde su móvil); Iván cambia una indicación
  const pf = live(api).find(p => p.data.title === 'Pierna A');
  const doneData = Object.assign({}, pf.data, { status: 'done', doneId: 'S1', doneAt: 123 });
  const pr = JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ token: 'T1', action: 'push', records: [{ store: 'plans', id: pf.id, userId: 'd1', updatedAt: Date.now() + 1, deleted: false, data: doneData }] }) } }));
  eq(pr.accepted.length, 1, 'el deportista marca hecho');
  api.actualizarEstados();
  eq(col(ss, 'estado')[0], 'Sin cambios · Hecho', 'actualizar estados');
  ss.sheets.programacion.data[1][COLS.indexOf('notas')] = 'Calienta 15 min';
  api.publicarProgramacion();
  const pf2 = live(api).find(p => p.data.title === 'Pierna A');
  eq([pf2.data.notes, pf2.data.status, pf2.data.doneId, col(ss, 'estado')[0]], ['Calienta 15 min', 'done', 'S1', '✓ Publicado · Hecho'], 'cambio conserva el estado Hecho');

  // H9: cambiar de deportista una sesión publicada
  const rid = col(ss, 'id')[7];
  ss.sheets.programacion.data[8][0] = 'DG';
  api.publicarProgramacion();
  const all = plans(api);
  ok(all.find(p => p.id === rid).deleted, 'el plan del deportista anterior se borra');
  const moved = live(api).filter(p => p.data.title === 'Rodaje');
  eq(moved.map(p => p.userId), ['d1'], 'aparece en el nuevo deportista');
  ok(col(ss, 'id')[7] !== rid, 'id nuevo');

  // H7: Borrar
  ss.sheets.programacion.data[7][COLS.indexOf('accion')] = 'Borrar';
  const sid = col(ss, 'id')[6];
  api.publicarProgramacion();
  ok(plans(api).find(p => p.id === sid).deleted, 'Borrar quita el plan');
  eq(col(ss, 'estado')[6], 'Borrado', 'estado Borrado');
  api.publicarProgramacion();
  eq(col(ss, 'estado')[6], 'Borrado', 'volver a publicar no hace nada más');
});

test('Errores (H8) y validaciones', () => {
  const { ss, api } = load();
  ss.sheets.deportistas.data.push(['d1', 'DG'], ['d2', 'X'], ['d9', 'X']);
  setProg(ss, [
    { apodo: 'Nadie', fecha: D('2026-10-20'), sesion: 'A', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'DG', fecha: '', sesion: 'B', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'C', tipo: 'Aeróbico', intervalos: '10 Z9' },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'D', ejercicio: 'Press', series: 3, reps: 5, kg: 50 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'D', ejercicio: 'Press', series: 1, reps: 3, pct: 90 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'E', tipo: 'Aeróbico', minutos: 30, zona: 2 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'E', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'X', fecha: D('2026-10-20'), sesion: 'F', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'DG', fecha: '31/02/2026', sesion: 'G', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'H', ejercicio: 'Press', series: 3, reps: 5, rpe: 12 },
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'OK', ejercicio: 'Press', series: 3, reps: 5 },
  ]);
  api.publicarProgramacion();
  const st = col(ss, 'estado');
  ok(/deportista desconocido/.test(st[0]), 'apodo desconocido: ' + st[0]);
  ok(/falta la fecha/.test(st[1]), 'sin fecha: ' + st[1]);
  ok(/intervalos no válidos/.test(st[2]), 'intervalos: ' + st[2]);
  ok(/kg o %1RM/.test(st[3]) && st[3] === st[4], 'mezcla kg y %: ' + st[3]);
  ok(/fuerza o aeróbica/.test(st[5]) && st[5] === st[6], 'mezcla de tipos: ' + st[5]);
  ok(/repetido/.test(st[7]), 'apodo repetido: ' + st[7]);
  ok(/fecha no válida/.test(st[8]), 'fecha imposible: ' + st[8]);
  ok(/RPE no válido/.test(st[9]), 'RPE fuera de rango: ' + st[9]);
  eq(st[10], '✓ Publicado · Pendiente', 'lo correcto se publica igualmente');
  eq(live(api).length, 1, 'solo un plan publicado');
  ok(ss.toasts.some(t => /8 con error/.test(t)), 'aviso final: ' + ss.toasts.slice(-1));
});

test('Copiar filas con su id oculto; juntar sesiones; deportista desvinculado; hoja ocupada', () => {
  const { ss, ctx, api } = load();
  ss.sheets.deportistas.data.push(['d1', 'DG'], ['d3', 'Viejo']);
  setProg(ss, [
    { apodo: 'DG', fecha: D('2026-10-20'), sesion: 'Pierna', ejercicio: 'Sentadilla', series: 3, reps: 5 },
    { apodo: 'DG', fecha: D('2026-10-21'), sesion: 'Torso', ejercicio: 'Press', series: 3, reps: 5 },
    { apodo: 'Viejo', fecha: D('2026-10-21'), sesion: 'Torso', ejercicio: 'Press', series: 3, reps: 5 },
  ]);
  api.publicarProgramacion();
  ok(/desvinculado/.test(col(ss, 'estado')[2]), 'desvinculado: ' + col(ss, 'estado')[2]);
  const [idP, idT] = col(ss, 'id');
  // Iván copia la fila de Pierna (con su id oculto) ENCIMA, para la semana siguiente
  const copy = ss.sheets.programacion.data[1].slice();
  copy[1] = D('2026-10-27');
  ss.sheets.programacion.data.splice(1, 0, copy);
  api.publicarProgramacion();
  const ids = col(ss, 'id');
  eq(ids[1], idP, 'el original conserva su id aunque la copia esté encima');
  ok(ids[0] && ids[0] !== idP, 'la copia recibe id nuevo');
  eq(live(api).map(p => p.data.date).sort(), ['2026-10-20', '2026-10-21', '2026-10-27'], 'tres planes, ninguno pisado');
  // Juntar Torso dentro de Pierna (cambiando su sesión y fecha): el plan de Torso sobra y se quita
  ss.sheets.programacion.data[3][1] = D('2026-10-20');
  ss.sheets.programacion.data[3][2] = 'Pierna';
  api.publicarProgramacion();
  ok(plans(api).find(p => p.id === idT).deleted, 'el plan que ya no tiene filas propias se quita');
  eq(live(api).find(p => p.id === idP).data.exercises.map(e => e.name), ['Sentadilla', 'Press'], 'sesión juntada');
  // Hoja ocupada
  ctx.__busy = true;
  api.publicarProgramacion();
  ok(/ALERT La hoja está ocupada/.test(ss.toasts.slice(-1)[0]), 'aviso de hoja ocupada');
});

console.log(`\n${passes} comprobaciones correctas, ${fails} fallos`);
process.exit(fails ? 1 : 0);
