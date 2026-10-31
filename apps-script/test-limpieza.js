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

Sheet.prototype.clear = function () { this.data = []; };
Sheet.prototype.autoResizeColumns = function () {};
Sheet.prototype.activate = function () {};
const _gr = Sheet.prototype.getRange;
Sheet.prototype.getRange = function (r, c, nr, nc) { if (typeof r === 'string') return { clearDataValidations() {}, setNote() {} , setFontWeight() { return this; }, setBackground() { return this; } }; return _gr.call(this, r, c, nr, nc); };
Range.prototype.clearContent = function () { for (let i = 0; i < this.nr; i++) { const R = this.sh.data[this.r - 1 + i]; if (R) for (let j = 0; j < this.nc; j++) R[this.c - 1 + j] = ''; } return this; };
SS.prototype.insertSheet = function (n) { return this.add(n, []); };

function load() {
  const ss = new SS();
  const props = { SHEET_ID: 'ss1', lastServerAt: '0' };
  const ctx = {
    console, Logger: { log: m => ctx.__log.push(m) }, __log: [],
    SpreadsheetApp: { openById: () => ss, getActiveSpreadsheet: () => ss, getActive: () => ss, flush() {},
      newDataValidation: () => { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
      getUi: () => ({ alert: m => ss.toasts.push('ALERT ' + m) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }), getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: { getUuid: () => crypto.randomUUID(), formatDate: d => d.toISOString().slice(0, 10) },
    ContentService: { createTextOutput: s => ({ setMimeType: () => s }), MimeType: {} },
  };
  vm.createContext(ctx);
  const src = ['Code.gs', 'Programacion.gs', 'Importar.gs', 'Diagnostico.gs'].map(f => fs.readFileSync(path.join(__dirname, f), 'utf8')).join('\n') +
    '\n;this.__api = {doPost, limpiezaPaso1Preparar, limpiezaPaso2Ejecutar, publicarProgramacion};';
  vm.runInContext(src, ctx, { filename: 'gs.js' });
  ss.add('config', [['key', 'value'], ['coachToken', 'COACH'], ['version', '1']]);
  ss.add('athletes', [['id_unico', 'token', 'createdAt', 'revoked'], ['dA', 'TA', 1, false], ['dB', 'TB', 2, false], ['dC', 'TC', 3, false]]);
  ['profiles', 'sessions', 'strength', 'tests', 'plans', 'wellness', 'exercises'].forEach(n => ss.add(n, [['id', 'userId', 'author', 'updatedAt', 'serverAt', 'deleted', 'data']]));
  ss.add('tracks', [['id', 'userId', 'part', 'parts', 'updatedAt', 'serverAt', 'deleted', 'data']]);
  ss.add('deportistas', [['codigo', 'apodo'], ['dA', 'Ana'], ['dB', 'Bea'], ['dC', 'Car']]);
  ss.add('fuentes', [['Apodo', 'Mes', 'Enlace', 'Importar', 'Estado', 'Última importación'],
    ['Ana', 'oct', 'L1', true, '', ''], ['Car', 'oct', 'L2', true, '', ''], ['A: Ana, C: Car', 'oct', 'L3', true, '', ''], ['Bea', 'oct', 'L4', true, '', '']]);
  ss.add('limpieza', []);
  ss.ui = null;
  return { ss, ctx, api: ctx.__api };
}
let fails = 0;
function ok(c, m) { if (!c) { fails++; console.log('  ✗ ' + m); } }
const call = (api, p) => JSON.parse(api.doPost({ postData: { contents: JSON.stringify(p) } }));

const { ss, ctx, api } = load();
// datos: cada deportista con fuerza, aeróbico (+track), test, plan, bienestar; Bea tiene un aeróbico hijo de una fuerza
const recs = [];
['dA', 'dB', 'dC'].forEach((u, k) => {
  recs.push({ store: 'profiles', id: u, userId: u, updatedAt: 10, data: { hrMax: 190 } });
  recs.push({ store: 'strength', id: 'st' + u, userId: u, updatedAt: 10, data: { exercises: [] } });
  recs.push({ store: 'sessions', id: 'se' + u, userId: u, updatedAt: 10, data: { status: 'done' } });
  recs.push({ store: 'tracks', id: 'se' + u, userId: u, updatedAt: 10, data: 'x'.repeat(100000) });
  recs.push({ store: 'tests', id: 'te' + u, userId: u, updatedAt: 10, data: {} });
  recs.push({ store: 'plans', id: 'pl' + u, userId: u, updatedAt: 10, data: {} });
  recs.push({ store: 'wellness', id: u + '_2026-10-01', userId: u, updatedAt: 10, data: {} });
});
recs.push({ store: 'sessions', id: 'hijoB', userId: 'dB', updatedAt: 10, data: { status: 'done', parentId: 'stdB', cardioPlan: 'x' } });
const pr = call(api, { token: 'COACH', action: 'push', records: recs });
ok(pr.ok && pr.accepted.length === recs.length, 'push inicial ' + JSON.stringify(pr.rejected));
const tracksRowsBefore = ss.sheets.tracks.getLastRow();

// programacion: Ana (2 filas), Car (2 filas heredando), Ana otra vez heredando fecha de Car
const H = ss.sheets.programacion ? null : null;
api.publicarProgramacion; // crea pestañas si hace falta
ss.add('programacion', [ctx.PROG_HEADERS ? ctx.PROG_HEADERS.slice() : []]);
const P = vm.runInContext('PROG_HEADERS', ctx);
const PCx = vm.runInContext('PC', ctx);
const mk = o => { const r = P.map(() => ''); Object.keys(o).forEach(k => { r[PCx[k]] = o[k]; }); return r; };
const d1 = new Date('2026-10-05T00:00:00Z'), d2 = new Date('2026-10-07T00:00:00Z');
ss.sheets.programacion.data = [P.slice(),
  mk({ apodo: 'Ana', fecha: d1, sesion: 'S1', ejercicio: 'Sentadilla' }),
  mk({ ejercicio: 'Press' }),
  mk({ apodo: 'Car', fecha: d2, ejercicio: 'Remo' }),
  mk({ ejercicio: 'Dominada' }),
  mk({ apodo: 'Ana', sesion: 'S2', ejercicio: 'Peso muerto' }),
  mk({ ejercicio: 'Hip thrust' })];

api.limpiezaPaso1Preparar();
const L = ss.sheets.limpieza.data;
ok(L.length === 4 && L[1][0] === 'dA' && L[1][1] === 'Ana' && L[1][3] === 1 && L[2][4] === 2, 'pestaña limpieza ' + JSON.stringify(L));
// sin elegir: no borra
let err = null; try { api.limpiezaPaso2Ejecutar(); } catch (e) { err = e.message; }
ok(err && /Falta elegir/.test(err), 'exige elegir: ' + err);
L[1][8] = 'Mantener todo'; L[2][8] = 'Solo aeróbico'; L[3][8] = 'Vaciar';
api.limpiezaPaso2Ejecutar();
console.log(ctx.__log[ctx.__log.length - 1]);

const pull = call(api, { token: 'COACH', action: 'pull', since: 0 }).records;
const st = (s, id) => pull.find(r => r.store === s && r.id === id);
ok(['strength', 'sessions', 'tests', 'plans'].every(s => !st(s, (s === 'strength' ? 'st' : s === 'sessions' ? 'se' : s === 'tests' ? 'te' : 'pl') + 'dA').deleted), 'Ana intacta');
ok(!st('tracks', 'sedA').deleted && st('tracks', 'sedA').data.length === 100000, 'track de Ana intacto');
ok(st('strength', 'stdB').deleted && st('tests', 'tedB').deleted && st('plans', 'pldB').deleted && st('wellness', 'dB_2026-10-01').deleted, 'Bea sin fuerza/test/plan/bienestar');
ok(!st('sessions', 'sedB').deleted && !st('tracks', 'sedB').deleted, 'Bea conserva aeróbico y track');
ok(!st('sessions', 'hijoB').deleted && !st('sessions', 'hijoB').data.parentId && st('sessions', 'hijoB').data.cardioPlan === 'x', 'hijo de Bea suelto');
ok(['strength', 'sessions', 'tests', 'plans', 'tracks'].every(s => st(s, (s === 'strength' ? 'st' : s === 'sessions' || s === 'tracks' ? 'se' : s === 'tests' ? 'te' : 'pl') + 'dC').deleted), 'Car vaciado');
ok(!st('profiles', 'dC').deleted && !st('profiles', 'dB').deleted, 'perfiles siguen');
ok(ss.sheets.tracks.getLastRow() < tracksRowsBefore, 'filas sobrantes de tracks quitadas');
// tombstone del track de Car en una sola fila
ok(ss.sheets.tracks.data.filter(r => r[0] === 'sedC').length === 1, 'track Car una fila');

const pg = ss.sheets.programacion.data.slice(1).filter(r => r.some(x => x !== '' && x !== false));
ok(pg.length === 4, 'programacion 4 filas: ' + pg.length);
ok(pg[2][PCx.apodo] === 'Ana' && +pg[2][PCx.fecha] === +d2 && pg[2][PCx.sesion] === 'S2' && pg[2][PCx.ejercicio] === 'Peso muerto', 'fila heredada completada ' + JSON.stringify(pg[2]));
ok(pg[1][PCx.apodo] === '' && pg[3][PCx.apodo] === '', 'las demás siguen heredando');

const F = ss.sheets.fuentes.data;
ok(F[1][3] === true && F[2][3] === false && F[3][3] === false && F[4][3] === false, 'fuentes: ' + JSON.stringify(F.map(r => r[3])));
console.log(fails ? fails + ' FALLOS' : 'Todo OK');
process.exit(fails ? 1 : 0);
