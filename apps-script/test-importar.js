// Pruebas de Importar.gs sin Google.  Uso: node apps-script/test-importar.js [ruta a un .json con el XML de un xlsx]
// El .json de ejemplo (un Excel real de un cliente) no está en el repositorio: se pasa como argumento.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const dir = __dirname;
const ctx = { console, Utilities: { formatDate: d => d.toISOString().slice(0, 10) } };
vm.createContext(ctx);
const src = ['Code.gs', 'Programacion.gs', 'Importar.gs'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') +
  '\n;this.__ = {parseQuienes_, valorDe_, readXlsx_, parseWorkbook_, mesDeNombre_, filasDeSesion_, planImportacion_, primerNumero_, buildPlan_, parseRows_, groupSessions_, PC, PROG_COLS};';
vm.runInContext(src, ctx);
const X = ctx.__;

let fails = 0, passes = 0;
const eq = (a, b, msg) => { const ok = JSON.stringify(a) === JSON.stringify(b); ok ? passes++ : (fails++, console.log('  ✗ ' + msg + '\n     esperado ' + JSON.stringify(b) + '\n     obtenido ' + JSON.stringify(a))); };
const ok = (c, msg) => eq(!!c, true, msg);

console.log('· Utilidades');
eq(X.mesDeNombre_('PEPE_OCTUBRE_26.xlsx'), { m: 10, y: 2026 }, 'mes y año del nombre');
eq(X.mesDeNombre_('Ana_Marzo_2027.xlsx'), { m: 3, y: 2027 }, 'año con 4 cifras');
eq(X.mesDeNombre_('PEPE_OCTUBRE.xlsx').m, 10, 'sin año');
eq([X.primerNumero_("17'5 lado", true), X.primerNumero_('50 kg', true), X.primerNumero_('poner lastre', true), X.primerNumero_('4-6rep', false), X.primerNumero_("3 x 4''", false)],
  [17.5, 50, null, 4, 3], 'números en texto');

const fx = process.argv[2];
if (fx && fs.existsSync(fx)) {
  console.log('· Excel de ejemplo: ' + path.basename(fx));
  const wb = X.readXlsx_(JSON.parse(fs.readFileSync(fx, 'utf8')));
  const res = X.parseWorkbook_(wb, 2026, 10);
  res.sesiones.forEach(s => console.log('   ' + s.fecha + ' ' + s.sesion.padEnd(16) + ' (' + s.hoja + ') ' + s.ejercicios.length + ' ej.: ' +
    s.ejercicios.map(e => e.name).join(' | ').slice(0, 110)));
  console.log('   avisos: ' + res.avisos.join(' / '));
  // Filas y planes
  const filas = [];
  res.sesiones.forEach(s => X.filasDeSesion_('Pepe', s, 'excel:x').forEach(r => filas.push(r)));
  const parsed = X.parseRows_([X.PC && Array(X.PROG_COLS).fill('h')].concat(filas), d => d);
  const groups = X.groupSessions_(parsed);
  let errs = 0;
  groups.forEach(g => { const b = X.buildPlan_(g, 'id', 'd1', null, 1); if (!b.ok) { errs++; console.log('   ✗ ' + g.key + ': ' + b.error); } });
  eq(errs, 0, 'todas las sesiones importadas se convierten en planes válidos');
  eq(groups.length, res.sesiones.length, 'una sesión por bloque');
  const g0 = groups[0], p0 = X.buildPlan_(g0, 'id', 'd1', null, 1).data;
  console.log('   ejemplo → ' + p0.date + ' «' + p0.title + '»: ' + p0.exercises.map(e => e.name + (e.variant ? ' [' + e.variant + ']' : '') + ' ' + e.sets.length + 'x' + e.sets[0].reps + (e.sets[0].kg ? '@' + e.sets[0].kg : '')).join('; ') + (p0.notes ? ' · notas: ' + p0.notes : ''));
} else console.log('· (sin Excel de ejemplo: se salta)');

console.log('· Archivo compartido (O: Olga, S: Sergio)');
{
  eq(X.parseQuienes_('O: Olga Ejemplo, S: Sergio Ejemplo'), [{ letra: 'O', apodo: 'Olga Ejemplo' }, { letra: 'S', apodo: 'Sergio Ejemplo' }], 'apodos con letra');
  eq(X.parseQuienes_('Pepe Ejemplo'), [{ letra: null, apodo: 'Pepe Ejemplo' }], 'un solo apodo');
  const L = ['O', 'S'];
  eq(['O: 25 S: 12', 'O:25kg / S:12kg', 's: 12 o: 25', "O: 17'5 lado S: 10", 'O 25', '20', 'S: 12'].map(x => X.valorDe_(x, 'O', L).txt),
     ['25', '25kg', '25', "17'5 lado", 'O 25', '20', ''], 'peso de Olga');
  eq(X.valorDe_('O: 25 S: 12', 'S', L).txt, '12', 'peso de Sergio');
  const s = { fecha: '2026-10-06', sesion: 'Pierna', ejercicios: [
    { name: 'Sentadilla', series: 3, repsTxt: '8', pesoTxt: 'O: 25 S: 12', rpe: 8, obs: '' },
    { name: 'Remo', series: 3, repsTxt: '10', pesoTxt: '20', rpe: 'O: 8 S: 7', obs: '' },
    { name: 'Press', series: 3, repsTxt: '10', pesoTxt: "O: 17'5 lado S: 10", rpe: '', obs: '' }] };
  const o = X.filasDeSesion_('Olga Ejemplo', s, 'excel:A#O', 'O', L), sa = X.filasDeSesion_('Sergio Ejemplo', s, 'excel:A#S', 'S', L);
  eq(o.map(r => [r[X.PC.kg], r[X.PC.rpe]]), [[25, 8], [20, 8], [17.5, '']], 'Olga: kg y RPE');
  eq(sa.map(r => [r[X.PC.kg], r[X.PC.rpe]]), [[12, 8], [20, 7], [10, '']], 'Sergio: kg y RPE');
  eq([o[0][X.PC.notas], sa[0][X.PC.notas]], ["Press: 17'5 lado", ''], 'notas solo con lo suyo');
}
console.log('· Reimportar conserva ids y marca para borrar lo que ya no está');
{
  const s1 = { fecha: '2026-10-06', sesion: 'Tracc', ejercicios: [{ name: 'Remo', series: 3, repsTxt: '8', pesoTxt: '', rpe: '', obs: '' }] };
  const s2 = { fecha: '2026-10-08', sesion: 'Empuje', ejercicios: [{ name: 'Press', series: 3, repsTxt: '6-8rep', pesoTxt: "17'5 lado", rpe: 8, obs: '' }] };
  const vals = [];
  [s1, s2].forEach(s => X.filasDeSesion_('Pepe', s, 'excel:A').forEach(r => vals.push(r)));
  vals.forEach((r, i) => { r[X.PC.id] = i === 0 ? 'hA' : 'hB'; r[X.PC.estado] = '✓ Publicado · Pendiente'; });
  const otra = Array(X.PROG_COLS).fill(''); otra[X.PC.apodo] = 'Ana'; otra[X.PC.origen] = ''; vals.push(otra);
  const s1b = { fecha: '2026-10-06', sesion: 'Tracc', ejercicios: [{ name: 'Remo', series: 4, repsTxt: '8', pesoTxt: '', rpe: '', obs: '' }] };
  const plan = X.planImportacion_(vals, 'excel:A', 'Pepe', [s1b], d => d);
  eq(plan.borrarFilas, [0, 1], 'solo se rehacen las filas de este Excel');
  eq(plan.nuevas.map(r => [r[X.PC.sesion], r[X.PC.series], r[X.PC.id], r[X.PC.accion]]), [['Tracc', 4, 'hA', ''], ['Empuje', '', 'hB', 'Borrar']], 'id conservado y sesión quitada marcada para borrar');
  eq(plan.borrar, 1, 'una para borrar');
  const v2 = X.filasDeSesion_('Pepe', s2, 'excel:A')[0];
  eq([v2[X.PC.reps], v2[X.PC.variante], v2[X.PC.kg], v2[X.PC.rpe], v2[X.PC.notas]], [6, '6-8rep', 17.5, 8, "Press: 17'5 lado"], 'texto de reps y peso');
}

console.log(`\n${passes} comprobaciones correctas, ${fails} fallos`);
process.exit(fails ? 1 : 0);
