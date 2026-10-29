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
  '\n;this.__ = {fechaExcel_, partesBiserie_, parseQuienes_, valorDe_, readXlsx_, parseWorkbook_, mesDeNombre_, filasDeSesion_, planImportacion_, primerNumero_, buildPlan_, parseRows_, groupSessions_, PC, PROG_COLS};';
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
console.log('· Biseries');
{
  eq(X.partesBiserie_('Remo barra + PM unipodal kb contralateral '), ['Remo barra', 'PM unipodal kb contralateral'], 'biserie');
  eq(X.partesBiserie_('Ext tronco pica + (Plancha comando + Rot cubanas disco)'), ['Ext tronco pica', 'Plancha comando', 'Rot cubanas disco'], 'triserie con paréntesis');
  const s = { fecha: '2026-10-06', sesion: 'T', ejercicios: [
    { name: 'Remo barra + PM unipodal', series: 3, repsTxt: '4-6rep / 6-8rep', pesoTxt: '40 / 12', rpe: 8, obs: '' },
    { name: 'Press inclinado + remo mentón', series: 3, repsTxt: '8, 12', pesoTxt: '20', rpe: '', obs: '' },
    { name: 'Sentadilla + plancha', series: 2, repsTxt: '6', pesoTxt: 'O: 50 S: 30', rpe: '', obs: '' }] };
  const f = X.filasDeSesion_('Pepe', s, 'excel:A', 'S', ['O', 'S']);
  eq(f.map(r => [r[X.PC.ejercicio], r[X.PC.series], r[X.PC.reps], r[X.PC.variante], r[X.PC.kg], r[X.PC.rpe]]), [
    ['Remo barra', 3, 4, 'biserie · 4-6rep', 40, 8], ['PM unipodal', 3, 6, 'biserie · 6-8rep', 12, 8],
    ['Press inclinado', 3, 8, 'biserie', 20, ''], ['remo mentón', 3, 12, 'biserie', '', ''],
    ['Sentadilla', 2, 6, 'biserie', 30, ''], ['plancha', 2, 6, 'biserie', '', '']], 'cada ejercicio con sus reps y kilos');
}
console.log('· Plantilla nueva (lista por semana)');
{
  eq([X.fechaExcel_(46328), X.fechaExcel_('02/11/2026')], ['2026-11-02', '2026-11-02'], 'fechas de Excel');
  const c = { A1: 'SEMANA DEL', B1: 46328, A4: 'Día', B4: 'Sesión', C4: 'Bloque', D4: 'Ejercicio', E4: 'Series', F4: 'Reps', G4: 'Nota reps', H4: 'Peso O', I4: 'Peso S', J4: 'RPE', K4: 'Unilateral', L4: 'Pliometría', M4: 'Observaciones',
    A5: 'Lunes', B5: 'Pierna', C5: 'Activación', D5: 'Puente', E5: 2, F5: 10,
    C6: 'P. Principal', D6: 'Sentadilla', E6: 3, F6: 8, G6: '8-10', H6: 24, I6: 12, J6: 7,
    D7: 'Zancada', E7: 3, F7: 8, H7: 10, I7: 6, K7: 'Sí',
    A8: 'Jueves', D8: 'Press + Remo', E8: 3, F8: 8, H8: 50,
    C9: 'Vuelta a la calma', D9: 'Estiramientos', N4: 'Minutos', O4: 'Zona', P4: 'Intervalos', C10: 'Aeróbico', D10: 'Bicicleta', N10: 20, O10: 'Z2' };
  const res = X.parseWorkbook_([{ name: 'Semana 1', cells: c }, { name: 'INSTRUCCIONES', cells: { A1: 'Cómo', D5: 'x' } }], null, null);
  eq(res.sesiones.map(s => s.fecha + ' ' + s.sesion + ': ' + s.ejercicios.map(e => e.name || e.aer.type).join(',')), ['2026-11-02 Pierna: Puente,Sentadilla,Zancada', '2026-11-05 Entreno jueves: Press + Remo,Bicicleta'], 'sesiones de la lista (con activación)');
  const f = X.filasDeSesion_('Sergio', res.sesiones[0], 'x', 'S', ['O', 'S']);
  eq(f.map(r => [r[X.PC.tipo], r[X.PC.ejercicio], r[X.PC.kg], r[X.PC.variante], r[X.PC.unilateral]]), [['Activación', 'Puente', '', '', false], ['Fuerza', 'Sentadilla', 12, '8-10', false], ['Fuerza', 'Zancada', 6, '', true]], 'activación, peso de su columna, nota y unilateral');
  const fj = X.filasDeSesion_('Sergio', res.sesiones[1], 'x', 'S', ['O', 'S']);
  eq(fj.map(r => [r[X.PC.tipo], r[X.PC.ejercicio], r[X.PC.minutos], r[X.PC.zona]]), [['Fuerza', 'Press', '', ''], ['Fuerza', 'Remo', '', ''], ['Aeróbico', 'Bicicleta', 20, 'Z2']], 'vuelta a la calma fuera y aeróbico al final');
}
console.log('· Plantilla de 4 tablas');
{
  const c = { A1: 'SEMANA DEL', B1: 46328,
    A3: 'SESIÓN 1', B3: 'Día:', C3: 'Lunes', E3: 'Nombre:', F3: 'Pierna',
    A4: 'Bloque', B4: 'Ejercicio', C4: 'Series', D4: 'Reps', E4: 'Nota reps', F4: 'Peso (kg)', G4: 'RPE', K4: 'Minutos', L4: 'Zona', M4: 'Intervalos',
    A5: 'Activación', B5: 'Puente', C5: 2, D5: 10, A6: 'P. Principal', B6: 'Sentadilla', C6: 3, D6: 8, F6: 60, A7: 'Aeróbico', B7: 'Bici', K7: 15, L7: 'Z2',
    A22: 'SESIÓN 2', B22: 'Día:', E22: 'Nombre:', A23: 'Bloque', B23: 'Ejercicio', C23: 'Series', D23: 'Reps', B24: 'Press', C24: 3, D24: 8,
    A41: 'SESIÓN 3', B41: 'Día:', C41: 'Jueves', A42: 'Bloque', B42: 'Ejercicio', C42: 'Series', D42: 'Reps' };
  const res = X.parseWorkbook_([{ name: 'Semana 1', cells: c }], null, null);
  eq(res.sesiones.map(s => s.fecha + ' ' + s.sesion + ': ' + s.ejercicios.map(e => e.name || e.aer.type).join(',')), ['2026-11-02 Pierna: Puente,Sentadilla,Bici'], 'tabla con día; vacía e incompleta fuera');
  ok(res.avisos.some(a => /sesión 2: falta el día/.test(a)), 'aviso: tabla con ejercicios sin día');
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
