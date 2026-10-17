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
