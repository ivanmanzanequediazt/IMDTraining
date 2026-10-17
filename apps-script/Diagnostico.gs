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

function moverDatosLuci() {
  moverDatos_('d19c2430a', 'd3729f74a');
}
