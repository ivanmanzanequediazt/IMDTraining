/**
 * IMDTraining · importar la programación de los Excel de los clientes
 *
 * Pestaña «fuentes»: una fila por archivo (Apodo · Mes · Enlace · Importar · Estado · Última importación).
 * Menú IMDTraining > Importar de Excel: descarga cada archivo marcado, lee sus hojas de semana
 * ("05-09", "13-16"…) y escribe las sesiones en la pestaña «programacion». No publica: después se
 * revisa y se pulsa Publicar. Las filas importadas se marcan en la columna oculta «Origen» y se
 * rehacen en cada importación (lo que se corrija a mano en ellas se pierde: hay que corregirlo en el Excel).
 *
 * Formato del Excel (plantilla del equipo):
 *   - Hojas de semana con nombre "día-día" (p. ej. 05-09). Las demás hojas se ignoran.
 *   - En cada hoja, bloques con una fila de cabecera EJERCICIOS · SERIES · REPETICIONES · PESO · RPE · OBSERVACIONES.
 *     Encima de la cabecera: el día (LUNES, Martes…) y, opcionalmente, el nombre de la sesión (Tracc, Empuje…).
 *   - Debajo, un ejercicio por fila hasta la primera fila vacía. ACTIVACIÓN, P. PRINCIPAL, PP… son secciones y se saltan.
 */

const FUENTES_TAB = 'fuentes';
const FUENTES_HEADERS = ['Apodo', 'Mes', 'Enlace', 'Importar', 'Estado', 'Última importación'];
const DIAS = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const PRINCIPAL = ['pprincipal', 'pp', 'parteprincipal', 'principal', 'ppal'];
const SECCIONES = ['activacion', 'pprincipal', 'pp', 'parteprincipal', 'principal', 'ppal', 'calentamiento', 'vueltaalacalma', 'enfriamiento',
  'movilidad', 'core', 'finisher', 'accesorios', 'complementarios', 'fuerza', 'potencia', 'pliometria'];

/* ================= Pestaña fuentes ================= */

function prepararFuentes_() {
  const ss = sheet_(PROG_TAB).getParent();
  let f = ss.getSheetByName(FUENTES_TAB);
  if (!f) {
    f = ss.insertSheet(FUENTES_TAB);
    f.getRange(1, 1, 1, FUENTES_HEADERS.length).setValues([FUENTES_HEADERS]).setFontWeight('bold').setBackground('#e8eaed');
    f.setFrozenRows(1);
    [110, 110, 420, 75, 360, 140].forEach(function (w, i) { f.setColumnWidth(i + 1, w); });
    const n = f.getMaxRows() - 1;
    f.getRange(2, 1, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName(DEP_TAB).getRange('B2:B'), true).setAllowInvalid(true).build());
    f.getRange(2, 2, n, 1).setNumberFormat('@');
    f.getRange(2, 3, n, 1).setNumberFormat('@');
    f.getRange(2, 4, n, 1).insertCheckboxes();
    f.getRange(2, 5, n, 2).setFontColor('#5f6368').setFontStyle('italic');
    f.getRange('A1').setNote('El apodo del cliente (el mismo de «deportistas»).\nSi el archivo es de varios clientes que entrenan juntos: «O: Olga Ejemplo, S: Sergio Ejemplo», y en el Excel escribid los pesos como «O: 25 S: 12».');
    f.getRange('B1').setNote('Mes y año del archivo: 10/2026 u «octubre 2026». Si lo dejas vacío, se saca del nombre del archivo (PEPE_OCTUBRE_26).');
    f.getRange('C1').setNote('En OneDrive: Compartir > «Cualquier persona con el vínculo puede ver» > Copiar vínculo, y pégalo aquí.');
    f.getRange('D1').setNote('Marcado = se importa al pulsar «Importar de Excel». Desmárcalo en los meses que ya han pasado.');
  }
  return f;
}

/* ================= Importar ================= */

function importarExcel() {
  // Un bloqueo propio (no el de la sincronización) para que el automático y el botón no importen a la vez
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) { toast_('Ya se está importando. Prueba otra vez en un minuto.'); return false; }
  try { importar_(); return true; } finally { lock.releaseLock(); }
}

// Botón «Importar y publicar ahora» y modo automático
function importarYPublicar() {
  if (importarExcel() !== false) publicarProgramacion();
}

// Lo que ejecuta el automático cada 30 minutos (de 7:00 a 23:59; de noche no hace nada)
function importacionAutomatica() {
  const tz = sheet_(PROG_TAB).getParent().getSpreadsheetTimeZone();
  const h = +Utilities.formatDate(new Date(), tz, 'H');
  if (h < 7) return;
  importarYPublicar();
}

function activarAutomatico() {
  quitarAutomatico_();
  ScriptApp.newTrigger('importacionAutomatica').timeBased().everyMinutes(30).create();
  PropertiesService.getScriptProperties().setProperty('autoImport', '1');
  alert_('Automático activado: cada 30 minutos (de 7:00 a 24:00) se importan los Excel marcados en «fuentes» y se publican en las apps.\n\n' +
    'Para quitarlo: IMDTraining > Automático: desactivar.');
}

function desactivarAutomatico() {
  quitarAutomatico_();
  PropertiesService.getScriptProperties().deleteProperty('autoImport');
  alert_('Automático desactivado. Puedes seguir usando «Importar de Excel» y «Publicar» a mano.');
}

function quitarAutomatico_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importacionAutomatica') ScriptApp.deleteTrigger(t);
  });
}

function importar_() {
  const f = prepararFuentes_();
  const prog = sheet_(PROG_TAB);
  ensureOrigenCol_(prog);
  const n = f.getLastRow() - 1;
  if (n < 1) { toast_('Pega los enlaces en la pestaña «fuentes» y vuelve a pulsar Importar.'); f.activate(); return; }
  const tz = sheet_(PROG_TAB).getParent().getSpreadsheetTimeZone();
  const dep = readDeportistas_();
  const rows = f.getRange(2, 1, n, 6).getValues();
  let total = 0, conError = 0;
  rows.forEach(function (v, i) {
    const apodoTxt = str_(v[0]), mesTxt = v[1], link = str_(v[2]);
    if (!link || !bool_(v[3])) return;
    let estado;
    try {
      const quienes = parseQuienes_(apodoTxt);
      if (!quienes) throw new Error('falta el apodo (o, si el archivo es de varios, «O: Olga Ejemplo, S: Sergio Ejemplo»)');
      quienes.forEach(function (q) {
        if (!dep.byApodo[norm_(q.apodo)]) throw new Error('el apodo «' + q.apodo + '» no está en «deportistas»');
      });
      const dl = descargarXlsx_(link);
      const ym = parseMes_(mesTxt, tz) || mesDeNombre_(dl.name) || { y: null, m: null };   // la plantilla nueva no lo necesita
      const wb = readXlsx_(unzipXlsx_(dl.blob));
      const res = parseWorkbook_(wb, ym.y, ym.m);
      let borrar = 0;
      const letras = quienes.map(function (q) { return q.letra; }).filter(function (x) { return x; });
      quienes.forEach(function (q) {
        borrar += escribirImportadas_(prog, 'excel:' + link + (q.letra ? '#' + q.letra : ''), q.apodo, res.sesiones, q.letra, letras).borrar;
      });
      total += res.sesiones.length * quienes.length;
      estado = '✓ ' + res.sesiones.length + ' sesiones' + (quienes.length > 1 ? ' × ' + quienes.length + ' deportistas' : '') +
        (borrar ? ' · ' + borrar + ' para borrar' : '') +
        (res.avisos.length ? ' · Avisos: ' + res.avisos.join('; ') : '');
    } catch (e) {
      estado = 'Error: ' + (e && e.message ? e.message : e);
      conError++;
    }
    f.getRange(i + 2, 5, 1, 2).setValues([[estado, Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')]]);
  });
  SpreadsheetApp.flush();
  toast_('Importadas ' + total + ' sesiones' + (conError ? ' · ' + conError + ' archivo(s) con error (mira «fuentes»)' : '') + '.');
}

function ensureOrigenCol_(prog) {
  if (prog.getMaxColumns() < PROG_COLS) prog.insertColumnsAfter(prog.getMaxColumns(), PROG_COLS - prog.getMaxColumns());
  const h = prog.getRange(1, PC.origen + 1);
  if (h.getValue() !== 'Origen') {
    h.setValue('Origen').setFontWeight('bold').setBackground('#e8eaed');
    prog.getRange(2, PC.origen + 1, prog.getMaxRows() - 1, 1).setNumberFormat('@');
    prog.hideColumns(PC.origen + 1);
  }
}

// Última fila con algo escrito (las casillas sin marcar no cuentan: llegan hasta el final de la hoja)
function ultimaFilaConDatos_(prog) {
  const last = prog.getLastRow();
  if (last < 2) return 1;
  const vals = prog.getRange(2, 1, last - 1, PC.accion + 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(function (x) { return x !== '' && x !== false && x !== null; })) return i + 2;
  }
  return 1;
}

// Sustituye las filas de este origen por las nuevas. Conserva el id (columna oculta) de las sesiones que siguen
// existiendo, para que Publicar las actualice en vez de crear otras; las que ya no están en el Excel y se habían
// publicado quedan en una fila con Acción «Borrar», para que Publicar las quite de la app.
function escribirImportadas_(prog, origen, apodo, sesiones, letra, letras) {
  const last = prog.getLastRow();
  const vals = last > 1 ? prog.getRange(2, 1, last - 1, PROG_COLS).getValues() : [];
  const fmt = function (d) { return Utilities.formatDate(d, prog.getParent().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); };
  const plan = planImportacion_(vals, origen, apodo, sesiones, fmt, letra, letras);
  // Borrar las filas antiguas de este origen (de abajo arriba, por tramos seguidos)
  const del = plan.borrarFilas.slice().sort(function (a, b) { return b - a; });
  let k = 0;
  while (k < del.length) {
    let end = k;
    while (end + 1 < del.length && del[end + 1] === del[end] - 1) end++;
    prog.deleteRows(del[end] + 2, end - k + 1);
    k = end + 1;
  }
  if (plan.nuevas.length) {
    const start = ultimaFilaConDatos_(prog) + 1;
    if (start + plan.nuevas.length - 1 > prog.getMaxRows()) prog.insertRowsAfter(prog.getMaxRows(), start + plan.nuevas.length - prog.getMaxRows());
    prog.getRange(start, 1, plan.nuevas.length, PROG_COLS).setValues(plan.nuevas);
  }
  return { borrar: plan.borrar };
}

/* ================= Descarga ================= */

// Prueba varias formas de descargar un archivo compartido de OneDrive con «Cualquier persona con el vínculo».
function descargarXlsx_(link) {
  const tries = [];
  const add = function (u) { if (u && tries.indexOf(u) < 0) tries.push(u); };
  const b64 = Utilities.base64EncodeWebSafe(link).replace(/=+$/, '');
  add('https://api.onedrive.com/v1.0/shares/u!' + b64 + '/root/content');
  add(link + (link.indexOf('?') >= 0 ? '&' : '?') + 'download=1');
  // Enlaces cortos 1drv.ms: seguir la redirección y pedir la descarga directa
  try {
    const r = UrlFetchApp.fetch(link, { followRedirects: false, muteHttpExceptions: true });
    const loc = r.getHeaders()['Location'] || r.getHeaders()['location'];
    if (loc) {
      add(loc.replace('/redir?', '/download?').replace('/embed?', '/download?'));
      add(loc + (loc.indexOf('?') >= 0 ? '&' : '?') + 'download=1');
    }
  } catch (e) { /* seguimos con lo demás */ }
  add(link);
  const fallos = [];
  // Enlaces nuevos de OneDrive personal: la web de OneDrive pide primero un permiso de invitado ("Badger") y con él
  // descarga el archivo compartido. No es una vía documentada por Microsoft: si deja de funcionar, avisará el error.
  try {
    const tk = UrlFetchApp.fetch('https://api-badgerp.svc.ms/v1.0/token', { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ appId: '5cbed6ac-a083-4e14-b191-b4ba07653de2' }), muteHttpExceptions: true });
    const token = tk.getResponseCode() === 200 ? JSON.parse(tk.getContentText()).token : null;
    if (token) {
      const h = { Authorization: 'Badger ' + token, Prefer: 'autoredeem' };
      const base = 'https://my.microsoftpersonalcontent.com/_api/v2.0/shares/u!' + b64 + '/driveitem';
      const r1 = UrlFetchApp.fetch(base + '/content', { headers: h, followRedirects: true, muteHttpExceptions: true });
      const ok1 = descargaValida_(r1);
      if (ok1) return ok1;
      const r2 = UrlFetchApp.fetch(base, { headers: h, muteHttpExceptions: true });
      if (r2.getResponseCode() === 200) {
        const it = JSON.parse(r2.getContentText());
        const url = it['@content.downloadUrl'] || it['@microsoft.graph.downloadUrl'];
        if (url) {
          const ok2 = descargaValida_(UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true }), it.name);
          if (ok2) return ok2;
        }
      }
      fallos.push('b' + r1.getResponseCode() + '/' + r2.getResponseCode());
    } else fallos.push('b' + tk.getResponseCode());
  } catch (e) { fallos.push('bx'); }
  for (let i = 0; i < tries.length; i++) {
    try {
      const r = UrlFetchApp.fetch(tries[i], { followRedirects: true, muteHttpExceptions: true });
      const ok = descargaValida_(r);
      if (ok) return ok;
      fallos.push(r.getResponseCode());
    } catch (e) { fallos.push('x'); }
  }
  throw new Error('no se pudo descargar (' + fallos.join(',') + '). Comprueba que el archivo está compartido como ' +
    '«Cualquier persona con el vínculo puede ver» y que es .xlsx');
}

// Respuesta que es de verdad un .xlsx (empieza por "PK", es un zip) -> {blob, name}; si no, null
function descargaValida_(r, nombre) {
  const bytes = r.getContent();
  if (r.getResponseCode() !== 200 || bytes.length < 5 || bytes[0] !== 0x50 || bytes[1] !== 0x4B) return null;
  const cd = String(r.getHeaders()['Content-Disposition'] || r.getHeaders()['content-disposition'] || '');
  const m = cd.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
  return { blob: r.getBlob(), name: nombre || (m ? decodeURIComponent(m[1]) : '') };
}

function unzipXlsx_(blob) {
  const files = {};
  Utilities.unzip(blob.setContentType('application/zip')).forEach(function (b) {
    const name = b.getName();
    if (/^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|worksheets\/sheet\d+\.xml)$/.test(name)) {
      files[name] = b.getDataAsString('UTF-8');
    }
  });
  if (!files['xl/workbook.xml']) throw new Error('el archivo no es un Excel .xlsx (los .xls antiguos hay que guardarlos como .xlsx)');
  return files;
}

/* ================= Núcleo (sin servicios de Google; se prueba con Node) ================= */

function xmlText_(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, function (a, h) { return String.fromCodePoint(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (a, d) { return String.fromCodePoint(+d); })
    .replace(/&amp;/g, '&');
}

function attr_(tag, name) {
  const m = tag.match(new RegExp('(?:^|\\s)' + name + '="([^"]*)"'));
  return m ? xmlText_(m[1]) : null;
}

// files: {ruta: texto XML} -> [{name, cells: {A1: valor}}]
function readXlsx_(files) {
  const shared = [];
  const sst = files['xl/sharedStrings.xml'] || '';
  (sst.match(/<si>[\s\S]*?<\/si>/g) || []).forEach(function (si) {
    const noPh = si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let t = '';
    (noPh.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).forEach(function (x) {
      const m = x.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);
      if (m) t += xmlText_(m[1]);
    });
    shared.push(t);
  });
  const rels = {};
  (String(files['xl/_rels/workbook.xml.rels'] || '').match(/<Relationship\b[^>]*>/g) || []).forEach(function (r) {
    let t = attr_(r, 'Target') || '';
    t = t.replace(/^\/?xl\//, '').replace(/^\//, '');
    rels[attr_(r, 'Id')] = 'xl/' + t;
  });
  const out = [];
  (String(files['xl/workbook.xml']).match(/<sheet\b[^>]*>/g) || []).forEach(function (tag) {
    const name = attr_(tag, 'name');
    const path = rels[attr_(tag, 'r:id')];
    const xml = path && files[path];
    const cells = {};
    if (xml) {
      const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let m;
      while ((m = re.exec(xml))) {
        const a = m[1], body = m[2] || '';
        const ref = attr_(a, 'r');
        const t = attr_(a, 't');
        let val = null;
        if (t === 'inlineStr') {
          const tt = body.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || [];
          val = tt.map(function (x) { return xmlText_(x.replace(/<[^>]+>/g, '')); }).join('');
        } else {
          const v = body.match(/<v>([\s\S]*?)<\/v>/);
          if (v) {
            if (t === 's') val = shared[+v[1]];
            else if (t === 'str' || t === 'e') val = xmlText_(v[1]);
            else if (t === 'b') val = v[1] === '1';
            else val = Number(v[1]);
          }
        }
        if (ref && val !== null && val !== '') cells[ref] = val;
      }
    }
    out.push({ name: name, cells: cells });
  });
  return out;
}

function colNum_(letters) { let n = 0; for (let i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n; }
function colLet_(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; }

function sinAcentos_(s) {
  return String(s).toLowerCase().replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u');
}

function diaDe_(v) {
  if (typeof v !== 'string') return -1;
  const s = sinAcentos_(v).replace(/[^a-z]/g, '');
  if (!s || s.length > 14) return -1;
  for (let i = 0; i < DIAS.length; i++) if (s === DIAS[i] || (s.length <= DIAS[i].length + 2 && s.indexOf(DIAS[i]) >= 0)) return i;
  return -1;
}

function esSeccion_(name) {
  const s = sinAcentos_(name).replace(/[^a-z]/g, '');
  return SECCIONES.indexOf(s) >= 0;
}

// '10/2026', '10-26', 'octubre 2026', fecha -> {y, m}
function parseMes_(v, tz) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return { y: +Utilities.formatDate(v, tz, 'yyyy'), m: +Utilities.formatDate(v, tz, 'M') };
  }
  const s = sinAcentos_(String(v)).trim();
  let m = s.match(/^(\d{1,2})\s*[\/\-. ]\s*(\d{2}|\d{4})$/);
  if (m && +m[1] >= 1 && +m[1] <= 12) return { m: +m[1], y: m[2].length === 2 ? 2000 + +m[2] : +m[2] };
  m = s.match(/([a-z]+)\D*(\d{2}|\d{4})?/);
  if (m) {
    const i = MESES.indexOf(m[1]) >= 0 ? MESES.indexOf(m[1]) : MESES.map(function (x) { return x.slice(0, 3); }).indexOf(m[1].slice(0, 3));
    if (i >= 0) return { m: i + 1, y: m[2] ? (m[2].length === 2 ? 2000 + +m[2] : +m[2]) : new Date().getFullYear() };
  }
  return null;
}

// PEPE_OCTUBRE_26.xlsx -> {y:2026, m:10}; sin año -> el año en curso
function mesDeNombre_(name) {
  const s = sinAcentos_(String(name || '')).replace(/\.xlsx?$/, '');
  const parts = s.split(/[_\-\s.]+/);
  for (let i = 0; i < parts.length; i++) {
    const k = MESES.indexOf(parts[i]);
    if (k >= 0) {
      const yy = parts[i + 1] && /^\d{2}(\d{2})?$/.test(parts[i + 1]) ? parts[i + 1] : '';
      return { m: k + 1, y: yy ? (yy.length === 2 ? 2000 + +yy : +yy) : new Date().getFullYear() };
    }
  }
  return null;
}

// Primer número de un texto ("4-6rep" -> 4; "17'5 lado" -> 17.5; "3 x 4''" -> 3)
function primerNumero_(v, decimales) {
  if (typeof v === 'number') return v;
  const s = String(v || '');
  const m = decimales ? s.match(/\d+(?:[.,']\d+)?(?!['´])/) || s.match(/\d+/) : s.match(/\d+/);
  if (!m) return null;
  return Number(m[0].replace(/[,']/g, '.'));
}

function ymdStr_(y, m, d) {
  const t = new Date(Date.UTC(y, m - 1, d));
  return t.getUTCFullYear() + '-' + ('0' + (t.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + t.getUTCDate()).slice(-2);
}

// Fecha del bloque: el día de la semana dentro del rango del nombre de la hoja ("05-09" de octubre de 2026)
function fechaBloque_(d1, d2, y, m, dow) {
  const out = [];
  const start = Date.UTC(y, m - 1, d1);
  const len = d2 >= d1 ? d2 - d1 : (new Date(Date.UTC(y, m, 0)).getUTCDate() - d1) + d2; // 29-02: cruza de mes
  for (let i = 0; i <= Math.min(len, 9); i++) {
    const t = new Date(start + i * 864e5);
    if (t.getUTCDay() === dow) out.push(ymdStr_(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()));
  }
  return out.length === 1 ? out[0] : null;
}

/* ---------- Plantilla nueva: una lista por semana ---------- */
// Fila 1: «SEMANA DEL (lunes)» + fecha. Cabecera: Día · Sesión · Bloque · Ejercicio · Series · Reps · Nota reps ·
// Peso (o «Peso X» por cliente) · RPE · Unilateral · Pliometría · Observaciones. Día, Sesión y Bloque vacíos = los de arriba.

function esHojaLista_(at) {
  for (let r = 1; r <= 3; r++) for (let c = 1; c <= 3; c++) {
    if (sinAcentos_(String(at(r, c))).replace(/[^a-z]/g, '').indexOf('semanadel') === 0) return true;
  }
  return false;
}

// Fecha de Excel (número de serie o texto dd/mm/aaaa) -> 'YYYY-MM-DD' | null
function fechaExcel_(v) {
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
    return ymdStr_(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
  }
  const m = String(v || '').trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/);
  if (m) return ymdStr_(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
  return null;
}

function parseLista_(hoja, at, txt, maxRow, bloques, avisos, sinPrincipal) {
  // Fecha de inicio de semana: la celda a la derecha de «SEMANA DEL»
  let inicio = null;
  for (let r = 1; r <= 3 && !inicio; r++) for (let c = 1; c <= 3 && !inicio; c++) {
    if (sinAcentos_(txt(r, c)).replace(/[^a-z]/g, '').indexOf('semanadel') === 0) {
      for (let k = c + 1; k <= c + 3 && !inicio; k++) inicio = fechaExcel_(at(r, k));
    }
  }
  // Cabecera
  let hr = 0;
  const col = {}, pesos = {};
  for (let r = 1; r <= 12 && !hr; r++) {
    const hs = {};
    for (let c = 1; c <= 30; c++) hs[c] = sinAcentos_(txt(r, c)).replace(/[^a-z0-9 ]/g, '').trim();
    const vals = Object.keys(hs).map(function (c) { return hs[c]; });
    if (vals.indexOf('ejercicio') < 0 && vals.indexOf('ejercicios') < 0) continue;
    if (vals.indexOf('dia') < 0) continue;
    hr = r;
    Object.keys(hs).forEach(function (c) {
      const h = hs[c], n = +c, k = h.replace(/ /g, '');
      if (k === 'dia') col.dia = n;
      else if (k === 'sesion') col.sesion = n;
      else if (k === 'bloque' || k === 'parte') col.bloque = n;
      else if (k === 'ejercicio' || k === 'ejercicios') col.name = n;
      else if (k === 'series') col.series = n;
      else if (k === 'reps' || k === 'repeticiones') col.reps = n;
      else if (k.indexOf('notarep') === 0) col.nota = n;
      else if (k === 'rpe') col.rpe = n;
      else if (k === 'unilateral') col.uni = n;
      else if (k.indexOf('pliometr') === 0) col.plyo = n;
      else if (k.indexOf('observ') === 0) col.obs = n;
      else if (/^peso/.test(k)) {
        const m = h.match(/^peso\s+([a-z])$/);
        if (m) pesos[m[1].toUpperCase()] = n; else col.peso = n;
      }
    });
  }
  if (!hr || !col.name) { avisos.push('hoja ' + hoja + ': no encuentro la fila de cabecera (Día · Sesión · … · Ejercicio)'); return; }
  const sesiones = {}, orden = [];
  let dia = -1, sesion = '', bloque = '', vacias = 0, haySin = false;
  for (let r = hr + 1; r <= maxRow; r++) {
    const name = txt(r, col.name);
    const fila = [col.dia, col.sesion, col.bloque, col.name, col.series, col.reps].some(function (c) { return c && txt(r, c) !== ''; });
    if (!fila) { if (++vacias >= 10) break; continue; }
    vacias = 0;
    // Un día nuevo empieza sesión nueva: su nombre y su bloque no se heredan del día anterior
    if (col.dia && txt(r, col.dia)) {
      const d = diaDe_(txt(r, col.dia));
      if (d >= 0) { dia = d; sesion = ''; bloque = ''; } else avisos.push('hoja ' + hoja + ' fila ' + r + ': día «' + txt(r, col.dia) + '» no reconocido');
    }
    if (col.sesion && txt(r, col.sesion)) { sesion = txt(r, col.sesion); if (!(col.bloque && txt(r, col.bloque))) bloque = ''; }
    if (col.bloque && txt(r, col.bloque)) bloque = sinAcentos_(txt(r, col.bloque)).replace(/[^a-z]/g, '');
    if (!name) continue;
    if (dia < 0) { haySin = true; continue; }
    const k = dia + '|' + norm_(sesion);
    if (!sesiones[k]) { sesiones[k] = { dia: dia, label: sesion, ejercicios: [], principal: false }; orden.push(k); }
    const ps = {};
    Object.keys(pesos).forEach(function (l) { const v = at(r, pesos[l]); if (v !== '') ps[l] = String(v); });
    const esPrincipal = PRINCIPAL.indexOf(bloque) >= 0;
    if (esPrincipal) sesiones[k].principal = true;
    const si = function (c) { return c ? /^(si|s|x|1|true)$/.test(sinAcentos_(txt(r, c)).replace(/[^a-z0-9]/g, '')) : false; };
    sesiones[k].ejercicios.push({ name: name.replace(/\s+/g, ' '), series: at(r, col.series || 0), repsTxt: col.reps ? txt(r, col.reps) : '',
      nota: col.nota ? txt(r, col.nota) : '', pesoTxt: col.peso ? txt(r, col.peso) : '', pesos: Object.keys(pesos).length ? ps : null,
      rpe: col.rpe ? at(r, col.rpe) : '', obs: col.obs ? txt(r, col.obs) : '', unilateral: si(col.uni), plyo: si(col.plyo), principal: esPrincipal });
  }
  if (haySin) avisos.push('hoja ' + hoja + ': hay ejercicios sin día (pon el día en la primera fila de cada sesión)');
  if (!orden.length) return;
  if (!inicio) { avisos.push('hoja ' + hoja + ': falta la fecha de «SEMANA DEL»'); return; }
  const t0 = new Date(inicio + 'T00:00:00Z');
  orden.forEach(function (k) {
    const s = sesiones[k];
    const off = (s.dia - t0.getUTCDay() + 7) % 7;
    const t = new Date(t0.getTime() + off * 864e5);
    let ej = s.ejercicios;
    if (s.principal) ej = ej.filter(function (e) { return e.principal; });
    else sinPrincipal.push(hoja + ' ' + DIAS[s.dia].replace('miercoles', 'miércoles').replace('sabado', 'sábado'));
    if (!ej.length) return;
    bloques.push({ hoja: hoja, fecha: ymdStr_(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()), dia: DIAS[s.dia], label: s.label, ejercicios: ej });
  });
}

// wb (readXlsx_) -> {sesiones:[{fecha, sesion, hoja, ejercicios:[{name, series, repsTxt, pesoTxt, rpe, obs}]}], avisos:[]}
function parseWorkbook_(wb, y, m) {
  const avisos = [];
  const bloques = [];
  const sinPrincipal = [];
  let sinNombre = 0;
  wb.forEach(function (sh) {
    const grid = {};
    let maxRow = 0;
    Object.keys(sh.cells).forEach(function (ref) {
      const p = ref.match(/^([A-Z]+)(\d+)$/);
      if (!p) return;
      const r = +p[2], c = colNum_(p[1]);
      (grid[r] = grid[r] || {})[c] = sh.cells[ref];
      if (r > maxRow) maxRow = r;
    });
    const at = function (r, c) { const v = grid[r] && grid[r][c]; return v === undefined || v === null ? '' : v; };
    const txt = function (r, c) { return String(at(r, c)).trim(); };
    // Plantilla nueva (una lista por semana): «SEMANA DEL (lunes)» arriba a la izquierda
    if (esHojaLista_(at)) { parseLista_(sh.name, at, txt, maxRow, bloques, avisos, sinPrincipal); return; }
    const nm = String(sh.name).trim().match(/^(\d{1,2})\s*[-–_]\s*(\d{1,2})$/);
    if (!nm) return;                                   // ANUAL, RPE SERIES, RM BÁSICOS, INSTRUCCIONES…
    if (!y || !m) { avisos.push('hoja ' + sh.name + ': pon el mes en la columna Mes de «fuentes»'); return; }
    const d1 = +nm[1], d2 = +nm[2];
    for (let r = 1; r <= maxRow; r++) {
      Object.keys(grid[r] || {}).forEach(function (cs) {
        const c = +cs;
        if (sinAcentos_(txt(r, c)).replace(/[^a-z]/g, '') !== 'ejercicios') return;
        // Columnas por su cabecera (con valores por defecto si alguna falta)
        const col = { name: c, series: c + 1, reps: c + 2, peso: c + 3, rpe: c + 4, obs: c + 5 };
        for (let k = c + 1; k <= c + 7; k++) {
          const h = sinAcentos_(txt(r, k)).replace(/[^a-z]/g, '');
          if (h === 'series') col.series = k;
          else if (h.indexOf('repet') === 0 || h === 'reps') col.reps = k;
          else if (h === 'peso' || h === 'kg' || h === 'carga') col.peso = k;
          else if (h === 'rpe') col.rpe = k;
          else if (h.indexOf('observ') === 0) col.obs = k;
        }
        // Día y nombre de la sesión: en las 3 filas de encima
        let dow = -1, label = '';
        for (let up = 1; up <= 3 && r - up >= 1; up++) {
          const v = txt(r - up, c);
          if (!v) continue;
          const dd = diaDe_(v);
          if (dd >= 0) { dow = dd; break; }
          if (!label) label = v;
        }
        const where = 'hoja ' + sh.name + ' celda ' + colLet_(c) + r;
        if (dow < 0) { avisos.push('bloque sin día (' + where + ')'); return; }
        const fecha = fechaBloque_(d1, d2, y, m, dow);
        if (!fecha) { avisos.push(DIAS[dow] + ' no cae entre los días ' + sh.name + ' (' + where + ')'); return; }
        const ejercicios = [];
        let vacias = 0, trasHueco = false, seccion = '', hayPrincipal = false;
        for (let rr = r + 1; rr <= Math.min(maxRow, r + 40); rr++) {
          const v = { name: txt(rr, col.name), series: at(rr, col.series), reps: at(rr, col.reps), peso: at(rr, col.peso),
            rpe: at(rr, col.rpe), obs: txt(rr, col.obs) };
          const otros = [v.series, v.reps, v.peso, v.rpe, v.obs].some(function (x) { return x !== ''; });
          if (!v.name && !otros) {                       // fila vacía: 3 seguidas = fin del bloque
            if (++vacias >= 3) break;
            trasHueco = true;
            continue;
          }
          vacias = 0;
          if (v.name && (diaDe_(v.name) >= 0 || sinAcentos_(v.name).replace(/[^a-z]/g, '') === 'ejercicios')) break;
          if (v.name && !otros && (esSeccion_(v.name) || (v.name === v.name.toUpperCase() && /[A-ZÁÉÍÓÚÑ]{2}/.test(v.name) && v.name.length <= 20))) {
            seccion = sinAcentos_(v.name).replace(/[^a-z]/g, '');
            if (PRINCIPAL.indexOf(seccion) >= 0) hayPrincipal = true;
            continue;
          }
          // Después de un hueco, solo cuentan las filas con series o repeticiones (no las notas sueltas de abajo)
          if (trasHueco && v.series === '' && v.reps === '') continue;
          if (!v.name) { sinNombre++; continue; }
          ejercicios.push({ name: v.name.replace(/\s+/g, ' '), series: v.series, repsTxt: String(v.reps).trim(),
            pesoTxt: String(v.peso).trim(), rpe: v.rpe, obs: v.obs, principal: PRINCIPAL.indexOf(seccion) >= 0 });
        }
        // Solo la parte principal; si el bloque no tiene «P. PRINCIPAL», se importa entero (y se avisa)
        if (hayPrincipal) {
          for (let k = ejercicios.length - 1; k >= 0; k--) if (!ejercicios[k].principal) ejercicios.splice(k, 1);
        } else if (ejercicios.length) {
          sinPrincipal.push(sh.name + ' ' + DIAS[dow].replace('miercoles', 'miércoles').replace('sabado', 'sábado'));
        }
        if (!ejercicios.length) return;
        bloques.push({ hoja: sh.name, fecha: fecha, dia: DIAS[dow], label: label, ejercicios: ejercicios });
      });
    }
  });
  // Hojas que dan el mismo día: no se importa ninguna de las dos
  const hojasPorFecha = {};
  bloques.forEach(function (b) { (hojasPorFecha[b.fecha] = hojasPorFecha[b.fecha] || {})[b.hoja] = true; });
  const malas = {};
  Object.keys(hojasPorFecha).forEach(function (f) {
    const hs = Object.keys(hojasPorFecha[f]);
    if (hs.length > 1) hs.forEach(function (h) { malas[h] = hs; });
  });
  Object.keys(malas).sort().forEach(function (h) {
    const otras = malas[h].filter(function (x) { return x !== h; });
    if (h < otras[0]) avisos.push('las hojas ' + malas[h].sort().join(' y ') + ' tienen días repetidos: no se importan');
  });
  const sesiones = [];
  const usados = {};
  bloques.filter(function (b) { return !malas[b.hoja]; }).forEach(function (b) {
    let sesion = b.label || ('Entreno ' + b.dia.replace('miercoles', 'miércoles').replace('sabado', 'sábado'));
    const k0 = b.fecha + '|' + norm_(sesion);
    if (usados[k0]) { let i = 2; while (usados[b.fecha + '|' + norm_(sesion + ' (' + i + ')')]) i++; sesion += ' (' + i + ')'; }
    usados[b.fecha + '|' + norm_(sesion)] = true;
    sesiones.push({ fecha: b.fecha, sesion: sesion, hoja: b.hoja, ejercicios: b.ejercicios });
  });
  sesiones.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
  if (sinPrincipal.length) avisos.push('sin «P. PRINCIPAL», importado entero: ' + sinPrincipal.join(', '));
  if (sinNombre) avisos.push(sinNombre + ' fila' + (sinNombre === 1 ? '' : 's') + ' con series pero sin ejercicio (saltada' + (sinNombre === 1 ? '' : 's') + ')');
  return { sesiones: sesiones, avisos: avisos };
}

// Columna Apodo de «fuentes»: "Pepe Ejemplo" o, si el archivo es de varios, "O: Olga Ejemplo, S: Sergio Ejemplo"
// -> [{letra:null, apodo}] | [{letra:'O', apodo}, {letra:'S', apodo}] | null
function parseQuienes_(txt) {
  const s = str_(txt);
  if (!s) return null;
  if (!/^\s*[A-Za-zÁÉÍÓÚÑáéíóúñ]\s*[:=]/.test(s)) return [{ letra: null, apodo: s }];
  const out = [];
  s.split(/[,;\n]+/).forEach(function (p) {
    const m = p.trim().match(/^([A-Za-zÁÉÍÓÚÑáéíóúñ])\s*[:=]\s*(.+)$/);
    if (m) out.push({ letra: m[1].toUpperCase(), apodo: m[2].trim() });
  });
  return out.length ? out : null;
}

// Celda compartida "O: 25 S: 12" -> lo de la letra pedida ("25"). Sin letras, vale para todos.
// Devuelve {txt, val}: txt = texto de esta persona (sin las partes de los demás), val = número si la celda es número.
function valorDe_(v, letra, letras) {
  if (typeof v === 'number') return { txt: String(v), val: v };
  const s = String(v === null || v === undefined ? '' : v).trim();
  if (!letra || !letras || !letras.length || !s) return { txt: s, val: null };
  const set = letras.join('') + letras.join('').toLowerCase();
  const re = new RegExp('(^|[^A-Za-zÁÉÍÓÚÑáéíóúñ])([' + set + '])\\s*[:=]', 'g');
  const marcas = [];
  let m;
  while ((m = re.exec(s))) marcas.push({ letra: m[2].toUpperCase(), ini: m.index + m[1].length, fin: re.lastIndex });
  if (!marcas.length) return { txt: s, val: null };
  const partes = {};
  marcas.forEach(function (k, i) {
    partes[k.letra] = s.slice(k.fin, i + 1 < marcas.length ? marcas[i + 1].ini : s.length).replace(/[\s,;\/|-]+$/, '').trim();
  });
  return { txt: partes[letra] || '', val: null };
}

// Sesiones -> filas de «programacion» (22 columnas + Origen)
// "Sentadilla + plancha" -> ["Sentadilla", "plancha"]; los paréntesis se quitan ("A + (B + C)" -> A, B, C)
function partesBiserie_(name) {
  return String(name).split('+').map(function (x) { return x.replace(/[()]/g, ' ').replace(/\s+/g, ' ').trim(); })
    .filter(function (x) { return x; });
}

// "8 rep / 6-8rep" o "8, 12" -> una parte por ejercicio de la biserie (si cuadra el número de partes; si no, null)
function repartir_(txt, n) {
  if (n < 2 || !txt) return null;
  const p = String(txt).split(/\s*[\/,;]\s*/).map(function (x) { return x.trim(); }).filter(function (x) { return x; });
  return p.length === n ? p : null;
}

function filasDeSesion_(apodo, s, origen, letra, letras) {
  const notas = [];
  const filas = [];
  s.ejercicios.forEach(function (e) {
    const nombres = partesBiserie_(e.name);
    const n = nombres.length || 1;
    const tipo = n === 2 ? 'biserie' : n === 3 ? 'triserie' : n > 3 ? 'circuito' : '';
    const series = typeof e.series === 'number' ? Math.round(e.series) : primerNumero_(e.series, false);
    const repsP = repartir_(e.repsTxt, n);
    const peso = e.pesos ? { txt: letra && e.pesos[letra] !== undefined ? String(e.pesos[letra]) : (letra ? '' : e.pesoTxt), val: null }
      : valorDe_(e.pesoTxt, letra, letras);
    const pesoP = repartir_(peso.txt, n);
    const rpe = valorDe_(e.rpe, letra, letras);
    const rpeN = typeof rpe.val === 'number' ? rpe.val : primerNumero_(rpe.txt, true);
    (nombres.length ? nombres : [e.name]).forEach(function (nombre, k) {
      const r = PROG_HEADERS.map(function () { return ''; });
      r[PC.apodo] = apodo; r[PC.fecha] = s.fecha; r[PC.sesion] = s.sesion; r[PC.tipo] = 'Fuerza';
      r[PC.ejercicio] = nombre;
      r[PC.unilateral] = !!e.unilateral; r[PC.plyo] = !!e.plyo;
      r[PC.series] = series && series >= 1 && series <= 20 ? series : 1;
      const repsTxt = repsP ? repsP[k] : e.repsTxt;
      const reps = primerNumero_(repsTxt, false);
      r[PC.reps] = reps && reps >= 1 ? reps : '';
      const textoReps = repsTxt && !/^\d+$/.test(repsTxt) ? repsTxt.replace(/\s+/g, ' ') : '';
      r[PC.variante] = [tipo, textoReps, e.nota || ''].filter(function (x) { return x; }).join(' · ');
      // Peso: repartido si viene separado ("40 / 10"); si es uno solo, es del primer ejercicio de la biserie
      const pesoTxt = pesoP ? pesoP[k] : (k === 0 ? peso.txt : '');
      const kg = pesoTxt === '' ? null : primerNumero_(pesoTxt, true);
      if (kg !== null && kg > 0 && kg <= 1000 && !e.plyo) r[PC.kg] = kg;
      if (rpeN !== null && rpeN >= 0 && rpeN <= 10 && rpe.txt !== '') r[PC.rpe] = rpeN;
      r[PC.origen] = origen;
      filas.push(r);
    });
    if (peso.txt && /[a-záéíóúñ]{2}/i.test(peso.txt)) notas.push(e.name + ': ' + peso.txt);
    if (e.obs) notas.push(e.name + ': ' + e.obs);
  });
  if (notas.length && filas.length) filas[0][PC.notas] = notas.join(' · ');
  return filas;
}

// vals: filas actuales de «programacion» (sin cabecera). Devuelve qué filas borrar (índices 0..) y las filas nuevas.
function planImportacion_(vals, origen, apodo, sesiones, fmt, letra, letras) {
  const key = function (a, f, s) { return norm_(a) + '|' + f + '|' + norm_(s); };
  const viejas = {};          // clave -> {id, estado}
  const borrarFilas = [];
  vals.forEach(function (v, i) {
    if (String(v[PC.origen]) !== origen) return;
    borrarFilas.push(i);
    const f = date_(v[PC.fecha], fmt);
    const k = key(v[PC.apodo], f, v[PC.sesion]);
    const id = str_(v[PC.id]);
    const yaBorrada = norm_(v[PC.accion]) === 'borrar' && /^Borrado/.test(String(v[PC.estado]));
    if (!viejas[k]) viejas[k] = { id: '', borrada: yaBorrada, apodo: str_(v[PC.apodo]), fecha: f, sesion: str_(v[PC.sesion]) };
    if (id && !viejas[k].id) viejas[k].id = id;
  });
  const nuevas = [];
  const vistas = {};
  sesiones.forEach(function (s) {
    const k = key(apodo, s.fecha, s.sesion);
    vistas[k] = true;
    const filas = filasDeSesion_(apodo, s, origen, letra, letras);
    const old = viejas[k];
    if (old && old.id) filas.forEach(function (r) { r[PC.id] = old.id; });
    filas.forEach(function (r) { nuevas.push(r); });
  });
  let borrar = 0;
  Object.keys(viejas).forEach(function (k) {
    const o = viejas[k];
    if (vistas[k] || !o.id || o.borrada) return;      // sigue en el Excel, nunca se publicó o ya se quitó de la app
    const r = PROG_HEADERS.map(function () { return ''; });
    r[PC.apodo] = o.apodo; r[PC.fecha] = o.fecha; r[PC.sesion] = o.sesion; r[PC.ejercicio] = '(quitada del Excel)';
    r[PC.unilateral] = false; r[PC.plyo] = false;
    r[PC.accion] = 'Borrar'; r[PC.id] = o.id; r[PC.origen] = origen;
    nuevas.push(r);
    borrar++;
  });
  return { borrarFilas: borrarFilas, nuevas: nuevas, borrar: borrar };
}

if (typeof module !== 'undefined') module.exports = { fechaExcel_: fechaExcel_, partesBiserie_: partesBiserie_, parseQuienes_: parseQuienes_, valorDe_: valorDe_, readXlsx_: readXlsx_, parseWorkbook_: parseWorkbook_, mesDeNombre_: mesDeNombre_,
  filasDeSesion_: filasDeSesion_, planImportacion_: planImportacion_, primerNumero_: primerNumero_ };
