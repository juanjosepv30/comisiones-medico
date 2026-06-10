// ═══════════════════════════════════════════════════════════════
// Animal Center — Comisiones Médico
// Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME   = 'Registros';
const SUMMARY_SHEET = 'Resumen';
const SS_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

function getSpreadsheet() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SHEET_NAME) {
      sh.getRange(1, 1, 1, 14).setValues([[
        'ID', 'Fecha', 'Paciente', 'Propietario', 'Factura', 'Notas',
        'Descripción ítem', 'Categoría', 'Valor ítem', '% Comisión',
        'Comisión ítem', 'Total venta registro', 'Total comisión registro', 'Mes'
      ]]);
      sh.getRange(1, 1, 1, 14)
        .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidths(1, 14, 130);
    }
  }
  return sh;
}

// ── CORS helper ──────────────────────────────────────────────
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function addCors(output) {
  return output; // GAS maneja CORS automáticamente con doGet/doPost
}

// ── Router principal ─────────────────────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if      (action === 'sync')   result = syncEntries(data.entries);
    else if (action === 'getAll') result = getAllEntries();
    else if (action === 'delete') result = deleteEntry(data.id);
    else result = { ok: false, error: 'Acción desconocida' };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';
    let result;
    if (action === 'getAll') result = getAllEntries();
    else result = { ok: false, error: 'Acción desconocida' };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Sincronizar: recibe TODOS los entries del cliente ─────────
function syncEntries(entries) {
  const sh = getOrCreateSheet(SHEET_NAME);

  // Limpiar todo (menos cabecera)
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 14).clearContent();

  if (!entries || !entries.length) return { ok: true, rows: 0 };

  const rows = [];
  entries.forEach(entry => {
    const mes = entry.date ? entry.date.slice(0, 7) : '';
    entry.items.forEach(item => {
      rows.push([
        entry.id,
        entry.date,
        entry.patient,
        entry.owner,
        entry.fact,
        entry.notes || '',
        item.desc,
        item.cat,
        item.val,
        item.pct,
        item.com,
        entry.totalVal,
        entry.totalCom,
        mes
      ]);
    });
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 14).setValues(rows);
  }

  updateSummary();
  return { ok: true, rows: rows.length };
}

// ── Leer todos ───────────────────────────────────────────────
function getAllEntries() {
  const sh = getOrCreateSheet(SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, entries: [] };

  const data = sh.getRange(2, 1, lastRow - 1, 14).getValues();

  // Reagrupar por ID
  const map = {};
  data.forEach(row => {
    const [id, date, patient, owner, fact, notes,
           desc, cat, val, pct, com, totalVal, totalCom, mes] = row;
    if (!id) return;
    if (!map[id]) {
      map[id] = {
        id: Number(id), date, patient, owner, fact: String(fact),
        notes, items: [], totalVal: Number(totalVal), totalCom: Number(totalCom),
        catTotals: {}
      };
    }
    map[id].items.push({
      desc, cat, val: Number(val), pct: Number(pct), com: Number(com)
    });
    map[id].catTotals[cat] = (map[id].catTotals[cat] || 0) + Number(com);
  });

  return { ok: true, entries: Object.values(map) };
}

// ── Eliminar una entrada por ID ───────────────────────────────
function deleteEntry(id) {
  const sh = getOrCreateSheet(SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true };

  const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  // Borrar de abajo hacia arriba para no desajustar índices
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
    }
  }
  updateSummary();
  return { ok: true };
}

// ── Hoja de resumen mensual ───────────────────────────────────
function updateSummary() {
  const srcSh = getOrCreateSheet(SHEET_NAME);
  const ss    = getSpreadsheet();
  let sumSh   = ss.getSheetByName(SUMMARY_SHEET);
  if (!sumSh) sumSh = ss.insertSheet(SUMMARY_SHEET);
  sumSh.clearContents();

  const lastRow = srcSh.getLastRow();
  if (lastRow < 2) return;

  const data = srcSh.getRange(2, 1, lastRow - 1, 14).getValues();

  // Agrupar por mes
  const byMonth = {};
  data.forEach(row => {
    const mes     = row[13] || 'Sin fecha';
    const cat     = row[7];
    const comItem = Number(row[10]);
    const comReg  = Number(row[12]);
    const id      = row[0];

    if (!byMonth[mes]) byMonth[mes] = { ids: new Set(), totalCom: 0, cats: {} };
    byMonth[mes].ids.add(id);
    byMonth[mes].cats[cat] = (byMonth[mes].cats[cat] || 0) + comItem;
  });
  // totalCom por mes (único por registro)
  const perReg = {};
  data.forEach(row => {
    const k = row[0] + '_' + row[13];
    if (!perReg[k]) { perReg[k] = { mes: row[13], com: Number(row[12]) }; }
  });
  Object.values(perReg).forEach(r => {
    if (!byMonth[r.mes]) return;
    byMonth[r.mes].totalCom = (byMonth[r.mes].totalCom || 0) + r.com;
  });
  // Corregir: sumar comisiones por ítem (más exacto)
  Object.keys(byMonth).forEach(mes => {
    byMonth[mes].totalCom = Object.values(byMonth[mes].cats).reduce((a,b)=>a+b, 0);
  });

  const headers = [['Mes', 'Registros', 'Total Comisión',
    'Servicios/CX (30%)', 'Hosp/Urocult (20%)', 'Productos/PCR (10%)', 'Otras']];
  sumSh.getRange(1, 1, 1, 7).setValues(headers)
    .setBackground('#34a853').setFontColor('#fff').setFontWeight('bold');

  const rows = Object.entries(byMonth).sort((a,b)=>a[0]<b[0]?-1:1).map(([mes, d]) => {
    const c30  = (d.cats['servicios']||0) + (d.cats['cirugia']||0);
    const c20  = (d.cats['hospitalizacion']||0) + (d.cats['urocultivo']||0) +
                 (d.cats['medicacion']||0) + (d.cats['inyectologia']||0);
    const c10  = (d.cats['productos']||0) + (d.cats['pcr']||0);
    const otras = d.totalCom - c30 - c20 - c10;
    return [mes, d.ids.size, d.totalCom, c30, c20, c10, otras < 0 ? 0 : otras];
  });

  if (rows.length) {
    sumSh.getRange(2, 1, rows.length, 7).setValues(rows);
    sumSh.getRange(2, 3, rows.length, 5).setNumberFormat('$#,##0');
  }
  sumSh.setColumnWidths(1, 7, 160);
}
