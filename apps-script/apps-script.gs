// Scora: guarda cada consulta del formulario de la web en esta planilla.
// Pegalo en Extensiones > Apps Script de tu Google Sheet y publicalo como aplicación web.

const SHEET_NAME = 'Consultas';
const HEADERS = ['Fecha', 'Nombre', 'Email', 'WhatsApp', 'Rubro', 'Paquete', 'Mensaje', 'Estado'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  // Campo trampa: los humanos no lo ven, los bots de spam lo completan.
  if (p.empresa) return respond({ ok: true });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    sheet.appendRow([
      new Date(),
      clean(p.nombre, 100),
      clean(p.email, 150),
      clean(p.whatsapp, 30),
      clean(p.rubro, 80),
      clean(p.paquete, 80),
      clean(p.mensaje, 1500),
      'Nueva',
    ]);
    return respond({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Recorta el texto y evita que un valor que empieza con = + - @ se lea como fórmula.
function clean(value, max) {
  const s = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
