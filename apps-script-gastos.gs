/**
 * Gastos del panel "Día" → tu Google Sheet.
 *
 * CÓMO ACTIVARLO (2 minutos):
 * 1. En tu hoja de Google: Extensiones → Apps Script.
 * 2. Pega TODO este archivo al final de tu código existente (no borres nada).
 * 3. En tu función doPost(e) ya existente, justo después de leer los datos
 *    (donde haces JSON.parse(e.postData.contents)), agrega estas 2 líneas:
 *
 *      var data = JSON.parse(e.postData.contents);
 *      if (data.kind === 'gasto') return handleGasto_(data);
 *
 * 4. Implementar → Administrar implementaciones → editar (lápiz) → Versión nueva → Implementar.
 *    (La URL /exec es la misma; no tienes que volver a pegarla en el panel.)
 *
 * Escribe cada gasto en la pestaña "Gastos". La pestaña "Plan" ya lee de ahí:
 * las columnas "Gastado (tu app)", "Te queda" y "Estado" se actualizan solas.
 */
function handleGasto_(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Gastos');
  if (!sh) {
    sh = ss.insertSheet('Gastos');
    sh.appendRow(['ID', 'Fecha', 'Categoría', 'Monto', 'Nota', 'Quincena', 'Tope', 'Restante al anotar', 'Registrado']);
    sh.setFrozenRows(1);
    sh.getRange('A1:I1').setFontWeight('bold');
  }
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(function (r) { return String(r[0]); });
  var idx = ids.indexOf(String(d.id));
  if (d.op === 'del') {
    if (idx >= 0) sh.deleteRow(idx + 2);
  } else if (idx < 0) {
    sh.appendRow([d.id, d.fecha, d.categoria, d.monto, d.nota, d.quincena, d.tope, d.restante, new Date(d.ts)]);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
