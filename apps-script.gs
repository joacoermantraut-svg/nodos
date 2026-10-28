// NODOS · registro de entradas
// Pegar en la planilla: Extensiones → Apps Script. Implementar → Nueva implementación →
// Tipo "Aplicación web", Ejecutar como "Yo", Acceso "Cualquier persona". Copiar la URL /exec.

const HOJA = 'Entradas';
const PRIMERA_FILA = 10;   // fila 9 tiene los encabezados

function doPost(e) {
  const p = (e && e.parameter) || {};
  const nombre = String(p.nombre || '').trim().slice(0, 80);
  const telefono = String(p.telefono || '').trim().slice(0, 30);
  const cantidad = Math.min(10, Math.max(1, parseInt(p.cantidad, 10) || 1));
  const tipo = p.tipo === 'Efectivo' ? 'Efectivo' : 'Transf';   // valores que suman las fórmulas de B4 y B5
  if (!nombre || !telefono) return ContentService.createTextOutput('falta');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA);
    // primera fila con la columna Nombre vacía (la columna F ya trae el valor de la entrada)
    const nombres = hoja.getRange(PRIMERA_FILA, 1, hoja.getMaxRows() - PRIMERA_FILA + 1, 1).getValues();
    let fila = nombres.findIndex(r => r[0] === '');
    fila = fila === -1 ? hoja.getLastRow() + 1 : PRIMERA_FILA + fila;
    // un apóstrofo adelante para que Sheets no interprete nombres o teléfonos como fórmulas
    hoja.getRange(fila, 1, 1, 5).setValues([["'" + nombre, "'" + telefono, cantidad, 'No', tipo]]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Correr UNA vez desde el editor (botón Ejecutar): crea la hoja "Invitaciones" para anotar las
// entradas regaladas y suma dos totales arriba de la hoja "Entradas", al lado de los que ya existen.
function crearInvitaciones() {
  const ss = SpreadsheetApp.getActive();
  let inv = ss.getSheetByName('Invitaciones');
  if (!inv) {
    inv = ss.insertSheet('Invitaciones');
    inv.getRange('A1:E1').setValues([['Nombre', 'Teléfono', 'Cantidad', 'Para', 'Notas']])
      .setFontWeight('bold').setBackground('#f1f1f1');
    inv.setFrozenRows(1);
    inv.getRange('C2:C500').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build());
    inv.getRange('D2:D500').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Cliente', 'Amigo', 'Otro'], true).build());
    inv.setColumnWidth(1, 220); inv.setColumnWidth(2, 140); inv.setColumnWidth(3, 90);
    inv.setColumnWidth(4, 110); inv.setColumnWidth(5, 260);
  }

  // totales en D3:E4 de "Entradas" (zona libre: no toca la tabla de ventas ni lo que escribe la web)
  const hoja = ss.getSheetByName(HOJA);
  hoja.getRange('D3:D4').setValues([['Entradas regaladas'], ['Total entregadas (vendidas + regaladas)']]).setFontWeight('bold');
  hoja.getRange('E3').setFormula("=SUM(Invitaciones!C2:C)");
  hoja.getRange('E4').setFormula('=B3+E3');
}
