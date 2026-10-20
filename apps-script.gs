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
