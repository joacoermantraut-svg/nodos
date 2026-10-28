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

  // totales en C3:D4 de "Entradas" (zona libre: no toca la tabla de ventas ni lo que escribe la web)
  const hoja = ss.getSheetByName(HOJA);
  hoja.getRange('C3:C4').setValues([['Entradas regaladas'], ['Total entregadas']]).setFontWeight('bold');
  hoja.getRange('D3').setFormula("=SUM(Invitaciones!C2:C)");
  hoja.getRange('D4').setFormula('=B3+D3');
  hoja.getRange('D3:D4').setFontWeight('bold').setHorizontalAlignment('right');
}

// Arma "Invitaciones" en dos tablas lado a lado, cada una con su subtotal, y el total de regaladas
// de "Entradas" suma las dos. Solo reorganiza si la hoja todavía no tiene invitados anotados.
function armarInvitacionesEnDosTablas() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Invitaciones');
  const anotados = inv.getRange('A2:E' + inv.getMaxRows()).getValues().flat().filter(v => v !== '').length;
  if (anotados) throw new Error('La hoja Invitaciones ya tiene datos: no la reorganizo para no perder nada.');

  inv.getRange(1, 1, inv.getMaxRows(), inv.getMaxColumns()).clear().clearDataValidations();
  inv.setFrozenRows(0);
  const tablas = [
    { col: 1, titulo: 'INFLUENCERS / INVITADOS ESPECIALES' },
    { col: 6, titulo: 'CLIENTES' },
  ];
  tablas.forEach(({ col, titulo }) => {
    const L = c => String.fromCharCode(64 + col + c);           // letra de columna relativa a la tabla
    inv.getRange(1, col).setValue(titulo).setFontWeight('bold').setFontSize(12);
    inv.getRange(2, col, 1, 2).setValues([['Subtotal', '']]).setFontWeight('bold');
    inv.getRange(2, col + 2).setFormula('=SUM(' + L(2) + '4:' + L(2) + ')').setFontWeight('bold');
    inv.getRange(3, col, 1, 4).setValues([['Nombre', 'Teléfono', 'Cantidad', 'Notas']])
      .setFontWeight('bold').setBackground('#f1f1f1');
    inv.getRange(4, col + 2, 497, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build());
    inv.setColumnWidth(col, 220); inv.setColumnWidth(col + 1, 130);
    inv.setColumnWidth(col + 2, 80); inv.setColumnWidth(col + 3, 200);
  });
  inv.setColumnWidth(5, 30);                                     // separación entre las dos tablas
  inv.setFrozenRows(3);

  ss.getSheetByName(HOJA).getRange('D3').setFormula('=SUM(Invitaciones!C4:C)+SUM(Invitaciones!H4:H)');
}

// Invitaciones con casilla "Entregada": cada tabla muestra cuántas anotaste (estimado) y cuántas ya
// entregaste; "Regaladas entregadas" de la hoja Entradas suma solo las tildadas. Conserva lo ya anotado.
// Las fórmulas van sin separadores (SUMPRODUCT con *) porque la planilla está en configuración regional en español.
function invitacionesConEntregada() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Invitaciones');
  const leer = rango => inv.getRange(rango).getValues()
    .filter(r => r.some(v => v !== '' && v !== false))
    .map(r => [r[0], r[1], r[2], false, r[3]]);
  // lo anotado con el formato anterior (Nombre | Teléfono | Cantidad | Notas, desde la fila 4)
  const previos = [leer('A4:D' + inv.getMaxRows()), leer('F4:I' + inv.getMaxRows())];

  inv.setFrozenRows(0);
  inv.getRange(1, 1, inv.getMaxRows(), inv.getMaxColumns()).clear().clearDataValidations();
  const tablas = [
    { col: 1, titulo: 'INFLUENCERS / INVITADOS ESPECIALES' },
    { col: 7, titulo: 'CLIENTES' },
  ];
  tablas.forEach(({ col, titulo }, i) => {
    const L = c => String.fromCharCode(64 + col + c);            // letra de columna relativa a la tabla
    const cant = L(2), ent = L(3);
    inv.getRange(1, col).setValue(titulo).setFontWeight('bold').setFontSize(12);
    inv.getRange(2, col).setValue('Anotadas (estimado)').setFontWeight('bold');
    inv.getRange(2, col + 2).setFormula('=SUM(' + cant + '5:' + cant + ')').setFontWeight('bold');
    inv.getRange(3, col).setValue('Entregadas').setFontWeight('bold');
    inv.getRange(3, col + 2).setFormula('=SUMPRODUCT(' + ent + '5:' + ent + '*' + cant + '5:' + cant + ')').setFontWeight('bold');
    inv.getRange(4, col, 1, 5).setValues([['Nombre', 'Teléfono', 'Cantidad', 'Entregada', 'Notas']])
      .setFontWeight('bold').setBackground('#f1f1f1');
    inv.getRange(5, col + 2, 496, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(1).setAllowInvalid(false).build());
    inv.getRange(5, col + 3, 496, 1).insertCheckboxes();
    if (previos[i].length) inv.getRange(5, col, previos[i].length, 5).setValues(previos[i]);
    inv.setColumnWidth(col, 220); inv.setColumnWidth(col + 1, 130);
    inv.setColumnWidth(col + 2, 80); inv.setColumnWidth(col + 3, 85); inv.setColumnWidth(col + 4, 200);
  });
  inv.setColumnWidth(6, 30);                                      // separación entre las dos tablas
  inv.setFrozenRows(4);

  const hoja = ss.getSheetByName(HOJA);
  hoja.getRange('C3').setValue('Regaladas entregadas');
  hoja.getRange('D3').setFormula('=SUMPRODUCT(Invitaciones!D5:D*Invitaciones!C5:C)+SUMPRODUCT(Invitaciones!J5:J*Invitaciones!I5:I)');
}

// Resumen arriba de Invitaciones: cupo para regalar, cuántas hay anotadas entre las dos tablas,
// cuántas quedan y cuántas ya se entregaron. Agrega 6 filas arriba; las fórmulas existentes se corren solas.
function totalInvitaciones() {
  const inv = SpreadsheetApp.getActive().getSheetByName('Invitaciones');
  if (inv.getRange('A1').getValue() === 'TOTAL INVITACIONES') return;   // ya está armado
  inv.setFrozenRows(0);
  inv.insertRowsBefore(1, 6);
  // las tablas ahora empiezan en la fila 7: anotadas en la fila 8, entregadas en la 9
  inv.getRange('A1').setValue('TOTAL INVITACIONES').setFontWeight('bold').setFontSize(12);
  inv.getRange('A2:A5').setValues([['Cupo para regalar'], ['Anotadas (estimado)'], ['Quedan para anotar'], ['Entregadas']])
    .setFontWeight('bold');
  inv.getRange('C2').setValue(80).setBackground('#fff6cc');
  inv.getRange('C3').setFormula('=C8+I8');
  inv.getRange('C4').setFormula('=C2-C3');
  inv.getRange('C5').setFormula('=C9+I9');
  inv.getRange('C2:C5').setFontWeight('bold');
  inv.getRange('A1:E5').setBorder(true, true, true, true, false, false, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
  const reglas = inv.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
    .setFontColor('#d93025').setRanges([inv.getRange('C4')]).build());
  inv.setConditionalFormatRules(reglas);
  inv.setFrozenRows(10);
}
