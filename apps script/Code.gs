function mostrarFormularioPersona() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioPersona')
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registro de Persona');
}

function guardarPersona(data) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  const fila = hoja.getLastRow() + 1;

  hoja.getRange(fila, 1, 1, 3).setValues([
    [data.codigo, data.nombre, data.edad]
  ]);

  const rango = hoja.getRange(fila, 1, 1, 3);
  rango.setHorizontalAlignment("center");
  rango.setBorder(true, true, true, true, true, true);
}