function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alumnes')
      .addItem("Gestiona informes d'avaluació", 'showSidebar')
      .addItem("Enviament massiu", 'showEmailSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar'); // Càrrega el fitxer HTML
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  html.append(`<script>buildSheetSelect(${JSON.stringify({sheets: sheets, selected: activeSheet.getName()})});</script>`);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showEmailSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('emailSidebar');
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  html.append(`<script>initEmailSidebar(${JSON.stringify({sheets: sheets, selected: activeSheet.getName()})});</script>`);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Funció per carregar les capçaleres de la pestanya seleccionada
function carregarCapcaleres(pestanya,headerRowIndex) {
  var full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pestanya);
  if (full) {
    var dades = full.getDataRange().getValues(); // Obtenir totes les dades del full
    var headers = dades[headerRowIndex-1];
    Logger.log("headerRowIndex: "+headerRowIndex);
    PropertiesService.getScriptProperties().setProperty("HEADER_ROW_INDEX",headerRowIndex);
    var capcalera = headers.filter(function(camp) {
      Logger.log(camp + "és un: "+typeof camp);    
      return (typeof camp) === "string" && camp !== "";
    });
    Logger.log("capcaleres valides: "+capcalera);
    return capcalera; // Retornar les capçaleres vàlides
  } else {
    Logger.log("La pestanya no existeix: " + pestanya);
  }
}

function showSheet(sheetName) {
  var full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (full) {
    SpreadsheetApp.setActiveSheet(full);
  }
}

// Helper function to load row data based on key header
function carregarDadesPestanyaHelper(sheet, keyHeader, headerRowIndex) {
  var dades = sheet.getDataRange().getValues();
  var headers = dades[headerRowIndex - 1];
  var headerColIndex = headers.indexOf(keyHeader);

  if (headerColIndex === -1) {
    throw new Error('No s\'ha trobat la columna "' + keyHeader + '".');
  }

  // Extreu les dades de la columna keyHeader
  var keyValues = [];
  for (var i = headerRowIndex; i < dades.length; i++) {
    var fila = {};
    fila[keyHeader] = dades[i][headerColIndex];
    keyValues.push(fila);
  }

  return keyValues;
}

// Funció per carregar les dades de la pestanya seleccionada
function carregarDadesPestanya(keyHeader) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var full = ss.getActiveSheet();
  const headerRowIndex = parseInt(PropertiesService.getScriptProperties().getProperty("HEADER_ROW_INDEX"));
  
  return carregarDadesPestanyaHelper(full, keyHeader, headerRowIndex);
}

// Funció per carregar les dades de la pestanya especificada per a enviament massiu
function carregarDadesPestanyaEmail(sheetName, keyHeader, headerRowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var full = ss.getSheetByName(sheetName);
  
  if (!full) {
    throw new Error('No s\'ha trobat la pestanya "' + sheetName + '".');
  }
  
  return carregarDadesPestanyaHelper(full, keyHeader, headerRowIndex);
}

function obtenirDadesBy(sheet,keyValue,searchHeader) {
  const dades = sheet.getDataRange().getValues();
  const headerRowIndex = parseInt(PropertiesService.getScriptProperties().getProperty("HEADER_ROW_INDEX"));
  const headers = dades[headerRowIndex-1];
  const indexKeyValue = headers.indexOf(searchHeader);
  Logger.log("indexKeyValue: "+indexKeyValue+" keyValue: "+keyValue)
  const foundIndex = dades.slice(headerRowIndex).findIndex(row => row[indexKeyValue] === keyValue);
  const rowIndex = foundIndex+headerRowIndex;
  if(foundIndex<0) throw "Error de cerca! ... "+keyValue+" rowIndex: "+rowIndex+" headerRowIndex: "+headerRowIndex;
  var header2colIndex = {};
  Logger.log("obtenirDadesBy ... "+keyValue+" rowIndex: "+rowIndex+" headerRowIndex: "+headerRowIndex);
  headers.forEach(function(value,index) {
    Logger.log( " value: "+value + " index: "+index)
    header2colIndex[value] = index;
  });
  Logger.log(dades[rowIndex]);
  return {sheet: sheet, keyValue: keyValue, rowIndex: rowIndex, dades: dades[rowIndex], headers: headers, colIndex: header2colIndex};
}
