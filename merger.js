function ownMergeSetup(pestanya) {
  var full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pestanya);
  if (full) {
    var dades = full.getDataRange().getValues(); // Obtenir totes les dades del full
    var headers = dades[0];
    return buscarOAfegirColumnaReport(full,headers);
  } else {
    Logger.log("La pestanya no existeix: " + pestanya);
  }
}

function resetTemplateId(id) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('TEMPLATE_ID','undefined');
  scriptProperties.setProperty('REPOSITORY_ID','undefined');
}

function setTemplateId(id,prop) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(prop,id);
  return prop;
}

function getTemplateId(prop) {
  return PropertiesService.getScriptProperties().getProperty(prop);
}

function getTemplateName(prop) {
  var id = getTemplateId(prop);
  if(id!='undefined') {
    var file = DriveApp.getFileById(id);
    if(file) {
      var fileData = {name: file.getName(), url: file.getUrl(), prop: prop};
      return JSON.stringify(fileData);
    }
  }
  return 'undefined';
}

// Funció per identificar TAGS i carregar-los
function carregarTags() {

  var templateId = getTemplateId('TEMPLATE_ID');
  
  // Obtenir el document
  var document = DocumentApp.openById(templateId);
  var cos = document.getBody();
  
  // Llegir el contingut del document
  var contingut = cos.getText();
  
  // Patró per identificar qualsevol cosa entre << i >>
  var patron = /<<(.*?)>>/g; // Captura qualsevol cosa entre << i >>
  var resultats;
  var tagsTrobats = [];

  // Buscar els TAGS
  while ((resultats = patron.exec(contingut)) !== null) {
    // Afegir el contingut trobat (sense els signes << i >>) a l'array
    tagsTrobats.push(resultats[1].trim()); // Trimar per eliminar espais innecessaris
  }

  return tagsTrobats; // Retornar els TAGS trobats
}

function mergeRow(keyValue,keyHeader,mapatge,pestanyaName) {
  
  // Obtenir la pestanya pel seu nom
  var full = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pestanyaName);
  if (!full) {
    Logger.log("No s'ha trobat la pestanya amb nom: " + pestanyaName);
    return;
  }

  var rowData = obtenirDadesBy(full,keyValue,keyHeader);

  // ... fins aquí obtenció de dades del keyValue indicat. Ara es podrà fer el que volguem amb aquestes dades, per exemple enviar un email, fer funcions de merge com l'Autocrat ...


  // TODO: a partir d'aquí nova funció per al cas de merge Autocrat:
  checkRowFileUrl(rowData);
  replaceTags(rowData,mapatge);

}

// Buscar la columna "MERGE_DOC_URL", i si no existeix, afegir-la
function buscarOAfegirColumnaReport(full, headerRow) {
  var reportColIndex = headerRow.indexOf("MERGE_DOC_URL");

  // Si no es troba la columna "MERGE_DOC_URL", afegir-la
  if (reportColIndex === -1) {
    full.insertColumnAfter(full.getLastColumn()); // Inserir la columna
    full.insertColumnAfter(full.getLastColumn()); // Inserir la columna
    var newReportColIndex = full.getLastColumn()+2;
    Logger.log(" ... inserint Columna 'MERGE_DOC_URL' a la col num: " + newReportColIndex);
    const headerRowIndex = parseInt(PropertiesService.getScriptProperties().getProperty("HEADER_ROW_INDEX"));
    full.getRange(headerRowIndex, newReportColIndex).setValue("MERGE_DOC_URL"); // Estableix el nom de la columna

    reportColIndex = newReportColIndex - 1; // Actualitza l'índex de la columna "MERGE_DOC_URL"
    Logger.log("Columna 'MERGE_DOC_URL' afegida a la columna " + newReportColIndex);
  }

  return reportColIndex;
}

function checkRowFileUrl(rowData) {

  const plantillaDocId = getTemplateId('TEMPLATE_ID');
  const carpetaId = getTemplateId('REPOSITORY_ID');

  var reportColIndex = rowData.headers.indexOf("MERGE_DOC_URL");
  Logger.log(" checkRowFIleUrl ... reportColIndex: "+reportColIndex)
  if(reportColIndex!=-1) {
    var url = rowData.dades[reportColIndex];
    Logger.log("   col iundex found, url: "+url)
    if(!esUrlValida(url) || !esFitxerDe(url,carpetaId)) {
      Logger.log("   url NO vàlida o canvi de repository ...");
      // Crear una còpia del document i escriure la URL a "MERGE_DOC_URL"
      var novaUrl = crearICopiarInforme(rowData,plantillaDocId,carpetaId);
      rowData.sheet.getRange(rowData.rowIndex + 1, reportColIndex + 1).setValue(novaUrl.replace(/\/edit.*/, "/preview?rm=minimal"));
      rowData.dades[reportColIndex] = novaUrl; // desem a dades perque si no existia la url tampoc la tenim a l'array
    }
    else {
      Logger.log("   url SÍ vàlida");
      replaceDocumentContent(url, plantillaDocId);
    }
  }
  else {
    Logger.log("  --- ups, MERGE_DOC_URL column not found"); // TODO: comunicar error
  }

}

function replaceDocumentContent(targetDocUrl, templateDocId) {
  // Obtenir el document objectiu mitjançant la URL
  var targetDoc = DocumentApp.openByUrl(targetDocUrl);
  
  // Obtenir el document template mitjançant l'ID
  var templateDoc = DocumentApp.openById(templateDocId);

  // Obtenir el cos de cada document
  var targetBody = targetDoc.getBody();
  var templateBody = templateDoc.getBody();

  // Esborrar tot el contingut del document objectiu
  targetBody.clear();

  // Copiar tots els elements del document template al document objectiu mantenint el format
  var totalElements = templateBody.getNumChildren();
  
  for (var i = 0; i < totalElements; i++) {
    var element = templateBody.getChild(i).copy(); // Copiar cada element (manté el format)
    
    // Comprovar el tipus d'element i utilitzar el mètode adequat
    if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(element);
    } else if (element.getType() == DocumentApp.ElementType.HEADER_SECTION) {
      targetBody.appendTable(element.asHeaderSection());
    } else if (element.getType() == DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(element);
    } else if (element.getType() == DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(element);
    } else {
      // Afegir altres tipus d'elements si cal (imatges, dibuixos, etc.)
      targetBody.appendElement(element);
    }
  }

  // Guardar els canvis
  targetDoc.saveAndClose();
}

// Comprovar si una URL és vàlida i apunta a un document existent que no està a la paperera
function esUrlValida(url) {
  try {
    var fileId = url.match(/[-\w]{25,}/); // Extreure l'ID del fitxer de la URL
    var file = DriveApp.getFileById(fileId);
    return file && !file.isTrashed(); // Comprovar si el fitxer existeix i no està a la paperera
  } catch (e) {
    return false;
  }
}

function esFitxerDe(url,carpetaId) {
  try {
    var fileId = url.match(/[-\w]{25,}/); // Extreure l'ID del fitxer de la URL
    var file = DriveApp.getFileById(fileId);
    if(file) {
      var parents = file.getParents();    
      while(parents.hasNext()) {
        var folder = parents.next();
        if(folder.getId() === carpetaId) {
          return true;
        }
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

// Crear una còpia de l'informe i retornar la URL
function crearICopiarInforme(rowData,plantillaDocId,carpetaId) {

  var keyValue = rowData.keyValue;
  var nomInforme = "Informe_" + keyValue;

  var plantillaDoc = DriveApp.getFileById(plantillaDocId);
  var carpeta = DriveApp.getFolderById(carpetaId);

  var novaCopia = plantillaDoc.makeCopy(nomInforme, carpeta);
  var novaUrl = novaCopia.getUrl();

  var emailIndex = rowData.colIndex["Email"];
  if(emailIndex>0) {
    var email = rowData.dades[emailIndex];
    novaCopia.addViewer(email);
  }

  Logger.log("Creat document: " + nomInforme + " amb URL: " + novaUrl);
  return novaUrl;
}

function replaceTags(rowData, mapatge) {

  const reportUrlIndex = rowData.colIndex["MERGE_DOC_URL"];
  const reportUrl = rowData.dades[reportUrlIndex];

  // Obtenir el document de text de Google Drive
  const document = DocumentApp.openByUrl(reportUrl);
  const body = document.getBody();

  // Substituir els tags en el document
  Logger.log("processant fitxer: "+document.getName())
  mapatge.forEach(function(pair) {
    var header = pair.capcalera;
    var tag = pair.tag;
    Logger.log( "`-> header: "+ header + " tag: "+ tag)
    const indexHeader = rowData.colIndex[header];
    if (indexHeader !== undefined) {
      Logger.log( "`-> header: "+ header + " tag: "+ tag + " SUBSTITUCIó!")
      var valor = rowData.dades[indexHeader];
      if(Object.prototype.toString.call(valor) === '[object Date]') valor = valor.toLocaleDateString();
      const regex = new RegExp(`<<${tag}>>`, 'g'); // Crear una expressió regular per trobar el tag
      body.replaceText(regex.source, valor); // Substituir el tag pel valor
      Logger.log( "`-> valor substitució: "+ valor);
    }
  });

  // Desar el document
  document.saveAndClose();
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Escull plantilla');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
