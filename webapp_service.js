function amIAdmin_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('webapp');
  var isAdmin = false;
  if(s) {
    const email = getUserEmail();
    isAdmin = s.getRange("H1:H"+ s.getLastRow()).getValues().some(function(row) {
      return row[0] == email;
    });
  }
  Logger.log("isAdmin: "+isAdmin);
  return isAdmin;
}

function getIndexPageName_() {
  return amIAdmin_()? "adminIndex" : "webapp"
}

function doGet(e) {
  const filename = getIndexPageName_();
  return HtmlService.createTemplateFromFile(filename)
                    .evaluate()
                    .setTitle("Informes d'avaluació")
                    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Opcional, permet que la pàgina es mostri dins d'un iframe, si cal
}

// funció per incrustar codi de fitxers externs
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function processUrls() {
  var userEmail = getUserEmail();
  return processUrlsByEmail(userEmail);
}

function processUrlsByEmail(userEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('webapp');
  var data = sheet.getDataRange().getValues();
  var results = [];

Logger.log(data);
Logger.log(data[1]);
Logger.log("forEach ...");
  data.slice(1).forEach(function(row) {
    var reportName = row[0];
    var url = row[2];
    var sheetName = row[3];
    var headerRowIndex = row[4]-1;
    Logger.log("row ****");
    Logger.log(row);
    Logger.log("reportName: "+reportName+" sheetName: "+sheetName);
    if (url && url.match(/docs.google.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)) {
      try {
        var spreadsheet = SpreadsheetApp.openByUrl(url);
        var mergeSheet = spreadsheet.getSheetByName(sheetName);
        var mergeData = mergeSheet.getDataRange().getValues();
        
        // Cerca l'email de l'usuari a la columna "EMAIL"
        var emailIndex = mergeData[headerRowIndex].indexOf("EMAIL");
        var reportIndex = mergeData[headerRowIndex].indexOf("MERGE_DOC_URL");
        Logger.log("mergeSheet: "+mergeSheet.getName()+" emailIndex: "+emailIndex + " reportIndex: "+reportIndex)
        for (var i = headerRowIndex; i < mergeData.length; i++) {
        Logger.log("llegit: "+mergeData[i][emailIndex] + " userEmail: "+userEmail);
        var rowEmail = mergeData[i][emailIndex];
          if (rowEmail === userEmail) {
            var reportUrl = mergeData[i][reportIndex];
            if (reportUrl && reportUrl.match(/docs.google.com\/document\/d\/([a-zA-Z0-9-_]+)/)) {
              results.push({ reportName: reportName, spreadsheetName: spreadsheet.getName(), reportUrl: reportUrl });
            }
            break;
          }
        }
      } catch (error) {
        Logger.log("Error processing URL: " + url + " - " + error);
      }
    }
  });

  return results;
}

function getUserEmail() {
  return Session.getActiveUser().getEmail(); // Obtenir l'email de l'usuari actual
}

function getUserName() {
    var email = Session.getActiveUser().getEmail();
    var name = email.split('@')[0].replace(/\./g, ' '); // Opcional: substituir punts per espais
    return name;
}

function getNamesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("webapp");
  const range = sheet.getRange("J2:L" + sheet.getLastRow()); // Obtenim les dades des de J2 fins a L (inclou Cognoms, Nom i Email)
  const values = range.getValues(); // Obtenim les dades del rang
  return values.map(function(row) {
    return {
      name: row[0],  // Cognoms, Nom (columna J)
      email: row[2]  // Email (columna L)
    };
  });
}
