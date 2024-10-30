function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('webapp')
    .setTitle('URL Reports') // Estableix el títol del document
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Opcional, permet que la pàgina es mostri dins d'un iframe, si cal
}

function processUrls() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('webapp');
  var data = sheet.getDataRange().getValues();
  var userEmail = getUserEmail();
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
        
        // Cerca l'email de l'usuari a la columna "Email"
        var emailIndex = mergeData[headerRowIndex].indexOf("Email");
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


