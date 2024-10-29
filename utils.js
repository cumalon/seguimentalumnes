function addViewerSilently(fileId, emailAddress) {
  var permission = {
    'type': 'user',
    'role': 'reader',
    'emailAddress': emailAddress
  };
  
  Drive.Permissions.create(permission, fileId, {
    'sendNotificationEmail': false
  });
}

function extractFileId(url) {
  return url.match(/[-\w]{25,}/); // null if no match is found
}

function getFileFromUrl(url) {
  try {
    var fileId = extractFileId(url);
    var file = DriveApp.getFileById(fileId);
    if(file && !file.isTrashed()) {
      return file;
    }
    return null;
  } catch (e) {
    return null;
  }
}

