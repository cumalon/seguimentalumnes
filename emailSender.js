/**
 * Mass Email Sender Module
 * Handles scheduling and sending of mass emails with attachments
 * 
 * IMPORTANT: This script requires the Drive API to be enabled
 * To enable: In the Apps Script editor, go to Resources > Advanced Google Services
 * and enable "Drive API". Also enable it in the Google Cloud Console.
 */

// Function to check email and attachment quotas
function checkEmailQuotas(config) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    var data = sheet.getDataRange().getValues();
    var headers = data[config.headerRowIndex - 1];
    var emailColumnIndex = headers.indexOf('EMAIL');
    
    if (emailColumnIndex === -1) {
      throw new Error('No s\'ha trobat la columna EMAIL');
    }
    
    // Count emails to send
    var emailCount = 0;
    for (var i = config.headerRowIndex; i < data.length; i++) {
      if (data[i][emailColumnIndex]) {
        emailCount++;
      }
    }
    
    // Count attachments per email
    var attachmentCount = config.attachmentColumns ? config.attachmentColumns.length : 0;
    var totalAttachments = emailCount * attachmentCount;
    
    // Get remaining email quota
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    
    // Google Workspace quotas (approximate - can vary by account type)
    // Standard: 100 emails/day (consumer), 2000 emails/day (Google Workspace)
    // Attachments: typically 25 MB total size per email
    var quotaInfo = {
      emailsToSend: emailCount,
      emailQuotaRemaining: emailQuotaRemaining,
      attachmentsPerEmail: attachmentCount,
      totalAttachments: totalAttachments,
      canProceed: emailCount <= emailQuotaRemaining,
      quotaExceeded: emailCount > emailQuotaRemaining,
      quotaPercentage: emailQuotaRemaining > 0 ? Math.round((emailCount / emailQuotaRemaining) * 100) : 100
    };
    
    return quotaInfo;
  } catch (error) {
    Logger.log('Error checking quotas: ' + error);
    throw error;
  }
}

// Function to schedule mass emails
function scheduleMassEmails(config) {
  try {
    // Validate configuration
    if (!config.sheetName || !config.subject || !config.body) {
      throw new Error('Configuració incompleta');
    }

    // Check quotas first
    var quotaInfo = checkEmailQuotas(config);
    
    // If quota exceeded, return error
    if (quotaInfo.quotaExceeded) {
      return {
        success: false,
        quotaInfo: quotaInfo,
        error: 'Quota diària excedida. No es poden enviar els correus.'
      };
    }

    // Check if immediate send (delay = 0)
    var isImmediate = config.scheduleInfo.type === 'delay' && config.scheduleInfo.delayMinutes === 0;
    
    if (isImmediate) {
      // Send immediately without creating trigger
      var result = sendMassEmailsNow(config);
      return { 
        success: true, 
        count: result.successCount, 
        executionTime: 'Immediat',
        errors: result.errorCount,
        quotaInfo: quotaInfo
      };
    }

    // Store configuration in script properties for later execution
    var scriptProps = PropertiesService.getScriptProperties();
    var configKey = 'EMAIL_CONFIG_' + new Date().getTime();
    scriptProps.setProperty(configKey, JSON.stringify(config));

    // Calculate execution time
    var executionTime = new Date();
    if (config.scheduleInfo.type === 'delay') {
      var delayMs = config.scheduleInfo.delayMinutes * 60 * 1000;
      executionTime = new Date(executionTime.getTime() + delayMs);
    } else {
      executionTime = new Date(config.scheduleInfo.datetime);
    }

    // Create time-based trigger
    ScriptApp.newTrigger('executeScheduledEmails')
      .timeBased()
      .at(executionTime)
      .create();

    // Add config key to trigger for later retrieval
    scriptProps.setProperty('PENDING_EMAIL_CONFIG', configKey);

    return { 
      success: true, 
      count: quotaInfo.emailsToSend, 
      executionTime: executionTime.toString(),
      quotaInfo: quotaInfo
    };
  } catch (error) {
    Logger.log('Error in scheduleMassEmails: ' + error);
    throw error;
  }
}

// Function executed by trigger to send emails
function executeScheduledEmails() {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var configKey = scriptProps.getProperty('PENDING_EMAIL_CONFIG');
    
    if (!configKey) {
      Logger.log('No pending email configuration found');
      return;
    }

    var configJson = scriptProps.getProperty(configKey);
    if (!configJson) {
      Logger.log('Configuration not found: ' + configKey);
      return;
    }

    var config = JSON.parse(configJson);
    
    // Send emails
    sendMassEmailsNow(config);

    // Clean up
    scriptProps.deleteProperty(configKey);
    scriptProps.deleteProperty('PENDING_EMAIL_CONFIG');

    // Delete the trigger
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'executeScheduledEmails') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  } catch (error) {
    Logger.log('Error in executeScheduledEmails: ' + error);
    throw error;
  }
}

// Function to send emails immediately
function sendMassEmailsNow(config) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    var data = sheet.getDataRange().getValues();
    var headers = data[config.headerRowIndex - 1];
    
    // Find column indices for EMAIL (required), EMAIL_CC and EMAIL_BCC (optional)
    var emailColumnIndex = headers.indexOf('EMAIL');
    var emailCcColumnIndex = headers.indexOf('EMAIL_CC');
    var emailBccColumnIndex = headers.indexOf('EMAIL_BCC');
    var replyToColumnIndex = headers.indexOf('REPLY_TO');
    var senderNameColumnIndex = headers.indexOf('SENDER_NAME');
    
    if (emailColumnIndex === -1) {
      throw new Error('Columna EMAIL no trobada. Cal una columna amb el nom exacte "EMAIL"');
    }

    // Find or create log column
    var logColumnIndex = headers.indexOf('EMAIL_LOG');
    if (logColumnIndex === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      logColumnIndex = sheet.getLastColumn() - 1;
      sheet.getRange(config.headerRowIndex, logColumnIndex + 1).setValue('EMAIL_LOG');
    }

    // Build header to index map
    var headerMap = {};
    headers.forEach(function(header, index) {
      headerMap[header] = index;
    });

    // Get attachment column indices
    var attachmentIndices = [];
    if (config.attachmentColumns && config.attachmentColumns.length > 0) {
      config.attachmentColumns.forEach(function(colName) {
        var index = headers.indexOf(colName);
        if (index !== -1) {
          attachmentIndices.push(index);
        }
      });
    }

    // Process each row
    var successCount = 0;
    var errorCount = 0;

    for (var i = config.headerRowIndex; i < data.length; i++) {
      var row = data[i];
      var emailValue = row[emailColumnIndex];
      
      if (!emailValue || !String(emailValue).trim()) {
        sheet.getRange(i + 1, logColumnIndex + 1).setValue('Error: Email buit');
        errorCount++;
        continue;
      }

      try {
        // Parse email addresses (can be comma or semicolon separated)
        var emails = parseEmailAddresses(String(emailValue));
        if (emails.length === 0) {
          sheet.getRange(i + 1, logColumnIndex + 1).setValue('Error: Email invàlid');
          errorCount++;
          continue;
        }
        
        // Parse CC emails if present
        var ccEmails = [];
        if (emailCcColumnIndex !== -1 && row[emailCcColumnIndex]) {
          ccEmails = parseEmailAddresses(String(row[emailCcColumnIndex]));
        }
        
        // Parse BCC emails if present
        var bccEmails = [];
        if (emailBccColumnIndex !== -1 && row[emailBccColumnIndex]) {
          bccEmails = parseEmailAddresses(String(row[emailBccColumnIndex]));
        }

        // Replace tags in subject and body
        var personalizedSubject = replaceMailTags(config.subject, row, config.tagMapping, headerMap);
        var personalizedBody = replaceMailTags(config.body, row, config.tagMapping, headerMap);

        // Get attachments
        var attachments = [];
        attachmentIndices.forEach(function(index) {
          var url = row[index];
          if (url && typeof url === 'string' && url.trim()) {
            try {
              var attachment = getAttachmentFromUrl(url);
              if (attachment) {
                attachments.push(attachment);
              }
            } catch (attachError) {
              Logger.log('Error getting attachment from ' + url + ': ' + attachError);
            }
          }
        });

        // Build email options
        var emailOptions = {
          name: 'Enviament massiu',
          attachments: attachments
        };
        
        // Set sender name from SENDER_NAME column if present and not empty
        if (senderNameColumnIndex !== -1 && row[senderNameColumnIndex]) {
          var senderName = String(row[senderNameColumnIndex]).trim();
          if (senderName) {
            emailOptions.name = senderName;
          }
        }
        
        // Set reply-to from REPLY_TO column if present and not empty
        if (replyToColumnIndex !== -1 && row[replyToColumnIndex]) {
          var replyTo = String(row[replyToColumnIndex]).trim();
          if (replyTo && isValidEmail(replyTo)) {
            emailOptions.replyTo = replyTo;
          }
        }
        
        if (ccEmails.length > 0) {
          emailOptions.cc = ccEmails.join(',');
        }
        
        if (bccEmails.length > 0) {
          emailOptions.bcc = bccEmails.join(',');
        }

        // Send email (main recipient is first email, rest are added to the list)
        var mainEmail = emails.join(',');
        MailApp.sendEmail(mainEmail, personalizedSubject, personalizedBody, emailOptions);
        
        // Log success
        var timestamp = new Date().toLocaleString('ca-ES');
        var logMessage = 'Enviat: ' + timestamp;
        if (ccEmails.length > 0) logMessage += ' (CC: ' + ccEmails.length + ')';
        if (bccEmails.length > 0) logMessage += ' (BCC: ' + bccEmails.length + ')';
        sheet.getRange(i + 1, logColumnIndex + 1).setValue(logMessage);
        successCount++;

      } catch (emailError) {
        Logger.log('Error sending email to ' + emailValue + ': ' + emailError);
        sheet.getRange(i + 1, logColumnIndex + 1).setValue('Error: ' + emailError.message);
        errorCount++;
      }

      // Add a small delay to avoid quota issues
      Utilities.sleep(100);
    }

    Logger.log('Mass email send completed. Success: ' + successCount + ', Errors: ' + errorCount);
    return { success: true, successCount: successCount, errorCount: errorCount };

  } catch (error) {
    Logger.log('Error in sendMassEmailsNow: ' + error);
    throw error;
  }
}

// Function to replace tags in text with row data
function replaceMailTags(text, rowData, tagMapping, headerMap) {
  var result = text;
  
  tagMapping.forEach(function(mapping) {
    var tag = mapping.tag;
    var header = mapping.header;
    var columnIndex = headerMap[header];
    
    if (columnIndex !== undefined) {
      var value = rowData[columnIndex];
      
      // Format dates
      if (Object.prototype.toString.call(value) === '[object Date]') {
        value = value.toLocaleDateString('ca-ES');
      }
      
      // Convert to string
      value = value !== null && value !== undefined ? String(value) : '';
      
      // Replace tag
      var regex = new RegExp('<<' + tag + '>>', 'g');
      result = result.replace(regex, value);
    }
  });
  
  return result;
}

// Function to get attachment from URL (converts Google Docs and MS Word to PDF)
function getAttachmentFromUrl(url) {
  try {
    var fileId = null;
    
    // Check if it's a Google Doc URL
    var docMatch = url.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/);
    if (docMatch) {
      fileId = docMatch[1];
    }
    
    // Check if it's a Google Drive file URL (various formats)
    var driveMatch = url.match(/drive\.google\.com\/.*[?&]id=([a-zA-Z0-9-_]+)/);
    if (driveMatch) {
      fileId = driveMatch[1];
    } else {
      // Try to match /file/d/{id}/ format
      var fileMatch = url.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
      if (fileMatch) {
        fileId = fileMatch[1];
      } else {
        // Try to match open?id={id} format
        var openMatch = url.match(/open\?id=([a-zA-Z0-9-_]+)/);
        if (openMatch) {
          fileId = openMatch[1];
        }
      }
    }
    
    // If we found a file ID, process the file
    if (fileId) {
      var file = DriveApp.getFileById(fileId);
      var mimeType = file.getMimeType();
      
      Logger.log('Processing file: ' + file.getName() + ' with MIME type: ' + mimeType);
      
      // Convert Google Docs to PDF
      if (mimeType === MimeType.GOOGLE_DOCS) {
        return convertGoogleDocToPdf(fileId);
      }
      
      // Convert MS Word documents to PDF
      if (mimeType === MimeType.MICROSOFT_WORD || 
          mimeType === MimeType.MICROSOFT_WORD_LEGACY ||
          mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
          mimeType === 'application/msword') {
        return convertWordToPdf(file);
      }
      
      // Convert MS Excel to PDF
      if (mimeType === MimeType.MICROSOFT_EXCEL || 
          mimeType === MimeType.MICROSOFT_EXCEL_LEGACY ||
          mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          mimeType === 'application/vnd.ms-excel') {
        return convertExcelToPdf(file);
      }
      
      // For PDFs and other files, return as-is
      return file.getBlob();
    }
    
    return null;
  } catch (error) {
    Logger.log('Error getting attachment from URL ' + url + ': ' + error);
    return null;
  }
}

// Function to convert Google Doc to PDF
function convertGoogleDocToPdf(docId) {
  try {
    var doc = DocumentApp.openById(docId);
    var docBlob = doc.getAs(MimeType.PDF);
    docBlob.setName(doc.getName() + '.pdf');
    return docBlob;
  } catch (error) {
    Logger.log('Error converting Google Doc to PDF: ' + error);
    return null;
  }
}

// Function to convert MS Word document to PDF
function convertWordToPdf(file) {
  try {
    // For MS Word files, we need to convert to Google Docs first, then to PDF
    // This is because blob.getAs(MimeType.PDF) doesn't work directly with Office formats
    
    var fileName = file.getName();
    var fileId = file.getId();
    
    // Create a temporary Google Doc from the Word file
    var folderId = file.getParents().next().getId();
    var folder = DriveApp.getFolderById(folderId);
    
    // Import the Word file as a Google Doc
    var resource = {
      title: fileName + '_temp',
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    var blob = file.getBlob();
    var doc = Drive.Files.insert(resource, blob, {convert: true});
    
    // Get the converted Google Doc and export as PDF
    var tempDocId = doc.id;
    var tempDoc = DocumentApp.openById(tempDocId);
    var pdfBlob = tempDoc.getAs(MimeType.PDF);
    
    // Set proper PDF name
    var pdfName = fileName.replace(/\.(docx?|DOCX?)$/, '') + '.pdf';
    pdfBlob.setName(pdfName);
    
    // Delete the temporary Google Doc
    DriveApp.getFileById(tempDocId).setTrashed(true);
    
    Logger.log('Converted Word document to PDF: ' + pdfName);
    return pdfBlob;
  } catch (error) {
    Logger.log('Error converting Word document to PDF: ' + error);
    // If conversion fails, return the original file
    try {
      return file.getBlob();
    } catch (e) {
      Logger.log('Error getting original file blob: ' + e);
      return null;
    }
  }
}

// Function to convert MS Excel spreadsheet to PDF
function convertExcelToPdf(file) {
  try {
    // For MS Excel files, we need to convert to Google Sheets first, then to PDF
    
    var fileName = file.getName();
    var fileId = file.getId();
    
    // Create a temporary Google Sheet from the Excel file
    var folderId = file.getParents().next().getId();
    var folder = DriveApp.getFolderById(folderId);
    
    // Import the Excel file as a Google Sheet
    var resource = {
      title: fileName + '_temp',
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    var blob = file.getBlob();
    var sheet = Drive.Files.insert(resource, blob, {convert: true});
    
    // Get the converted Google Sheet and export as PDF
    var tempSheetId = sheet.id;
    var tempFile = DriveApp.getFileById(tempSheetId);
    var pdfBlob = tempFile.getAs(MimeType.PDF);
    
    // Set proper PDF name
    var pdfName = fileName.replace(/\.(xlsx?|XLSX?)$/, '') + '.pdf';
    pdfBlob.setName(pdfName);
    
    // Delete the temporary Google Sheet
    tempFile.setTrashed(true);
    
    Logger.log('Converted Excel spreadsheet to PDF: ' + pdfName);
    return pdfBlob;
  } catch (error) {
    Logger.log('Error converting Excel spreadsheet to PDF: ' + error);
    // If conversion fails, return the original file
    try {
      return file.getBlob();
    } catch (e) {
      Logger.log('Error getting original file blob: ' + e);
      return null;
    }
  }
}

// Function to parse email addresses from a string (comma or semicolon separated)
function parseEmailAddresses(emailString) {
  if (!emailString || typeof emailString !== 'string') {
    return [];
  }
  
  // Split by comma or semicolon
  var emails = emailString.split(/[,;]/).map(function(email) {
    return email.trim();
  }).filter(function(email) {
    return email && isValidEmail(email);
  });
  
  return emails;
}

// Function to validate email address
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
