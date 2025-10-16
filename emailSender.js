/**
 * Mass Email Sender Module
 * Handles scheduling and sending of mass emails with attachments
 */

// Function to check email and attachment quotas
function checkEmailQuotas(config) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    var data = sheet.getDataRange().getValues();
    var emailColumnIndex = data[config.headerRowIndex - 1].indexOf(config.emailColumn);
    
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
    if (!config.sheetName || !config.emailColumn || !config.subject || !config.body) {
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
    
    // Find column indices
    var emailColumnIndex = headers.indexOf(config.emailColumn);
    if (emailColumnIndex === -1) {
      throw new Error('Columna d\'email no trobada: ' + config.emailColumn);
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
      var email = row[emailColumnIndex];
      
      if (!email || !isValidEmail(email)) {
        sheet.getRange(i + 1, logColumnIndex + 1).setValue('Error: Email invàlid');
        errorCount++;
        continue;
      }

      try {
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

        // Send email
        var emailOptions = {
          name: 'Enviament massiu',
          attachments: attachments
        };

        MailApp.sendEmail(email, personalizedSubject, personalizedBody, emailOptions);
        
        // Log success
        var timestamp = new Date().toLocaleString('ca-ES');
        sheet.getRange(i + 1, logColumnIndex + 1).setValue('Enviat: ' + timestamp);
        successCount++;

      } catch (emailError) {
        Logger.log('Error sending email to ' + email + ': ' + emailError);
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

// Function to get attachment from URL (converts Google Docs to PDF)
function getAttachmentFromUrl(url) {
  try {
    // Check if it's a Google Doc
    var docMatch = url.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)/);
    if (docMatch) {
      var docId = docMatch[1];
      return convertGoogleDocToPdf(docId);
    }

    // Check if it's a Google Drive file
    var driveMatch = url.match(/drive\.google\.com\/.*[?&]id=([a-zA-Z0-9-_]+)/);
    if (driveMatch) {
      var fileId = driveMatch[1];
      var file = DriveApp.getFileById(fileId);
      return file.getAs(MimeType.PDF);
    }

    // Try to get file by extracting ID
    var fileId = extractFileId(url);
    if (fileId) {
      var file = DriveApp.getFileById(fileId);
      
      // Convert Google Docs to PDF
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        return convertGoogleDocToPdf(fileId);
      }
      
      // Return other files as-is
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

// Function to validate email address
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
