function getEmailsAndWriteToSpreadsheet() {
    var targetEmailDomain = 'dsr-incogni.com';
    var batchSize = 500; // Number of threads to fetch per request
    var maxEmails = 500; // Maximum emails to retrieve
  
    // Log in to Gmail using the advanced Gmail service
    var threads = [];
    var pageToken = null;
  
     // Create a new Google Sheet or use an existing one
    var spreadsheet = SpreadsheetApp.openById('15FQFkGTlojpxAyV3RH_Dwi92IvDKrOphw8eypesZXZg');
    var sheet = spreadsheet.getSheetByName('Log&Status_dsr-incongni'); // Change the sheet name as neede
  
    do {
      var options = {
        q: 'from:*@' + targetEmailDomain, // Modify the query as needed
        maxResults: Math.min(batchSize, maxEmails - threads.length),
        pageToken: pageToken,
      };
  
      var response = Gmail.Users.Messages.list('me', options);
      var messages = response.messages;
      console.log("length of messages", messages.length);
  
      // Write headers to the spreadsheet
      console.log("number of emails", threads.length);
      sheet.getRange(1, 1).setValue('Subject');
      sheet.getRange(1, 2).setValue('Date');
      sheet.getRange(1, 3).setValue('Name');
      sheet.getRange(1, 4).setValue('Email Address');
      sheet.getRange(1, 5).setValue('Address');
      sheet.getRange(1, 6).setValue('Request Received via');
   
      for (var i = 0; i < messages.length; i++) {
        var message = GmailApp.getMessageById(messages[i].id);
        var subject = message.getSubject();
        var date = message.getDate();
        var body = message.getPlainBody();
        var senderEmail = extractSenderEmail(message.getFrom());
  
  
        // Search for "Name" and "Email-address" in the email body
        var nameMatch = /Name: ([^\n]+)/i.exec(body);
        var emailMatch = /E-mail address: ([^\n]+)/i.exec(body);
        var addressMatch = /Address: ([^\n]+)/i.exec(body);
  
  
  
        // Check if both "Name" and "Email-address" were found
        if (nameMatch && emailMatch) {
          var name = nameMatch[1].trim();
          var email = emailMatch[1].trim();
          var address = addressMatch ? addressMatch[1].trim() : '';
  
   
          // Write email details to the spreadsheet
          // Make sure to update the code here to write to your spreadsheet
         // Logger.log('Subject: ' + subject);
          // Logger.log('Date: ' + date);
          // Logger.log('Name: ' + name);
          // Logger.log('Email Address: ' + email);
          // Logger.log('Address: ' + address);
          // Logger.log('Sender Email: ' + senderEmail);
          // Write email details to the spreadsheet
          sheet.appendRow([subject, date, name, email, address, senderEmail]);
          // Add the email to the 'threads' array
          threads.push(message);
          if(threads.length >= 500 && threads.length <=550) {
            Logger.log('Date: ' + date);
          }
  
  
          if (threads.length >= maxEmails) {
            // Stop processing if the maximum number of emails is reached
            return;
          }
        }
      }
  
  
      pageToken = response.nextPageToken;
    } while (pageToken);
  
  
    Logger.log('Total emails processed: ' + threads.length);
  }
  
  
  // Function to extract sender's email address from the sender information
  function extractSenderEmail(senderInfo) {
    var emailMatch = /<([^>]+)>/.exec(senderInfo);
    if (emailMatch) {
      return emailMatch[1];
    } else {
      return senderInfo;
    }
  }
  
  function getEmailsFromLabelsAndWriteToSpreadsheet() {
    var labelName = 'MK/US1';
    var maxResults = 500; // Maximum number of emails to retrieve per page
  
    // Create a new Google Sheet or use an existing one
    var spreadsheet = SpreadsheetApp.openById('1Ral6IpcNOv4B3nzsyjC22kH8c0yPO2o28kgajonM_T0');
    var sheet = spreadsheet.getSheetByName('UnitedStates'); // Change the sheet name as needed
  
    // Get the list of processed message IDs from Script Properties
    var scriptProperties = PropertiesService.getScriptProperties();
    var processedMessageIds = scriptProperties.getProperty('processedMessageIds');
    processedMessageIds = processedMessageIds ? processedMessageIds.split(',') : [];
  
    // Write headers to the spreadsheet if they don't exist
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 6).setValues([['Subject', 'Date', 'Name', 'Email Address', 'Address', 'Sender Email']]);
    }
  
    var pageToken = '';
  
    do {
      var threads = Gmail.Users.Threads.list('me', {
        q: 'label:' + labelName,
        maxResults: maxResults,
        pageToken: pageToken
      });
  
      var messages = [];
  
      for (var i = 0; i < threads.threads.length; i++) {
        var threadId = threads.threads[i].id;
        var thread = GmailApp.getThreadById(threadId);
        var threadMessages = thread.getMessages();
        messages = messages.concat(threadMessages);
      }
  
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var messageId = message.getId();
  
        if (!processedMessageIds.includes(messageId)) {
          var subject = message.getSubject();
          var date = message.getDate();
          var body = message.getPlainBody();
          var senderEmail = extractSenderEmail(message.getFrom());
  
          var nameMatch = /Name: ([^\n]+)/i.exec(body);
          var emailMatch = /E-mail address: ([^\n]+)/i.exec(body);
          var addressMatch = /Address: ([^\n]+)/i.exec(body);
  
          if (nameMatch && emailMatch && addressMatch) {
            var name = nameMatch[1].trim();
            var email = emailMatch[1].trim();
            var address = addressMatch[1].trim();
  
            sheet.appendRow([subject, date, name, email, address, senderEmail]);
            processedMessageIds.push(messageId);
          }
        }
      }
  
      // Set the page token for the next page of threads
      pageToken = threads.nextPageToken;
    } while (pageToken);
  
    // Update the list of processed message IDs in Script Properties
    scriptProperties.setProperty('processedMessageIds', processedMessageIds.join(','));
  }
  
  function extractSenderEmail(senderInfo) {
    var emailMatch = /<([^>]+)>/.exec(senderInfo);
    if (emailMatch) {
      return emailMatch[1];
    } else {
      return senderInfo;
    }
  }
  
  
  
  