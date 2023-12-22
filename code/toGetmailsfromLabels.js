// Function to extract mails from the Labels

function getEmailsFromLabelsAndWriteToSpreadsheet() {
    var targetDomain = 'domainName';
    var labelName = 'LabelName';
    var maxResults = 500; // Maximum number of emails to retrieve per page

    // Log in to Gmail
    var gmailApp = GmailApp;
    var threads = gmailApp.search('from:' + targetDomain);
  
  
    // Create a new Google Sheet or use an existing one
    var spreadsheet = SpreadsheetApp.openById('sheetID');
    var sheet = spreadsheet.getSheetByName('sheetName'); // Change the sheet name as needed
  
  
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
  
  // Function to extract sender info
  
  function extractSenderEmail(senderInfo) {
    var emailMatch = /<([^>]+)>/.exec(senderInfo);
    if (emailMatch) {
      return emailMatch[1];
    } else {
      return senderInfo;
    }
  }