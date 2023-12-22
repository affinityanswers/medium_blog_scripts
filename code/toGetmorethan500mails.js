// Function to extract more than 500 mails using pagination

function getEmailsAndWriteToSpreadsheet() {
    var targetEmailDomain = 'domainName';
    var batchSize = 500; // Number of threads to fetch per request
    var maxEmails = 500; // Maximum emails to retrieve
  
    // Log in to Gmail using the advanced Gmail service
    var threads = [];
    var pageToken = null;
  
     // Create a new Google Sheet or use an existing one
    var spreadsheet = SpreadsheetApp.openById('sheetID');
    var sheet = spreadsheet.getSheetByName('sheetName'); // Change the sheet name as neede
  
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
  