// Author: Vineet JK
// github : https://github.com/vineetjk


function getDataSheet() {

  sheet = SpreadsheetApp.getActiveSheet();

  startRow = 2;  // First row of data to process
  numRows = 100;   // Number of rows to process
  startCol = 1;  //First column of data to process
  numCols = 6;    // Number of columns to process 

  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols);

  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  return data;
}

function getMessage(name) {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Message');

  var message = htmlOutput.getContent()
  message = message.replace("%name", name);
  

  return message;
}

function sendEmail() {

  var emailSent = "Yes";

  var emailCol = 3;

  var data = getDataSheet();

  for (var i = 0; i < data.length; i++) {

    var row = data[i];

   
    var isEmailSent = row[2];
    var name = row[0];
    

    if (isEmailSent != emailSent) {

      var subject = "Get Your own Website at affordable price.";
      var message = getMessage(name);

      var recipientEmail = row[1];


      MailApp.sendEmail(recipientEmail, subject, message, { htmlBody: message, name: "Infinity Web Design" });

      sheet.getRange(startRow + i, emailCol).setValue(emailSent);
    }
  }
}