/*

This code adds the ability to export the invoice as a PDF and sends as an email. 
It will also save a copy of the PDF to a specific folder within Google Drive.

1. If you don't see the "Email and Save" button inside of the Google Sheet template, run the "onOpenTwo" function first.

2. To run, from inside the Google Sheet click "Email and Save", then click "Email Invoice and Save Copy".

3. This will export only the Invoice tab within the Sheet to a PDF and email it to the email address included in the code below.
   It will then save a copy of the PDF to the designated Google Drive folder in the code below.

*/

//Code to add the buttons to the Google Sheet
function onOpenTwo() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email and Save')
    .addItem('Email Invoice and Save Copy','emailSpreadsheetAsPDF')
    .addToUi();
}

/* Send Spreadsheet in an email as PDF, automatically */
function emailSpreadsheetAsPDF() {
  
  // Send the PDF of the spreadsheet to this email address
  var email = "user@domain.com"; 
  
  // Get the currently active spreadsheet URL (link)
  // Or use SpreadsheetApp.openByUrl("<<SPREADSHEET URL>>");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
  //Get only the Invoice sheet from the Google Sheet
  var sheet = ss.getSheetByName("Invoice");
  
  // Subject of email message
  var subject = "Your requested Streak invoice from Weston" //+ ss.getName(); 

  // Email Body can  be HTML too with your logo image - see ctrlq.org/html-mail
  var body = "Here's a copy of your latest invoice! <br><br>Please reply back if anything is incorrect and/or if you need anything else.<br><br><b><u>Note for Andrew</u></b>: This email and the attached PDF invoice was generated via one-click from the Google Sheets menu bar. Let me know when you get it!<br><br>Best,<br>Weston<br>Streak Support";
  
  // Base URL
  var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
  
  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&size=letter'                       // paper size legal / letter / A4
  + '&portrait=true'                    // orientation, false for landscape
  + '&fitw=true&source=labnol'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=';                             // the sheet's Id
  
  var token = ScriptApp.getOAuthToken();

  var response = UrlFetchApp.fetch(url + url_ext + sheet.getSheetId(), {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    }).getBlob().setName(sheet.getName() + ".pdf");

  /* Code to download a copy of the invoice as a PDF and save to a specific folder within Google Drive */

  //Adds PDF copy of the invoice to the root Drive directory
  var createPDF = DriveApp.createFile(response).getId();
  
  //Grabs a specific Drive folder
  var driveFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxxxxxxxxxx");
  
  //Grabs the PDF
  var file = DriveApp.getFileById(createPDF);
  
  //Moves the PDF to the specified directory
  var moveFile = driveFolder.addFile(file);
     
  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0) 
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[response]     
    });
}