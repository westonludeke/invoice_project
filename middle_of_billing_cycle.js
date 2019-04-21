//**You must be on the "Raw API Data" tab in the Sheet as the active tab before running the code**

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stripe')
    .addItem('Retrieve Invoice Data','getInvoiceObj')
    .addToUi();
}

function getInvoiceObj() 
  {
    var apiKey, content, options, response, secret, url;
    
    //**Stripe Test Keys / For Testing Purposes Only**
    //secret = "sk_test_xxxxxxxxxxxxxxxxxxxxxxxxxx";
    //apiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxx";
    
    //**Live Key**
    secret = "rk_live_xxxxxxxxxxxxxxxxxxxxxxxxxx";
    apiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxx";
    
    //Stripe API invoice URL ***only update the trailing code after /invoices/***
    url = "https://api.stripe.com/v1/invoices/in_xxxxxxxxxxxxxxxxxxxxxxxxxx";
    
    //Direct Stipe dashboard URL to access a test invoice: https://dashboard.stripe.com/test/invoices/in_xxxxxxxxxxxxxxxxxxxxxxxxxx
    
    options = {
      "method" : "GET",
      "headers": {
        "Authorization": "Bearer " + secret 
      },
      "muteHttpExceptions":true
    };
    
    response = UrlFetchApp.fetch(url, options);
    
    //Push data to Sheet from invoice. **Writes over existing Sheet data**
    content = JSON.parse(response.getContentText());
    var sheet = SpreadsheetApp.getActiveSheet();
   
    //Retrieves currency type and adds to Sheet
    sheet.getRange(5,2).setValue([content.currency.toUpperCase()]);
    
    //Invoice amount due
    sheet.getRange(3,2).setValue([content.amount_due]);
    
    //Invoice date
    sheet.getRange(1,2).setValue([content.date]);
    
    //Invoice period begin / New Billing Cycle Only
    //sheet.getRange(6,2).setValue([content.lines.data[0].period.start]);
    
    //Yearly invoice period begin
    sheet.getRange(6,2).setValue([content.period_start]);
    
    //Invoice period end / New Billing Cycle Only
    //sheet.getRange(7,2).setValue([content.lines.data[0].period.end]);
    
    //Invoice period end yearly
    sheet.getRange(11,2).setValue([content.period_end]);
    
    //Number of licenses / New Billing Cycle Only
    //sheet.getRange(10,2).setValue([content.lines.data[0].quantity]);
    
    //Invoice number
    sheet.getRange(13,2).setValue([content.number]);
    
    //**Logs for testing**
    //Logger.log("Yes, this was logged!");
    Logger.log(response);
    //Logger.log(content);
}