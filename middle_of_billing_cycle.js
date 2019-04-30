/* HOW TO USE

1. Update the URL from the Stripe API with the information from the unique invoice URL
2. Save the code
3. Make sure "getInvoiceObj" is selected in the toolbar above
4. Go to the Google Sheet and make sure you're on the "Raw API Data" tab in the Sheet as the active tab
5. To run, either click "Stripe" > "Retreive Invoice Data" from the top of the Google Sheet.
   Or you can click the play icon from the scripts page.

*/

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
    
      /* ~~~~~~~IMPORTANT~~~~~~
    
    With the invoice API URL below, only update the code of the key after /invoices/ and before ?expand[]=charge&expand[]=customer";   
    
    You will replace the unique invoice ID, for example "in_1xxxxxxxxxxxxxxxxxxxxxxxxxx", with the ID of the invoice you want to pull.
    
    ~~~~~~~IMPORTANT~~~~~~*/
    url = "https://api.stripe.com/v1/invoices/in_xxxxxxxxxxxxxxxxxxxxxxxxxx?expand[]=charge&expand[]=customer";
    
   
      
    /* The URL below is for the Invoice Object only:
      url = "https://api.stripe.com/v1/invoices/in_xxxxxxxxxxxxxxxxxxxxxxxxxx";
    */
    
    /* The URL below is for the Charge Object only:
      url = "https://api.stripe.com/v1/charges/ch_xxxxxxxxxxxxxxxxxxxxxxxxxx"
    */
    
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
    
    //Payer's email address
    sheet.getRange(14,2).setValue([content.charge.card.name]);
    
    //Credit card type
    sheet.getRange(15,2).setValue([content.charge.card.brand]);
    
    //Credit card last four
    sheet.getRange(16,2).setValue([content.charge.card.last4]);
    
    //Credit card expiry month
    sheet.getRange(17,2).setValue([content.charge.card.exp_month]);
    
    //Credit card expiry year
    sheet.getRange(18,2).setValue([content.charge.card.exp_year]);
    
    //Is the invoice paid?
    sheet.getRange(23,2).setValue([content.charge.paid]);
    
    /*       Team Discount
    Notes:
    
    1. This seems to be the Team's current discount, not the discount at the time of the invoice
    2. If the team has no discount, when running the script you'll see a "TypeError" error message,
       which just means the value for discount in the API code equals "null".
    3. If you receive this error, the script will stop running at this point and you will need to manually
       remove the old discount from the Sheet if one is left there.
    */
    sheet.getRange(21,2).setValue([content.customer.subscriptions.data[0].discount.coupon.percent_off]);
    
    
    //**Logs for testing**
    //Logger.log("Yes, this was logged!");
    Logger.log(response);
    //Logger.log(content);
}
                        

