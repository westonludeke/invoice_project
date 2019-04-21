# Stripe Invoice to Google Sheet Billing Template 

### This project is a Google Apps Script to retrieve billing data from the Stripe API and import it into an invoice template within Google Sheets. To use:

1. Open up the invoice template within Google Sheets

2. Go to Tools > Script Editor

3. Paste in the code and save. Then make sure "getInvoiceObj" is selected in the dropdown in the menu bar.

4. Add in your unique API key next to "secret" and "apiKey" on lines 14 and 15.

5. Go to Stripe and open up a paid invoice that you want to track

6. In the URL section, grab the unique invoice ID

7. Open up the script page and paste the invoice ID in the section in the URL section of the script. **Note: Only update the section of the URL after /invoices/**

8. Go back to the Google Sheet and make sure you have "Raw API Data" selected as the active tab.

9. Click "Stripe" in the menu bar, then "Retreive Invoice Data".

10. The data should fill out the "Raw API Data" tab. You will see rows in this tab that have the raw data from the API and other rows have the converted data (ex: Converting epoch/UNIX time to standard MM/DD/YYYY date format).

11. After the script runs, the "Raw API Data" converted data will automatically update the second tab called "Invoice". You can click the "Invoice" tab to confirm.

12. **Important! - This project is still under "code review", meaning it's not 100% foolproof yet. Be sure after running the script to match up your invoice in the Sheet with the invoice in Stripe and make sure the data matches**

13. Once you confirm the data is correct, you can fill out the remaining customer data, such as their billing address, VAT, etc. You will need to do this by hand.

14. Once the invoice is completed, go to File > Download As > PDF to export the Sheet to a PDF for the customer.

### If you don't see "Stripe" in the Google Sheet menu bar:

1. Open the Script Editor in the Google Sheet.

2. Click "onOpen" in the dropdown on the top menu bar. 

3. Click to run the script.

4. This should add the "Stripe" menu bar to the Google Sheet.

**Note: You will need to click back on "getInvoiceObj" in the Script Editor before running the invoice script.**
