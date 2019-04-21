# Project: Pull a Stripe Invoice into a Google Sheet Billing Template 

## TL;DR: This project is a Google Apps Script written in JavaScript to use the Stripe API to pull a Stripe invoice into a premade Google Sheet invoice template. 

*Note: There is separate code for net-new billing cycles, and invoices generated in the middle of existing billing cycles, such as adding a user in the middle of a yearly billing cycle.*

## How To Use:

1. Open up the invoice template within Google Sheets

2. Go to Tools > Script Editor

3. In the code editor where it says *"Code.gs"*, paste in the code from the GitHub file and save. Then make sure *"getInvoiceObj"* is selected in the dropdown in the menu bar (instead of either *"Select function"* or *"onOpen"*).

4. Add in your unique API key next to *"secret"* and *"apiKey"*. This is roughly on lines 19 and 20.

5. Go to Stripe and open up a paid invoice that you want to track

6. In the URL section, grab the unique invoice ID

7. Open up the Google Sheet's Script Editor page and paste the invoice ID in the section in the URL section of the script. This is roughly on 23 of the code **Note: Only update the section of the URL after _/invoices/_**

8. Go back to the Google Sheet (the actual Sheet) and make sure you have *"Raw API Data"* selected as the active tab.

9. Click *"Stripe"* in the menu bar, then *"Retreive Invoice Data"*.

10. The script will fill out the data from the invoice you selected in the *"Raw API Data"* tab. **Do Not Edit This Data Manually** You will see rows in this tab that have the raw data from the API and other rows have the converted data (ex: Converting epoch/UNIX time to standard MM/DD/YYYY date format).

11. After the script runs, the *"Raw API Data"* converted data will automatically update the second tab called *"Invoice"*. You can click the *"Invoice"* tab to confirm.

12. **Important! - This project is still under "code review", meaning it's not 100% foolproof yet. There's still a potential for bugs. Be sure after running the script to match up your invoice in the Sheet with the invoice in Stripe and make sure the data matches**

13. Once you confirm the data is correct, you can fill out the remaining customer data, such as their billing address, VAT, etc. You will need to do this by hand.

14. Once the invoice is completed, go to File > Download As > PDF to export the Sheet to a PDF for the customer.

### Troubleshooting - If you don't see "Stripe" in the Google Sheet menu bar:

1. Open the Script Editor in the Google Sheet.

2. Click *"onOpen"* in the dropdown on the top menu bar on the Scipt Editor page. 

3. Click to run the script (play button in the menu bar).

4. This should add the *"Stripe"* menu bar to the Google Sheet.

**Note: You will need to click back on "getInvoiceObj" in the Script Editor before running the invoice script.**

### If all else fails ask the program's developer! (•‿•)
