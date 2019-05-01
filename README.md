# Project: Generate Custom Invoices from Stripe in Google Sheets

**What This Is**: This project contains two Google Apps Scripts written in JavaScript. The first script retrieves an invoice from Stripe via the Stripe API and populates the data into premade Google Sheet invoice template. The second script saves the invoice from Sheets to a PDF and sends an email to a designated recipient, while at the same time saving the PDF of the invoice in a designated Google Drive folder.

## Script to Retreive Invoice Data From Stripe

*Note: There is separate code for invoices generated at the beginning of a billing cycle (such as a user's monthly payment), and invoices generated in the middle of existing billing cycles (such as adding a user in the middle of a yearly billing cycle).*

## How To Use:

1. Open up the invoice template within Google Sheets

2. Go to Tools > Script Editor

3. In the code editor where it says *"Code.gs"*, paste in the code from the GitHub file and save. Then make sure *"getInvoiceObj"* is selected in the dropdown in the menu bar (instead of either *"Select function"* or *"onOpen"*).

4. Add in your unique API key next to *"secret"* and *"apiKey"*. 

5. Go to Stripe and open up a paid invoice that you want to track.

6. In the URL section, grab the unique invoice ID.

7. Open up the Google Sheet's Script Editor page and paste the invoice ID in the section in the URL section of the script.  **Note: Only update the section of the URL after _/invoices/_**

8. Go back to the Google Sheet (the actual Sheet) and make sure you have *"Raw API Data"* selected as the active tab.

9. Click *"Stripe"* in the menu bar, then *"Retrieve Invoice Data"*.

10. The script will fill out the data from the invoice you selected in the *"Raw API Data"* tab. **Do Not Edit This Data Manually.** You will see rows in this tab that have the raw data from the API and other rows have the converted data (ex: Converting epoch/UNIX time to standard MM/DD/YYYY date format).

11. After the script runs, the *"Raw API Data"* converted data will automatically update the second tab called something like *"Invoice"*. You can click the *"Invoice"* tab to confirm the data was populated correctly.

12. **Important! - This project is still under "code review", meaning it's not 100% foolproof yet. There's still a potential for bugs. Be sure after running the script to match up your invoice in the Sheet with the invoice in Stripe and make sure the data is the same.**

13. Once you confirm the data is correct you can fill out the remaining customer data, such as their billing address, VAT, etc. You will need to do this by hand.

14. Once the invoice is completed, you can proceed to using the next script.

### Troubleshooting - If you don't see "Stripe" in the Google Sheet menu bar:

1. Open the Script Editor in the Google Sheet.

2. Click *"onOpen"* in the dropdown on the top menu bar on the Scipt Editor page. 

3. Click to run the script (play button in the menu bar).

4. This should add the *"Stripe"* menu bar to the Google Sheet.

**Note: You will need to click back on "_getInvoiceObj_" in the Script Editor before running the invoice script.**

## Script to Email as PDF and Save a PDF to Google Drive:

1. Open the code inside of Sheets from Tools > Script Editor.

2. Make sure you're viewing the section on "automation.gs".

3. Under the comment "Send the PDF of the spreadsheet to this email address", add in an email address of where you'd like your email to be sent to.

4. Open the folder in Google Drive where you'd like to save a copy of the invoice.

5. Grab the folder ID, which is the trailing string after /folders/ in the URL bar: https://drive.google.com/drive/folders/xxxxxxxxxxxxxxx

6. Under the commnent "Grabs a specific Drive folder", add in the folder ID to this section between the quotation marks.

7. Run this script by opening the Google Sheet and selecting "Email and Save", then "Email Invoice and Save Copy".

This will send a PDF copy of the invoice to the email address you specifed in step #3 and will save a PDF copy of the invoice in the Google Drive folder specified in step #5.

**Note** You can change the subject line of the email udner the "Subject of email message" comment and change the body of the email, including using HTML, on the next line under "Email Body (HTML can be used)".

### If all else fails ask the program's developer! (•‿•)
