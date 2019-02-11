# MTG_Prices
Pulls MTG pricing info from Scryfall then updates a google sheet with that info
_________________
Do these steps:
1. go here, click the button https://www.python.org/downloads/
2. update the variables in the script
3. go to the script you want to run and type: python scriptname.py
_________________


Local Xlsx:
Need the full path to the file you're using, as well as the column indexes. Those variables at the top should be the only ones you need to change.

Google don't forget these steps:
There is one last required step to authorize your app, and it’s easy to miss!

Find the client_email inside client_secret.json. Back in your spreadsheet, click the Share button in the top right, and paste the client email into the People field to give it edit rights. Hit Send.

If you skip this step, you’ll get a gspread.exceptions.SpreadsheetNotFound error when you try to access the spreadsheet from Python.
