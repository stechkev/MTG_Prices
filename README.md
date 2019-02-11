# MTG_Prices
Pulls MTG pricing info from Scryfall then updates a google sheet with that info

Google don't forget these steps:
There is one last required step to authorize your app, and it’s easy to miss!

Find the client_email inside client_secret.json. Back in your spreadsheet, click the Share button in the top right, and paste the client email into the People field to give it edit rights. Hit Send.

If you skip this step, you’ll get a gspread.exceptions.SpreadsheetNotFound error when you try to access the spreadsheet from Python.
