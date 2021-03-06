# For use with a google drive spreadsheet.
import json
import math
import sys

import gspread
import requests
from oauth2client.service_account import ServiceAccountCredentials

# User Data:
# Edit the following variable to match your sheet
# The columns should correspond to those in your sheet, with column 'a' being '1' and so on.
spreadsheetName = "Copy of Magic Collection"
cardNameColumn = 2
foilFlagColumn = 3
priceColumn = 4

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)


def main():
    sheet = client.open(spreadsheetName).sheet1

    cardNames = sheet.col_values(cardNameColumn)
    foil = sheet.col_values(foilFlagColumn)

    header = True
    numCards = len(cardNames)

    for i in range(len(cardNames)):
        if (header):
            header = False
        else:
            currCard = make_card(cardNames[i], checkFoil(foil[i]))
            price = "$" + getPriceFor(currCard.name, currCard.foil)
            print("cost of " + currCard.name + " is " + price)
            sys.stdout.write('\r')
            sys.stdout.write(
                "[{:{}}] {:.1f}%".format("=" * math.floor(100 / (numCards - 1) * i), 100, (100 / (numCards - 1) * i)))
            sys.stdout.flush()
            sheet.update_cell(i + 1, priceColumn, price)


class Card(object):
    name = ""
    foilFlag = bool
    foil = bool

    # The class "constructor" - It's actually an initializer
    def __init__(self, name, foilFlag):
        self.name = name
        self.foil = foilFlag


def make_card(name, foil):
    card = Card(name, foil)
    return card


def getPriceFor(cardName, foilFlag):
    resp = requests.get('https://api.scryfall.com/cards/named?fuzzy=' + cardName)
    json_resp = resp.json()
    dump = json.dumps(json_resp)
    data = json.loads(dump)
    if ((foilFlag) and (data['prices']['usd_foil'] != None)):
        return data['prices']['usd_foil']
    elif ((not foilFlag) and (data['prices']['usd'] != None)):
        return data['prices']['usd']
    else:
        return "No Price Data"


def checkFoil(foilFlag):
    if (foilFlag == "Yes"):
        return True
    else:
        return False


if __name__ == '__main__':
    main()
