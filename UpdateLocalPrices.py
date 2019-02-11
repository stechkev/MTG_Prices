import json
import os
import sys

import requests
import xlrd
from xlutils.copy import copy

# 0 indexed
# The price column will be overwritten with new prices
spreadsheetPath = "Test_Price.xlsx"
cardNameColumn = 0
foilFlagColumn = 1
priceColumn = 2


def main():
    originalBook = xlrd.open_workbook(spreadsheetPath)
    originalSheet = originalBook.sheet_by_index(0)
    tempBook = copy(originalBook)
    tempSheet = tempBook.get_sheet(0)

    cardNames = originalSheet.col_values(cardNameColumn)
    foil = originalSheet.col_values(foilFlagColumn)

    header = True
    numCards = len(cardNames)

    for i in range(numCards):
        if (header):
            header = False
        else:
            currCard = make_card(cardNames[i], checkFoil(foil[i]))
            price = "$" + getPriceFor(currCard.name, currCard.foil)
            # print("cost of " + currCard.name + " is " + price)
            tempSheet.write(i, priceColumn, price)
            sys.stdout.write('\r')
            sys.stdout.write("[{:{}}] {:.1f}%".format("=" * i, numCards - 1, (100 / (numCards - 1) * i)))
            sys.stdout.flush()
    tempBook.save(spreadsheetPath + '.new' + os.path.splitext(spreadsheetPath)[-1])


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
    if (resp.status_code != 200):
        return "Card Not Found"
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


# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()
