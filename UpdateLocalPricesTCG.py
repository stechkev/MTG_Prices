import json
import math
import os
import sys

import requests
import xlrd
from xlutils.copy import copy

# 0 indexed
# The price column will be overwritten with new prices
# spreadsheetPath = "Test_Price_short.xlsx"
spreadsheetPath = "Copy of Magic Collection.xlsx"
cardNameColumn = 1
foilFlagColumn = 2
priceColumn = 3
setNameColumn = 6
sheetIndex = 0


def main():
    accessToken = "enHtRLYSOOw6N01rcGfUeFvsq4MDnLFdQHHg_EuPnC-Prpa6XqamPZGnHgaQx0mcvj0xCDQsTwLLEPB6_oLhr6PHiBckm-yZIRpBjlXNyOHabw3yMyflBYDbUP6uMeLBhVwH1j1BPyZ9ohIyBPUH8RvsV0gKvXm8gdP8ub4YAkdVI1-ptA6eGVNKpKAGKAmyuJwEsfwGkEWEY8Tmei9wb2N9A_UuMeOGxIe9bsy3jDAABp3mv1Di1hMaxyXn3x6BoQ_qVauoPx2eGBr9KN95pqiwF85cY1k_HpSkTxhMOXvDU-yptHR4UNRrxPHnQKtQ1kLxVg"
    originalBook = xlrd.open_workbook(spreadsheetPath)
    originalSheet = originalBook.sheet_by_index(sheetIndex)
    tempBook = copy(originalBook)
    tempSheet = tempBook.get_sheet(sheetIndex)

    cardNames = originalSheet.col_values(cardNameColumn)
    setName = originalSheet.col_values(setNameColumn)
    foil = originalSheet.col_values(foilFlagColumn)

    header = True
    numCards = len(cardNames)

    for i in range(numCards):
        if (header):
            header = False
        else:
            currCard = make_card(cardNames[i], checkFoil(foil[i]), setName[i])
            productId = getProductId(accessToken, cardNames[i], setName[i])
            price = "$" + str(getPriceFor(accessToken, currCard, productId))
            # print("cost of " + currCard.name + " is " + price)
            tempSheet.write(i, priceColumn, price)
            sys.stdout.write('\r')
            sys.stdout.write(
                "[{:{}}] {:.1f}%".format("=" * math.floor(100 / (numCards - 1) * i), 100, (100 / (numCards - 1) * i)))
            sys.stdout.flush()
    tempBook.save(spreadsheetPath + '.new' + os.path.splitext(spreadsheetPath)[-1])


class Card(object):
    name = ""
    foilFlag = bool
    foil = bool
    setName = ""

    # The class "constructor" - It's actually an initializer
    def __init__(self, name, foilFlag, set):
        self.name = name
        self.foil = foilFlag
        self.setName = set


def make_card(name, foil, set):
    card = Card(name, foil, set)
    return card


def getPriceFor(accessToken, card, productId):
    headers = {
        'Accept': 'application/json',
        'Authorization': 'bearer ' + accessToken,
    }
    url = "http://api.tcgplayer.com/v1.19.0/pricing/product/" + str(productId)

    querystring = {"getExtendedFields": "true"}

    resp = requests.request("GET", url, headers=headers, params=querystring)
    if (resp.status_code != 200):
        return "Card Not Found"
    json_resp = resp.json()
    dump = json.dumps(json_resp)
    data = json.loads(dump)
    if ((card.foil) and (data['results'][1]['marketPrice'] != None)):
        return data['results'][1]['marketPrice']
    elif ((not card.foil) and (data['results'][0]['marketPrice'] != None)):
        return data['results'][0]['marketPrice']
    else:
        return "No Price Data"


def authenticate():
    headers = {
        '': 'application/x-www-form-urlencoded',
    }

    data = 'grant_type=client_credentials&client_id=76F01709-0882-4CAF-94A4-221AE01440FB&client_secret=402F2156-B054-4585-BD1E-D949E6453F38'

    response = requests.post('https://api.tcgplayer.com/token', headers=headers, data=data)

    json_resp = response.json()
    dump = json.dumps(json_resp)
    data = json.loads(dump)

    return data['access_token']


def getProductId(accessToken, cardName, setName):
    headers = {
        'Accept': 'application/json',
        'Authorization': 'bearer ' + accessToken,
    }
    url = "http://api.tcgplayer.com/v1.19.0/catalog/products"

    querystring = {"categoryId": "1", "groupName": setName, "productName": cardName,
                   "getExtendedFields": "false"}

    response = requests.request("GET", url, headers=headers, params=querystring)
    if (response.status_code != 200):
        return "Card Not Found"
    json_resp = response.json()
    dump = json.dumps(json_resp)
    data = json.loads(dump)
    return data['results'][0]['productId']


def checkFoil(foilFlag):
    if (foilFlag == "Yes"):
        return True
    else:
        return False


# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()
