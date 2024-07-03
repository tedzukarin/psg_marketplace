import time

import requests
import pandas as pd
import PySimpleGUI as sg

def get_card(token, updatedAt=None, nm_ID=None):
    header = {'Authorization': token}
    url = 'https://suppliers-api.wildberries.ru/content/v2/get/cards/list'
    se =         {
          "settings": {
          "cursor": {
              "updatedAt": updatedAt,
              "nmID": nm_ID,
              "limit": 100
          },
          "filter": {
            "withPhoto": -1
            }
          }
        }
    response = requests.post(json=se, url=url, headers=header)
    if response.status_code == 200:
        if 'updatedAt' in response.json()['cursor'].keys():
            updatedAt = response.json()['cursor']['updatedAt']
            nm_ID = response.json()['cursor']['nmID']
            return response.json(), updatedAt, nm_ID
        else:
            return response.json(), None, None

    else:
        sg.Print(response.status_code, response.text)
        time.sleep(5)
        return None


def change_size(art, cards, out_data):
    width, length, height = out_data[0], out_data[1], out_data[2]

    for i in cards:
        if i['vendorCode'] == str(art):
            new_art = i
            new_art['dimensions']['length'] = length
            new_art['dimensions']['width'] = width
            new_art['dimensions']['height'] = height
            return new_art
    return None


def post_size(token, data):
    header = {'Authorization': token}
    url = 'https://suppliers-api.wildberries.ru/content/v2/cards/update'
    response = requests.post(url=url, headers=header, json=data)
    return response

def art_nmID(date):
    art_nmID_set = {}
    date = pd.read_excel(date, converters={'Артикул продавца': str})
    for ind, row in date.iterrows():
        art_nmID_set[row['Артикул WB']] = row['Артикул продавца']
    return(art_nmID_set)

def get_num_sizes(report, nmID):
    sizes = [
        report['Ширина (фактические габариты)'][report['Номенклатура'] == nmID].tolist()[0],
        report['Высота (фактические габариты)'][report['Номенклатура'] == nmID].tolist()[0],
        report['Длина (фактические габариты)'][report['Номенклатура'] == nmID].tolist()[0]
    ]
    return sizes

if __name__ == '__main__':
    pass