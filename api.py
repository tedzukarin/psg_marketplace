import requests
import analitics
import PySimpleGUI as sg
import datetime
from base64 import b64decode
import json

def get_orders(file, fixed_time):
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    url = 'https://suppliers-api.wildberries.ru/api/v3/orders/new'
    sg.Print('Запрос получения заказов')
    response = requests.get(url=url, headers=header)
    sg.Print('Заказы получены')
    resp_json = response.json()
    orders_dict = resp_json['orders']
    orders = {}
    for i in orders_dict:
        orders[i['id']] = [(i['skus'][0]), i['createdAt']]
    skus_name_dict = analitics.connected_skus_with_name(file)
    id_name_list = []
    if type(fixed_time) == str:
        try:
            fixed_time = requests.get(url='https://www.timeapi.io/api/Time/current/zone?timeZone=Europe/Amsterdam')
            fixed_time = fixed_time.json()['dateTime'] + '+00:00'
            fixed_time = datetime.datetime.fromisoformat(fixed_time)
            sg.Print('Текущее время получено')
        except:
            sg.Print('Не удалось получить текущее время с сервера timeapi')
            return

    for key in orders.keys():
        orders_time = datetime.datetime.fromisoformat(orders[key][1])
        if orders_time <= fixed_time:
            if orders[key][0] in skus_name_dict.keys():
                if len(skus_name_dict[orders[key][0]][0]) > 10:
                    skus_name_dict[orders[key][0]][0] = skus_name_dict[orders[key][0]][0][:11]
                id_name_list.append([key, skus_name_dict[orders[key][0]], orders[key][1]])
            else:
                flag = False
                for key_1 in skus_name_dict.keys():
                    if orders[key][0] in key_1:
                        if len(skus_name_dict[key_1][0]) > 10:
                            skus_name_dict[key_1][0] = skus_name_dict[key_1][0][:11]
                        id_name_list.append([key, skus_name_dict[key_1], orders[key][1]])
                        flag = True
                        break
                if not flag:
                    id_name_list.append([key, ['', '', '', ''], orders[key][1]])
    id_name_list.sort(key=lambda x: (x[1][0], x[1][1], x[1][2], x[1][3]))
    return id_name_list

def get_supplies():
    supplies_list = []
    next_step = 0
    len_json_resp_suppl = 1
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    while len_json_resp_suppl > 0:
        params = {'limit': '1000', 'next': f'{next_step}'}
        url = 'https://suppliers-api.wildberries.ru/api/v3/supplies'
        sg.Print('Get-запрос для получения пакета отгрузок')
        response = requests.get(url=url, headers=header, params=params)
        sg.Print('Отгрузки получены')
        resp_json = response.json()
        next_step = resp_json['next']
        len_json_resp_suppl = len(resp_json['supplies'])
        for i in resp_json['supplies']:
            if not i['done']:
                supplies_list.append([i['id'], i['name']])
    return supplies_list

def make_supply(name):
    url = 'https://suppliers-api.wildberries.ru/api/v3/supplies'
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    resp = requests.post(url=url, headers=header, data={'name': name})
    return resp.json()

def add_order_into_supply(orderId, supplyId):
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    url = f'https://suppliers-api.wildberries.ru/api/v3/supplies/{supplyId}/orders/{orderId}'
    resp = requests.patch(url=url, headers=header)
    return resp

def get_info_supply(supplyId):
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    url = f'https://suppliers-api.wildberries.ru/api/v3/supplies/{supplyId}/orders'
    resp = requests.get(url=url, headers=header)
    return resp

def get_orders_into_supply(supplyID):
    orders_id_list = [[i['id'], i['nmId'], i['article'], i['skus'][0]] for i in get_info_supply(supplyID).json()['orders']]
    return orders_id_list

def get_labels_of_order(order_info):
    standart_token = open('standart_token.txt').read()
    header = {'Authorization': standart_token}
    url = 'https://suppliers-api.wildberries.ru/api/v3/orders/stickers'
    query = {'type': 'png', 'width': 40, 'height': 30}
    data = {'orders': [order_info[0]]}
    response = requests.post(url=url, json=data, headers=header, params=query)
    file_64 = response.json()['stickers'][0]['file']
    image_file = b64decode(file_64)
    return image_file

def get_price_wb():
    price_token = open('price_token.txt').read()
    header = {'Authorization': price_token}
    url = 'https://discounts-prices-api.wb.ru/api/v2/list/goods/filter'
    params = {'limit': 1000}
    response = requests.get(url=url, headers=header, params=params)
    resp_json = response.json()
    resp_listgood = resp_json['data']['listGoods']
    return resp_listgood

def new_post_price_and_discount(price_list, token):
    header = {'Authorization': token}
    url = 'https://discounts-prices-api.wb.ru/api/v2/upload/task'
    response = requests.post(url=url, headers=header, json={'data': price_list})
    return response

def update_ozon_price(product_id: int, old_price: str, price: str, token: str):
    seller_id = '468742'
    url = 'https://api-seller.ozon.ru/v1/product/import/prices'
    data = {'prices': [{
        'auto_action_enabled': 'DISABLED',
        'currency_code': 'RUB',
        'min_price': price,
        'old_price': old_price,
        'price': price,
        'price_strategy_enabled': 'DISABLED',
        'product_id': product_id
    }]}
    response = requests.post(url=url, headers={'Client-Id': seller_id, 'Api-Key': token}, json=data)
    return response

def update_yandex_price(token, bussines_id, article, price: int, fake_price: int, cofinance_price: int):
    url = f'https://api.partner.market.yandex.ru/businesses/{bussines_id}/offer-prices/updates'
    headers = {'Authorization': f'Bearer {token}'}
    data = {"offers": [{
      "offerId": article,
      "price": {
        "value": price,
        "currencyId": "RUR",
        "discountBase": fake_price
      }}]}
    response = requests.post(url=url, headers=headers, json=data)

    '''Изменение цен с софинансированной скидкой'''
    url_cofinance = f'https://api.partner.market.yandex.ru/businesses/{bussines_id}/offer-mappings/update'
    data_cofinance = {
    "offerMappings": [{"offer": {
                "offerId": article,
                "cofinancePrice": {
                    "value": cofinance_price,
                    "currencyId": "RUR"}
            }}]}
    response_cofinance = requests.post(url=url_cofinance, json=data_cofinance, headers=headers)
    return response, response_cofinance

def send_ostatki_wb(items_ostatki_list):
    token = open('standart_token.txt').read()
    items_ostatki_list = items_ostatki_list[1:]
    items_ostatki_ = [{'sku': i[0], 'amount': i[1]} for i in items_ostatki_list]

    warehouseId = 566230
    send_ost = {
        'stocks': items_ostatki_
    }
    header = {'Authorization': token}
    url = f'https://suppliers-api.wildberries.ru/api/v3/stocks/{warehouseId}'
    response = requests.put(url=url, headers=header, json=send_ost)
    if response.status_code == 204:
        sg.Print('Остатки WB обновлены')
    else:
        sg.Print(f'Обновить остатки WB е удалось.\nКод ошибки: {response.status_code}\nТекст ошибки: {response.text}\n')

def send_sb_price_api(api_list, token):
    url = 'https://api.megamarket.tech/api/merchantIntegration/v1/offerService/manualPrice/save'
    prices = [{'offerId': str(i[0]), 'price': int(i[1]), 'isDeleted': True} for i in api_list]
    data = {
        'meta': {},
        'data': {
            'token': token,
            'prices': prices
        }
    }
    response = requests.post(url, json=data)
    return response

def update_stocks_ozon(stocks_list):
    token = open('ozon_ostatki_token.txt').read()
    client_id = '468742'
    header = {'Client-Id': client_id,
             'Api-Key': token}
    url = 'https://api-seller.ozon.ru/v2/products/stocks'
    response = requests.post(url, headers=header, json={'stocks': stocks_list})
    errors = [{'art': i['offer_id'], 'errors': i['errors']} for i in response.json()['result'] if len(i['errors']) > 0]
    return [response.status_code, errors]


def send_ostatki_sb(sb_ost, token):
    url = 'https://api.megamarket.tech/api/merchantIntegration/v1/offerService/stock/update'
    stocks = [{'offerId': str(i[0]), 'quantity': int(i[1]), 'isDeleted': False} for i in sb_ost]
    data = {
        'meta': {},
        'data': {
            'token': token,
            'stocks': stocks
        }
    }
    response = requests.post(url, json=data)
    return response


def update_yandex_stocks(token, stock_list):
    headers = {'Authorization': f'Bearer {token}'}
    url = 'https://api.partner.market.yandex.ru/campaigns'
    resp = requests.get(url=url, headers=headers)
    campaignId = resp.json()['campaigns'][0]['id']
    stock_list = stock_list[2:]
    json = {'skus': [{'sku': f'{i[2]}', 'items': [{'count': i[4]}]}] for i in stock_list}
    url = f'https://api.partner.market.yandex.ru/campaigns/{campaignId}/offers/stocks'

    response = requests.put(url=url, headers=headers, json=json)
    return response.status_code, response.json()


if __name__ == '__main__':
    token = 'y0_AgAAAABj80bLAAtJrwAAAAD69ifVAABm8bv0jatHKLwtONPGv-UHRyOd1w'
    client_id = '16343256'
    resp = update_yandex_price(token, client_id, '00000018365', 370, 650)
    print(resp)

    pass
