import pandas as pd
import PySimpleGUI as sg
import api


def get_wb_price(tovar_file, usloviya_file, api_enabled, token):
    sg.Print('Изменяем цены WB')
    tovar_df = pd.read_excel(tovar_file, converters={'Артикул продавца WB': str, 'Артикул WB': str,
                                                     'Баркод товара WB': str})
    tovar_df['Дополнительные расходы'] = tovar_df['Дополнительные расходы'].fillna(0)
    uslovia_df = pd.read_excel(usloviya_file, skiprows=1)
    articles_wb = tovar_df['Артикул WB'].dropna().unique()
    columns_new_df = ['Бренд', 'Категория', 'Артикул WB', 'Артикул продавца', 'Последний баркод', 'Остатки WB',
                      'Остатки продавца', 'Оборачиваемость', 'Текущая цена', 'Новая цена', 'Текущая скидка',
                      'Новая скидка']
    new_list = []
    if_api_list = []
    for i in articles_wb:
        brand = tovar_df['Бренд WB'][tovar_df['Артикул WB'] == i].values[0]
        category = tovar_df['Предмет WB'][tovar_df['Артикул WB'] == i].values[0]
        article_wb = i
        article_seller = tovar_df['Артикул продавца WB'][tovar_df['Артикул WB'] == i].values[0]
        last_barcode = tovar_df['Баркод товара WB'][tovar_df['Артикул WB'] == i].values[0].split(';')[0]
        ostatki = 0
        ostatki_seller = 0
        obor = 0
        current_price = 0
        current_discount = 0
        new_discount = 50
        '''Рассчёт новой цены'''
        '''Проверим, есть ли принудительная цена'''
        price = tovar_df['Цена WB'][tovar_df['Артикул WB'] == i].values[0]
        if price == 0:
            '''0 ставится, если товар работает в боте по удержанию цены'''
            continue
        if pd.isna(price):
            '''Если цена не указана, высчитываем её'''
            comission = uslovia_df['Склад продавца - везу на склад WB, %'][uslovia_df['Предмет wb'] == category].values[0]
            comission = comission.replace(',', '.')
            comission = float(comission)
            volume = tovar_df['Ширина упаковки WB'][tovar_df['Артикул WB'] == i].values[0] *\
                     tovar_df['Длина упаковки WB'][tovar_df['Артикул WB'] == i].values[0] *\
                     tovar_df['Высота упаковки WB'][tovar_df['Артикул WB'] == i].values[0] /1000
            if volume <= 1:
                logistic = uslovia_df['логистика min объём wb'].unique()[0]
            else:
                '''Конструкция с целочисленным делением и двумя минусами позволяет провести округление в большую сторону'''
                number_of_steps = -(-(volume - 1) // uslovia_df['шаг объёма wb'].unique()[0])
                logistic = uslovia_df['логистика min объём wb'].unique()[0] + number_of_steps * \
                           uslovia_df['стоимость шага wb'].unique()[0]
            zakupka = tovar_df['Закупка'][tovar_df['Артикул WB'] == i].values[0]
            dop_rash = tovar_df['Дополнительные расходы'][tovar_df['Артикул WB'] == i].values[0]
            fasovka = tovar_df['Фасовка WB'][tovar_df['Артикул WB'] == i].values[0]
            marzha = tovar_df['Маржинальность WB'][tovar_df['Артикул WB'] == i].values[0]
            percent_na_vozvraty = uslovia_df['% на возвраты'].unique()[0]
            price = ((zakupka * fasovka * (1 + marzha / 100)) + logistic + dop_rash) / (1-comission/100) * (1 + percent_na_vozvraty/100)
        price_with_fake_discount = price / 0.5

        new_list.append([brand, category, article_wb, article_seller, last_barcode, ostatki, ostatki_seller, obor,
                         current_price, price_with_fake_discount, current_discount, new_discount])

        if api_enabled:
            if_api_list.append({'nmID': int(article_wb), 'price': round(price_with_fake_discount), 'discount': 50})
    if api_enabled:
        ppad = api.new_post_price_and_discount(if_api_list, token)
        if ppad.status_code == 200:
            sg.Print('Цены на WB успешно изменены')
        else:
            sg.Print(f'Что-то пошло не так. Код ответа от сервера:{ppad, ppad.json()}')


    new_df = pd.DataFrame(columns=columns_new_df, data=new_list)
    new_df.to_excel('WB_new_price.xlsx', sheet_name='Отчет - цены и скидки на товары', index=False)


def get_ozon_price(tovar_file, usloviya_file, api_enabled, token):
    sg.Print('Изменяем цены OZON')
    tovar_df = pd.read_excel(tovar_file, converters={'Артикул продавца WB': str, 'Артикул WB': str})
    uslovia_df = pd.read_excel(usloviya_file, skiprows=1)
    tovar_df['Дополнительные расходы'] = tovar_df['Дополнительные расходы'].fillna(0)

    articles_oz = tovar_df['Артикул продавца OZON'].dropna().unique()

    new_list = []
    for i in articles_oz:
        article = i
        oz_sku = tovar_df['FBS OZON SKU ID'][tovar_df['Артикул продавца OZON'] == i].values[0]
        barcode = tovar_df['Barcode OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]

        '''Рассчёт новой цены'''
        '''Проверим, есть ли принудительная цена'''
        price = tovar_df['Цена OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]

        if pd.isna(price):
            category = tovar_df['Категория OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]
            comission = uslovia_df['Вознаграждение на FBS ozon'][
                uslovia_df['Категория товаров ozon'] == category].values[0] * 100
            value_mass = tovar_df['Объемный вес, кг OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]
            if value_mass <= uslovia_df['min объём ozon'].values[0]:
                logistic = uslovia_df['логистика min объём ozon'].values[0]
            else:
                '''Конструкция с целочисленным делением и двумя минусами позволяет провести округление в большую сторону'''
                number_of_steps = -(-(value_mass - 1) // uslovia_df['шаг объёма ozon'].unique()[0])
                logistic = uslovia_df['логистика min объём ozon'].unique()[0] + number_of_steps * \
                           uslovia_df['стоимость шага ozon'].unique()[0]
            last_mile = uslovia_df['последняя миля ozon'].values[0] * 100
            equairing = uslovia_df['эквайринг ozon'].values[0] * 100
            priemka = uslovia_df['Приём товара ozon'].values[0]
            zakupka = tovar_df['Закупка'][tovar_df['Артикул продавца OZON'] == i].values[0]
            fasovka = tovar_df['Фасовка OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]
            dop_rash = tovar_df['Дополнительные расходы'][tovar_df['Артикул продавца OZON'] == i].values[0]
            marzha = tovar_df['Маржинальность OZON'][tovar_df['Артикул продавца OZON'] == i].values[0]
            percent_na_vozvraty = uslovia_df['% на возвраты'].unique()[0]
            zatraty = logistic + priemka + (zakupka * fasovka) * (1 + marzha/100) + dop_rash
            percent_of_zatraty = last_mile + comission + equairing + percent_na_vozvraty
            temp_cost = round(zatraty / (1 - percent_of_zatraty / 100))
            if temp_cost / 100 * 5.5 > uslovia_df['макс последней мили ozon'].values[0]:
                last_mile = uslovia_df['макс последней мили ozon'].values[0]
                zatraty = logistic + last_mile + priemka + (zakupka * fasovka) * (1 + marzha / 100) + dop_rash
                percent_of_zatraty = comission + equairing + percent_na_vozvraty
                price = round(zatraty / (1 - percent_of_zatraty / 100))
            else:
                price = temp_cost
        fake_price = round(price / 0.7)
        if api_enabled:
            ozon_product_id = tovar_df['Ozon Product ID'][tovar_df['Артикул продавца OZON'] == article].values[0]
            uop = api.update_ozon_price(int(ozon_product_id), str(fake_price), str(price), token)
            if uop.status_code == 200:
                sg.Print(f'{article} - новая цена: {price}')
            else:
                sg.Print(f'Не удалось обновить цену: {article}')

        new_row = [article, oz_sku, None,  None, None, None, None, barcode, None, None, None, None, None, None, None,
                   None,  None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                   None,  None,  None, None, None, None, None, None, fake_price, price, None, None, None, price, None]
        new_list.append(new_row)
    new_df = pd.DataFrame(data=new_list)
    new_df.to_excel('OZ_new_price_копировать в исходный документ.xlsx', index=False)

def get_yandex_price(tovar_file, usloviya_file, api_enabled, token, bussines_id):
    sg.Print('Изменяем цены Яндекс')
    tovar_df = pd.read_excel(tovar_file, converters={'Артикул продавца WB': str, 'Артикул WB': str})
    uslovia_df = pd.read_excel(usloviya_file, skiprows=1, decimal=',')
    tovar_df['Дополнительные расходы'] = tovar_df['Дополнительные расходы'].fillna(0)

    articles_ya = tovar_df['Артикул Яндекс'].dropna().unique()

    new_list = []
    values_gabarity = uslovia_df['Объемный вес или масса, кг яндекс'].dropna().values
    cost_gabarity = uslovia_df['Стоимость услуги яндекс'].dropna().values

    for i in articles_ya:
        article = i
        '''Проверим, есть ли принудительная цена'''
        minimal_price = tovar_df['Цена Яндекс'][tovar_df['Артикул Яндекс'] == i].values[0]
        if pd.isna(minimal_price):
            weight = tovar_df['Вес Яндекс'][tovar_df['Артикул Яндекс'] == article].values[0]
            volume = tovar_df['Габариты Яндекс'][tovar_df['Артикул Яндекс'] == article].values[0].split('/')
            volume_weight = float(volume[0]) * float(volume[1]) * float(volume[2]) / 5000
            if volume_weight > weight:
                gabarity = volume_weight
            else:
                gabarity = weight
            ind = 0
            logistic = values_gabarity[0]
            while gabarity > float(values_gabarity[ind]):
                if ind == len(values_gabarity):
                    logistic = uslovia_df['Объемный вес или масса, кг max  яндекс'].values[0]
                    break
                logistic = int(cost_gabarity[ind + 1])
                ind += 1

            category = tovar_df['Категория  Яндекс'][tovar_df['Артикул Яндекс'] == article].values[0].split('\\')
            category = [category[0], category[-1]]
            uslovia_df['Тарифы FBS, Экспресс яндекс'] = uslovia_df['Тарифы FBS, Экспресс яндекс'].astype(str)
            uslovia_df['Тарифы FBS, Экспресс яндекс'] = uslovia_df['Тарифы FBS, Экспресс яндекс'].str.replace(
                '%', '', regex=False).astype(float)
            comission = uslovia_df['Тарифы FBS, Экспресс яндекс'][
                (uslovia_df['Родительская категория яндекс'] == category[0]) &
                (uslovia_df['Категория яндекс'].str.contains(category[1]))
            ].values[0]
            if comission < 1:
                comission *= 100

            dost = uslovia_df['Доставка яндекс'].values[0] * 100
            priemka = uslovia_df['Приём товара яндекс'].values[0]
            equairing = uslovia_df['Эквайринг яндекс'].values[0] * 100
            zakupka = tovar_df['Закупка'][tovar_df['Артикул Яндекс'] == i].values[0]
            dop_rash = tovar_df['Дополнительные расходы'][tovar_df['Артикул Яндекс'] == i].values[0]
            fasovka = tovar_df['Фасовка Яндекс'][tovar_df['Артикул Яндекс'] == i].values[0]
            marzha = tovar_df['Маржинальность Яндекс'][tovar_df['Артикул Яндекс'] == i].values[0]
            percent_na_vozvraty = uslovia_df['% на возвраты'].unique()[0]
            zatraty = (zakupka * fasovka) * (1 + marzha/100) + logistic + priemka + dop_rash
            percent_of_zatraty = percent_na_vozvraty + equairing + comission + dost
            temp_price = round(zatraty / (1 - percent_of_zatraty / 100))
            if temp_price * uslovia_df['Доставка яндекс'].values[0] > uslovia_df['Доставка max яндекс'].values[0]:
                dost = uslovia_df['Доставка max яндекс'].values[0]
                zatraty = (zakupka * fasovka) * (1 + marzha / 100) + logistic + priemka + dost + dop_rash
                percent_of_zatraty = percent_na_vozvraty + equairing + comission
                minimal_price = round(zatraty / (1 - percent_of_zatraty / 100))
            elif temp_price * uslovia_df['Доставка яндекс'].values[0] < uslovia_df['Доставка min яндекс'].values[0]:
                dost = uslovia_df['Доставка min яндекс'].values[0]
                zatraty = (zakupka * fasovka) * (1 + marzha / 100) + logistic + priemka + dost + dop_rash
                percent_of_zatraty = percent_na_vozvraty + equairing + comission
                minimal_price = round(zatraty / (1 - percent_of_zatraty / 100))
            else:
                minimal_price = temp_price
        price = round(minimal_price * 1.02)
        fake_price = round(price / 0.7)
        if api_enabled:
            uyp = api.update_yandex_price(token, bussines_id, article, price=price, fake_price=fake_price,
                                    cofinance_price=minimal_price)[0]
            if uyp.status_code == 200:
                sg.Print(f'{article} - новая цена: {price}')
            else:
                sg.Print(f'Не удалось обновить цену: {article}')
        new_row = [None, None, article, None, price, fake_price, None, None, minimal_price, None, None, None, None]
        new_list.append(new_row)
    new_df = pd.DataFrame(data=new_list)
    new_df.to_excel('YA_new_price_копировать в исходный документ.xlsx', index=False)


def get_sber_price(tovar_file, usloviya_file, api_enabled):
    sg.Print('Изменяем цены Сбера')
    try:
        token = open('sber_token.txt').read()
    except:
        sg.Print('Токен не найден')
        api_enabled = False
    tovar_df = pd.read_excel(tovar_file, converters={'Артикул продавца WB': str, 'Артикул WB': str})
    uslovia_df = pd.read_excel(usloviya_file, skiprows=1, decimal=',')
    tovar_df['Дополнительные расходы'] = tovar_df['Дополнительные расходы'].fillna(0)

    articles_sb = tovar_df['Вендор код (артикул производителя)'].dropna().unique()
    columns = ['id', 'Доступность товара', 'Категория', 'Производитель (Бренд)',' Артикул', 'Модель', 'Название',
                   'Цена(руб)', 'Старая цена(руб)', 'Остаток', 'НДС', 'Штрихкод', 'Ссылка на картинку',' Описание',
                   'Ссылка на товар на сайте магазина', 'Время заказа До', 'Дней на отгрузку']
    new_list = [['offer_id', 'available', 'category', 'vendor', 'vendor_code', 'model', 'name', 'price', 'old_price',
                'instock',	'vat', 'barcode', 'picture','description', 'url', 'order-before', 'days']]
    values_gabarity = uslovia_df['объём, сбер'].dropna().values
    cost_gabarity = uslovia_df['логистика, сбер'].dropna().values
    api_list = []
    for i in articles_sb:
        article = i
        '''Проверим, есть ли принудительная цена'''
        minimal_price = tovar_df['Цена Сбер'][tovar_df['Вендор код (артикул производителя)'] == i].values[0]

        if pd.isna(minimal_price):
            gabarity = tovar_df['Длина × ширина × высота, см'][tovar_df['Вендор код (артикул производителя)'] == article].values[0].split(' x ')
            gabarity = (float(gabarity[0]) * float(gabarity[1]) * float(gabarity[2])) / 1000
            ind = 0
            logistic = int(cost_gabarity[ind])
            while gabarity > float(values_gabarity[ind]):
                if ind == len(values_gabarity):
                    logistic = int(cost_gabarity[ind])
                    break
                logistic = int(cost_gabarity[ind + 1])
                ind += 1


            category = tovar_df['Категория Мегамаркет'][tovar_df['Вендор код (артикул производителя)'] == article].values[0].split('\\')
            category = category[-1]

            comission = uslovia_df['Тариф, сбер'][uslovia_df['Товарная категория, сбер'] == category].values[0]

            dost = uslovia_df['доставка, сбер'].values[0] * 100
            priemka = uslovia_df['приём товра, сбер'].values[0]
            equairing = uslovia_df['экайринг, сбер'].values[0] * 100
            zakupka = tovar_df['Закупка'][tovar_df['Вендор код (артикул производителя)'] == i].values[0]
            dop_rash = tovar_df['Дополнительные расходы'][tovar_df['Вендор код (артикул производителя)'] == i].values[0]
            fasovka = tovar_df['Фасовка сбер'][tovar_df['Вендор код (артикул производителя)'] == i].values[0]
            marzha = tovar_df['Маржинальность Сбер'][tovar_df['Вендор код (артикул производителя)'] == i].values[0]
            percent_na_vozvraty = uslovia_df['% на возвраты'].unique()[0]
            zatraty = (zakupka * fasovka) * (1 + marzha/100) + logistic + priemka + dop_rash
            percent_of_zatraty = percent_na_vozvraty + equairing + comission + dost
            temp_price = round(zatraty / (1 - percent_of_zatraty / 100))
            if temp_price * uslovia_df['доставка, сбер'].values[0] > uslovia_df['доставка max, сбер'].values[0]:
                dost = uslovia_df['доставка max, сбер'].values[0]
                zatraty = (zakupka * fasovka) * (1 + marzha / 100) + logistic + priemka + dost + dop_rash
                percent_of_zatraty = percent_na_vozvraty + equairing + comission
                minimal_price = round(zatraty / (1 - percent_of_zatraty / 100))
            else:
                minimal_price = temp_price
        price = minimal_price
        fake_price = round(price / 0.7)

        if api_enabled:
            api_list.append([article, price])


        new_row = [article, None, None, None, article, None, None, price, fake_price, None, None, None, None, None, None, None, None]
        new_list.append(new_row)

    if api_enabled:
        ppad = api.send_sb_price_api(api_list, token)
        if ppad.json()['success'] == 1:
            sg.Print('Цены на SB успешно изменены')
        else:
            sg.Print(f'Что-то пошло не так. Код ответа от сервера:{ppad, ppad.json()}')

    new_df = pd.DataFrame(data=new_list, columns=columns)
    new_df.to_excel('SB_new_price_копировать в исходный документ.xlsx', index=False)



if __name__ == '__main__':
    get_yandex_price('товары.xlsx', 'условия.xlsx')

