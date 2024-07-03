import datetime
import pandas as pd
import re
import json
import PySimpleGUI as sg


def get_prod(file):
    prod_set = open('settings_ost_prod.ini', 'r')
    prod_set = json.load(prod_set)

    prod_skip_row = int(prod_set['prod_skip_row'])
    data_column = int(prod_set['data_column'])
    cur_column = []

    if prod_set['data_']:
        data_ = int(prod_set['data_'])
    else:
        data_ = False

    if prod_set['nai']:
        nai = int(prod_set['nai'])
        cur_column.append(nai)
    else:
        sg.Print('Укажите колонку с наименованием')
        return

    if prod_set['art']:
        art = int(prod_set['art'])
        cur_column.append(art)
    else:
        sg.Print('Укажите колонку с артикулами')
        return

    if prod_set['bar']:
        bar = int(prod_set['bar'])
        cur_column.append(bar)

    priority_column = prod_set['priority_column']
    priority_column = priority_column.split(',')
    for i in range(len(priority_column)):
        if priority_column[i].isdigit():
            priority_column[i] = int(priority_column[i])-1
    if prod_set['prod']:
        prod = int(prod_set['prod'])
        cur_column.append(prod)
    else:
        sg.Print('Укажите колонку с продажами')
        return
    delim = False
    if prod_set['delim']:
        delim = prod_set['delim']

    fas = prod_set['fas']

    ostatki_xlsx = pd.read_excel(file)
    ostatki = []
    ost_with_fas = []
    n = 0

    if data_:
        ito = data_
    else:
        data_str = ostatki_xlsx.columns[data_column-1]
        patt_st = r'(?<=с )\d\d\.\d\d\.\d\d\d\d'
        patt_fin = r'(?<=по )\d\d\.\d\d\.\d\d\d\d'
        match_st = re.search(patt_st, data_str)
        if not match_st:
            sg.Print('Не удалось считать дату. Проверьте файл или задайте период вручную')
            return
        st_date = datetime.datetime.strptime(match_st[0], "%d.%m.%Y")
        match_fin = re.search(patt_fin, data_str)
        fin_date = datetime.datetime.strptime(match_fin[0], "%d.%m.%Y")
        ito = (fin_date-st_date).days
    for i, row in ostatki_xlsx.iterrows():
        n += 1
        if n > prod_skip_row:
            ostatki.append([row[j-1] for j in cur_column])
    ostatki = pd.DataFrame(ostatki, index=None)
    ostatki = ostatki.groupby(priority_column).sum().reset_index()
    ostatki = ostatki.values.tolist()

    if fas and delim:
        patt = r'(?<=\()\d+\D*(?=\))'
        for i in ostatki:
            art = i[1].split(delim)[0]
            i[1] = art
            match = re.search(patt, i[0])
            if match:
                kol = ''
                for n in match[0]:
                    if n.isdigit():
                        kol += n
            else:
                kol = 1
            i.append(int(kol))
            ost_with_fas.append(i)

    elif not delim and fas:
        sg.Print('Укажите специальный разделитель, что бы получить фасовку из артикула')

    ost_new = []
    if fas:
        for ind in range(1, len(ost_with_fas)):
            ost_with_fas[ind-1][-2] = ost_with_fas[ind-1][-2] * ost_with_fas[ind-1][-1]
            if ost_with_fas[ind-1][1] == ost_with_fas[ind-2][1]:
                ost_with_fas[ind-1][-2] += ost_with_fas[ind-2][-2]
                ost_new.pop(-1)
                ost_new.append(ost_with_fas[ind-1])
            else:
                ost_new.append(ost_with_fas[ind-1])


    return ost_new, ito


def get_ost(file, art, kol, skip_row = 0):
    ostatki_xlsx = pd.read_excel(file)
    art = int(art)
    kol = int(kol)
    ostatki = []
    n = 0
    if skip_row:
        skip_row = int(skip_row)

    for i, row in ostatki_xlsx.iterrows():
        n += 1
        if n > skip_row:
            ostatki.append([str(row[art-1]), row[kol-1]])

    return ostatki

def make_tab(prod_list, ost_list, days, out_name):
    itog_list = []
    prod_set = open('settings_ost_prod.ini', 'r')
    prod_set = json.load(prod_set)
    if prod_set['bar']:
        columns = ['Наименование', 'Артикул', 'штрихкод', f'Продажи за период: {days} дней', 'Текущие остатки',
               'На сколько дней хватит остатков']
    else:
        columns = ['Наименование', 'Артикул', f'Продажи за период: {days} дней', 'Текущие остатки',
               'На сколько дней хватит остатков']

    for i in prod_list:
        for j in ost_list:
            if i[1] == j[0]:
                if prod_set['bar']:
                    prodazhi = i[3]
                else:
                    prodazhi = i[2]
                if prodazhi == 0 or j[1] == 'nan' or j[1] == 0:
                    ost = 99999
                else:
                    ost = round((j[1])/(prodazhi/days))
                if prod_set['bar']:
                    new_row = [i[0], i[1], i[2], prodazhi, j[1], ost]
                else:
                    new_row = [i[0], i[1], prodazhi, j[1], ost]
                itog_list.append(new_row)

    itog = pd.DataFrame(itog_list, index=None, columns=columns)
    itog.to_excel(f'{out_name}.xlsx', index=False)
