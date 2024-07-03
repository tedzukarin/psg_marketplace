import pandas as pd
import PySimpleGUI as sg
import matplotlib.pyplot as plt
import datetime

def open_database(file):
    main_df = pd.read_excel(file, converters={'Артикул продавца': str})
    min_date = main_df['День'].min()
    max_date = main_df['День'].max()
    return main_df, min_date, max_date

def get_unique_values(df, name_column, brand=None):
    if brand:

        uni_list = df[name_column][df['Бренд'].isin(brand)].dropna().unique()
    else:
        uni_list = df[name_column].dropna().unique()
    uni_list = uni_list.tolist()
    uni_list = sorted(uni_list, reverse=False)
    return uni_list

def connected_art_with_name_dont_using(file):
    brand_art_name_dict = dict()
    df = pd.read_excel(file)
    for ind, i in df.iterrows():
        nai = i['Наименование']
        color = i['Цвет']
        if type(color) == float or color == "0":
            color = ''
        size = i['Размер']
        if type(size) == float or size == "0":
            size = ''
        art = i['Артикул продавца']
        if not i['Бренд'] in brand_art_name_dict:
            brand_art_name_dict[i['Бренд']] = [[str(f'{nai} {color} {size}'), art]]
        else:
            brand_art_name_dict[i['Бренд']].append([str(f'{nai} {color} {size}'), art])
    return brand_art_name_dict

def connected_art_with_name(file):
    art_name_dict = dict()
    df = pd.read_excel(file)
    for ind, i in df.iterrows():
        brand = i['Бренд']
        nai = i['Наименование']
        color = i['Цвет']
        if type(color) == float or color == "0":
            color = ''
        size = i['Размер']
        if type(size) == float or size == "0":
            size = ''
        art = i['Артикул продавца']
        art_name_dict[art] = [brand, nai, color, size]
    return art_name_dict

def connected_skus_with_name(file):
    skus_name_dict = dict()
    df = pd.read_excel(file)
    for ind, i in df.iterrows():
        brand = i['Бренд']
        if len(brand) > 10:
            brand = brand[:11]
        nai = i['Наименование']
        color = i['Цвет']
        if type(color) == float or color == "0":
            color = ''
        size = i['Размер']
        if type(size) == float or size == "0":
            size = ''
        skus = i['Баркод товара']
        skus_name_dict[skus] = [brand, nai, color, size]
    return skus_name_dict
def make_axis(df, period, need_columns, item):
    art_or_brand_flag = False
    operation_flag = False
    for i in need_columns:
        if i == 'Артикул продавца' or i == 'Бренд':
            art_or_brand = i
            art_or_brand_flag = True

        if i in ['Выкупили, шт.', 'К перечислению за товар, руб.', 'Заказано, шт.', 'Сумма заказов минус комиссия WB, руб.']:
            operation = i
            operation_flag = True

    if not art_or_brand_flag:
        sg.Print(
            'Не найдена колонка бренда или артикула. Переименуйте нужную в формате: "Артикул продавца" или "Бренд"')
        return
    if not operation_flag:
        sg.Print('Не найдена колонка с операцией')
        return

    new_df = df[['День', art_or_brand, operation]]

    if art_or_brand == 'Бренд':
        new_df = new_df.groupby([art_or_brand, 'День']).sum().reset_index()
        new_df = new_df[new_df[art_or_brand] == item]

    elif art_or_brand == 'Артикул продавца':
        new_df = new_df[new_df[art_or_brand] == item]

    new_df = new_df[['День', operation]]
    new_df = new_df[(period[0] <= new_df['День']) & (new_df['День'] <= period[1])]
    min_day = new_df['День'].min()
    max_day = new_df['День'].max()
    day_axis = pd.date_range(min_day, max_day)
    return (new_df, item), day_axis

def make_graph(df, days_axis, art_name=False, sum_flag=False):
    fig = plt.figure(figsize=(15, 8))
    axes = fig.add_axes([0.05, 0.1, 0.8, 0.9])
    n = 0
    x = days_axis.strftime('%d.%m').tolist()
    sell_mean = []
    sell_sum = []
    all_y = []
    items = []
    for d, item in df:
        y = []
        current = 0
        for day in x:
            for ind, row in d.iterrows():
                if row[0].strftime('%d.%m') == day:
                    current = row[1]
                    break
                else:
                    current = 0
            y.append(current)
        sell_sum.append([item, round(sum(y), 2)])
        sell_mean.append([item, round(sum(y)/len(y), 2)])
        if not sum_flag:
            if art_name:
                axes.plot(x, y, label=f'{item} - {" ".join(art_name[item])}')
            else:
                axes.plot(x, y, label=item)
        else:
            all_y.append(y)
            items.append(item)
        n += 1

    if sum_flag:
        all_y = list(map(sum, zip(*all_y)))
        sell_sum.append(['Общая сумма', round(sum(all_y), 2)])
        sell_mean.append(['Общее среднее', round(sum(all_y)/len(all_y), 2)])
        axes.plot(x, all_y, label=f'{[str(i) for i in items]}')

    m = 'Седнее:'
    for i in sell_mean:
        m += '\n'
        m += i[0]
        m += ' - '
        m += str(i[1])

    s = 'Сумма:'
    for i in sell_sum:
        s += '\n'
        s += i[0]
        s += ' - '
        s += str(i[1])

    plt.text(1, 0.6, m, transform=axes.transAxes,
             bbox={"fill": True, "facecolor": "white", "linestyle": "dotted", "linewidth": 2.0})
    plt.text(1, 0.3, s, transform=axes.transAxes,
             bbox={"fill": True, "facecolor": "white", "linestyle": "dotted", "linewidth": 2.0})
    plt.xticks(rotation=50, fontsize = 8)
    axes.legend(loc='upper left')
    axes.grid(linestyle='--')
    axes.set_xlabel('Период')
    axes.set_ylabel(f'{d.columns[1]}')

    plt.show()


def check_main_file_on_limits(file, limit, bigger_flag, art_name_dict, period, rubli_shtuki_flag, file_with_all_items):
    main_df = pd.read_excel(file, converters={'Артикул продавца': str})
    main_df = main_df[(period[0] <= main_df['День']) & (main_df['День'] <= period[1])]
    if rubli_shtuki_flag == 'Рубли': itog = 'Сумма продаж'
    else: itog = 'Всего штук'

    if rubli_shtuki_flag == 'Рубли': raschet_column = 'Сумма заказов минус комиссия WB, руб.'
    else: raschet_column = 'Заказано, шт.'

    all_items_df = pd.read_excel(file_with_all_items, converters={'Артикул продавца': str})
    all_items = all_items_df['Артикул продавца'].unique()



    unique_name = get_unique_values(main_df, 'Артикул продавца')
    new_df_column = ['Артикул', 'Наименование', itog]
    new_df = []
    sg.Print('Товары отвечающие заданному лимиту:')
    for i in unique_name:
        if i in art_name_dict.keys():
            name = art_name_dict[i]
        else:
            name = ''

        summ = main_df[raschet_column][main_df['Артикул продавца'] == i].sum()
        summ = round(summ, 1)
        if not bigger_flag:
            if summ > limit:
                new_df.append([i, ' '.join(name), summ])
                sg.Print(new_df[-1])
        else:
            if summ < limit:
                new_df.append([i, ' '.join(name), summ])
                sg.Print(new_df[-1])
        if i not in all_items:
            new_df.append([i, ' '.join(name), 0])
    new_df.sort(key=lambda x: (x[2], x[1]))
    if bigger_flag:
        compare = 'меньше'
    else:
        compare = 'больше'
    new_df_ = pd.DataFrame(columns=new_df_column, data=new_df, index=None)
    new_df_.to_excel(f'{period[0].date()}-{period[1].date()} {compare} {limit}.xlsx', index=False)

if __name__ == '__main__':
    fi, min_day, max_day = open_database('report_2023_6_27.xlsx.XLSX')
    need_columns = ['Артикул продавца', 'Сумма заказов минус комиссия WB, руб.']
    period = [datetime.datetime.strptime('11.05.2023', '%d.%m.%Y'), datetime.datetime.strptime('16.06.2023', '%d.%m.%Y')]
    df1, day_axis = make_axis(fi, period, need_columns, 'УТ000009928')
    df2, day_axis2 = make_axis(fi, period, need_columns, '00000020691')
    df = []
    df.append(df1)
    df.append(df2)
    make_graph(df, day_axis)

