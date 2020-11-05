import pygsheets
import configparser
import os
import sys
import re
import pandas as pd

valuecolumns = ['адрес_заведения', 'Название_заведения', 'product', 'Выберите_дату_доставки']


def make_cfg_template():
    config = configparser.ConfigParser()
    config.add_section("In")
    config.add_section("Out")
    config.set("In", "Table", '')
    config.set("In", "Sheet", '')
    config.set("Out", "Table", '')
    config.set("Out", "Sheet", '')
    with open('config.cfg', "w") as config_file:
        config.write(config_file)


def parse_items(in_str):
    lst = [item.strip() for item in in_str[0].split(';')]
    res = {}
    for item in lst:
        name = re.search(r'[\w ]* -', item)
        qty = re.search(r'[\d]*x', item)
        SKU = name.group(0)[:-2]
        quantity = int(qty.group(0)[:-1])
        res.update({SKU: quantity})
    return res


def save_frame(client, table, sheet, frame):
    sh = client.open(table)
    wks = sh.worksheet_by_title(sheet)
    headers = [''] + list(frame.columns)
    range = pygsheets.datarange.DataRange(start=(1, 1), end=(1, len(headers)), worksheet=wks)
    range.update_values(values=[headers])
    count = 2
    for index, row in frame.iterrows():
        write_range = [index] + list(row)
        range = pygsheets.datarange.DataRange(start=(count, 1), end=(count, len(headers)), worksheet=wks)
        range.update_values(values=[write_range])
        count += 1


if __name__ == '__main__':
    # если нет config, создается шаблон
    if not os.path.exists('config.cfg'):
        make_cfg_template()
        sys.exit(0)

    # если есть config, чтение параметров
    config = configparser.ConfigParser()
    config.read('config.cfg', encoding='windows-1251')
    in_table = config.get("In", "Table")
    in_sheet = config.get('In', 'Sheet')
    out_table = config.get("Out", "Table")
    out_sheet = config.get('Out', 'Sheet')

    # авторизация в GDisk
    try:
        gc = pygsheets.authorize()
        sh = gc.open(in_table)
        wks = sh.worksheet_by_title(in_sheet)
    except:
        print('Authorization unsuccessful')
        sys.exit(1)

    cols = wks.cols
    headers = wks.get_values((1, 1), (1, cols), returnas='matrix')
    headers = headers[0].copy()
    head_dict = {item: headers.index(item) for item in headers if item != ''}
    data = [wks.get_values((2, head_dict[valuecolumns[i]]+1), (wks.rows, head_dict[valuecolumns[i]]+1),
            returnas='matrix') for i in range(0, 3)]
    SKUs = []
    for item in data[2]:
        SKUs.append(parse_items(item))
    size = min(len(data[0]), len(data[1]), len(SKUs))
    res, new_headers, newSKUs = [], [], []
    for i in range(0, size):
        res.append((data[0][i][0] + ' / ' + data[1][i][0], SKUs[i]))
        new_headers.append(data[0][i][0] + ' / ' + data[1][i][0])
        newSKUs.append(SKUs[i])
    new_headers = sorted(list(set(new_headers)))
    new_SKUs = [list(item.keys()) for item in newSKUs]
    new_SKUs = sorted(list(set(sum(new_SKUs, []))))
    out_frame = pd.DataFrame(index=new_SKUs, columns=new_headers)
    out_frame.fillna(0, inplace=True)
    for item in res:
        place = item[0]
        for SKU in item[1].keys():
            out_frame.at[SKU, place] = out_frame.at[SKU, place] + item[1][SKU]
    out_frame['Итоговое значение'] = out_frame.sum(axis=1)
    save_frame(gc, out_table, out_sheet, out_frame)




