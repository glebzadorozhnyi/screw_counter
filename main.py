import re


def make_list_of_df(data, dictionary):  # создаёт глобал df с крепежом если не пустые. возвращает список с названиями df
    data_list = []
    for items, values in dictionary.items():
        globals()[items] = data.query(
            'Наименование.str.lower().str.contains(@values) and ~Наименование.str.lower().str.contains("пакб")')
        if len(globals()[items]) > 0:
            data_list.append(items)
    return data_list


def M_17475(row):  # диаметр винта, да и гайки тоже
    temp = re.search('[МM]\d+\.*[56]*', row)
    anwser = temp.group()[1:]
    if anwser[0] > '2':
        anwser = anwser[0]
    return anwser


def screw_format():
    def L_17475(row):  # длина винта
        temp = re.search('[xх]\d+', row)
        return int(temp.group()[1:])

    def screw_by_gost(container, part0, part2, part3):
        global screws
        screw_gost = screws.query('Наименование.str.contains(@container)')
        check_sum1 = screw_gost['Количество'].sum()
        screws = screws[~screws.index.isin(screw_gost.index)]
        screw_gost['diameter'] = screw_gost['Наименование'].apply(M_17475)
        screw_gost['length'] = screw_gost['Наименование'].apply(L_17475)
        screw_grouped = screw_gost.groupby(['diameter', 'length'])['Количество'].sum().reset_index()
        screw_grouped['Наименование'] = part0 + screw_grouped['diameter'] + part2 + \
                                screw_grouped['length'].apply(str) + part3
        screw_grouped['Размер'] = 'M' + screw_grouped['diameter'] + 'x' + screw_grouped[
            'length'].apply(
            str)
        screw_grouped['ГОСТ/ОСТ'] = container
        check_sum2 = screw_grouped['Количество'].sum()
        screw_grouped['type'] = 'винт'
        screw_grouped.loc[screw_grouped['length'] == 0, 'Размер'] = 'мелкий шаг'
        if check_sum1 != check_sum2:
            raise ValueError('потеряли винт', container)
        return screw_grouped

    screws1745 = screw_by_gost('17475', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 17475-80')

    screws1491 = screw_by_gost('1491', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 1491-80')

    screws4762 = screw_by_gost('4762', 'Винт ГОСТ Р ИСО 4762-M', 'x', '.A2-70')

    screw1477 = screw_by_gost('1477', r'Винт A.M', r'-6gx', r'.22H.35.016 ГОСТ 1477-93')

    screw1476 = screw_by_gost('1476', r'Винт A.M', r'-6gx', r'.23.20Х13 ГОСТ 1476-93')

    screw2009 = screw_by_gost('2009', r'Винт с потайной головкой ГОСТ Р ИСО 2009 - М', 'x', r'.A2-70')

    screw11644 = screw_by_gost('11644', r'Винт А.М', r'-6gx', r'.23.20Х13.11 ГОСТ 11644-75')

    screw11074 = screw_by_gost('11074', r'Винт А.М', r'-6gx', r'.23.20Х13.11 ГОСТ 11074-93')

    screw1479 = screw_by_gost('1479', r'Винт A.M', r'-6gx', r'.23.20Х13 ГОСТ 1479-93')

    all_screws = pd.concat([screws1745, screws1491, screws4762, screw1477, screw1476, screw2009, screw11644, screw11074,
                            screw1479]).reset_index(drop=True)
    return all_screws


def nuts_format():
    def nuts_by_gost(container, part0, part2):
        global nuts
        nuts_gost = nuts.query('Наименование.str.contains(@container)')
        checs_sum1 = nuts_gost['Количество'].sum()
        nuts = nuts[~nuts.index.isin(nuts_gost.index)]
        nuts_gost['diameter'] = nuts_gost['Наименование'].apply(M_17475)
        nuts_grouped = nuts_gost.groupby(['diameter'])['Количество'].sum().reset_index()
        nuts_grouped['Наименование'] = part0 + nuts_grouped['diameter'] + part2
        nuts_grouped['Размер'] = 'M' + nuts_grouped['diameter']
        nuts_grouped['ГОСТ/ОСТ'] = container
        checs_sum2 = nuts_grouped['Количество'].sum()
        if checs_sum1 != checs_sum2:
            raise ValueError('потеряли гайку', container)
        nuts_grouped['type'] = 'гайка'
        return nuts_grouped

    nuts5927 = nuts_by_gost('5927', 'Гайка М', '-6H.5.20.016 ГОСТ 5927-70')
    all_nuts = nuts5927
    return all_nuts


def washer_format():
    def M_washer(row):
        temp = re.search('[AА]\.\d+\.*[56]*', row)
        anwser = temp.group()[2:]

        if anwser[0] > '2':
            anwser = anwser[0]
        if anwser.endswith('.'):
            anwser = anwser[0:-1]
        return anwser

    def M_washer6402(row):
        temp = re.search('йба \d+\.*[56]*', row)
        anwser = temp.group()[4:]
        if anwser[0] > '2':
            anwser = anwser[0]
        if anwser.endswith('.'):
            anwser = anwser[0:-1]
        return anwser

    def washer_by_gost(container, part0, part2, apply_function):
        global washer
        washer_gost = washer.query('Наименование.str.contains(@container)')
        check_sum1 = washer_gost['Количество'].sum()
        washer = washer[~washer.index.isin(washer_gost.index)]
        washer_gost['diameter'] = washer_gost['Наименование'].apply(apply_function)
        washer_grouped = washer_gost.groupby(['diameter'])['Количество'].sum().reset_index()
        washer_grouped['Наименование'] = part0 + washer_grouped['diameter'] + part2
        washer_grouped['Размер'] = 'M' + washer_grouped['diameter']
        washer_grouped['ГОСТ/ОСТ'] = container
        check_sum2 = washer_grouped['Количество'].sum()
        if check_sum1 != check_sum2:
            raise ValueError('потеряли шайбу', container)
        washer_grouped['type'] = 'шайба'
        return washer_grouped

    washer10450 = washer_by_gost('10450', 'Шайба А', '.01.019 ГОСТ 10405-78', M_washer)

    washer11371 = washer_by_gost('11371', 'Шайба А', '.01.019 ГОСТ 11371-78', M_washer)

    washer6402 = washer_by_gost('6402', 'Шайба А', '.01.019 ГОСТ 6402-70', M_washer6402)

    all_washer = pd.concat([washer10450, washer11371, washer6402]).reset_index(drop=True)
    return all_washer


def pins_format():
    def M(row):  # диаметр штифта
        temp = re.search('ифт \d+\.*[56]*', row)
        anwser = temp.group()[4:]
        if anwser[0] > '2':
            anwser = anwser[0]
        return anwser

    def L(row):  # длина штифта
        temp = re.search('[xх]\d+', row)
        return int(temp.group()[1:])

    def pins_by_gost(container, part0, part2, part3):
        global pins
        pins_gost = pins.query('Наименование.str.contains(@container)')
        check_sum1 = pins_gost['Количество'].sum()
        pins = pins[~pins.index.isin(pins_gost.index)]
        pins_gost['diameter'] = pins_gost['Наименование'].apply(M)
        pins_gost['length'] = pins_gost['Наименование'].apply(L)
        pins_grouped = pins_gost.groupby(['diameter', 'length'])['Количество'].sum().reset_index()
        pins_grouped['Наименование'] = part0 + pins_grouped['diameter'] + part2 + \
                               pins_grouped['length'].apply(str) + part3
        pins_grouped['Размер'] = 'M' + pins_grouped['diameter'] + 'x' + pins_grouped['length'].apply(str)
        pins_grouped['ГОСТ/ОСТ'] = container
        check_sum2 = pins_grouped['Количество'].sum()
        if check_sum1 != check_sum2:
            raise ValueError('потеряли штифт', container)
        pins_grouped['type'] = 'штифт'
        return pins_grouped

    pins3128 = pins_by_gost('3128', r'Штифт ', 'x', r'.Хим.Окс.прм. ГОСТ 3128-70')

    all_pins = pins3128
    return all_pins


def vint_translit_format():
    def M(row):  # диаметр штифта
        temp = re.search('M\d+[-_]*[56]*', row)
        anwser = temp.group()[1:]
        anwser = anwser.replace('_', '.')
        anwser = anwser.replace('-', '.')
        if anwser[0] > '2':
            anwser = anwser[0]
        return anwser

    def L(row):  # длина штифта
        temp = re.search('X\d+', row)
        return int(temp.group()[1:])

    def vint_by_gost(container, part0, part2, part3):
        global vint
        vint_gost = vint.query('Наименование.str.contains(@container)')
        check_sum1 = vint_gost['Количество'].sum()
        vint = vint[~vint.index.isin(vint_gost.index)]
        vint_gost['diameter'] = vint_gost['Наименование'].apply(M)
        vint_gost['length'] = vint_gost['Наименование'].apply(L)
        vint_grouped = vint_gost.groupby(['diameter', 'length'])['Количество'].sum().reset_index()
        vint_grouped['Наименование'] = part0 + vint_grouped['diameter'] + part2 + \
                               vint_grouped['length'].apply(str) + part3
        vint_grouped['Размер'] = 'M' + vint_grouped['diameter'] + 'x' + vint_grouped['length'].apply(str)
        vint_grouped['ГОСТ/ОСТ'] = container
        check_sum2 = vint_grouped['Количество'].sum()
        if check_sum1 != check_sum2:
            raise ValueError('потеряли vint', container)
        vint_grouped['type'] = 'винт'
        return vint_grouped

    vint4762 = vint_by_gost('4762', 'Винт ГОСТ Р ИСО 4762-M', 'x', '.A2-70')

    vint1491 = vint_by_gost('1491', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 1491-80')

    vint17475 = vint_by_gost('17475', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 17475-80')

    vint1477 = vint_by_gost('1477', r'Винт A.M', r'-6gx', r'.22H.35.016 ГОСТ 1477-93')

    vint1476 = vint_by_gost('1476', r'Винт A.M', r'-6gx', r'.23.20Х13 ГОСТ 1476-93')

    vint11738 = vint_by_gost('11738', r'Винт M', r'-6gx', r'.23.20Х13 ГОСТ 11738-84')

    all_vint = pd.concat([vint4762, vint1491, vint17475, vint1477, vint1476, vint11738]).reset_index(drop=True)

    return all_vint


import pandas as pd

filename = 'sop_oe.csv'
data = pd.read_csv(filename, encoding='ANSI', sep=';')
data.columns = ['Наименование', 'Количество']
dictionary = {'screws': "винт", 'nuts': "гайка", 'washer': "шайба", 'pins': "штифт", 'vint': "vint", 'gajka': "gajka",
              'shajba': "shajba", 'shtift': "shtift"}
data_list = make_list_of_df(data, dictionary)
all_screws = screw_format()
all_nuts = nuts_format()
print(len(all_nuts), 'all_nuts')
all_washer = washer_format()
print(len(all_washer), 'all_washer')
all_pins = pins_format()
print(len(all_pins), 'all_pins')
all_vint = vint_translit_format()
print(all_screws['Количество'].sum() + all_vint['Количество'].sum())
# объединяем латинские винты и обычные

all_screws_and_vints = pd.concat([all_screws, all_vint]).reset_index(drop=True)
duplicates = all_screws_and_vints.groupby('Наименование')['Количество'].sum()
all_screws_and_vints = all_screws_and_vints.merge(duplicates, on='Наименование', how='left')
all_screws_and_vints = all_screws_and_vints[['type', 'diameter', 'length', 'Размер', 'ГОСТ/ОСТ', 'Количество_y', 'Наименование']]
all_screws_and_vints.columns = ['type', 'diameter', 'length', 'Размер', 'ГОСТ/ОСТ', 'Количество', 'Наименование']
all_screws_and_vints = all_screws_and_vints.drop_duplicates('Наименование')
all_screws_and_vints = all_screws_and_vints.sort_values(by=['ГОСТ/ОСТ','diameter','length']).reset_index(drop=True)
print(all_screws_and_vints['Количество'].sum())
print(len(all_screws_and_vints), 'all_screws')
bad = pd.concat([screws, nuts, washer, pins, vint, gajka, shajba, shtift]).reset_index(drop=True)

fileout = filename[0:-4] + '_out.xlsx'
writer = pd.ExcelWriter(fileout, engine='xlsxwriter')
all_screws_and_vints[['Наименование', 'ГОСТ/ОСТ', 'Размер', 'Количество']].to_excel(writer, 'Винты')
all_nuts[['Наименование', 'ГОСТ/ОСТ', 'Размер', 'Количество']].to_excel(writer, 'Гайки')
all_washer[['Наименование', 'ГОСТ/ОСТ', 'Размер', 'Количество']].to_excel(writer, 'Шайбы')
all_pins[['Наименование', 'ГОСТ/ОСТ', 'Размер', 'Количество']].to_excel(writer, 'Штифты')
bad.to_excel(writer, 'Не распозналось')
for sheet_name in ['Винты', 'Гайки', 'Шайбы', 'Штифты', 'Не распозналось']:
    worksheet = writer.sheets[sheet_name]
    worksheet.set_column(1, 1, 50)
    worksheet.set_column(2, 4, 10)
   # worksheet.set_column(3, 3, 10)
    worksheet.set_default_row(20)
writer.save()
