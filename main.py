import re
import pandas as pd

def make_list_of_samples(filename):
    # читаем список шаблонов крепежа в df
    return pd.read_csv(filename, encoding='ANSI', sep=';')


def parsing_df_of_fasteners(filename, dictionary):
    # проходим весь входной csv файл
    # выбираем строки, в которых есть ключевые слова-type из dictionary и нет "пакб"
    # возвращаем df где есть только строки с крепежём
    data = pd.read_csv(filename, encoding='ANSI', sep=';')
    data.columns = ['Наименование', 'Кол.']
    list_of_df = []
    for type in dictionary:
        temp_df = data.query('Наименование.str.lower().str.contains(@type) and ~Наименование.str.lower().str.contains("пакб")')
        temp_df['type'] = type
        if len(temp_df) > 0:
            list_of_df.append(temp_df)
    return pd.concat([*list_of_df]).reset_index(drop=True)


def get_diameter(row, regular_expression, symbols_to_delete):
    # диаметр крепежа
    # regular_expression - регулярка для поиска диаметра крепежа
    # symbols_to_delete - количество первых символов, которые нужно удалять из найденной регулярки
    temp = re.search(regular_expression, row)
    anwser = temp.group()[symbols_to_delete:]
    if anwser[0] > '2':
        anwser = anwser[0]
    if anwser.endswith('.'):
        anwser = anwser[0:-1]
    return anwser


def get_length(row):
    # длина винта или штифта
    temp = re.search('[xх]\d+', row)
    return int(temp.group()[1:])


def normalization(list_of_samples, df):
    list_2d = list_of_samples.query('type == ("винт", "штифт", "болт")') # двумерный крепёж (диаметр + длина)
    list_1d = list_of_samples.query('type == ("гайка", "шайба")') # одномерный крепёж (диаметр)

    def normalization_2d(list_2d, df):
        list_of_df = []  # массив из разобранных df относящихся к разным гостам крепежа
        for row in list_2d.itertuples(index=True):
            temp_df = df.query('Наименование.str.contains(@row.container)') # временный df с крепежём госта container
            check_sum1 = temp_df['Кол.'].sum() # пересчитаем весь крепеж этого госта, чтобы потом убедиться, что ничего не потеряли
            df = df[~df.index.isin(temp_df.index)] # выкидываем крепёж из исходного df, чтобы потом сделать список из неразобранного крепежа
            temp_df['diameter'] = temp_df['Наименование'].apply(get_diameter, args=(row.regular_expression, row.symbols_to_delete))
            temp_df['length'] = temp_df['Наименование'].apply(get_length)
            temp_df_grouped = temp_df.groupby(['diameter', 'length'])['Кол.'].sum().reset_index()
            temp_df_grouped['Наименование'] = row.part0 + temp_df_grouped['diameter'] + row.part2 + \
                                            temp_df_grouped['length'].apply(str) + row.part3
            temp_df_grouped['Размер'] = 'M' + temp_df_grouped['diameter'] + 'x' + temp_df_grouped[
                'length'].apply(
                str)
            temp_df_grouped['ГОСТ/ОСТ'] = row.container
            check_sum2 = temp_df_grouped['Кол.'].sum()
            temp_df_grouped.loc[temp_df_grouped['length'] == 0, 'Размер'] = 'мелкий шаг'
            if check_sum1 != check_sum2:
                raise ValueError('потеряли', row.container)
            list_of_df.append(temp_df_grouped)
        return pd.concat([*list_of_df]).reset_index(drop=True), df

    def normalization_1d(list_1d,df):
        list_of_df = []
        for row in list_1d.itertuples(index=True):
            temp_df = df.query('Наименование.str.contains(@row.container)')
            check_sum1 = temp_df['Кол.'].sum()
            df = df[~df.index.isin(temp_df.index)]
            temp_df['diameter'] = temp_df['Наименование'].apply(get_diameter,
                                                                args=(row.regular_expression, row.symbols_to_delete))
            temp_df_grouped = temp_df.groupby(['diameter'])['Кол.'].sum().reset_index() # отличия
            temp_df_grouped['Наименование'] = row.part0 + temp_df_grouped['diameter'] + row.part2 # отличия
            temp_df_grouped['Размер'] = 'M' + temp_df_grouped['diameter'] # отличия
            temp_df_grouped['ГОСТ/ОСТ'] = row.container
            check_sum2 = temp_df_grouped['Кол.'].sum()
            # отличия
            if check_sum1 != check_sum2:
                raise ValueError('потеряли', row.container)
            list_of_df.append(temp_df_grouped)
        return pd.concat([*list_of_df]).reset_index(drop=True), df

    fasteners_2d,df = normalization_2d(list_2d,df)
    fasteners_1d,df = normalization_1d(list_1d,df)
    fasteners = pd.concat([fasteners_2d,fasteners_1d]).reset_index(drop=True)
    return fasteners,df


def vint_translit_format():  # разбор старых винтов из Creo записаных транслитом
    def M(row):  # диаметр винта
        temp = re.search('M\d+[-_]*[56]*', row)
        anwser = temp.group()[1:]
        anwser = anwser.replace('_', '.')
        anwser = anwser.replace('-', '.')
        if anwser[0] > '2':
            anwser = anwser[0]
        return anwser

    def L(row):  # длина винта
        temp = re.search('X\d+', row)
        return int(temp.group()[1:])

    def vint_by_gost(container, part0, part2, part3):
        global vint
        vint_gost = vint.query('Наименование.str.contains(@container)')
        check_sum1 = vint_gost['Кол.'].sum()
        vint = vint[~vint.index.isin(vint_gost.index)]
        vint_gost['diameter'] = vint_gost['Наименование'].apply(M)
        vint_gost['length'] = vint_gost['Наименование'].apply(L)
        vint_grouped = vint_gost.groupby(['diameter', 'length'])['Кол.'].sum().reset_index()
        vint_grouped['Наименование'] = part0 + vint_grouped['diameter'] + part2 + \
                                       vint_grouped['length'].apply(str) + part3
        vint_grouped['Размер'] = 'M' + vint_grouped['diameter'] + 'x' + vint_grouped['length'].apply(str)
        vint_grouped['ГОСТ/ОСТ'] = container
        check_sum2 = vint_grouped['Кол.'].sum()
        if check_sum1 != check_sum2:
            raise ValueError('потеряли vint', container)
        vint_grouped['type'] = 'винт'
        return vint_grouped

    vint4762 = vint_by_gost('4762', 'Винт ГОСТ Р ИСО 4762-M', 'x', '.A2-70')

    vint1491 = vint_by_gost('1491-80', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 1491-80')

    vint17475 = vint_by_gost('17475-80', r'Винт А.М', r'-6gx', r'.58.20.016 ГОСТ 17475-80')

    vint1477 = vint_by_gost('1477-93', r'Винт A.M', r'-6gx', r'.22H.35.016 ГОСТ 1477-93')

    vint1476 = vint_by_gost('1476-93', r'Винт A.M', r'-6gx', r'.23.20Х13 ГОСТ 1476-93')

    vint11738 = vint_by_gost('11738-84', r'Винт M', r'-6gx', r'.23.20Х13 ГОСТ 11738-84')

    all_vint = pd.concat([vint4762, vint1491, vint17475, vint1477, vint1476, vint11738]).reset_index(drop=True)

    return all_vint

#начало мэйна

filename = 'sop_oe.csv'
true_dictionary = ['винт', 'гайка', 'шайба', 'штифт', 'vint', 'gajka', 'shajba', 'shtift']
data = parsing_df_of_fasteners(filename, true_dictionary)
print(normalization(make_list_of_samples('formats.csv'),data))

