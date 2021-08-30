import re
import pandas as pd
import xlsxwriter


def read_list_of_keywords(filename):
    with open(filename) as f:
        return f.readline().split()


def make_list_of_samples(filename):
    # читаем список шаблонов крепежа в df
    return pd.read_csv(filename, encoding='ANSI', sep=';').fillna('')


def parsing_df_of_fasteners(filename, dictionary):
    # проходим весь входной csv файл
    # выбираем строки, в которых есть ключевые слова-type из dictionary и нет "пакб"
    # возвращаем df где есть только строки с крепежём
    data = pd.read_csv(filename, encoding='ANSI', sep=';')
    data.columns = ['Наименование', 'Кол']
    list_of_df = []
    for type in dictionary:
        temp_df = data.query(
            'Наименование.str.lower().str.contains(@type) and ~Наименование.str.lower().str.contains("пакб")')
        temp_df['type'] = type
        if len(temp_df) > 0:
            list_of_df.append(temp_df)
    return pd.concat([*list_of_df]).reset_index(drop=True)


def get_diameter(row, regular_expression, symbols_to_delete):
    # диаметр крепежа
    # regular_expression - регулярка для поиска диаметра крепежа
    # symbols_to_delete - Колчество первых символов, которые нужно удалять из найденной регулярки
    temp = re.search(regular_expression, row)
    anwser = temp.group()[symbols_to_delete:]
    if anwser[0] > '2':
        anwser = anwser[0]
    if anwser.endswith('.'):
        anwser = anwser[0:-1]
    return anwser


def get_length(row):
    temp = re.search('[xх]\d+', row)
    return int(temp.group()[1:])


def normalization(list_of_samples, df):
    # массив из разобранных df относящихся к разным гостам крепежа
    list_of_samples = list_of_samples[list_of_samples['type'] != 'vint']
    list_of_df = []
    for row in list_of_samples.itertuples(index=True):
        # временный df с крепежём госта container
        temp_df = df.query(
            'Наименование.str.contains(@row.container) and ~Наименование.str.lower().str.contains("vint")')
        # пересчитаем весь крепеж этого госта, чтобы потом убедиться, что ничего не потеряли
        check_sum1 = temp_df['Кол'].sum()
        # выкидываем крепёж из исходного df, чтобы потом сделать список из неразобранного крепежа
        df = df[~df.index.isin(temp_df.index)]
        temp_df['diameter'] = temp_df['Наименование'].apply(get_diameter,
                                                            args=(row.regular_expression, row.symbols_to_delete))
        if row.part3 != '':
            # если крепёж "двумерный"
            temp_df['length'] = temp_df['Наименование'].apply(get_length)
            temp_df_grouped = temp_df.groupby(['diameter', 'length'])['Кол'].sum().reset_index()
            temp_df_grouped['Наименование'] = row.part0 + temp_df_grouped['diameter'] + row.part2 + \
                                              temp_df_grouped['length'].apply(str) + row.part3
            temp_df_grouped['Размер'] = 'M' + temp_df_grouped['diameter'] + 'x' + temp_df_grouped['length'].apply(str)
        else:
            # если крепёж "одномерный"
            temp_df_grouped = temp_df.groupby(['diameter'])['Кол'].sum().reset_index()
            temp_df_grouped['Наименование'] = row.part0 + temp_df_grouped['diameter'] + row.part2
            temp_df_grouped['Размер'] = 'M' + temp_df_grouped['diameter']
            temp_df_grouped['length'] = ''
        temp_df_grouped['ГОСТ/ОСТ'] = row.container
        temp_df_grouped['type'] = row.type
        check_sum2 = temp_df_grouped['Кол'].sum()
        temp_df_grouped.loc[temp_df_grouped['length'] == 0, 'Размер'] = 'мелкий шаг'
        if check_sum1 != check_sum2:
            raise ValueError('потеряли', row.container)
        list_of_df.append(temp_df_grouped)
    return pd.concat([*list_of_df]).reset_index(drop=True), df


def vint_translit_normalization(list_of_samples, df):
    # разбор старых винтов из Creo записаных транслитом
    # костыльным винтам костыльная функция
    # нужно вызывать после normalization, иначе сломается
    def M(row):
        # диаметр винта
        temp = re.search('M\d+[-_]*[56]*', row)
        anwser = temp.group()[1:]
        anwser = anwser.replace('_', '.')
        anwser = anwser.replace('-', '.')
        if anwser[0] > '2':
            anwser = anwser[0]
        return anwser

    def L(row):
        # длина винта
        temp = re.search('X\d+', row)
        return int(temp.group()[1:])

    list_of_samples = list_of_samples[list_of_samples['type'] == 'vint']
    list_of_df = []

    for row in list_of_samples.itertuples(index=True):
        if len(row.container) > 5:
            slice = row.container[:-3]
        else:
            slice = row.container
        # временный df с крепежём госта container
        temp_df = df.query('Наименование.str.contains(@slice)')
        # пересчитаем весь крепеж этого госта, чтобы потом убедиться, что ничего не потеряли
        check_sum1 = temp_df['Кол'].sum()
        # выкидываем крепёж из исходного df, чтобы потом сделать список из неразобранного крепежа
        df = df[~df.index.isin(temp_df.index)]
        temp_df['diameter'] = temp_df['Наименование'].apply(M)
        temp_df['length'] = temp_df['Наименование'].apply(L)
        temp_df_grouped = temp_df.groupby(['diameter', 'length'])['Кол'].sum().reset_index()
        temp_df_grouped['Наименование'] = row.part0 + temp_df_grouped['diameter'] + row.part2 + \
                                          temp_df_grouped['length'].apply(str) + row.part3
        temp_df_grouped['Размер'] = 'M' + temp_df_grouped['diameter'] + 'x' + temp_df_grouped['length'].apply(str)
        temp_df_grouped['ГОСТ/ОСТ'] = row.container
        check_sum2 = temp_df_grouped['Кол'].sum()
        temp_df_grouped.loc[temp_df_grouped['length'] == 0, 'Размер'] = 'мелкий шаг'
        if check_sum1 != check_sum2:
            raise ValueError('потеряли', row.container)
        temp_df_grouped['type'] = 'винт'
        list_of_df.append(temp_df_grouped)
    return pd.concat([*list_of_df]).reset_index(drop=True), df


def concat_translit_screws(rus_df, en_df):
    # присоединяем костыльные винты к нормальным
    all_screws = pd.concat([rus_df, en_df]).reset_index(drop=True)
    duplicates = all_screws.groupby('Наименование')['Кол'].sum()
    all_screws = all_screws.merge(duplicates, on='Наименование', how='left')
    all_screws = all_screws[['type', 'diameter', 'length', 'Размер', 'ГОСТ/ОСТ', 'Кол_y', 'Наименование']]
    all_screws.columns = ['type', 'diameter', 'length', 'Размер', 'ГОСТ/ОСТ', 'Кол', 'Наименование']
    all_screws = all_screws.drop_duplicates('Наименование')
    all_screws = all_screws.reset_index(drop=True)
    return all_screws


def sort_and_delete(df):
    def custom_sorting(col: pd.Series):
        to_ret = col
        if col.name == "ГОСТ/ОСТ":
            order = col.value_counts()
            order_number = 0
            for index, value in order.items():
                to_ret.loc[to_ret == index] = order_number
                order_number += 1
        return to_ret

    df['Прим'] = ''
    df = df.sort_values(by=['type', 'ГОСТ/ОСТ', 'diameter', 'length'], ascending=True, key=custom_sorting).reset_index(
        drop=True)
    df.loc[
        df['Размер'] == 'мелкий шаг', ['Прим']] = 'скрипт не обрабатывает мелкий шаг. нужно вручную вбивать этот винт'
    df.index += 1
    df.index.rename('№', inplace=True)
    return df[['Наименование', 'ГОСТ/ОСТ', 'Размер', 'Кол', 'Прим']]


def reformate_bad_df(df):
    df.columns = ['Наименование', 'Кол', 'Тип']
    df = df.reset_index(drop=True)
    df.index.rename('№', inplace=True)
    df.index += 1
    return df


def create_out_xls(df, df_bad):
    fileout = filename[0:-4] + '_out.xlsx'
    writer = pd.ExcelWriter(fileout, engine='xlsxwriter')
    workbook = writer.book
    df.to_excel(writer, 'Крепёж', startrow=1)
    df_bad.to_excel(writer, 'Не распозналось')
    border_fmt1 = workbook.add_format({'bottom': 1, 'top': 1, 'left': 1, 'right': 1})
    border_fmt2 = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'right': 2})
    cell_format = workbook.add_format({'font_size': 14})
    cell_format.set_align('vcenter')

    for sheet_name in ['Крепёж', 'Не распозналось']:
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column(1, 1, 50)
        worksheet.set_column(2, 2, 10)
        worksheet.set_default_row(20)

    # Первый лист

    worksheet = writer.sheets['Крепёж']
    worksheet.set_row(0, 30)
    max_row = len(df) + 1
    max_col = len(df.columns)

    worksheet.conditional_format(xlsxwriter.utility.xl_range(2, 1, max_row, max_col),
                                 {'type': 'no_errors', 'format': border_fmt1})

    worksheet.conditional_format(xlsxwriter.utility.xl_range(1, 0, max_row, max_col),
                                 {'type': 'no_errors', 'format': border_fmt2})

    worksheet.conditional_format(xlsxwriter.utility.xl_range(1, 0, max_row, max_col),
                                 {'type': 'no_errors', 'format': border_fmt2})

    worksheet.write_string(max_row + 2, 1, 'Главный конструктор ОКР', cell_format=cell_format)
    worksheet.write_string(0, 0, 'Перечень крепежных изделий для сборки опытного образца', cell_format=cell_format)

    # Второй лист

    worksheet = writer.sheets['Не распозналось']
    max_row = len(df_bad)
    max_col = len(df_bad.columns)
    worksheet.conditional_format(xlsxwriter.utility.xl_range(1, 1, max_row, max_col),
                                 {'type': 'no_errors', 'format': border_fmt1})
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, max_row, max_col),
                                 {'type': 'no_errors', 'format': border_fmt2})

    writer.save()


# начало мэйна

filename = 'locator.csv'
keywords = read_list_of_keywords('keywords.txt')
data = parsing_df_of_fasteners(filename, keywords)
list_of_samples = make_list_of_samples('formats.csv')
data, bad_data = normalization(list_of_samples, data)
vint, bad_data = vint_translit_normalization(list_of_samples, bad_data)
data = concat_translit_screws(data, vint)
data = sort_and_delete(data)
bad_data = reformate_bad_df(bad_data)
create_out_xls(data, bad_data)
