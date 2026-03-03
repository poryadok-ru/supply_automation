from time import time
import pandas as pd
import numpy as np
import json
import csv
import glob
import os
from datetime import datetime
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()




csv.field_size_limit(2147483647)
pd.options.mode.chained_assignment = None
os.chdir(os.path.dirname(os.path.abspath(__file__)))


def addServerMessage(message):
  return {'message': message}


def calculate_eb(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    t = time()
    add_message('Стартуем!')
    add_message('Считываем файлы')
    
    def unic_list(df, column_name):
        listik = []
        df_1 = df.drop_duplicates(subset=[column_name])
        for i in range(0, len(df_1.index)):
            if df_1[column_name].iloc[i] in listik:
                continue
            else:
                listik.append(df_1[column_name].iloc[i])

        listik = pd.Series(listik).dropna().tolist()
        return listik

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")
    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.csv")}
    file_name_csv = max(d, key=lambda i: d[i])
    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.xlsx")}
    file_name_xlsx = max(d, key=lambda i: d[i])

    df = pd.read_excel(file_name_xlsx, engine="calamine")
    df['Склад(название)'] = df['Склад(название)'].replace('Порядок мг42_Н.Новгород_РЦ_Подольск',
                                                            'Порядок мг42_ННовгород_РЦ_Подольск')
    df['Склад(название)'] = df['Склад(название)'].replace('Порядок мг43__Рассказовка, ТЦ Сказка _РЦ_Подольск',
                                                            'Порядок мг43__Рассказовка ТЦ Сказка _РЦ_Подольск')
    prk_list = unic_list(df, 'Склад(название)')

    
    glav_prav_list_1 = unic_list(df, 'Код товара(доп.)')
    glav_prav_list = [str(x) for x in glav_prav_list_1]

    pravila_dict = dict()
    for i in prk_list:
        key = ('блок_СНАБ_' + i)
        pravila_dict[key] = i
    pravila_dict_vrem = dict()
    for j in prk_list:
        key = ('блок_ВРЕМ_' + j)
        pravila_dict_vrem[key] = j
    add_message('Создаем список временных блокировок')
    # Создаем общий список уникальных кодов для временных блокировок
    df_vrem_block = df[df['Временная блокировка по дефициту'] == 1]
    vrem_prav_list_1 = unic_list(df_vrem_block, 'Код товара(доп.)')
    vrem_prav_list = [str(x) for x in vrem_prav_list_1]
    vrem_pust = ['127323']

    slovar_razblock = dict()
    slovar_zablock = dict()
    slovar_vrem_lock = dict()
    add_message('Создаем списки блокировок по каждому ПРК')
    for each_prk in prk_list:
        df_each_prk = df[df['Склад(название)'] == each_prk]
        df_ep_razblock = df_each_prk[(df_each_prk['Направления балансировки'] == 'Заблокировано снабжением') &
                                        (df_each_prk['Разблок'] == 1)]
        list_razbl = df_ep_razblock['Код товара(доп.)'].apply(json.dumps).tolist()
        slovar_razblock[each_prk] = list_razbl

        df_ep_zabl = df_each_prk[df_each_prk['Заблок'] == 1]
        list_zabl = df_ep_zabl['Код товара(доп.)'].apply(json.dumps).tolist()
        slovar_zablock[each_prk] = list_zabl

        df_ep_vrem = df_each_prk[df_each_prk['Временная блокировка по дефициту'] == 1]
        list_vrem = df_ep_vrem['Код товара(доп.)'].apply(json.dumps).tolist()
        slovar_vrem_lock[each_prk] = list_vrem

    add_message('Загружаем блокировки в файл с правилами')
    with open(file_name_csv, newline='', encoding='utf-8') as infile, \
            open((file_path + 'правила нов ' + now1 + '.csv'), 'w', newline='', encoding='utf-8') as outfile:
        reader = csv.reader(infile, delimiter=';')
        writer = csv.writer(outfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for row in reader:
            if row[0] in pravila_dict.keys():
                data = json.loads(row[1])
                values = data["properties"][0]["values"]
                for code in slovar_razblock[pravila_dict[row[0]]]:
                    if code in values:
                        values.remove(code)
                for code in slovar_zablock[pravila_dict[row[0]]]:
                    if code not in values:
                        values.append(code)
                data["properties"][0]["values"] = values
                row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            elif row[0] == 'Еженедельные блокировки - ПО КОДУ установка РЦ-ПРК(обновлять)':
                data = json.loads(row[1])
                data["properties"][0]["values"] = glav_prav_list
                row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            elif row[0] in pravila_dict_vrem.keys():
                prk = pravila_dict_vrem[row[0]]
                codes = slovar_vrem_lock[prk] if slovar_vrem_lock[prk] else ['127323']

                data = json.loads(row[1])
                data["properties"][0]["values"] = codes  # ← Вот здесь происходит простое присваивание нового списка
                row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
                # if len(vrem_prav_list) > 0:
                #     data = json.loads(row[1])
                #     data["properties"][0]["values"] = vrem_prav_list
                #     row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
                # else:
                #     data = json.loads(row[1])
                #     data["properties"][0]["values"] = vrem_pust
                #     row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            writer.writerow(row)

    add_message('Сверяем старый и новый файл с правилами на отличия')
    with open(file_name_csv, newline='', encoding='utf-8') as file1, open((file_path + 'правила нов ' + now1 + '.csv'),
                                                                            newline='', encoding='utf-8') as file2:
        reader1 = csv.reader(file1)
        reader2 = csv.reader(file2)
        for row1, row2 in zip(reader1, reader2):
            if row1 != row2:
                add_message(f"Строка в файле 1: {row1[0].split(';')[0]}")
                add_message(f"Строка в файле 2: {row2[0].split(';')[0]}")  

    add_message('Сохраняем')
    add_message('Готово за ' + str(time() - t))

    


def calculate_bn(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    t = time()
    
    add_message('Стартуем!')
    add_message('Считываем файлы')
    def unic_list(df, column_name):
        listik = []
        df_1 = df.drop_duplicates(subset=[column_name])
        for i in range(0, len(df_1.index)):
            if df_1[column_name].iloc[i] in listik:
                continue
            else:
                listik.append(df_1[column_name].iloc[i])

        listik = pd.Series(listik).dropna().tolist()
        return listik

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")
    print(file_path)

    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.csv")}
    print(d)
    file_name_csv = max(d, key=lambda i: d[i])
    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.xlsx")}
    print(d)
    file_name_xlsx = max(d, key=lambda i: d[i])

    df = pd.read_excel(file_name_xlsx, engine="calamine")
    df['Склад(Название)'] = df['Склад(Название)'].replace('Порядок мг42_Н.Новгород_РЦ_Подольск', 'Порядок мг42_ННовгород_РЦ_Подольск')
    df['Склад(Название)'] = df['Склад(Название)'].replace('Порядок мг43__Рассказовка, ТЦ Сказка _РЦ_Подольск',
                                        'Порядок мг43__Рассказовка ТЦ Сказка _РЦ_Подольск')
    prk_list = unic_list(df, 'Склад(Название)')

    pravila_dict = dict()
    for i in prk_list:
        key = ('Блокнел_' + i)
        pravila_dict[key] = i

    add_message('Создаем списки блокировок по каждому ПРК')
    slovar_zablock = dict()

    for each_prk in prk_list:
        df_each_prk = df[df['Склад(Название)'] == each_prk]
        list_zabl = df_each_prk['Код'].apply(json.dumps).tolist()
        slovar_zablock[each_prk] = list_zabl

    add_message('Загружаем блокировки в файл с правилами')
    with open(file_name_csv, newline='', encoding='utf-8') as infile, \
            open((file_path + 'правила нов ' + now1 + '.csv'), 'w', newline='', encoding='utf-8') as outfile:
        reader = csv.reader(infile, delimiter=';')
        writer = csv.writer(outfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for row in reader:
            if row[0] in pravila_dict.keys():
                data = json.loads(row[1])
                data["properties"][0]["values"] = slovar_zablock[pravila_dict[row[0]]]
                row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            writer.writerow(row)

    add_message('Сверяем старый и новый файл с правилами на отличия')
    with open(file_name_csv, newline='', encoding='utf-8') as file1, open((file_path + 'правила нов ' + now1 + '.csv'),
                                                                            newline='', encoding='utf-8') as file2:
        reader1 = csv.reader(file1)
        reader2 = csv.reader(file2)
        for row1, row2 in zip(reader1, reader2):
            if row1 != row2:
                add_message(f"Строка в файле 1: {row1[0].split(';')[0]}")
                add_message(f"Строка в файле 2: {row2[0].split(';')[0]}")

    add_message('Сохраняем')
    add_message('Готово за ' + str(time() - t))



def calculate_pb(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    t = time()
    
    add_message('Стартуем!')
    add_message('Считываем файлы')
    def unic_list(df, column_name):
        listik = []
        df_1 = df.drop_duplicates(subset=[column_name])
        for i in range(0, len(df_1.index)):
            if df_1[column_name].iloc[i] in listik:
                continue
            else:
                listik.append(df_1[column_name].iloc[i])

        listik = pd.Series(listik).dropna().tolist()
        return listik

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")
    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.csv")}
    print(file_path)
    
    file_name_csv = max(d, key=lambda i: d[i])
    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для установки\\*.xlsx")}
    file_name_xlsx = max(d, key=lambda i: d[i])

    df = pd.read_excel(file_name_xlsx, engine="calamine")
    df['Склад'] = df['Склад'].replace('Порядок мг42_Н.Новгород_РЦ_Подольск', 'Порядок мг42_ННовгород_РЦ_Подольск')
    df['Склад'] = df['Склад'].replace('Порядок мг43__Рассказовка, ТЦ Сказка _РЦ_Подольск', 'Порядок мг43__Рассказовка ТЦ Сказка _РЦ_Подольск')
    prk_list = unic_list(df, 'Склад')

    pravila_dict = dict()
    for i in prk_list:
        key = ('блок_СНАБ_' + i)
        pravila_dict[key] = i

    slovar_zablock = dict()
    add_message('Создаем списки блокировок по каждому ПРК')

    for each_prk in prk_list:
        df_each_prk = df[df['Склад'] == each_prk]
        df_ep_zabl = df_each_prk[df_each_prk['К блокировке'] == 1]
        list_zabl = df_ep_zabl['Код (доп.)'].apply(json.dumps).tolist()
        slovar_zablock[each_prk] = list_zabl

    print(file_name_csv)

    add_message('Загружаем блокировки в файл с правилами')
    with open(file_name_csv, newline='', encoding='utf-8') as infile, \
            open((file_path + 'правила нов ' + now1 + '.csv'), 'w', newline='', encoding='utf-8') as outfile:
        reader = csv.reader(infile, delimiter=';')
        writer = csv.writer(outfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for row in reader:
            if row[0] in pravila_dict.keys():
                data = json.loads(row[1])
                values = data["properties"][0]["values"]
                for code in slovar_zablock[pravila_dict[row[0]]]:
                    if code not in values:
                        values.append(code)
                data["properties"][0]["values"] = values
                row[1] = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
            writer.writerow(row)
    
    add_message('Сверяем старый и новый файл с правилами на отличия')

    with open(file_name_csv, newline='', encoding='utf-8') as file1, open((file_path + 'правила нов ' + now1 + '.csv'),
                                                                            newline='', encoding='utf-8') as file2:
        reader1 = csv.reader(file1)
        reader2 = csv.reader(file2)
        for row1, row2 in zip(reader1, reader2):
            if row1 != row2:
                add_message(f"Строка в файле 1: {row1[0].split(';')[0]}")
                add_message(f"Строка в файле 2: {row2[0].split(';')[0]}")       

    add_message('Сохраняем')
    add_message('Готово за ' + str(time() - t))



def calculate_rb(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    t = time()
    
    add_message('Стартуем!')
    add_message('Считываем файлы')
    def unic_list(df, column_name):
        listik = []
        df_1 = df.drop_duplicates(subset=[column_name])
        for i in range(0, len(df_1.index)):
            if df_1[column_name].iloc[i] in listik:
                continue
            else:
                listik.append(df_1[column_name].iloc[i])
        listik = pd.Series(listik).dropna().sort_values().tolist()
        return listik


    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники для расчета\\*.xlsx")}

    files = []
    for i in d.keys():
        files.append(i)

    df_1 = pd.read_excel(files[0], engine="calamine")
    df_2 = pd.read_excel(files[1], engine="calamine")

    add_message('Прочитали файлы')

    if 'Стеллажное хранение - Состояние запасов' in df_1.values:
        df_main_stel = df_2.copy()
        df_sthr = df_1.copy()
    else:
        df_sthr = df_2.copy()
        df_main_stel = df_1.copy()

    add_message('Определили кто из них кто')

    prk = unic_list(df_main_stel, 'Склад(название)')

    df_main1 = pd.DataFrame()
    df_itog = pd.DataFrame()

    for each_prk in prk:
        df_main_n = (df_main_stel[df_main_stel['Склад(название)'] == each_prk])
        i, j = np.where(df_sthr.values == each_prk)
        column_prk_stel = (int(j + 7))

        df_stel = df_sthr[['Unnamed: 0', ('Unnamed: ' + str(column_prk_stel))]]
        df_stel = df_stel.truncate(before=(i[0]))
        df_stel = df_stel.rename(columns={'Unnamed: 0': 'Сегмент', ('Unnamed: ' + str(column_prk_stel)): '+/-'})
        df_wstel = df_main_n.merge(df_stel, on='Сегмент', how='left')

        df_main1 = pd.concat([df_main1, df_wstel])

    first_column = df_main1.pop("+/-")
    df_main1.insert(7, "+/-", first_column)

    add_message('Стеллажка ВПР готово')
    add_message('Делаем коллекции')

    df_main1['+/-'] = pd.to_numeric(df_main1['+/-'], errors='coerce')
    df_main1['+/-'].fillna(-99999, inplace=True)
    df_main1['Разблок'].fillna(0, inplace=True)
    df_main1['Заблок'].fillna(0, inplace=True)
    df_main1['Коллекция'].fillna(0, inplace=True)

    # df_main1.to_excel("123.xlsx", index=False)
    df_main1.loc[(df_main1['Сегмент'].isin(['Садовый электро/бензо-инструмент, силовая техника', 'Электро-инструмент'])) & (df_main1['Остатки базы'] >= 1), 'Остатки базы'] = 10
    # df_main1.to_excel("1232.xlsx", index=False)

    for each_prk in prk:
        df_main = (df_main1[df_main1['Склад(название)'] == each_prk]).copy()
        df_main = df_main.sort_values(by=['Коллекция', 'Приоритет', 'СТМ']).copy()
        df_main_collection = df_main[df_main['Коллекция'] != 0].copy()
        df_without_collection = df_main[df_main['Коллекция'] == 0].copy()

        list_col = unic_list(df_main_collection, 'Коллекция')


        if len(list_col) != 0:
            df = pd.DataFrame()
            for collection in list_col:
                df_c = df_main_collection[df_main_collection['Коллекция'] == collection]
                count = 0
                # sumost = 0
                for j in range(len(df_c.index)):
                    if df_c['+/-'].iloc[j] > 0:
                        count += 1
                # for s in range(len(df_c.index)):
                #     if df_c['Остатки базы'].iloc[s] >= 5:
                #         sumost += 1
                for n in range(len(df_c.index)):
                    if count >= (len(df_c) / 2):
                        df_c['Разблок'].iloc[n] = 1
                    else:
                        continue
                df = pd.concat([df, df_c])
            df1 = df[df['Разблок'] == 1]
            df2 = df[df['Разблок'] != 1]
        else:
            df1 = df_main[df_main['Разблок'] == 1]
            df2 = df_main[df_main['Разблок'] != 1]
        
        

        segment1 = unic_list(df1, 'Сегмент')

        df3 = pd.DataFrame()
        df_s3 = pd.DataFrame()
        df_s10 = pd.DataFrame()
        collection_del = []

        for segment_col in segment1:
            df_s2 = (df1[df1['Сегмент'] == segment_col])
            df_s2 = df_s2.sort_values(by=['Коллекция', 'Приоритет', 'СТМ'])
            # добавленная логика:
            collection2 = unic_list(df_s2, 'Коллекция')
            if (len(collection2) == 1) and (((df_s2['+/-'].sum() / df_s2['Разблок'].sum())) >= df_s2['Разблок'].sum()/2):
                df_s3 = pd.concat([df_s3, df_s2])
            else:
                df_s1 = pd.DataFrame()
                df_s4 = pd.DataFrame()
                if df_s2['Разблок'].sum() <= (df_s2['+/-'].sum() / df_s2['Разблок'].sum()):
                    df_s3 = pd.concat([df_s3, df_s2])
                else:
                    collection2 = unic_list(df_s2, 'Коллекция')
                    ostatok_segmenta = df_s2['+/-'].sum() / df_s2['Разблок'].sum()
                    for collection3 in collection2:
                        df_s = df_s2[df_s2['Коллекция'] == collection3]
                        df_s = df_s.sort_values(by=['Коллекция', 'Приоритет', 'СТМ'])
                        razblock_collection_sum = (df_s['Разблок'].sum())
                        if ostatok_segmenta >= (razblock_collection_sum - 1):
                            df_s['Разблок'] = 1
                            ostatok_segmenta = ostatok_segmenta - razblock_collection_sum
                        else:
                            df_s['Разблок'] = 0
                            collection_del.append(collection3)
                        df_s1 = pd.concat([df_s1, df_s])
                df_s3 = pd.concat([df_s3, df_s1])
        
        

        df_s7 = pd.DataFrame()
        if len(collection_del) == 0:
            df_s7 = df_s3
        else:
            df_s7 = pd.DataFrame()
            for col_2 in list_col:
                if col_2 in collection_del:
                    df_s6 = df_s3[df_s3['Коллекция'] == col_2]
                    razblock_collection_sum1 = (df_s6['Разблок'].sum())
                    if razblock_collection_sum1 >= (len(df_s6) / 2):
                        df_s6['Разблок'] = 1
                    else:
                        df_s6['Разблок'] = 0
                    df_s7 = pd.concat([df_s6, df_s7])
                else:
                    df_s7 = pd.concat([df_s7, df_s3[df_s3['Коллекция'] == col_2]])

        collection_del = []
        df3 = pd.concat([df2, df_s7])
        df_collection_ready = pd.concat([df3, df_without_collection])
        df_itog = pd.concat([df_collection_ready, df_itog])
    df_itog1 = df_itog.sort_values(by=['Склад(название)', 'Сегмент', 'Коллекция'])

    add_message('Коллекции готовы')
    add_message('Делаем остальные блокировки')

    df_itog2 = pd.DataFrame()

    for each_prk in prk:
        df_main7 = (df_itog1[df_itog1['Склад(название)'] == each_prk])
        df_menshe_nulya = df_main7[(df_main7['+/-'] <= 0)]
        df111 = df_menshe_nulya[(df_menshe_nulya['Направления балансировки'] != 'Заблокировано снабжением') & (
                    df_menshe_nulya['Разблок'] != 1)]
        df111['Заблок'] = 1
        df411 = df_menshe_nulya[(df_menshe_nulya['Направления балансировки'] != 'Заблокировано снабжением') & (
                    df_menshe_nulya['Разблок'] == 1)]
        df211 = df_menshe_nulya[(df_menshe_nulya['Направления балансировки'] == 'Заблокировано снабжением')]
        df311 = pd.concat([df111, df411, df211])

        df_bolshe_nulya = df_main7[(df_main7['+/-'] > 0)]

        df_main7 = pd.concat([df311, df_bolshe_nulya])

        df_main_1 = df_main7[df_main7['Остатки базы'] > 4]
        df_main_2 = df_main7[df_main7['Остатки базы'] <= 4]

        # создаем список сегментов для создания цикла
        segment2 = unic_list(df_main_1, 'Сегмент')

        df_s14_main = pd.DataFrame()
        df_s14_main_2 = pd.DataFrame()

        for segment_col in segment2:
            df_s14_1 = (df_main_1[df_main_1['Сегмент'] == segment_col])
            l = 0
            l1 = 0
            df_s14_1 = df_s14_1.sort_values(by=['Приоритет', 'СТМ'])
            while (l1 < len(df_s14_1)) and (df_s14_1['Разблок'].sum() < df_s14_1['+/-'].iloc[0]):
                if (df_s14_1['Разблок'].iloc[l] == 1) or (df_s14_1['Коллекция'].iloc[l] != 0):
                    l += 1
                    l1 += 1
                else:
                    df_s14_1['Разблок'].iloc[l] = 1
                    l += 1
                    l1 += 1
            df_s14_main_2 = pd.concat([df_s14_1, df_s14_main_2])

        df_s14_main_3 = df_s14_main_2[(df_s14_main_2['Направления балансировки'] != 'Заблокировано снабжением') & (
                    df_s14_main_2['Разблок'] != 0)]
        df_s14_main_4 = df_s14_main_2[(df_s14_main_2['Направления балансировки'] != 'Заблокировано снабжением') & (
                    df_s14_main_2['Разблок'] == 0)]
        df_s14_main_5 = df_s14_main_2[(df_s14_main_2['Направления балансировки'] == 'Заблокировано снабжением')]
        df_s14_main_4['Заблок'] = 1
        df_s14_main_6 = pd.concat([df_s14_main_3, df_s14_main_4, df_s14_main_5])
        df_17 = pd.concat([df_s14_main_6, df_main_2])

        df_20 = df_17[(df_17['Разблок'] != 1) & (df_17['Заблок'] != 1) & (
                    df_17['Направления балансировки'] != 'Заблокировано снабжением')]
        df_23 = df_17[(df_17['Разблок'] != 1) & (df_17['Заблок'] != 1) & (
                    df_17['Направления балансировки'] == 'Заблокировано снабжением')]
        df_20['Временная блокировка по дефициту'] = 1
        df_21 = df_17[(df_17['Разблок'] == 1) | (df_17['Заблок'] == 1)]
        df_22 = pd.concat([df_20, df_23, df_21])

        # >>> НОВЫЙ БЛОК — расчет остатка и перевод части "Заблокировано снабжением" во временную блокировку

        for seg in df_22['Сегмент'].dropna().unique():
            seg_mask = (df_22['Сегмент'] == seg)
            seg_df = df_22.loc[seg_mask]

            # Берем значение "+/-" для сегмента
            plusminus = int(seg_df['+/-'].iloc[0]) if len(seg_df) else 0
            # Сколько уже разблокировано
            unlocked_cnt = int(seg_df['Разблок'].sum())

            # Временные блокировки (старые + текущие), НЕ разблокированные
            temp_not_unlocked_cnt = int((
                (seg_df['Разблок'] != 1) &
                (
                    (seg_df['Направления балансировки'] == 'Временная блокировка по дефициту') |
                    (seg_df['Временная блокировка по дефициту'] == 1)
                )
            ).sum())

            # Считаем, сколько еще можно выставить временных блокировок
            free_slots = plusminus - unlocked_cnt - temp_not_unlocked_cnt
            if free_slots <= 0:
                continue  # Свободных мест нет, переходим к следующему сегменту

            # Кандидаты на перевод во "временную": заблокировано, не разблокировано, без коллекции
            cand_mask = (
                seg_mask &
                (df_22['Направления балансировки'] == 'Заблокировано снабжением') &
                (df_22['Разблок'] != 1) &
                (df_22['Коллекция'] == 0)
            )

            candidates = df_22.loc[cand_mask].sort_values(by=['Приоритет', 'СТМ'], ascending=[True, True])
            if candidates.empty:
                continue  # Нет подходящих кандидатов

            to_convert_idx = candidates.index[:free_slots]
            # Переводим найденные строки в "Временная блокировка по дефициту"
            df_22.loc[to_convert_idx, 'Временная блокировка по дефициту'] = 1
            df_22.loc[to_convert_idx, 'Направления балансировки'] = 'Временная блокировка по дефициту'
            df_22.loc[to_convert_idx, 'Заблок'] = 0  # При необходимости сбрасываем флаг
        # <<< КОНЕЦ НОВОГО БЛОКА

        
        df_itog2 = pd.concat([df_22, df_itog2])

    add_message('Блокировки готовы')
    add_message('Экспорт в файл')

    now = datetime.now()
    now = now.strftime("%d-%m-%Y")

    df_itog2 = df_itog2.sort_values(
        by=['Склад(название)', 'Сегмент', 'Приоритет', 'СТМ', 'Название', 'Направления балансировки'])

    df_itog2.to_excel(file_path + 'Блокировки ' + now + '.xlsx', index=False)

    add_message('Готово за ' + str(time() - t) + '')

    