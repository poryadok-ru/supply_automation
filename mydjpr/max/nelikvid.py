import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime
import time
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()


pd.options.mode.chained_assignment = None


os.chdir(os.path.dirname(os.path.abspath(__file__)))


def nelikvids(file_path, porog1, porog2, porog3, porog4, porog5, porog6, add_message):

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

    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    t1 = time.time()

    df_prod = pd.read_excel(file_path + "Исходники\\продажи.xlsx", engine="calamine")
    df_prod = df_prod.drop([0, 1, len(df_prod.index)-1])
    df_prod['Unnamed: 0'] = pd.to_numeric(df_prod['Unnamed: 0'], errors='coerce')
    df_klaster = pd.read_excel(file_path + "Исходники\\Кластеры.xlsx", engine="calamine")
    df_rozn_cena = pd.read_excel(file_path + "Исходники\\цена розн.xlsx", usecols=['Номенклатура.Код', 'Розничная'], engine="calamine")
    df_isklu = pd.read_excel(file_path + "Исходники\\исключения.xlsx", engine="calamine")

    t2 = time.time()
    print('прочитали файлы: продажи, кластеры, цена розн, исключения. время выполнения: ' + str(t2-t1))

    # d = {f: os.stat(f).st_mtime for f in glob.iglob("*.csv")}
    files = []
    files1 = glob.glob(file_path + "Исходники\\*.csv")
    for i in files1:
        if i == file_path + "Исходники\\сезонность.csv":
            continue
        else:
            files.append(i)

    df_effectiv = pd.DataFrame(columns=['Код (доп.)', 'Продукт', 'Склад',
                                        'Остаток текущий, ед.', 'ЗакупочнаяЦена',
                                        'КомментарийТО', 'Наличие товара, дней'])
    for k in range(0, len(files)):
        df_af = pd.read_csv(files[k], delimiter=';', dtype='unicode', usecols=['Код (доп.)', 'Продукт', 'Склад',
                                                                            'Остаток текущий, ед.',
                                                                            'ЗакупочнаяЦена',
                                                                            'КомментарийТО', 'Наличие товара, дней'])
        df_af['Остаток текущий, ед.'] = df_af['Остаток текущий, ед.'].str.replace(' ', '')
        df_af['Остаток текущий, ед.'] = df_af['Остаток текущий, ед.'].str.replace(',', '.')
        df_af['Остаток текущий, ед.'] = pd.to_numeric(df_af['Остаток текущий, ед.'], errors='coerce')
        df_af['ЗакупочнаяЦена'] = pd.to_numeric(df_af['ЗакупочнаяЦена'], errors='coerce')
        df_af['Наличие товара, дней'] = pd.to_numeric(df_af['Наличие товара, дней'], errors='coerce')
        df_af = df_af[df_af['Остаток текущий, ед.'] > 0]
        df_effectiv = pd.concat([df_effectiv, df_af])

    print(files)
    t3 = time.time()
    print('прочитали файлы из анализа эффективности. время выполнения: ' + str(t3-t2))

    df_effectiv['Остаток текущий, д.е.'] = df_effectiv['Остаток текущий, ед.']*df_effectiv['ЗакупочнаяЦена']
    df_effectiv = df_effectiv[['Код (доп.)', 'Продукт', 'Склад',  'Остаток текущий, ед.', 'Остаток текущий, д.е.',
                            'ЗакупочнаяЦена', 'КомментарийТО', 'Наличие товара, дней']]

    df_effectiv.rename(columns={'Код (доп.)': 'Код'}, inplace=True)

    df_effectiv_wklaster = df_effectiv.merge(df_klaster, on='Склад', how='left')
    df_effectiv_wklaster = df_effectiv_wklaster[df_effectiv_wklaster['КомментарийТО'] != "-"]
    df_effectiv_wklaster['Код'] = pd.to_numeric(df_effectiv_wklaster['Код'], errors='coerce')

    t4 = time.time()
    print('впр кластера. время выполнения: ' + str(t4-t3))

    prk_list = unic_list(df_effectiv_wklaster, 'Склад')
    # for prk in range(len(df_effectiv_wklaster.index)):
    #     if df_effectiv_wklaster['Склад'].iloc[prk] in prk_list:
    #         continue
    #     else:
    #         prk_list.append(df_effectiv_wklaster['Склад'].iloc[prk])

    # подтягиваем продажи
    # prk_list.sort()
    df_effectiv_wprod = pd.DataFrame()

    for each_prk in prk_list:
        # print(each_prk)
        df_y = (df_effectiv_wklaster[df_effectiv_wklaster['Склад'] == each_prk])
        # определяем строку и столбец названия магазина в файле продаж
        i, j = np.where(df_prod.values == each_prk)
        column_prk_prod = int(j)
        df_prod_y = df_prod[['Unnamed: 0', ('Unnamed: ' + str(column_prk_prod))]]
        df_prod_y = df_prod_y.rename(columns={'Unnamed: 0': 'Код', ('Unnamed: '+str(column_prk_prod)): 'Продажи 4 мес'})
        df_each_prk = df_y.merge(df_prod_y, on='Код', how='left')
        df_each_prk['Продажи 4 мес'].fillna(0, inplace=True)
        df_effectiv_wprod = pd.concat([df_effectiv_wprod, df_each_prk])

    t5 = time.time()
    print('впр продаж. время выполнения: ' + str(t5-t4))

    df_rozn_cena = df_rozn_cena.rename(columns={'Номенклатура.Код': 'Код', 'Розничная': 'Цена Магазин Воронеж'})
    df_effectiv_wrozncena = df_effectiv_wprod.merge(df_rozn_cena, on='Код', how='left')
    df_effectiv_wrozncena['Цена Магазин Воронеж'].fillna(0, inplace=True)

    t6 = time.time()
    print('впр розн цен. время выполнения: ' + str(t6-t5))



    # подтягиваем коэф сезонности
    months = {'01': 'Янв',
            '02': 'Фев',
            '03': 'Мар',
            '04': 'Апр',
            '05': 'Май',
            '06': 'Июн',
            '07': 'Июл',
            '08': 'Авг',
            '09': 'Сен',
            '10': 'Окт',
            '11': 'Ноя',
            '12': 'Дек',
            }

    now = datetime.now()
    now_month = now.strftime("%m")

    month1, year1 = (now.month-1, now.year) if now.month != 1 else (12, now.year-1)
    prev_month_1 = now.replace(day=1, month=month1, year=year1).strftime("%m")

    month2, year2 = (month1-1, now.year) if month1 != 1 else (12, now.year-1)
    prev_month_2 = now.replace(day=1, month=month2, year=year2).strftime("%m")

    month3, year3 = (month2-1, now.year) if month2 != 1 else (12, now.year-1)
    prev_month_3 = now.replace(day=1, month=month3, year=year3).strftime("%m")



    df_sez = pd.read_csv(file_path + "Исходники\\сезонность.csv", delimiter=';', dtype='unicode', usecols=['Код (доп.)',
                                                                                months[now_month],
                                                                                months[prev_month_1],
                                                                                months[prev_month_2],
                                                                                months[prev_month_3]])
    df_sez = df_sez.rename(columns={'Код (доп.)': 'Код'})
    df_sez['Код'] = pd.to_numeric(df_sez['Код'], errors='coerce')

    df_sez[months[now_month]] = df_sez[months[now_month]].str.replace(',', '.')
    df_sez[months[now_month]] = pd.to_numeric(df_sez[months[now_month]], errors='coerce')
    df_sez[months[prev_month_1]] = df_sez[months[prev_month_1]].str.replace(',', '.')
    df_sez[months[prev_month_1]] = pd.to_numeric(df_sez[months[prev_month_1]], errors='coerce')
    df_sez[months[prev_month_2]] = df_sez[months[prev_month_2]].str.replace(',', '.')
    df_sez[months[prev_month_2]] = pd.to_numeric(df_sez[months[prev_month_2]], errors='coerce')
    df_sez[months[prev_month_3]] = df_sez[months[prev_month_3]].str.replace(',', '.')
    df_sez[months[prev_month_3]] = pd.to_numeric(df_sez[months[prev_month_3]], errors='coerce')

    df_sez['Суммарный коэф сез'] = df_sez[months[now_month]] + df_sez[months[prev_month_1]] +\
                                df_sez[months[prev_month_2]] + df_sez[months[prev_month_3]]
    df_sez = df_sez[['Код', 'Суммарный коэф сез']]

    t7 = time.time()
    print('определение суммарного коэффициента сезонности. время выполнения: ' + str(t7-t6))



    df_effectiv_wks = df_effectiv_wrozncena.merge(df_sez, on='Код', how='left')
    df_effectiv_wks_iskl = df_effectiv_wks.merge(df_isklu, on='Код', how='left')

    t8 = time.time()
    print('впр коэф сезонности. время выполнения: ' + str(t8-t7))

    print('начинаем расчет')

    df_effectiv_wks_iskl['Смоделированные продажи 4мес'] = \
        (df_effectiv_wks_iskl['Продажи 4 мес']/df_effectiv_wks_iskl['Суммарный коэф сез'])*4
    df_effectiv_wks_iskl.replace([np.inf, -np.inf], np.nan, inplace=True)
    df_effectiv_wks_iskl['Смоделированные продажи 4мес'].fillna(0, inplace=True)

    df_effectiv_wks_iskl['Средне-дневные продажи (восстановленные)'] = \
        (df_effectiv_wks_iskl['Смоделированные продажи 4мес']/df_effectiv_wks_iskl['Наличие товара, дней'])

    df_effectiv_wks_iskl['Продажи за месяц в реализации (с учётом дней в наличии)'] = \
        (df_effectiv_wks_iskl['Средне-дневные продажи (восстановленные)']*df_effectiv_wks_iskl['Цена Магазин Воронеж']*30)

    # определяем неликвиды в зависимости от суммы продаж покластерно 150 100 75 рублей

    df_nel_kl = pd.DataFrame()

    df_effectiv_wks_1 = df_effectiv_wks_iskl[df_effectiv_wks_iskl['Кластер'] == 1]
    df_effectiv_wks_1_m150 = \
        df_effectiv_wks_1[df_effectiv_wks_1['Продажи за месяц в реализации (с учётом дней в наличии)'] < porog1]
    df_effectiv_wks_1_m150['Неликвиды, расчет покластерно по порогам'] = 'неликвид'
    df_effectiv_wks_1_b150 = \
        df_effectiv_wks_1[df_effectiv_wks_1['Продажи за месяц в реализации (с учётом дней в наличии)'] >= porog1]

    df_effectiv_wks_2 = df_effectiv_wks_iskl[df_effectiv_wks_iskl['Кластер'] == 2]
    df_effectiv_wks_2_m100 = \
        df_effectiv_wks_2[df_effectiv_wks_2['Продажи за месяц в реализации (с учётом дней в наличии)'] < porog2]
    df_effectiv_wks_2_m100['Неликвиды, расчет покластерно по порогам'] = 'неликвид'
    df_effectiv_wks_2_b100 = \
        df_effectiv_wks_2[df_effectiv_wks_2['Продажи за месяц в реализации (с учётом дней в наличии)'] >= porog2]

    df_effectiv_wks_3 = df_effectiv_wks_iskl[df_effectiv_wks_iskl['Кластер'] == 3]
    df_effectiv_wks_3_m75 = \
        df_effectiv_wks_3[df_effectiv_wks_3['Продажи за месяц в реализации (с учётом дней в наличии)'] < porog3]
    df_effectiv_wks_3_m75['Неликвиды, расчет покластерно по порогам'] = 'неликвид'
    df_effectiv_wks_3_b75 = \
        df_effectiv_wks_3[df_effectiv_wks_3['Продажи за месяц в реализации (с учётом дней в наличии)'] >= porog3]

    df_nel_kl = pd.concat([df_effectiv_wks_1_m150, df_effectiv_wks_1_b150,
                        df_effectiv_wks_2_m100, df_effectiv_wks_2_b100,
                        df_effectiv_wks_3_m75, df_effectiv_wks_3_b75])

    # исключаем неликвиды по дням в наличии менее 60 дней и списку исключений

    df_nel_iskl_dn = pd.DataFrame()

    df_nel_iskl = df_nel_kl[df_nel_kl['Неликвиды, расчет покластерно по порогам'] == 'неликвид']
    df_nel_iskl_b60 = df_nel_iskl[df_nel_iskl['Наличие товара, дней'] >= 60]
    df_nel_iskl_b60['Неликвиды Итог'] = "неликвид"
    df_nel_iskl_m60 = df_nel_iskl[df_nel_iskl['Наличие товара, дней'] < 60]
    df_nel_neiskl = df_nel_kl[df_nel_kl['Неликвиды, расчет покластерно по порогам'] != 'неликвид']

    df_nel_iskl_dn = pd.concat([df_nel_neiskl, df_nel_iskl_b60, df_nel_iskl_m60])

    df_nel_iskl_list = pd.DataFrame()

    df_nel_iskl_list_1 = df_nel_iskl_dn[df_nel_iskl_dn['Неликвиды, расчет покластерно по порогам'] == 'неликвид']
    df_nel_iskl_list_1_d = df_nel_iskl_list_1[df_nel_iskl_list_1['Исключение'] == 'да']
    df_nel_iskl_list_1_d['Неликвиды Итог'] = ""
    df_nel_iskl_list_1_n = df_nel_iskl_list_1[df_nel_iskl_list_1['Исключение'] != 'да']

    df_nel_iskl_list_2 = df_nel_iskl_dn[df_nel_iskl_dn['Неликвиды, расчет покластерно по порогам'] != 'неликвид']

    df_nel_iskl_list = pd.concat([df_nel_iskl_list_1_d, df_nel_iskl_list_1_n, df_nel_iskl_list_2])

    df_nel_itog = pd.DataFrame()

    df_nel_itog_1 = df_nel_iskl_list[df_nel_iskl_list['Неликвиды Итог'] == 'неликвид']
    df_nel_itog_1['Неликвиды, д.е'] = df_nel_itog_1['ЗакупочнаяЦена']*df_nel_itog_1['Остаток текущий, ед.']
    df_nel_itog_2 = df_nel_iskl_list[df_nel_iskl_list['Неликвиды Итог'] != 'неликвид']

    df_nel_itog = pd.concat([df_nel_itog_1, df_nel_itog_2])

    df_nel_itog = df_nel_itog.sort_values(by=['Склад', 'Продукт'])

    t9 = time.time()
    print('рассчитано. время выполнения: ' + str(t9-t8))

    # разделяем на два файла итоговые списки, тк есть ограничения эксель

    prk_list = unic_list(df_nel_itog, 'Склад')
    # for prk in range(len(df_nel_itog.index)):
    #     if df_nel_itog['Склад'].iloc[prk] in prk_list:
    #         continue
    #     else:
    #         prk_list.append(df_nel_itog['Склад'].iloc[prk])

    # prk_list.sort()

    df_nel_itog_1_1 = pd.DataFrame()
    df_nel_itog_1_2 = pd.DataFrame()

    for each_prk_1 in prk_list:
        df_each_prk_1 = df_nel_itog[df_nel_itog['Склад'] == each_prk_1]
        if (len(df_nel_itog_1_1) + len(df_each_prk_1) + 150_000) < 1_048_000:
            df_nel_itog_1_1 = pd.concat([df_each_prk_1, df_nel_itog_1_1])
        else:
            df_nel_itog_1_2 = pd.concat([df_each_prk_1, df_nel_itog_1_2])

    t10 = time.time()
    print('разделили на отдельные листы. время выполнения: ' + str(t10-t9))

    df_tolnel = df_nel_itog[df_nel_itog['Неликвиды Итог'] == 'неликвид']

    # определяем блокировки по неликвидам

    df_tolnel_block = df_tolnel[['Код', 'Продукт', 'Склад', 'Кластер', 'Цена Магазин Воронеж']]


    # d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники\\блок\\*.csv")}
    files2 = glob.glob(file_path + "Исходники\\блок\\*.csv")
    # files1 = []
    # for i in d.keys():
    #     files1.append(i)

    print(files2)

    df_nel_bl = pd.DataFrame()

    for k in range(0, len(files2)):
        df_af = pd.read_csv(files2[k], delimiter=';', low_memory=False)
        df_nel_bl = pd.concat([df_nel_bl, df_af])

    df_nel_bl["ABC"] = df_nel_bl["ABC 01. ЛКМ, клея, пропитки, растворители"] +\
                    df_nel_bl["ABC 02. Пена монтажная, герметики"] +\
                    df_nel_bl["ABC 03. Строительные и отделочные материалы"] +\
                    df_nel_bl["ABC 04. Инструмент"] +\
                    df_nel_bl["ABC 05. Товары для сада и огорода"] +\
                    df_nel_bl["ABC 06. Сантехника. Газ. Вентиляция"] +\
                    df_nel_bl["ABC 07. Посуда"] +\
                    df_nel_bl["ABC 08. Пластмассовые изделия"] +\
                    df_nel_bl["ABC 09. Хозтовары"] +\
                    df_nel_bl["ABC 10. Текстиль"] +\
                    df_nel_bl["ABC 11. Предметы интерьера"] +\
                    df_nel_bl["ABC 12. Замочно-скобяные изделия"] +\
                    df_nel_bl["ABC 13. Товары для спорта и отдыха"] +\
                    df_nel_bl["ABC 14. Бытовая техника"] +\
                    df_nel_bl["ABC 15. Электротовары"] +\
                    df_nel_bl["ABC 16. Бытовая химия"] +\
                    df_nel_bl["ABC 17. Товары для авто"]

    df_nel_bl = df_nel_bl[['Артикул (доп)', 'Склад(Название)', 'Прогноз спроса', 'ABC']]
    df_nel_bl["ABC"] = df_nel_bl["ABC"].str.replace('-', '')

    df_nel_bl.rename(columns={'Артикул (доп)': 'Код'}, inplace=True)

    df_nel_wpr = pd.DataFrame()

    for each_prk in prk_list:
        df_e = (df_nel_bl[df_nel_bl['Склад(Название)'] == each_prk])
        df_each_prk11 = df_e.merge(df_tolnel_block[df_tolnel_block['Склад'] == each_prk], on='Код', how='left')
        df_nel_wpr = pd.concat([df_nel_wpr, df_each_prk11])


    df_nel_wpr = df_nel_wpr[df_nel_wpr['Кластер'] != '']
    df_nel_wpr['Сумма продаж прогнозная, руб'] = df_nel_wpr['Прогноз спроса']*df_nel_wpr['Цена Магазин Воронеж']

    df_nel_wpr_1 = df_nel_wpr[df_nel_wpr['Кластер'] == 1]
    df_nel_wpr_1_1 = df_nel_wpr_1[df_nel_wpr_1['Сумма продаж прогнозная, руб'] < porog4]
    df_nel_wpr_2 = df_nel_wpr[df_nel_wpr['Кластер'] == 2]
    df_nel_wpr_2_1 = df_nel_wpr_2[df_nel_wpr_2['Сумма продаж прогнозная, руб'] < porog5]
    df_nel_wpr_3 = df_nel_wpr[df_nel_wpr['Кластер'] == 3]
    df_nel_wpr_3_1 = df_nel_wpr_3[df_nel_wpr_3['Сумма продаж прогнозная, руб'] < porog6]

    df_nel_wpr_itog = pd.concat([df_nel_wpr_1_1, df_nel_wpr_2_1, df_nel_wpr_3_1])

    nel_abc = df_nel_wpr_itog[(df_nel_wpr_itog["ABC"] != "") &
                (df_nel_wpr_itog["ABC"].str.contains("Новый") != True)]

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")

    nel_abc.to_excel(file_path + 'Блокировки по неликвидам ' + now1 + '.xlsx', index=False)

    path_for_dop_zakaz = r'\\lan.sct.ru\x\Воронеж\Подразделения\Коммерческий\Т.Шафиев\Общая\supply\Список_нел\\'

    df_for_dop_zakaz = df_tolnel[['Код', 'Склад']]
    # df_for_dop_zakaz.to_csv(path_for_dop_zakaz + 'Список неликвидов ' + now1 + '.csv', sep=";", index=False)
    df_for_dop_zakaz.to_parquet(path_for_dop_zakaz + 'Список неликвидов ' + now1 + '.parquet', engine='pyarrow')

    print('начинаем сохранять в файл')

    if len(df_nel_itog_1_2) != 0:
        with pd.ExcelWriter(file_path + 'Неликвиды общий файл ' + now1 + '.xlsx') as writer:
            df_nel_itog_1_1.to_excel(writer, sheet_name='общий часть 1', index=False)
            df_nel_itog_1_2.to_excel(writer, sheet_name='общий часть 2', index=False)
            df_tolnel.to_excel(writer, sheet_name='НЕЛИКВИДЫ', index=False)

    else:

        with pd.ExcelWriter(file_path + 'Неликвиды общий файл ' + now1 + '.xlsx') as writer:
            df_nel_itog.to_excel(writer, sheet_name='общий', index=False)
            df_tolnel.to_excel(writer, sheet_name='НЕЛИКВИДЫ', index=False)

    t11 = time.time()
    print('сохранено. время выполнения: ' + str(t11-t10))

    print('время расчета без оформления: ' + str(t11-t1))

    add_message('готово ' + str(t11-t1))