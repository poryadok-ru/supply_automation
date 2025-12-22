import sys
import pandas as pd
import numpy as np
import glob
import os
import shutil
from time import time
from datetime import datetime, timedelta
from openpyxl.styles import Alignment, Border, Side, PatternFill
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
import xml.etree.ElementTree as ET
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()


pd.options.mode.chained_assignment = None

os.chdir(os.path.dirname(os.path.abspath(__file__)))


def minpartyf(file_path, porog1, porog2, porog3, semena, melk, pod_zakup, koef_okrugl, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep
    
    t = time()

    def int_r(num):
        num = int(num + 0.5)
        return num


    def int_10(num):
        num = int(num / 10 + 0.5)
        return num * 10


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

    # porog1 = int(porog_1.get())
    # porog2 = int(porog_2.get())
    # porog3 = int(porog_3.get())
    # semena = int(porog_s.get())
    # melk = int(porog_m.get())
    # pod_zakup = int(porog_pz.get())
    # koef_okrugl = float(koef_b.get())

    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "Исходники\\*.csv")}
    files = []
    for i in d.keys():
        files.append(i)

    df_effectiv = pd.DataFrame(
        columns=['Артикул (доп)', 'Продукт', 'Склад(Название)', 'Прогноз спроса', 'КаналПоставки',
                 'Код', 'Склад(Код)', 'ЗакупочнаяЦена', 'Кратность ,ед.'])

    for k in range(0, len(files)):
        df_af = pd.read_csv(files[k], delimiter=';', low_memory=False,
                            usecols=['Артикул (доп)', 'Продукт', 'Склад(Название)', 'Прогноз спроса', 'КаналПоставки',
                                     'Код', 'Склад(Код)', 'ЗакупочнаяЦена', 'Кратность ,ед.'])
        df_effectiv = pd.concat([df_effectiv, df_af])

    df_klaster = pd.read_excel(file_path + "Исходники\\Кластеры.xlsx")
    df_rozn_cena = pd.read_excel(file_path + "Исходники\\цена розн.xlsx", usecols=['Номенклатура.Код', 'Розничная'])
    df_fizminparty = pd.read_excel(file_path + "Исходники\\физминпартия.xlsx", usecols=['Код', 'Минимальная партия отгрузки'])

    df_rozn_cena.rename(columns={'Номенклатура.Код': 'Артикул (доп)', 'Розничная': 'Цена'}, inplace=True)
    df_fizminparty.rename(columns={'Код': 'Артикул (доп)'}, inplace=True)
    df_effectiv.rename(columns={'Код': 'Код позиции', 'Склад(Название)': 'Склад'}, inplace=True)

    add_message('прочитали файлы')

    df_effectiv_wklaster = df_effectiv.merge(df_klaster, on='Склад', how='left')
    df_effectiv_wrozncena = df_effectiv_wklaster.merge(df_rozn_cena, on='Артикул (доп)', how='left')
    df_effectiv_wfmp = df_effectiv_wrozncena.merge(df_fizminparty, on='Артикул (доп)', how='left')
    df_effectiv_wfmp['Цена'].fillna(99999, inplace=True)
    df_effectiv_wfmp['Минимальная партия отгрузки'].fillna(1, inplace=True)
    df_effectiv_wfmp['ЗакупочнаяЦена'] = pd.to_numeric(df_effectiv_wfmp['ЗакупочнаяЦена'], errors='coerce')
    df_effectiv_wfmp['ЗакупочнаяЦена'] = df_effectiv_wfmp['ЗакупочнаяЦена'].fillna(0)
    df_effectiv_wfmp = df_effectiv_wfmp[df_effectiv_wfmp['ЗакупочнаяЦена'] != 0]

    add_message('Объединили данные')

    df_effectiv_wfmp.loc[(porog1 / df_effectiv_wfmp['Цена']) < 0.5, 'под ' + str(porog1)] = 1
    df_effectiv_wfmp.loc[(porog1 / df_effectiv_wfmp['Цена']) >= 0.5, 'под ' + str(porog1)] = (
                porog1 / df_effectiv_wfmp['Цена']).apply(
        int_r)

    df_effectiv_wfmp.loc[(porog2 / df_effectiv_wfmp['Цена']) < 0.5, 'под ' + str(porog2)] = 1
    df_effectiv_wfmp.loc[(porog2 / df_effectiv_wfmp['Цена']) >= 0.5, 'под ' + str(porog2)] = (
                porog2 / df_effectiv_wfmp['Цена']).apply(
        int_r)

    df_effectiv_wfmp.loc[(porog3 / df_effectiv_wfmp['Цена']) < 0.5, 'под ' + str(porog3)] = 1
    df_effectiv_wfmp.loc[(porog3 / df_effectiv_wfmp['Цена']) >= 0.5, 'под ' + str(porog3)] = (
                porog3 / df_effectiv_wfmp['Цена']).apply(
        int_r)

    df_effectiv_wfmp.loc[df_effectiv_wfmp['Прогноз спроса'] < 0.5, 'под прогноз спроса'] = 1
    df_effectiv_wfmp.loc[df_effectiv_wfmp['Прогноз спроса'] >= 0.5, 'под прогноз спроса'] = \
        df_effectiv_wfmp['Прогноз спроса'].apply(int_r)

    # расчет для 1-2 кластеров
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] != 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] >= df_effectiv_wfmp[
                             'под ' + str(porog1)]), 'под 1-2 кластеры'] = \
        df_effectiv_wfmp['под ' + str(porog1)]

    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] != 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] < df_effectiv_wfmp['под ' + str(porog1)]) &
                         (df_effectiv_wfmp['под прогноз спроса'] >= df_effectiv_wfmp[
                             'под ' + str(porog2)]), 'под 1-2 кластеры'] = \
        df_effectiv_wfmp['под прогноз спроса']

    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] != 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] < df_effectiv_wfmp[
                             'под ' + str(porog2)]), 'под 1-2 кластеры'] = \
        df_effectiv_wfmp['под ' + str(porog2)]

    # расчет для 3 кластера
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] == 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] >= df_effectiv_wfmp[
                             'под ' + str(porog1)]), 'под 3 кластер'] = \
        df_effectiv_wfmp['под ' + str(porog1)]

    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] == 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] < df_effectiv_wfmp['под ' + str(porog1)]) &
                         (df_effectiv_wfmp['под прогноз спроса'] >= df_effectiv_wfmp[
                             'под ' + str(porog3)]), 'под 3 кластер'] = \
        df_effectiv_wfmp['под прогноз спроса']

    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] == 3) &
                         (df_effectiv_wfmp['под прогноз спроса'] < df_effectiv_wfmp[
                             'под ' + str(porog3)]), 'под 3 кластер'] = \
        df_effectiv_wfmp['под ' + str(porog3)]

    add_message('Определили минпартию по-кластерно')

    # расчет для семян под 200р
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          ((semena / df_effectiv_wfmp['Цена']).apply(int_r) < 0.5)), 'Семена под' + str(semena)] = 1
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          ((semena / df_effectiv_wfmp['Цена']).apply(int_r) >= 0.5)), 'Семена под' + str(semena)] = \
        (semena / df_effectiv_wfmp['Цена']).apply(int_r)

    add_message('Определили минпартию для семян')

    # под закупку
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Цена'] == 99999) |
                          (df_effectiv_wfmp['Цена'] < df_effectiv_wfmp['ЗакупочнаяЦена'])) &
                         (pod_zakup / df_effectiv_wfmp['ЗакупочнаяЦена'] < 0.5), 'порог под ' + str(pod_zakup)] = 1
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Цена'] == 99999) |
                          (df_effectiv_wfmp['Цена'] < df_effectiv_wfmp['ЗакупочнаяЦена'])) &
                         (pod_zakup / df_effectiv_wfmp['ЗакупочнаяЦена'] >= 0.5), 'порог под ' + str(pod_zakup)] = \
        (pod_zakup / df_effectiv_wfmp['ЗакупочнаяЦена']).apply(int_r)
    df_effectiv_wfmp['порог под ' + str(pod_zakup)].fillna(0, inplace=True)

    add_message('Определили минпартию по закупке')

    # минпартия по-кластерно в один столбец
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] != 3), 'минпартия по-кластерно'] = \
        df_effectiv_wfmp['под 1-2 кластеры']
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Кластер'] == 3), 'минпартия по-кластерно'] = \
        df_effectiv_wfmp['под 3 кластер']
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Продукт'].str.contains("Семена") == True), 'минпартия по-кластерно'] = \
        df_effectiv_wfmp['Семена под' + str(semena)]

    # канал поставки мелкий
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['КаналПоставки'] == "мелкий") &
                          (df_effectiv_wfmp['Цена'] != 99999) &
                          (melk / df_effectiv_wfmp['Цена'] >= 0.5)), 'минпартия по-кластерно'] = \
        (melk / df_effectiv_wfmp['Цена']).apply(int_r)
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['КаналПоставки'] == "мелкий") &
                          (df_effectiv_wfmp['Цена'] != 99999) &
                          (melk / df_effectiv_wfmp['Цена'] < 0.5)), 'минпартия по-кластерно'] = 1

    add_message('Определили минпартию для канала поставки мелкий')

    # если минпартия больше 20, то округляем ее кратно 10
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['минпартия по-кластерно'] >= 20), 'уменьшение вариаций установки'] = \
        df_effectiv_wfmp['минпартия по-кластерно'].apply(int_10)
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['минпартия по-кластерно'] < 20), 'уменьшение вариаций установки'] = \
        df_effectiv_wfmp['минпартия по-кластерно']

    df_effectiv_wfmp.loc[
        (df_effectiv_wfmp['порог под ' + str(pod_zakup)] >= 20), 'Минпартия под закупочную, вариация'] = \
        df_effectiv_wfmp['порог под ' + str(pod_zakup)].apply(int_10)
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['порог под ' + str(pod_zakup)] < 20) &
                          (df_effectiv_wfmp[
                               'порог под ' + str(pod_zakup)] != "")), 'Минпартия под закупочную, вариация'] = \
        df_effectiv_wfmp['порог под ' + str(pod_zakup)]

    # для семян кратно 10

    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          (df_effectiv_wfmp['минпартия по-кластерно'] >= 10)), 'уменьшение вариаций установки'] = \
        df_effectiv_wfmp['минпартия по-кластерно'].apply(int_10)
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          (df_effectiv_wfmp['минпартия по-кластерно'] < 10)), 'уменьшение вариаций установки'] = \
        df_effectiv_wfmp['минпартия по-кластерно']

    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          (df_effectiv_wfmp[
                               'порог под ' + str(pod_zakup)] >= 10)), 'Минпартия под закупочную, вариация'] = \
        df_effectiv_wfmp['порог под ' + str(pod_zakup)].apply(int_10)
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Продукт'].str.contains("Семена") == True) &
                          (df_effectiv_wfmp['порог под ' + str(pod_zakup)] < 10) &
                          (df_effectiv_wfmp[
                               'порог под ' + str(pod_zakup)] != "")), 'Минпартия под закупочную, вариация'] = \
        df_effectiv_wfmp['порог под ' + str(pod_zakup)]

    add_message('Округлили минпартию для больших значений (кратно 10)')

    # физминпартия
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['уменьшение вариаций установки'] <
                           df_effectiv_wfmp['Минимальная партия отгрузки']) &
                          (df_effectiv_wfmp['Минимальная партия отгрузки'] != 1)), 'Мин партия ИТОГ'] = \
        df_effectiv_wfmp['Минимальная партия отгрузки']
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['уменьшение вариаций установки'] >=
                           df_effectiv_wfmp['Минимальная партия отгрузки']) &
                          (df_effectiv_wfmp['Минимальная партия отгрузки'] != 1)), 'Мин партия ИТОГ'] = \
        (df_effectiv_wfmp['уменьшение вариаций установки'] / df_effectiv_wfmp['Минимальная партия отгрузки']). \
            apply(int_r) * df_effectiv_wfmp['Минимальная партия отгрузки']
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Минимальная партия отгрузки'] == 1), 'Мин партия ИТОГ'] = \
        df_effectiv_wfmp['уменьшение вариаций установки']

    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Минпартия под закупочную, вариация'] <
                           df_effectiv_wfmp['Минимальная партия отгрузки']) &
                          (df_effectiv_wfmp['Минпартия под закупочную, вариация'] != 0) &
                          (df_effectiv_wfmp['Минимальная партия отгрузки'] != 1)), 'Мин партия ИТОГ'] = \
        df_effectiv_wfmp['Минимальная партия отгрузки']
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Минпартия под закупочную, вариация'] >=
                           df_effectiv_wfmp['Минимальная партия отгрузки']) &
                          (df_effectiv_wfmp['Минпартия под закупочную, вариация'] != 0) &
                          (df_effectiv_wfmp['Минимальная партия отгрузки'] != 1)), 'Мин партия ИТОГ'] = \
        (df_effectiv_wfmp['Минпартия под закупочную, вариация'] / df_effectiv_wfmp['Минимальная партия отгрузки']). \
            apply(int_r) * df_effectiv_wfmp['Минимальная партия отгрузки']
    df_effectiv_wfmp.loc[((df_effectiv_wfmp['Минимальная партия отгрузки'] == 1) &
                          (df_effectiv_wfmp['Минпартия под закупочную, вариация'] != 0)), 'Мин партия ИТОГ'] = \
        df_effectiv_wfmp['Минпартия под закупочную, вариация']

    add_message('Учли физическую минпартию')

    # если вдруг цена продажная меньше закупочной цены
    df_effectiv_wfmp.loc[(df_effectiv_wfmp['Цена'] < df_effectiv_wfmp['ЗакупочнаяЦена']), 'Мин партия ИТОГ'] = \
        df_effectiv_wfmp['Минпартия под закупочную, вариация']

    add_message('Учли варианты, где продажная цена меньше закупки')

    df_effectiv_wfmp['Мин партия ИТОГ'] = df_effectiv_wfmp['Мин партия ИТОГ'].apply(int)

    # Создание функций для каждой части формулы
    def part1(row):
        if row['Мин партия ИТОГ'] == 1 or row['Мин партия ИТОГ'] % row['Кратность ,ед.'] == 0:
            return row['Мин партия ИТОГ']
        else:
            return np.nan

    def part2(row):
        if row['Мин партия ИТОГ'] / row['Кратность ,ед.'] > 1:
            if row['Мин партия ИТОГ'] % row['Кратность ,ед.'] < koef_okrugl * row['Кратность ,ед.']:
                return np.floor(row['Мин партия ИТОГ'] / row['Кратность ,ед.']) * row['Кратность ,ед.']
            else:
                return np.ceil(row['Мин партия ИТОГ'] / row['Кратность ,ед.']) * row['Кратность ,ед.']
        else:
            if row['Мин партия ИТОГ'] < koef_okrugl * row['Кратность ,ед.']:
                return row['Мин партия ИТОГ']
            else:
                return row['Кратность ,ед.']

    # Применение функций к DataFrame
    df_effectiv_wfmp['для загрузки'] = \
        df_effectiv_wfmp.apply(lambda row: part1(row) if pd.notnull(row['Мин партия ИТОГ']) else np.nan, axis=1)
    df_effectiv_wfmp['для загрузки'] = \
        df_effectiv_wfmp.apply(lambda row: part2(row) if pd.isnull(row['для загрузки']) else row['для загрузки'],
                               axis=1)

    add_message('Округлили мин партию по упаковке')

    prk_list = unic_list(df_effectiv_wfmp, 'Склад')

    df_nel_itog_1_1 = pd.DataFrame()
    df_nel_itog_1_2 = pd.DataFrame()

    for each_prk_1 in prk_list:
        df_each_prk_1 = df_effectiv_wfmp[df_effectiv_wfmp['Склад'] == each_prk_1]
        if (len(df_nel_itog_1_1) + len(df_each_prk_1) + 300_000) < 1_048_000:
            df_nel_itog_1_1 = pd.concat([df_each_prk_1, df_nel_itog_1_1])
        else:
            df_nel_itog_1_2 = pd.concat([df_each_prk_1, df_nel_itog_1_2])

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")

    df_effectiv_wfmp['Название параметра'] = "Минимальная партия,ед."
    df_csv = df_effectiv_wfmp[['Склад(Код)', 'Код позиции', 'Название параметра', 'для загрузки']]

    df_csv.to_csv(file_path + 'Минимальная партия для заливки ' + now1 + '.csv', sep=";", header=False, index=False)

    with pd.ExcelWriter(file_path + 'Минимальная партия общий файл ' + now1 + '.xlsx') as writer:
        df_nel_itog_1_1.to_excel(writer, sheet_name='общий часть 1', index=False)
        df_nel_itog_1_2.to_excel(writer, sheet_name='общий часть 2', index=False)

    
    add_message(f'Готово за {time() - t}')