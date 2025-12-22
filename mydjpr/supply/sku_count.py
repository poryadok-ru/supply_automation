import tkinter as tk
import pandas as pd
import os
from datetime import datetime
import glob
import threading
import sys
from time import time
import openpyxl
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()



def sku_countw(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep
        
    t = time()

    d = {f: os.stat(f).st_mtime for f in glob.iglob(file_path + "исходники\\*.xlsx")}

    files = []
    for i in d.keys():
        files.append(i)

    df1 = pd.DataFrame()

    for k in range(0, len(files)):
        df11 = pd.read_excel(files[k], engine='calamine')
        df11 = df11.truncate(before=8)
        df11 = df11[['Unnamed: 0', 'Unnamed: 4']]
        df11['Unnamed: 4'].fillna(0, inplace=True)
        df11 = df11[df11['Unnamed: 0'] != "Итого"]
        df11.rename(columns={'Unnamed: 0': 'Продукт', 'Unnamed: 4': 'Код'}, inplace=True)
        df1 = pd.concat([df1, df11])

    count = 0

    df1['номер перем'] = ''

    for i in range(0, len(df1.index)):
        if (df1['Код'].iloc[i] != 0):
            df1['номер перем'].iloc[i] = count
        else:
            count += 1

    df2 = df1[df1['Код'] == 0]
    prk_list = []

    for prk in range(0, len(df2.index)):
        if df2['Продукт'].iloc[prk] in prk_list:
            continue
        else:
            prk_list.append(df2['Продукт'].iloc[prk])

    for i in range(1, len(prk_list) + 1):
        df1.loc[df1['номер перем'] == i, 'perem'] = prk_list[i - 1]

    sku = []
    for k in prk_list:
        df3 = df1[df1['perem'] == k]
        sku.append(len(df3.index))

    df4 = pd.DataFrame(columns=['Перемещение', 'Количество SKU'])
    for l in range(0, len(prk_list)):
        df4.loc[l, 'Перемещение'] = prk_list[l]
        df4.loc[l, 'Количество SKU'] = sku[l]

    df4 = df4[(df4['Количество SKU'] > 50) | (df4['Количество SKU'] == 0)]

    df6 = df4

    df6.reset_index(drop=True, inplace=True)

    df6['schet'] = ''

    pos = 0
    schet = 0

    for i in range(1, len(df6.index)):
        if df6.loc[i, 'Количество SKU'] != 0:
            df6.loc[i, 'schet'] = df6.loc[pos, 'Перемещение']
            schet += 1
        else:
            schet += 1
            pos += schet
            i += 1
            schet = 0

    df7 = df6[df6['Перемещение'].str.contains("Порядок") == True]

    prk_list11 = []
    for prk in range(0, len(df7.index)):
        if df7['Перемещение'].iloc[prk] in prk_list11:
            continue
        else:
            prk_list11.append(df7['Перемещение'].iloc[prk])

    df_srzn = df6.copy()

    for prk in prk_list11:
        df_prk = df_srzn[df_srzn['schet'] == prk]
        sum = df_prk['Количество SKU'].sum()
        col_per = len((df_prk))
        srzn = int(sum / col_per) if col_per !=0 else ''
        df_srzn.loc[df_srzn['Перемещение'] == prk, 'Количество SKU'] = srzn

    for prk in prk_list11:
        df_prk = df6[df6['schet'] == prk]
        sum = df_prk['Количество SKU'].sum()
        df6.loc[df6['Перемещение'] == prk, 'Количество SKU'] = sum

    df6 = df6[['Перемещение', 'Количество SKU']]
    df_tolprk = df6[df6['Перемещение'].str.contains("Порядок") == True]
    df_srzn = df_srzn[['Перемещение', 'Количество SKU']]
    df_tolprk_srzn = df_srzn[df_srzn['Перемещение'].str.contains("Порядок") == True]

    df_volume = pd.read_excel(file_path + "объем перемещений.xlsx", usecols=['Номер', 'Объем', 'Склад-получатель'], engine='calamine')
    df77 = df6[df6['Перемещение'].str.contains("Порядок") != True]
    df77["Номер"] = df77["Перемещение"].str.extract(r'([A-Za-z]{4}\d+)')

    # Объединяем датафреймы по номеру
    df77 = df77.merge(df_volume, on="Номер", how="left")

    # Заменяем запятые на точки и преобразуем объем в float
    df77['Объем'] = df77['Объем'].astype(str).str.replace(',', '.').astype(float)

    # 1. Расчет общих показателей
    total_sku = df77['Количество SKU'].sum()
    total_volume = df77['Объем'].sum()
    avg_sku_per_m3 = total_sku / total_volume
    avg_sku_per_shipment = df77['Количество SKU'].mean()

    # 2. Разделение на Воронеж и регионы
    df_voronezh = df77[df77['Склад-получатель'].str.contains('Воронеж', case=False, na=False)]
    df_regions = df77[~df77['Склад-получатель'].str.contains('Воронеж', case=False, na=False)]

    # 3. Расчет показателей для Воронежа
    voronezh_sku = df_voronezh['Количество SKU'].sum()
    voronezh_volume = df_voronezh['Объем'].sum()
    voronezh_avg = voronezh_sku / voronezh_volume if voronezh_volume > 0 else 0
    voronezh_avg_ship = df_voronezh['Количество SKU'].mean()

    # 4. Расчет показателей для регионов
    regions_sku = df_regions['Количество SKU'].sum()
    regions_volume = df_regions['Объем'].sum()
    regions_avg = regions_sku / regions_volume if regions_volume > 0 else 0
    regions_avg_ship = df_regions['Количество SKU'].mean()

    # 5. Создание итоговой таблицы
    df_sum_itog = pd.DataFrame({
        '7 дней с даты': ['1', '1', '1'],
        'Код': ['город', 'регионы', 'всего'],
        'Суммарное количество SKU': [voronezh_sku, regions_sku, total_sku],
        'Суммарный объем, м3': [voronezh_volume, regions_volume, total_volume],
        'Среднее кол-во СКЮ/1 м3': [round(voronezh_avg, 1), round(regions_avg, 1), round(avg_sku_per_m3, 1)],
        'среднее кол-во скю на рейс': [round(voronezh_avg_ship, 1), round(regions_avg_ship, 1),
                                       round(avg_sku_per_shipment, 1)]
    })

    # Форматирование чисел с пробелами в качестве разделителей тысяч
    df_sum_itog['Суммарное количество SKU'] = df_sum_itog['Суммарное количество SKU'].apply(
        lambda x: f"{x:,.0f}".replace(',', ' '))
    df_sum_itog['Суммарный объем, м3'] = df_sum_itog['Суммарный объем, м3'].apply(lambda x: f"{x:,.0f}".replace(',', ' '))

    now1 = datetime.now()
    now1 = now1.strftime("%d-%m-%Y")

    with pd.ExcelWriter(file_path + 'Количество SKU ' + now1 + '.xlsx') as writer:
        df6.to_excel(writer, sheet_name='общий список перемещений', index=False)
        df_tolprk.to_excel(writer, sheet_name='Суммарные', index=False)
        df_tolprk_srzn.to_excel(writer, sheet_name='Средние', index=False)
        df77.to_excel(writer, sheet_name='объем', index=False)
        df_sum_itog.to_excel(writer, sheet_name='сводный_итог', index=False)

    wb = openpyxl.load_workbook(file_path + 'Количество SKU ' + now1 + '.xlsx')
    sheet = wb['общий список перемещений']
    sheet2 = wb['Суммарные']
    sheet3 = wb['Средние']


    sheet.row_dimensions[1].height = 20
    sheet.column_dimensions['A'].width = 57
    sheet.column_dimensions['B'].width = 17

    sheet2['A1'] = 'ПРК'
    sheet2.row_dimensions[1].height = 20
    sheet2.column_dimensions['A'].width = 57
    sheet2.column_dimensions['B'].width = 17

    sheet3['A1'] = 'ПРК'
    sheet3.row_dimensions[1].height = 20
    sheet3.column_dimensions['A'].width = 57
    sheet3.column_dimensions['B'].width = 17


    wb.save(file_path + 'Количество SKU ' + now1 + '.xlsx')

    add_message('готово')