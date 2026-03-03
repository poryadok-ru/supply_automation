import datetime
from datetime import datetime as dt
import calendar
import os
import numpy as np
import pandas as pd
from time import time
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()




pd.options.mode.chained_assignment = None

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def maks(file_path):

    t = time()
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep
    # Определяем сегодняшнюю дату
    today = datetime.date.today()
    
    # Определяем количество дней в текущем месяце
    days_in_month = calendar.monthrange(today.year, today.month)[1]

    # Определяем количество дней, оставшихся до конца месяца
    days_left = days_in_month - today.day + 1
    
    def int_r(num):
        num = int(num + 0.5)
        return num

    k_o_1 = 0.8
    k_o_2 = 0.7
    k_o_3 = 0.6
    k_o_4 = 0.5
    koef_iznach = 2

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

    now = dt.now()
    now = now.replace(day=1)
    # print(now)
    now_month = now.strftime("%m")
    future_month = ((now.month+1) if (now.month+1) < 12 else 1)
    future_month_1 = now.replace(month=future_month).strftime("%m")
    cols = [months[now_month], months[future_month_1]]


    # df1 = pd.read_excel('C:/Users/a.merkulov/Desktop/Макросы и Шаблоны/Питон/Максимальный запас/макс пр-01  2809.xlsx')

    df1 = pd.read_csv(file_path + "Балансировка.csv", delimiter=";", low_memory=False, thousands=" ")
    # df1.to_excel("123.xlsx", index=False)

    col_balans = ['Артикул(доп.)',	'Наименование',	'В', 'ОТЗ (В) на норму запаса',	'Минимальная партия,ед. (В)', 'Сезон',
                   'Изначальный запас (В)',	'Текущие акции (В)',	'СегментСтелажногоХранения', 'Каталог (В)',	'Маячки (В)', 'Код', 'В (код)']

    prod_cols_1 = [col for col in df1.columns if 'Продажи' in col]
    # print(prod_cols_1)
    prod_cols_2 = [col for col in prod_cols_1 if '(В)' in col]
    # print(prod_cols_2)
    count_prod = len([item for item in prod_cols_2])
    # print(count_prod)
    # prod_cols_3= prod_cols_2[:-1]
    # print(prod_cols_3)

    col_from_balans = col_balans + prod_cols_2

    # print(col_from_balans)

    df2 = df1[col_from_balans]
    df2.rename(columns = {'Код': 'Код fnow'}, inplace = True)
    df2.fillna({'Маячки (В)': 0}, inplace=True)
    df2.fillna({'Каталог (В)': 0}, inplace=True)


    df2.rename(columns = {'Артикул(доп.)':'Код'}, inplace = True)
    # print(df2)

    df3 = pd.read_excel(file_path + 'Прогноз.xlsx', usecols=['Код', 'Прогноз менеджера отдела закупок'], engine="calamine")
    # df3['Код'] = pd.to_numeric(df3['Код'])
    data_pd = pd.merge(df2, df3, how='left', on='Код')

    # df4 = pd.read_csv(file_path + 'сезонность.csv', delimiter=";", usecols=['Код (доп.)'] + cols)
    df4 = pd.read_excel(file_path + "сезонность.xlsx", usecols=['Код (доп.)'] + cols, engine="calamine")
    df4[cols] = df4[cols].replace(',', '.', regex=True)
    df4[cols] = df4[cols].astype(float)
    df4 = df4.drop_duplicates(subset=['Код (доп.)'])

    df4.rename(columns = {'Код (доп.)':'Код'}, inplace = True)
    df4['Код'] = pd.to_numeric(df4['Код'], errors='coerce')
    data_pd = pd.merge(data_pd, df4, how='left', on='Код')

    df5 = pd.read_excel(file_path + '/Кластеры.xlsx', engine="calamine")
    df5.rename(columns = {'Склад': 'В'}, inplace = True)
    data_pd = pd.merge(data_pd, df5, how='left', on='В')


    if count_prod > 24:
        prod_cols_2= prod_cols_2[:-1]

    def calculate_mean(row):
        non_null = row[row != 0]
        max_value = non_null.max()
        mean = non_null.mean()
        if max_value > mean * 4:
            non_null = non_null[non_null != max_value]
            mean = non_null.mean()
        return mean

    data_pd['Средние'] = data_pd[prod_cols_2].apply(calculate_mean, axis=1)


    data_pd['Изначальный запас * базовый коэф-т'] = data_pd['Изначальный запас (В)'] * koef_iznach

    # Средн. продажи * усредн.сезон. коэф-т
    data_pd['Средн. продажи * усредн.сезон. коэф-т'] = (data_pd['Средние'] * ((data_pd[months[now_month]] * days_left +
                                                                    data_pd[months[future_month_1]] * (30 - days_left)) / 30))\
                                                .apply(np.ceil)

    # '80% от ОТЗ при прогнозе, 70% в ином случае' - общий
    data_pd['80% от ОТЗ при прогнозе, 70% в ином случае'] = (data_pd['ОТЗ (В) на норму запаса'] * k_o_2).apply(int_r)
    # '80% от ОТЗ при прогнозе, 70% в ином случае' - учитываем прогноз закупок
    data_pd.loc[data_pd['Прогноз менеджера отдела закупок'] > 0, '80% от ОТЗ при прогнозе, 70% в ином случае'] = \
        (data_pd['ОТЗ (В) на норму запаса'] * k_o_1).apply(int_r)
    # '80% от ОТЗ при прогнозе, 70% в ином случае' - учитываем сегмент полотенца с прогнозом
    data_pd.loc[((data_pd['СегментСтелажногоХранения'] == 'Полотенца') &
            (data_pd['Прогноз менеджера отдела закупок'] > 0)), '80% от ОТЗ при прогнозе, 70% в ином случае'] = \
        (data_pd['ОТЗ (В) на норму запаса'] * k_o_3).apply(int_r)
    # '80% от ОТЗ при прогнозе, 70% в ином случае' - учитываем сегмент полотенца без прогноза
    data_pd.loc[((data_pd['СегментСтелажногоХранения'] == 'Полотенца') &
            (data_pd['Прогноз менеджера отдела закупок'] == 0)), '80% от ОТЗ при прогнозе, 70% в ином случае'] = \
        (data_pd['ОТЗ (В) на норму запаса'] * k_o_4).apply(int_r)

    data_pd.loc[data_pd['Текущие акции (В)'] != '-', 'Текущая акция'] = 'Доп.акция'
    data_pd.loc[data_pd['Маячки (В)'] != 0, 'Текущая акция'] = 'Маячки'
    data_pd.loc[data_pd['Каталог (В)'] != 0, 'Текущая акция'] = 'Каталог'
    data_pd.fillna({'Текущая акция': 0}, inplace=True)

    data_pd.loc[((data_pd['Текущая акция'] == 0) | (data_pd['Текущая акция'] == 'Доп.акция')), 'Макс_расчётный'] = data_pd[['Средн. продажи * усредн.сезон. коэф-т', 'Изначальный запас * базовый коэф-т', 'Минимальная партия,ед. (В)']].max(axis=1)

    data_pd.loc[((data_pd['Текущая акция'] == 'Маячки') | (data_pd['Текущая акция'] == 'Каталог')), 'Макс_расчётный'] = data_pd[['Средн. продажи * усредн.сезон. коэф-т', 'Изначальный запас * базовый коэф-т', 'Минимальная партия,ед. (В)', '80% от ОТЗ при прогнозе, 70% в ином случае']].max(axis=1)

    # df6 = pd.read_excel('C:/Users/a.merkulov/Desktop/Макросы и Шаблоны/Питон/Максимальный запас/Ограничения.xlsx', sheet_name='Исключения', usecols=['Код', 'Исключение'])
    data = pd.read_excel(file_path + 'Ограничения.xlsx', sheet_name=['Исключения', 'ПРК', 'Сегменты_доля', 'Сегменты_значение', 'Сезон'], engine="calamine")
    df6 = data['Исключения'][['Код', 'Исключение']]
    data_pd = pd.merge(data_pd, df6, how='left', on='Код')

    df7 = data['ПРК']
    df7['Свод_ограничения_ПРК'] = df7['ПРК'] + df7['Сегменты ограничения']
    data_pd['Свод_ограничения_ПРК'] = data_pd['В'] + data_pd['СегментСтелажногоХранения']
    data_pd = pd.merge(data_pd, df7[['Свод_ограничения_ПРК', 'Значение_ПРК']], how='left', on='Свод_ограничения_ПРК')


    df8 = data['Сегменты_доля']
    df8['Свод_ограничения_Сегменты'] = df8['Сегменты ограничения'] + df8['Кластер'].apply(str)
    data_pd['Свод_ограничения_Сегменты'] = data_pd['СегментСтелажногоХранения'] + data_pd['Кластер'].apply(str)
    data_pd = pd.merge(data_pd, df8[['Свод_ограничения_Сегменты', 'Метод', 'Доля']], how='left', on='Свод_ограничения_Сегменты')

    def multiply(row):
        if row['Метод'] == 'Изначальный запас * базовый коэф-т':
            return row['Изначальный запас (В)'] * row['Доля']
        elif row['Метод'] == 'ОТЗ (В) на норму запаса':
            return row['ОТЗ (В) на норму запаса'] * row['Доля']
        else:
            return None

    data_pd['Расчёт по методу'] = data_pd.apply(multiply, axis=1)

    df9 = data['Сегменты_значение']
    df9['Свод_ограничения_Сегменты'] = df9['Сегменты ограничения'] + df9['Кластер'].apply(str)
    data_pd = pd.merge(data_pd, df9[['Свод_ограничения_Сегменты', 'Значение_сегмент']], how='left', on='Свод_ограничения_Сегменты')

    df10 = data['Сезон']
    df10['Свод_ограничения_Сезон'] = df10['Сезон'] + df10['Кластер'].apply(str)
    data_pd['Свод_ограничения_Сезон'] = data_pd['Сезон'] + data_pd['Кластер'].apply(str)
    data_pd = pd.merge(data_pd, df10[['Свод_ограничения_Сезон', 'Значение_сезон']], how='left', on='Свод_ограничения_Сезон')

    data_pd['Мин_ограничение'] = data_pd[['Значение_сезон','Значение_сегмент','Расчёт по методу','Значение_ПРК']].min(axis=1, skipna=True)

    data_pd['Макс_ИТОГовый'] = data_pd['Макс_расчётный']

    data_pd['Мин_ограничение'].replace(r'^s*$', np.nan, regex = True)

    data_pd.loc[pd.isna(data_pd['Мин_ограничение']) == True, 'Мин_ограничение'] = 0

    data_pd.loc[((data_pd['Мин_ограничение'] != 0) & (data_pd['Мин_ограничение'] < data_pd['Макс_ИТОГовый'])), 'Макс_ИТОГовый'] = data_pd['Мин_ограничение']

    data_pd.loc[data_pd['Изначальный запас (В)'] == 0, 'Макс_ИТОГовый'] = 0
    data_pd.loc[data_pd['Изначальный запас (В)'] > data_pd['Макс_ИТОГовый'], 'Макс_ИТОГовый'] = data_pd['Изначальный запас (В)']

    data_pd.loc[data_pd['Исключение'] == 0, 'Макс_ИТОГовый'] = 0

    df_vr = data_pd['Макс_ИТОГовый'] != 0
    data_pd.loc[df_vr , 'ПРОВЕРКА'] = data_pd['ОТЗ (В) на норму запаса'] / data_pd.loc[df_vr , 'Макс_ИТОГовый']
    data_pd.loc[data_pd['Макс_ИТОГовый'] == 0, 'ПРОВЕРКА'] = 0

    data_pd.sort_values(by=['Наименование'], inplace=True)

    writer = pd.ExcelWriter(file_path + 'Шаблон.xlsx', engine='xlsxwriter')

    data_pd.to_excel(writer, index=False)

    writer.close()

    writer2 = pd.ExcelWriter(file_path + 'Шаблон_1С_КА.xlsx', engine='xlsxwriter')

    data_pd.to_excel(writer2, index=False, columns=['Код', 'Наименование', 'В', 'Макс_ИТОГовый'])
    data_pd['Параметр'] = 'Максимальный запас, ед.'
    data_pd_fnow = data_pd[['В (код)', 'Код fnow', 'Параметр', 'Макс_ИТОГовый']]
    data_pd_fnow['Макс_ИТОГовый'] = data_pd_fnow['Макс_ИТОГовый'].astype(str)
    data_pd_fnow['Макс_ИТОГовый'] = data_pd_fnow['Макс_ИТОГовый'].str.replace('.0', '').astype(int)
    data_pd_fnow.to_csv(file_path + 'Максимальный запас FNOW.csv', sep=";", header=False, index=False)

    writer2.close()

    output_filename = "Шаблон_1С_КА.xlsx"
    print(time() - t)

    return output_filename

