import pandas as pd
import os
import glob
from datetime import datetime, timedelta
from time import time
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()
pd.options.mode.chained_assignment = None
os.chdir(os.path.dirname(os.path.abspath(__file__)))


def run_all_nalichie_analysis(file_path, add_message):
    """
    Объединенная функция для запуска всех трех анализов наличия
    """
    results = []
    
    def combined_add_message(msg):
        add_message(msg)
        results.append(msg)
    
    # Запускаем все три функции последовательно
    
    nalichie_rozn(file_path, combined_add_message)
    
    
    nalichie_comp(file_path, combined_add_message)
    
    
    nalichie_comp_RF(file_path, combined_add_message)
    
    
    
    return results


# Оригинальные функции остаются без изменений:

def nalichie_rozn(file_path, add_message):
    # ... весь оригинальный код функции nalichie_rozn ...
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

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

    t = time()
    # add_message('Стартуем!')
    # add_message('Считываем файлы')

    files = glob.glob(file_path + "Исходники для наличия розницы\\*.csv")

    df_ABC = pd.read_csv(files[0], nrows=0)  # Чтение только заголовков
    headers = df_ABC.columns.tolist()

    list_ABC_SH = ["ABC 01. ЛКМ, клея, пропитки, растворители (ШафиевТ)",
            "ABC 02. Пена монтажная, герметики (ШафиевТ)",
            "ABC 03. Строительные и отделочные материалы (ШафиевТ)",
            "ABC 04. Инструмент (ШафиевТ)",
            "ABC 05. Товары для сада и огорода (ШафиевТ)",
            "ABC 06. Сантехника. Газ. Вентиляция (ШафиевТ)",
            "ABC 07. Посуда (ШафиевТ)",
            "ABC 08. Пластмассовые изделия (ШафиевТ)",
            "ABC 09. Хозтовары (ШафиевТ)",
            "ABC 10. Текстиль (ШафиевТ)",
            "ABC 11. Предметы интерьера (ШафиевТ)",
            "ABC 12. Замочно-скобяные изделия (ШафиевТ)",
            "ABC 13. Товары для спорта и отдыха (ШафиевТ)",
            "ABC 14. Бытовая техника (ШафиевТ)",
            "ABC 15. Электротовары (ШафиевТ)",
            "ABC 16. Бытовая химия (ШафиевТ)",
            "ABC 17. Товары для авто (ШафиевТ)",
            "ABC 19. Праздничные товары (ШафиевТ)"]

    if list_ABC_SH in headers:
        list_ABC = list_ABC_SH
    else:
        list_ABC = ["ABC 01. ЛКМ, клея, пропитки, растворители",
            "ABC 02. Пена монтажная, герметики",
            "ABC 03. Строительные и отделочные материалы",
            "ABC 04. Инструмент",
            "ABC 05. Товары для сада и огорода",
            "ABC 06. Сантехника. Газ. Вентиляция",
            "ABC 07. Посуда",
            "ABC 08. Пластмассовые изделия",
            "ABC 09. Хозтовары",
            "ABC 10. Текстиль",
            "ABC 11. Предметы интерьера",
            "ABC 12. Замочно-скобяные изделия",
            "ABC 13. Товары для спорта и отдыха",
            "ABC 14. Бытовая техника",
            "ABC 15. Электротовары",
            "ABC 16. Бытовая химия",
            "ABC 17. Товары для авто",
            "ABC 19. Праздничные товары"]


    df1 = pd.DataFrame()
    for k in range(0, len(files)):
        df_af = pd.read_csv(files[k], delimiter=';', dtype='unicode', usecols=(['Артикул (доп)', 'Склад(Название)',
                                                                            'Фактический остаток',
                                                                            'Прогноз спроса'] + list_ABC))

        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(' ', '')
        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(',', '.')
        df_af['Фактический остаток'] = pd.to_numeric(df_af['Фактический остаток'], errors='coerce')
        # df_af['Прогноз спроса'] = df_af['Прогноз спроса'].str.replace(' ', '')
        # df_af['Прогноз спроса'] = df_af['Прогноз спроса'].str.replace(',', '.')
        df_af['Прогноз спроса'] = pd.to_numeric(df_af['Прогноз спроса'], errors='coerce')
        df_af['Артикул (доп)'] = pd.to_numeric(df_af['Артикул (доп)'], errors='coerce')
        df1 = pd.concat([df1, df_af])

    for col in list_ABC:
        df1[col] = df1[col].fillna('')
    df1.columns = df1.columns.str.replace(r"\s*\(ШафиевТ\)", "", regex=True)



    df1["ABC"] = df1["ABC 01. ЛКМ, клея, пропитки, растворители"] +\
                    df1["ABC 02. Пена монтажная, герметики"] +\
                    df1["ABC 03. Строительные и отделочные материалы"] +\
                    df1["ABC 04. Инструмент"] +\
                    df1["ABC 05. Товары для сада и огорода"] +\
                    df1["ABC 06. Сантехника. Газ. Вентиляция"] +\
                    df1["ABC 07. Посуда"] +\
                    df1["ABC 08. Пластмассовые изделия"] +\
                    df1["ABC 09. Хозтовары"] +\
                    df1["ABC 10. Текстиль"] +\
                    df1["ABC 11. Предметы интерьера"] +\
                    df1["ABC 12. Замочно-скобяные изделия"] +\
                    df1["ABC 13. Товары для спорта и отдыха"] +\
                    df1["ABC 14. Бытовая техника"] +\
                    df1["ABC 15. Электротовары"] +\
                    df1["ABC 16. Бытовая химия"] +\
                    df1["ABC 17. Товары для авто"] +\
                    df1["ABC 19. Праздничные товары"]


    df1 = df1[['Артикул (доп)', 'Склад(Название)', 'Фактический остаток', 'Прогноз спроса', 'ABC']]

    prk_list = unic_list(df1, 'Склад(Название)')
    list_ABC_2 = ['A-A', 'A-B', 'A-C', 'B-A', 'B-B', 'C-A']

    df1 = df1[df1['ABC'].str.contains('|'.join(list_ABC_2))]
    df1['Наличие'] = ''

    for prk in prk_list:
        df_prk = df1[df1['Склад(Название)'] == prk]
        df1.loc[df1['Склад(Название)'] == prk, 'Наличие'] = (
                len(df_prk[df_prk['Фактический остаток'] >= df_prk['Прогноз спроса']])/len(df_prk))

    df1 = df1[['Склад(Название)', 'Наличие']]
    df1 = df1.drop_duplicates(subset=['Склад(Название)']).sort_values(by='Склад(Название)')
    srznach = df1['Наличие'].mean()
    now = datetime.now()
    now = now.strftime("%d-%m-%Y")

    df1.to_excel(file_path + 'Наличие Розница ' + now + '.xlsx', index=False)

    add_message('Наличие Розница ' + str(f'{srznach*100:.1f}'))


def nalichie_comp(file_path, add_message):
    # ... весь оригинальный код функции nalichie_comp ...
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

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

    t = time()
    # add_message('Стартуем!')
    # add_message('Считываем файлы')

    now = datetime.now() - timedelta(days=1)
    now_month = now.strftime("%m-%Y")
    month1, year1 = (now.month-1, now.year) if now.month != 1 else (12, now.year-1)
    prev_month_1 = now.replace(day=1, month=month1, year=year1).strftime("%m-%Y")
    month2, year2 = (month1-1, year1) if month1 != 1 else (12, now.year-1)
    prev_month_2 = now.replace(day=1, month=month2, year=year2).strftime("%m-%Y")
    month3, year3 = (month2-1, year2) if month2 != 1 else (12, now.year-1)
    prev_month_3 = now.replace(day=1, month=month3, year=year3).strftime("%m-%Y")


    files = glob.glob(file_path + "Исходники для наличия компании\\*.csv")
    

    df1 = pd.DataFrame()
    for k in range(0, len(files)):
        df_af = pd.read_csv(files[k], delimiter=';', dtype='unicode', low_memory=False, thousands=' ',
                            usecols=['Артикул (доп)', 'Склад(Название)',
                        'Фактический остаток', 'Заказано',
                        'Прогноз спроса', 'ЗакупочнаяЦена',
                        'Продажи за ' + prev_month_1,
                        'Продажи за ' + prev_month_2,
                        'Продажи за ' + prev_month_3,
                        'Поставщик для заказа (Название)'])

        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(' ', '')
        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(',', '.')
        df_af['Заказано'] = df_af['Заказано'].str.replace(' ', '')
        df_af['Заказано'] = df_af['Заказано'].str.replace(',', '.')
        df_af['Фактический остаток'] = pd.to_numeric(df_af['Фактический остаток'], errors='coerce')
        df_af['Заказано'] = pd.to_numeric(df_af['Заказано'], errors='coerce')
        df_af['Прогноз спроса'] = pd.to_numeric(df_af['Прогноз спроса'], errors='coerce')
        df_af['Артикул (доп)'] = pd.to_numeric(df_af['Артикул (доп)'], errors='coerce')
        df_af['ЗакупочнаяЦена'] = pd.to_numeric(df_af['ЗакупочнаяЦена'], errors='coerce')
        df1 = pd.concat([df1, df_af])


    sales_col = ['Продажи за ' + prev_month_1, 'Продажи за ' + prev_month_2, 'Продажи за ' + prev_month_3]
    df1[sales_col] = df1[sales_col].apply(pd.to_numeric, errors='coerce')
    df1['Продажи 3мес'] = df1[sales_col].sum(axis=1)
    df1 = df1.dropna(subset=['Артикул (доп)'])
    
	
    df_AL = df1[df1['Склад(Название)'] == 'Александровка']

    # определяем сколько не хватает для каждого склада
    df2 = df1.copy()
    df2.reset_index(drop=True, inplace=True)
    df2.loc[df2['Склад(Название)'] != 'Александровка', 'Не хватает на остатке'] =\
        (df2['Фактический остаток'] + df2['Заказано'] - df2['Прогноз спроса'])

    df2 = df2[df2['Не хватает на остатке'] < 0]
    df3 = df2.groupby('Артикул (доп)').agg(
        Не_хватает_на_остатке=('Не хватает на остатке', 'sum')).reset_index()
		

    # вся компания
    df1 = df1.groupby('Артикул (доп)').agg(
        Фактический_остаток=('Фактический остаток', 'sum'),
        Прогноз_спроса=('Прогноз спроса', 'sum'),
        Общие_продажи=('Продажи 3мес', 'sum'),
        ЗакупочнаяЦена=('ЗакупочнаяЦена', 'first')  # Используем 'first' для первого значения
    ).reset_index()

    df1 = df1[df1['ЗакупочнаяЦена'].notna()]

    df1['Сумма общих продаж'] = df1['Общие_продажи'] * df1['ЗакупочнаяЦена']

    df1['Доля ABC(Выручка)ОПТ'] = df1['Сумма общих продаж'] / df1['Сумма общих продаж'].sum()
    df1 = df1.sort_values('Доля ABC(Выручка)ОПТ')
    df1['Накопительная долявырОПТ'] = df1['Доля ABC(Выручка)ОПТ'].cumsum() * 100
    df1.loc[df1['Сумма общих продаж'] == 0, 'Накопительная долявырОПТ'] = 0
    df1.loc[df1['Накопительная долявырОПТ'] >= 20, 'АВСвырОПТ'] = 'A'
    df1.loc[(df1['Накопительная долявырОПТ'] < 20) & (df1['Накопительная долявырОПТ'] >= 5), 'АВСвырОПТ'] = 'B'
    df1.loc[(df1['Накопительная долявырОПТ'] < 5) & (df1['Накопительная долявырОПТ'] > 0), 'АВСвырОПТ'] = 'C'
    df1.loc[df1['Накопительная долявырОПТ'] == 0, 'АВСвырОПТ'] = 'D'
    df1['Доля ABC(Продажи, ед.)ОПТ'] = df1['Прогноз_спроса'] / df1['Прогноз_спроса'].sum()
    df1 = df1.sort_values('Доля ABC(Продажи, ед.)ОПТ')
    df1['Накопительная доляштОПТ'] = df1['Доля ABC(Продажи, ед.)ОПТ'].cumsum() * 100
    df1.loc[df1['Прогноз_спроса'] == 0, 'Накопительная доляштОПТ'] = 0
    df1.loc[df1['Накопительная доляштОПТ'] >= 20, 'АВСштОПТ'] = '-A'
    df1.loc[(df1['Накопительная доляштОПТ'] < 20) & (df1['Накопительная доляштОПТ'] >= 5), 'АВСштОПТ'] = '-B'
    df1.loc[(df1['Накопительная доляштОПТ'] < 5) & (df1['Накопительная доляштОПТ'] > 0), 'АВСштОПТ'] = '-C'
    df1.loc[df1['Накопительная доляштОПТ'] == 0, 'АВСштОПТ'] = '-D'
    df1['ABC'] = df1['АВСвырОПТ'] + df1['АВСштОПТ']


    list_ABC_2 = ['A-A', 'A-B', 'A-C', 'B-A', 'B-B', 'C-A']

    df1 = df1[df1['ABC'].str.contains('|'.join(list_ABC_2))]
    df1['Наличие'] = (
                len(df1[df1['Фактический_остаток'] >= df1['Прогноз_спроса']])/len(df1))
    df1 = df1.merge(df_AL[['Артикул (доп)', 'Фактический остаток', 'Прогноз спроса']], on='Артикул (доп)', how='left')
    df1 = df1.merge(df3, on='Артикул (доп)', how='left')
    df1[['Фактический остаток', 'Прогноз спроса', 'Не_хватает_на_остатке']] = (
        df1[['Фактический остаток', 'Прогноз спроса','Не_хватает_на_остатке']].fillna(0))
    df1['для_Наличие_компания'] = df1['Фактический остаток'] - df1['Прогноз спроса'] + df1['Не_хватает_на_остатке']
    df1['Наличие_компания'] = (
                len(df1[df1['для_Наличие_компания'] >= 0])/len(df1))
    l1 = len(df1)
    l2 = len(df1[df1['для_Наличие_компания'] >= 0])
    print(len(df1), len(df1[df1['для_Наличие_компания'] >= 0]))

    df1 = df1[['Наличие_компания']]
    df1 = df1.drop_duplicates(subset=['Наличие_компания'])
    nalichie = df1['Наличие_компания'].iloc[0]

    now = datetime.now()
    now = now.strftime("%d-%m-%Y")


    df1.to_excel(file_path + 'Наличие вся компания ' + now + '.xlsx', index=False)

    df_AL = df_AL.groupby('Артикул (доп)').agg(
        Фактический_остаток=('Фактический остаток', 'sum'),
        Прогноз_спроса=('Прогноз спроса', 'sum'),
        Общие_продажи=('Продажи 3мес', 'sum'),
        ЗакупочнаяЦена=('ЗакупочнаяЦена', 'first')  # Используем 'first' для первого значения
    ).reset_index()

    df_AL = df_AL[df_AL['ЗакупочнаяЦена'].notna()]

    df_AL['Сумма общих продаж'] = df_AL['Общие_продажи'] * df_AL['ЗакупочнаяЦена']

    df_AL['Доля ABC(Выручка)ОПТ'] = df_AL['Сумма общих продаж'] / df_AL['Сумма общих продаж'].sum()
    df_AL = df_AL.sort_values('Доля ABC(Выручка)ОПТ')
    df_AL['Накопительная долявырОПТ'] = df_AL['Доля ABC(Выручка)ОПТ'].cumsum() * 100
    df_AL.loc[df_AL['Сумма общих продаж'] == 0, 'Накопительная долявырОПТ'] = 0
    df_AL.loc[df_AL['Накопительная долявырОПТ'] >= 20, 'АВСвырОПТ'] = 'A'
    df_AL.loc[(df_AL['Накопительная долявырОПТ'] < 20) & (df_AL['Накопительная долявырОПТ'] >= 5), 'АВСвырОПТ'] = 'B'
    df_AL.loc[(df_AL['Накопительная долявырОПТ'] < 5) & (df_AL['Накопительная долявырОПТ'] > 0), 'АВСвырОПТ'] = 'C'
    df_AL.loc[df_AL['Накопительная долявырОПТ'] == 0, 'АВСвырОПТ'] = 'D'
    df_AL['Доля ABC(Продажи, ед.)ОПТ'] = df_AL['Прогноз_спроса'] / df_AL['Прогноз_спроса'].sum()
    df_AL = df_AL.sort_values('Доля ABC(Продажи, ед.)ОПТ')
    df_AL['Накопительная доляштОПТ'] = df_AL['Доля ABC(Продажи, ед.)ОПТ'].cumsum() * 100
    df_AL.loc[df_AL['Прогноз_спроса'] == 0, 'Накопительная доляштОПТ'] = 0
    df_AL.loc[df_AL['Накопительная доляштОПТ'] >= 20, 'АВСштОПТ'] = '-A'
    df_AL.loc[(df_AL['Накопительная доляштОПТ'] < 20) & (df_AL['Накопительная доляштОПТ'] >= 5), 'АВСштОПТ'] = '-B'
    df_AL.loc[(df_AL['Накопительная доляштОПТ'] < 5) & (df_AL['Накопительная доляштОПТ'] > 0), 'АВСштОПТ'] = '-C'
    df_AL.loc[df_AL['Накопительная доляштОПТ'] == 0, 'АВСштОПТ'] = '-D'
    df_AL['ABC'] = df_AL['АВСвырОПТ'] + df_AL['АВСштОПТ']


    list_ABC_2 = ['A-A', 'A-B', 'A-C', 'B-A', 'B-B', 'C-A']

    df_AL = df_AL[df_AL['ABC'].str.contains('|'.join(list_ABC_2))]
    df_AL['Наличие'] = (
                len(df_AL[df_AL['Фактический_остаток'] >= df_AL['Прогноз_спроса']])/len(df_AL))
    


    df_AL = df_AL[['Наличие']]
    df_AL = df_AL.drop_duplicates(subset=['Наличие'])
    

    
    df_AL.to_excel(file_path + 'Наличие Александровка ' + now + '.xlsx', index=False)

    add_message('Наличие компания ' + str(f'{nalichie*100:.1f}') + ' ' +  str(l2) + '/' +  str(l1))



def nalichie_comp_RF(file_path, add_message):
    # ... весь оригинальный код функции nalichie_comp_RF ...
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

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

    t = time()
    # add_message('Стартуем!')
    # add_message('Считываем файлы')

    now = datetime.now() - timedelta(days=1)
    now_month = now.strftime("%m-%Y")
    month1, year1 = (now.month-1, now.year) if now.month != 1 else (12, now.year-1)
    prev_month_1 = now.replace(day=1, month=month1, year=year1).strftime("%m-%Y")
    month2, year2 = (month1-1, year1) if month1 != 1 else (12, now.year-1)
    prev_month_2 = now.replace(day=1, month=month2, year=year2).strftime("%m-%Y")
    month3, year3 = (month2-1, year2) if month2 != 1 else (12, now.year-1)
    prev_month_3 = now.replace(day=1, month=month3, year=year3).strftime("%m-%Y")


    files = glob.glob(file_path + "Исходники для наличия компании\\*.csv")
    

    df1 = pd.DataFrame()
    for k in range(0, len(files)):
        df_af = pd.read_csv(files[k], delimiter=';', dtype='unicode', low_memory=False, thousands=' ',
                            usecols=['Артикул (доп)', 'Склад(Название)',
                        'Фактический остаток', 'Заказано',
                        'Прогноз спроса', 'ЗакупочнаяЦена',
                        'Продажи за ' + prev_month_1,
                        'Продажи за ' + prev_month_2,
                        'Продажи за ' + prev_month_3,
                        'Поставщик для заказа (Название)'])

        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(' ', '')
        df_af['Фактический остаток'] = df_af['Фактический остаток'].str.replace(',', '.')
        df_af['Заказано'] = df_af['Заказано'].str.replace(' ', '')
        df_af['Заказано'] = df_af['Заказано'].str.replace(',', '.')
        df_af['Фактический остаток'] = pd.to_numeric(df_af['Фактический остаток'], errors='coerce')
        df_af['Заказано'] = pd.to_numeric(df_af['Заказано'], errors='coerce')
        df_af['Прогноз спроса'] = pd.to_numeric(df_af['Прогноз спроса'], errors='coerce')
        df_af['Артикул (доп)'] = pd.to_numeric(df_af['Артикул (доп)'], errors='coerce')
        df_af['ЗакупочнаяЦена'] = pd.to_numeric(df_af['ЗакупочнаяЦена'], errors='coerce')
        df1 = pd.concat([df1, df_af])


    sales_col = ['Продажи за ' + prev_month_1, 'Продажи за ' + prev_month_2, 'Продажи за ' + prev_month_3]
    df1[sales_col] = df1[sales_col].apply(pd.to_numeric, errors='coerce')
    df1['Продажи 3мес'] = df1[sales_col].sum(axis=1)
    df1 = df1.dropna(subset=['Артикул (доп)'])

    df_AL = df1[df1['Склад(Название)'] == 'Александровка']
    

    # определяем сколько не хватает для каждого склада
    df2 = df1.copy()
    df2.reset_index(drop=True, inplace=True)
    df2.loc[df2['Склад(Название)'] != 'Александровка', 'Не хватает на остатке'] =\
        (df2['Фактический остаток'] + df2['Заказано'] - df2['Прогноз спроса'])

    df2 = df2[df2['Не хватает на остатке'] < 0]
    df3 = df2.groupby('Артикул (доп)').agg(
        Не_хватает_на_остатке=('Не хватает на остатке', 'sum')).reset_index()
        
    df_not_RF = pd.read_excel(file_path + 'Поставщики не РФ.xlsx', engine="calamine")
    not_RF_list = df_not_RF['Поставщик'].tolist()
    
    print(len(df1))
    df1 = df1[~df1['Поставщик для заказа (Название)'].isin(not_RF_list)]
    print(len(df1))
    # df1.to_excel("123.xlsx", index=False)

    # вся компания
    df1 = df1.groupby('Артикул (доп)').agg(
        Фактический_остаток=('Фактический остаток', 'sum'),
        Прогноз_спроса=('Прогноз спроса', 'sum'),
        Общие_продажи=('Продажи 3мес', 'sum'),
        ЗакупочнаяЦена=('ЗакупочнаяЦена', 'first')  # Используем 'first' для первого значения
    ).reset_index()

    df1 = df1[df1['ЗакупочнаяЦена'].notna()]

    df1['Сумма общих продаж'] = df1['Общие_продажи'] * df1['ЗакупочнаяЦена']

    df1['Доля ABC(Выручка)ОПТ'] = df1['Сумма общих продаж'] / df1['Сумма общих продаж'].sum()
    df1 = df1.sort_values('Доля ABC(Выручка)ОПТ')
    df1['Накопительная долявырОПТ'] = df1['Доля ABC(Выручка)ОПТ'].cumsum() * 100
    df1.loc[df1['Сумма общих продаж'] == 0, 'Накопительная долявырОПТ'] = 0
    df1.loc[df1['Накопительная долявырОПТ'] >= 20, 'АВСвырОПТ'] = 'A'
    df1.loc[(df1['Накопительная долявырОПТ'] < 20) & (df1['Накопительная долявырОПТ'] >= 5), 'АВСвырОПТ'] = 'B'
    df1.loc[(df1['Накопительная долявырОПТ'] < 5) & (df1['Накопительная долявырОПТ'] > 0), 'АВСвырОПТ'] = 'C'
    df1.loc[df1['Накопительная долявырОПТ'] == 0, 'АВСвырОПТ'] = 'D'
    df1['Доля ABC(Продажи, ед.)ОПТ'] = df1['Прогноз_спроса'] / df1['Прогноз_спроса'].sum()
    df1 = df1.sort_values('Доля ABC(Продажи, ед.)ОПТ')
    df1['Накопительная доляштОПТ'] = df1['Доля ABC(Продажи, ед.)ОПТ'].cumsum() * 100
    df1.loc[df1['Прогноз_спроса'] == 0, 'Накопительная доляштОПТ'] = 0
    df1.loc[df1['Накопительная доляштОПТ'] >= 20, 'АВСштОПТ'] = '-A'
    df1.loc[(df1['Накопительная доляштОПТ'] < 20) & (df1['Накопительная доляштОПТ'] >= 5), 'АВСштОПТ'] = '-B'
    df1.loc[(df1['Накопительная доляштОПТ'] < 5) & (df1['Накопительная доляштОПТ'] > 0), 'АВСштОПТ'] = '-C'
    df1.loc[df1['Накопительная доляштОПТ'] == 0, 'АВСштОПТ'] = '-D'
    df1['ABC'] = df1['АВСвырОПТ'] + df1['АВСштОПТ']


    list_ABC_2 = ['A-A', 'A-B', 'A-C', 'B-A', 'B-B', 'C-A']

    df1 = df1[df1['ABC'].str.contains('|'.join(list_ABC_2))]
    df1['Наличие'] = (
                len(df1[df1['Фактический_остаток'] >= df1['Прогноз_спроса']])/len(df1))
    df1 = df1.merge(df_AL[['Артикул (доп)', 'Фактический остаток', 'Прогноз спроса']], on='Артикул (доп)', how='left')
    df1 = df1.merge(df3, on='Артикул (доп)', how='left')
    df1[['Фактический остаток', 'Прогноз спроса', 'Не_хватает_на_остатке']] = (
        df1[['Фактический остаток', 'Прогноз спроса','Не_хватает_на_остатке']].fillna(0))
    df1['для_Наличие_компания'] = df1['Фактический остаток'] - df1['Прогноз спроса'] + df1['Не_хватает_на_остатке']
    df1['Наличие_компания'] = (
                len(df1[df1['для_Наличие_компания'] >= 0])/len(df1))
    l1 = len(df1)
    l2 = len(df1[df1['для_Наличие_компания'] >= 0])
    print(len(df1), len(df1[df1['для_Наличие_компания'] >= 0]))

    df1 = df1[['Наличие_компания']]
    df1 = df1.drop_duplicates(subset=['Наличие_компания'])
    nalichie = df1['Наличие_компания'].iloc[0]

    now = datetime.now()
    now = now.strftime("%d-%m-%Y")


    df1.to_excel(file_path + 'Наличие вся компания РФ' + now + '.xlsx', index=False)

    

    add_message('Наличие компания по поставщикам РФ ' +  str(f'{nalichie*100:.1f}') + ' ' +  str(l2) + '/' +  str(l1))



