import glob
import pandas as pd
import re
import math
import os
from datetime import datetime
import openpyxl
from collections import defaultdict
from python_calamine.pandas import pandas_monkeypatch

pandas_monkeypatch()

def process_transport_data(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    def parse_volume(value):
        if pd.isna(value) or value == '':
            return None
        value = str(value).strip()
        matches = re.findall(r'\d+', value)
        if matches:
            return [int(match) for match in matches]
        return None

    def get_target_day(grafik_text, volume):
        if pd.isna(grafik_text) or not isinstance(grafik_text, str):
            return None
        grafik_text = grafik_text.lower()

        days_map = {
            'понедельник': 'Пн',
            'вторник': 'Вт',
            'среду': 'Ср',
            'четверг': 'Чт',
            'пятницу': 'Пт',
            'субботу': 'Сб',
            'воскресенье': 'Вс'
        }

        if 'на ' in grafik_text and 'если' not in grafik_text:
            for day_name, day_short in days_map.items():
                if day_name in grafik_text:
                    return day_short
            return None

        parts = [p.strip() for p in grafik_text.split(';')]
        for part in parts:
            if str(volume) in part:
                for day_name, day_short in days_map.items():
                    if day_name in part:
                        return day_short
        return None

    grafik_file = glob.glob(file_path + "Исходники\\график сборки.xlsx")[0]
    fact_file = glob.glob(file_path + "Исходники\\график доставок.xlsx")[0]
    print("Файл графика:", grafik_file)
    print("Файл факта:", fact_file)

    grafik = pd.read_excel(grafik_file, sheet_name="График сборки", header=None)
    fact = pd.read_excel(fact_file, header=None)

    # Загружаем вкладку "Дальние"
    try:
        dalnie_df = pd.read_excel(grafik_file, sheet_name="Дальние")
        dalnie_prk = set(dalnie_df.iloc[:, 0].dropna().astype(int).tolist())
        print("Дальние магазины:", dalnie_prk)
    except Exception as e:
        dalnie_prk = set()
        print("⚠ Не удалось загрузить вкладку 'Дальние':", e)

    gorodskie = {1, 2, 3, 9, 19, 26, 30, 46, 54}
    podolskie = {11, 20, 36, 38, 41, 42, 43, 48, 50, 51, 52, 56}
    regionalnye = {
        4, 5, 6, 7, 8, 10, 12, 14, 15, 16, 17, 18, 21, 23, 24, 25,
        27, 28, 29, 31, 32, 33, 37, 39, 40, 53, 55, 57
    }

    weekly_results = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    shop_schedule = defaultdict(lambda: defaultdict(list))  # для второй вкладки

    days_order = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']
    day_to_num = {day: i + 1 for i, day in enumerate(days_order)}

    # Определение недель
    weeks = []
    current_week = {}
    week_num = 1
    max_weeks = 2

    all_days = []
    for col in range(1, len(fact.columns)):
        day = str(fact.iloc[0, col]).strip()
        if day in days_order:
            all_days.append((day, col))

    for i, (day, col) in enumerate(all_days):
        current_week[day] = col
        if day == 'Вс' or i == len(all_days) - 1 or len(weeks) >= max_weeks - 1:
            if current_week:
                weeks.append((f'Неделя {week_num}', current_week))
                week_num += 1
                current_week = {}
                if len(weeks) >= max_weeks:
                    break

    full_weeks = []
    for week in weeks:
        if len(week[1]) == 7 or week == weeks[-1]:
            full_weeks.append(week)
    weeks = full_weeks

    # Обработка данных
    for idx in range(1, len(grafik)):
        row = grafik.iloc[idx]
        prk_raw = row[0]

        if pd.isna(prk_raw):
            continue

        try:
            prk = int(prk_raw)
        except:
            continue

        if prk in gorodskie:
            prefix = 'Г'
        elif prk in podolskie:
            prefix = 'П'
        elif prk in regionalnye:
            prefix = 'В'
        else:
            continue

        fact_row = None
        for f_idx in range(1, len(fact)):
            try:
                if int(fact.iloc[f_idx, 0]) == prk:
                    fact_row = fact.iloc[f_idx]
                    break
            except:
                continue
        if fact_row is None:
            continue

        for week_name, week_days in weeks:
            if week_name == weeks[-1][0]:
                week_days = {day: col for day, col in week_days.items() if col < len(fact.columns)}

            for day, col in week_days.items():
                cell_value = fact_row[col]
                volumes = parse_volume(cell_value)
                if not volumes:
                    continue

                for volume in volumes:
                    if prefix == 'Г' and volume == 10:
                        machine_type = 'Г10'
                        count = 1
                    elif prefix == 'П' and volume == 14:
                        machine_type = 'П14'
                        count = 1
                    elif prefix == 'П' and (volume == 17 or volume == 35):
                        machine_type = 'П35'
                        count = 0.5 if volume == 17 else 1
                    elif prefix == 'В' and volume == 14:
                        machine_type = 'В14'
                        count = 1
                    elif prefix == 'В' and (volume == 17 or volume == 35):
                        machine_type = 'В35'
                        count = 0.5 if volume == 17 else 1
                    else:
                        continue

                    grafik_day_col = day_to_num[day]
                    grafik_day_plan = row[grafik_day_col]
                    target_day = get_target_day(grafik_day_plan, volume)
                    if not target_day:
                        continue

                    assembly_day = None
                    for d in days_order:
                        grafik_col = day_to_num[d]
                        grafik_text = row[grafik_col]
                        if get_target_day(grafik_text, volume) == day:
                            assembly_day = d
                            break
                    if not assembly_day:
                        continue

                    if assembly_day == 'Вс':
                        if week_name == 'Неделя 1':
                            weekly_results['Вс0'][machine_type]['Вс'] += count
                            shop_schedule[machine_type]['Вс0'].append(prk)
                        continue

                    week_label = f"{assembly_day}{weeks.index((week_name, week_days)) + 1}"
                    weekly_results[week_name][machine_type][assembly_day] += count
                    shop_schedule[machine_type][week_label].append(prk)

    # Формируем таблицу: недели → колонки
    machine_types = ['Г10', 'П14', 'П35', 'В14', 'В35']
    combined_results = defaultdict(lambda: defaultdict(float))

    # Вс0
    for m_type in machine_types:
        val = weekly_results.get('Вс0', {}).get(m_type, {}).get('Вс', 0)
        combined_results[m_type]['Вс0'] = round(val, 1)

    # Остальные недели (Пн1..Вс1, Пн2..)
    week_counter = 1
    for week_name, week_data in sorted(weekly_results.items(), key=lambda x: x[0]):
        if week_name == 'Вс0':
            continue
        for m_type in machine_types:
            for day in days_order:
                val = week_data.get(m_type, {}).get(day, 0)
                combined_results[m_type][f"{day}{week_counter}"] = round(val, 1)
        week_counter += 1

    df = pd.DataFrame.from_dict(combined_results, orient="index").reset_index()
    df.rename(columns={"index": "Тип"}, inplace=True)

    # Считаем итого по строкам
    day_columns = [col for col in df.columns if col != "Тип"]
    if day_columns:
        df["Итого"] = df[day_columns].sum(axis=1)
    else:
        df["Итого"] = 0

    # Добавляем строку "Дальние"
    dalnie_row = {"Тип": "Дальние"}
    for col in day_columns:
        shops = []
        # теперь берём только машины 35 м3
        for m_type in ["П35", "В35"]:
            prk_list = shop_schedule[m_type].get(col, [])
            shops.extend(prk_list)
        dalnie_count = sum(1 for s in shops if s in dalnie_prk)
        dalnie_row[col] = dalnie_count
    dalnie_row["Итого"] = sum(dalnie_row[col] for col in day_columns)
    df = pd.concat([df, pd.DataFrame([dalnie_row])], ignore_index=True)

    # ИТОГО по столбцам
    totals_row = {"Тип": "ИТОГО"}
    for col in day_columns:
        p14_total = df.loc[df["Тип"] == "П14", col].sum() if "П14" in df["Тип"].values else 0
        v14_total = df.loc[df["Тип"] == "В14", col].sum() if "В14" in df["Тип"].values else 0
        p35_total = df.loc[df["Тип"] == "П35", col].sum() if "П35" in df["Тип"].values else 0
        v35_total = df.loc[df["Тип"] == "В35", col].sum() if "В35" in df["Тип"].values else 0

        total14 = int(round(p14_total + v14_total))
        total35 = p35_total + v35_total

        if total14 == 0:
            result14 = 0
        elif total14 == 1:
            result14 = 1
        else:
            result14 = total14 // 2

        result35 = int(math.floor(total35))

        total_machines = result14 + result35
        totals_row[col] = int(total_machines)

    totals_row["Итого"] = sum(totals_row[col] for col in day_columns) if day_columns else 0
    df = pd.concat([df, pd.DataFrame([totals_row])], ignore_index=True)

    # --- Вторая вкладка с магазинами ---
    shop_df = pd.DataFrame(index=machine_types)
    all_day_labels = list(day_columns)
    for m_type in machine_types:
        for day_label in all_day_labels:
            shops = shop_schedule[m_type].get(day_label, [])
            shop_df.loc[m_type, day_label] = ", ".join(map(str, sorted(set(shops)))) if shops else ""

    # --- Добавляем строку "Дальние" для второй вкладки ---
    dalnie_shops_row = {}
    for day_label in all_day_labels:
        shops = []
        for m_type in ["П35", "В35"]:
            prk_list = shop_schedule[m_type].get(day_label, [])
            shops.extend(prk_list)
        dalnie_shops = [str(s) for s in sorted(set(shops)) if s in dalnie_prk]
        dalnie_shops_row[day_label] = ", ".join(dalnie_shops) if dalnie_shops else ""
    shop_df.loc["Дальние"] = dalnie_shops_row

    now = datetime.now().strftime("%d-%m-%Y")
    output_file = file_path + f'Результат расчета транспорта {now}.xlsx'

    # --- Сохранение в две вкладки ---
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Объемы", index=False)
        shop_df.to_excel(writer, sheet_name="Магазины")

    add_message('Готово! Результат сохранен в файл: ' + output_file)
