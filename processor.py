import pandas as pd
from collections import defaultdict

# Настройки (легко менять)
MAX_LINES_PER_RK = 30
RK_CIRCUIT_BREAKER_RATING = 25
MAX_FEEDERS_PIPE_CONTROL = 8
MAX_FEEDERS_LIMITER = 6

def process_data(df):
    # Индексы колонок
    col_name = 0
    col_process = 7
    col_control = 16
    col_op_current = 17
    col_start_current = 18
    col_load = 19
    col_line_number = 22

    all_rk = []
    rk_counter = 1

    # Компоновка распределительных коробок (по типу процесса отдельно)
    for process_type in ["EP", "EW"]:
        lines = df[df.iloc[:, col_process] == process_type]
        rk_groups = []
        current_group = []
        current_sum = 0

        for _, row in lines.iterrows():
            start_current = row.iloc[col_start_current]
            if pd.isna(start_current):
                continue

            if start_current > RK_CIRCUIT_BREAKER_RATING:
                rk_groups.append([row])
                continue

            if current_sum + start_current <= RK_CIRCUIT_BREAKER_RATING and len(current_group) < MAX_LINES_PER_RK:
                current_group.append(row)
                current_sum += start_current
            else:
                rk_groups.append(current_group)
                current_group = [row]
                current_sum = start_current

        if current_group:
            rk_groups.append(current_group)

        # Формируем распределительные коробки
        for group in rk_groups:
            rk_name = f"RK-{rk_counter}"
            for row in group:
                all_rk.append({
                    "Название линии": row.iloc[col_name],
                    "Тип процесса": row.iloc[col_process],
                    "Тип управления": row.iloc[col_control],
                    "Рабочий ток": row.iloc[col_op_current],
                    "Пусковой ток": row.iloc[col_start_current],
                    "Нагрузка": row.iloc[col_load],
                    "№ линии": row.iloc[col_line_number],
                    "Распред. коробка": rk_name,
                    "Требует подбора автомата": "Да" if row.iloc[col_start_current] > RK_CIRCUIT_BREAKER_RATING else "Нет"
                })
            rk_counter += 1

    # Переводим список в DataFrame
    df_rk = pd.DataFrame(all_rk)

    # Теперь группируем распределительные коробки по шкафам
    rk_info = df_rk.groupby("Распред. коробка").agg({
        "Тип управления": lambda x: x.mode().iloc[0],
        "Рабочий ток": "sum"
    }).reset_index()

    # Определение шкафов
    cabinet_counter = 1
    cabinets = []
    rk_to_cabinet = {}

    pipe_count = 0
    limiter_count = 0

    for _, row in rk_info.iterrows():
        rk = row["Распред. коробка"]
        control = row["Тип управления"]

        control_type = "pipe" if "ambient" in control.lower() or "pipe" in control.lower() else "limiter"

        if (
            (control_type == "pipe" and pipe_count >= MAX_FEEDERS_PIPE_CONTROL) or
            (control_type == "limiter" and limiter_count >= MAX_FEEDERS_LIMITER)
        ):
            # новый шкаф
            cabinet_counter += 1
            pipe_count = 0
            limiter_count = 0

        rk_to_cabinet[rk] = f"Cabinet-{cabinet_counter}"

        if control_type == "pipe":
            pipe_count += 1
        else:
            limiter_count += 1

    # Добавим колонку Шкаф в итоговую таблицу
    df_rk["Шкаф"] = df_rk["Распред. коробка"].map(rk_to_cabinet)

    # Доп. таблица: ведомость РК
    rk_registry = []
    for rk_name, group in df_rk.groupby("Распред. коробка"):
        rk_registry.append({
            "Распред. коробка": rk_name,
            "Линии": list(group["Название линии"]),
            "Суммарный ток": group["Пусковой ток"].sum(),
            "Кол-во линий": len(group)
        })
    df_rk_list = pd.DataFrame(rk_registry)

    # Доп. таблица: ведомость шкафов
    cabinet_registry = []
    for cab_name, group in df_rk.groupby("Шкаф"):
        unique_rk = group["Распред. коробка"].unique()
        rk_types = group.groupby("Распред. коробка")["Тип управления"].agg(lambda x: x.mode().iloc[0])
        pipe = sum(1 for v in rk_types if "pipe" in v.lower() or "ambient" in v.lower())
        limiter = sum(1 for v in rk_types if "limiter" in v.lower())
        cabinet_registry.append({
            "Шкаф": cab_name,
            "Распред. коробки": list(unique_rk),
            "Кол-во RK: pipe/ambient": pipe,
            "Кол-во RK: limiter": limiter,
            "Суммарный ток шкафа": group["Пусковой ток"].sum()
        })
    df_cabinet_list = pd.DataFrame(cabinet_registry)

    return {
        "Результат компоновки": df_rk,
        "Ведомость распределительных коробок": df_rk_list,
        "Ведомость шкафов": df_cabinet_list
    }
