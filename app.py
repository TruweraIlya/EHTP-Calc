import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime

# Пользовательские настройки
st.title("📦 Компоновка EHT линий по распределительным коробкам и шкафам")

with st.sidebar:
    st.header("⚙️ Настройки компоновки")
    max_lines_per_rk = st.number_input("Макс. количество линий в одной РК", min_value=1, value=4)
    rk_circuit_breaker_rating = st.number_input("Номинал автомата для РК (А)", min_value=1, value=25)
    max_feeders_pipe_control = st.number_input("Макс. коробок в шкафу (Pipe Control)", min_value=1, value=8)
    max_feeders_limiter = st.number_input("Макс. коробок в шкафу (Limiter)", min_value=1, value=6)
    rk_prefix = st.text_input("Префикс РК", value="RK")
    panel_prefix = st.text_input("Префикс Шкафа", value="P")

uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=7)

    # Получаем нужные данные по индексам
    df_extracted = pd.DataFrame({
        "Название линии": df.iloc[:, 0],      # A
        "Тип процесса": df.iloc[:, 7],        # H
        "Тип управления": df.iloc[:, 16],     # Q
        "Рабочий ток": df.iloc[:, 17],        # R
        "Пусковой ток": df.iloc[:, 18],       # S
        "Нагрузка": df.iloc[:, 19],           # T
        "№ линии": df.iloc[:, 22]             # V
    })

    st.subheader("📄 Предварительный просмотр данных")
    st.dataframe(df_extracted.head(), use_container_width=True)

    rk_counter = 1
    panel_counter = 1
    rk_list = []
    rk_assignments = []
    unassigned_lines = []

    # Группировка по типу процесса
    for process_type in df_extracted["Тип процесса"].unique():
        process_df = df_extracted[df_extracted["Тип процесса"] == process_type].copy()

        # Распределение по РК
        current_rk = []
        current_rk_type = ""
        current_rk_current = 0.0

        for _, row in process_df.iterrows():
            regulation = row["Тип управления"]
            current = row["Рабочий ток"]
            line_name = row["Название линии"]

            regulation_type = "LIMITER" if "limiter" in str(regulation).lower() else "PIPE"

            if current > rk_circuit_breaker_rating:
                rk_assignments.append({
                    "Название линии": line_name,
                    "Тип процесса": process_type,
                    "Тип управления": regulation,
                    "Распред. коробка": f"{rk_prefix}{rk_counter}",
                    "Шкаф": None,
                    "Требует подбора автомата": "Да"
                })
                rk_list.append({
                    "Номер": f"{rk_prefix}{rk_counter}",
                    "Тип": regulation_type,
                    "Линии": [line_name],
                    "Суммарный ток": current,
                    "Количество линий": 1
                })
                rk_counter += 1
                continue

            if (len(current_rk) < max_lines_per_rk) and (current_rk_type in ["", regulation_type]) and (current_rk_current + current <= rk_circuit_breaker_rating):
                current_rk.append(row)
                current_rk_type = regulation_type
                current_rk_current += current
            else:
                if current_rk:
                    lines = [r["Название линии"] for _, r in pd.DataFrame(current_rk).iterrows()]
                    rk_list.append({
                        "Номер": f"{rk_prefix}{rk_counter}",
                        "Тип": current_rk_type,
                        "Линии": lines,
                        "Суммарный ток": current_rk_current,
                        "Количество линий": len(lines)
                    })
                    for r in current_rk:
                        rk_assignments.append({
                            "Название линии": r["Название линии"],
                            "Тип процесса": process_type,
                            "Тип управления": r["Тип управления"],
                            "Распред. коробка": f"{rk_prefix}{rk_counter}",
                            "Шкаф": None,
                            "Требует подбора автомата": ""
                        })
                    rk_counter += 1
                current_rk = [row]
                current_rk_type = regulation_type
                current_rk_current = current

        if current_rk:
            lines = [r["Название линии"] for _, r in pd.DataFrame(current_rk).iterrows()]
            rk_list.append({
                "Номер": f"{rk_prefix}{rk_counter}",
                "Тип": current_rk_type,
                "Линии": lines,
                "Суммарный ток": current_rk_current,
                "Количество линий": len(lines)
            })
            for r in current_rk:
                rk_assignments.append({
                    "Название линии": r["Название линии"],
                    "Тип процесса": process_type,
                    "Тип управления": r["Тип управления"],
                    "Распред. коробка": f"{rk_prefix}{rk_counter}",
                    "Шкаф": None,
                    "Требует подбора автомата": ""
                })
            rk_counter += 1

    # Распределение РК по шкафам
    panels = []
    current_panel = {"Номер": f"{panel_prefix}{panel_counter}", "Коробки": [], "PIPE": 0, "LIMITER": 0, "Суммарный ток": 0.0}

    for rk in rk_list:
        rktype = rk["Тип"]
        allowed = max_feeders_pipe_control if rktype == "PIPE" else max_feeders_limiter

        if current_panel[rktype] < allowed:
            current_panel["Коробки"].append(rk["Номер"])
            current_panel[rktype] += 1
            current_panel["Суммарный ток"] += rk["Суммарный ток"]
        else:
            panels.append(current_panel)
            panel_counter += 1
            current_panel = {"Номер": f"{panel_prefix}{panel_counter}", "Коробки": [rk["Номер"]], "PIPE": 0, "LIMITER": 0, "Суммарный ток": rk["Суммарный ток"]}
            current_panel[rktype] = 1

    if current_panel["Коробки"]:
        panels.append(current_panel)

    rk_to_panel = {rk: panel["Номер"] for panel in panels for rk in panel["Коробки"]}

    for assign in rk_assignments:
        rk = assign["Распред. коробка"]
        assign["Шкаф"] = rk_to_panel.get(rk)

    df_result = pd.DataFrame(rk_assignments)
    df_rk_sheet = pd.DataFrame([{
        "№ РК": rk["Номер"],
        "Тип регулирования": rk["Тип"],
        "Суммарный ток, А": rk["Суммарный ток"],
        "Количество линий": rk["Количество линий"],
        **{f"Линия {i+1}": rk["Линии"][i] for i in range(len(rk["Линии"]))}
    } for rk in rk_list])

    df_panel_sheet = pd.DataFrame([{
        "№ Шкафа": panel["Номер"],
        "Кол-во РК (Pipe)": panel["PIPE"],
        "Кол-во РК (Limiter)": panel["LIMITER"],
        "Суммарный ток, А": panel["Суммарный ток"],
        **{f"РК {i+1}": panel["Коробки"][i] for i in range(len(panel["Коробки"]))}
    } for panel in panels])

    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_result.to_excel(writer, sheet_name="Результат", index=False)
        df_rk_sheet.to_excel(writer, sheet_name="Распред. коробки", index=False)
        df_panel_sheet.to_excel(writer, sheet_name="Шкафы", index=False)

        for sheet in writer.book.worksheets:
            for col in sheet.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                sheet.column_dimensions[col_letter].width = max_length + 3

    st.success("✅ Расчёт завершён")
    st.download_button("📥 Скачать результат", output.getvalue(), file_name=f"eht_result_{now}.xlsx")
