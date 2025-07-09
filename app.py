import streamlit as st
import pandas as pd
from processor import process_data
from utils import save_excel_with_timestamp, read_excel_lines

st.set_page_config(layout="wide")

st.title("📦 Компоновка линий в распределительные коробки и шкафы")

st.subheader("🔧 Настройки компоновки")
max_lines_per_box = st.number_input("Максимум линий в одной распределительной коробке", min_value=1, value=6, step=1)
breaker_nominal = st.number_input("Номинал автомата для распределительной коробки (А)", min_value=0.1, value=25.0, step=0.1)
max_pipe = st.number_input("Максимум коробок в шкафу (Pipe Control)", min_value=1, value=8, step=1)
max_limiter = st.number_input("Максимум коробок в шкафу (Limiter)", min_value=1, value=6, step=1)
rk_prefix = st.text_input("Префикс для распределительных коробок", value="РК")
panel_prefix = st.text_input("Префикс для шкафов", value="Шкаф")

uploaded_file = st.file_uploader("📂 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    df = read_excel_lines(uploaded_file)
    st.success("Файл успешно загружен!")

    st.dataframe(df.head(), use_container_width=True)

    if st.button("🚀 Выполнить компоновку"):
        result_dfs = process_data(
            df,
            max_lines_per_box,
            breaker_nominal,
            max_pipe,
            max_limiter,
            rk_prefix,
            panel_prefix,
        )

        output_file = save_excel_with_timestamp(result_dfs)

        st.success("✅ Компоновка завершена!")
        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Скачать результат",
                data=f,
                file_name=output_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
