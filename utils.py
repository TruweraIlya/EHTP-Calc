import pandas as pd
from pathlib import Path
from datetime import datetime

def read_excel_lines(uploaded_file):
    return pd.read_excel(uploaded_file, header=7)

def save_excel_with_timestamp(dataframes):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = Path(f"Результат_компоновки_{timestamp}.xlsx")
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        sheet_names = ["Ведомость распределения", "Ведомость РКЭО", "Ведомость ШУЭО"]
        for df, name in zip(dataframes, sheet_names):
            df.to_excel(writer, index=False, sheet_name=name)
    return file_path
