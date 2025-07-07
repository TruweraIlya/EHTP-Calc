import datetime
import pandas as pd
from pathlib import Path

def save_excel_with_timestamp(dataframes_dict):
    now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = Path(f"output_result_{now}.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in dataframes_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_path