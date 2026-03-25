import pandas as pd
import json
import sys

file_path = r"d:\OneDrive - CA Sri Lanka\Salinda Wickramasinghe\Projects\Project\bom tool\SAB Campus Dashboard.xlsx"
try:
    excel_file = pd.ExcelFile(file_path)
    output = {}
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=3)
        df = df.where(pd.notnull(df), None)
        df_dict = {}
        for col in df.columns:
            val = df.iloc[0][col] if len(df) > 0 else None
            if pd.isna(val):
                val = None
            elif hasattr(val, 'isoformat'):
                val = val.isoformat()
            elif isinstance(val, (int, float, str, bool)):
                pass
            else:
                val = str(val)
            df_dict[str(col)] = val
            
        output[sheet_name] = {
            "columns": [str(c) for c in df.columns.tolist()],
            "first_row": df_dict
        }
    with open("out.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)
except Exception as e:
    print("Error:", e)
    sys.exit(1)
