import pandas as pd
import sys

try:
    df = pd.read_excel('/home/tomo/work_office/zaimu_data_dl_20260223_1832.xlsx', sheet_name=0, header=None)
    print("Shape:", df.shape)
    for i, row in df.head(50).iterrows():
        print(f"Row {i}:", row.tolist()[:10]) # 先頭10列だけ表示
except Exception as e:
    print("Error:", e)
