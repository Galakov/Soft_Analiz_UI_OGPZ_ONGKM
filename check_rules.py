import pandas as pd
import sys
try:
    df = pd.read_excel(r'd:\1. Progects\Soft_Analiz_UI_OGPZ_ONGKM\analytics_ui\Правила названия столбцов.xlsx', engine='openpyxl')
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")
except Exception as e:
    print(e)
