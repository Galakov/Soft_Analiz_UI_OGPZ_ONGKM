import pandas as pd

try:
    res_df = pd.read_excel("/Users/gall/Documents/1. Progect/Soft_Analiz_UI_OGPZ_ONGKM/Результат_сопоставления_СЗСК_2026-03-04_4.8%.xlsx", sheet_name="Сопоставление")
    print("\nRESULTS FILE (Failed matches sorted):")
    low_matches = res_df[res_df['Процент совпадения (%)'] < 100]
    for idx, row in low_matches.head(10).iterrows():
        print(f"Orig: {repr(row['Наименование УИ'])}")
        print(f"Match: {repr(row['Похожее название (Инфотех)'])}")
        print(f"Score: {row['Процент совпадения (%)']}")
        print(f"Env: {row['Среда']}\n")
except Exception as e:
    print(f"Error: {e}")
