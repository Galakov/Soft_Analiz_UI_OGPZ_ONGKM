import pandas as pd
import sys

try:
    df1 = pd.read_excel("/Users/gall/Documents/1. Progect/Soft_Analiz_UI_OGPZ_ONGKM/2026.03.03 УИЖУ Инфотех Tolko-nazvaniia.xlsx")
    surgut_liquid = df1[df1['Филиал'].astype(str).str.contains('Сургут', case=False, na=False)]
    print("ЖИДКИЕ Сургут (всего:", len(surgut_liquid), "):")
    print(surgut_liquid['Наименование УИ'].head(20).tolist())
    
    df2 = pd.read_excel("/Users/gall/Documents/1. Progect/Soft_Analiz_UI_OGPZ_ONGKM/2026.03.03 УИРГ Инфотех Tolko-nazvaniia.xlsx")
    surgut_gas = df2[df2['Филиал'].astype(str).str.contains('Сургут', case=False, na=False)]
    print("\nГАЗОВЫЕ Сургут (всего:", len(surgut_gas), "):")
    print(surgut_gas['Наименование УИ'].head(20).tolist())
    
    # Also check the results file briefly
    res_df = pd.read_excel("/Users/gall/Documents/1. Progect/Soft_Analiz_UI_OGPZ_ONGKM/Результат_сопоставления_СЗСК_2026-03-04_4.8%.xlsx", sheet_name="Сопоставление")
    print("\nRESULTS FILE (Failed matches sorted):")
    low_matches = res_df[res_df['Процент совпадения (%)'] < 70]
    print(low_matches[['Наименование УИ', 'Похожее название (Инфотех)', 'Процент совпадения (%)']].head(20))
except Exception as e:
    print(f"Error: {e}")
