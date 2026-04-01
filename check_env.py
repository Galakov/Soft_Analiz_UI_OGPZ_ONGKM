import sys
import pandas as pd
sys.path.append("/Users/gall/Documents/1. Progect/Soft_Finding_differences_in_UI")
from Conv.convert_ui_lists import UIListConverter

converter = UIListConverter("/Users/gall/Documents/1. Progect/Soft_Finding_differences_in_UI/Правила преобразований пречней УИ.xlsx")
input_file = "/Users/gall/Documents/1. Progect/Soft_Analiz_UI_OGPZ_ONGKM/УИ_СЗСК_Перечни_СТО 101.1-2015_12.01.2026.xlsx"

sdb_df, msg = converter.process_file(input_file)
if sdb_df is not None:
    print("SDB Extracted (First 20 items):")
    print(sdb_df[["Наименование УИ", "Среда"]].head(20))
    print("\nSDB Total specific counts of Среда:")
    print(sdb_df["Среда"].value_counts())
else:
    print("Failed extracting SDB")
    
# Let's see the rules mapping for Сургутский ЗСК
rules_df = converter.rules_df
surgut_rules = rules_df[rules_df['Филиал'].astype(str).str.contains('Сургут', case=False, na=False)]
print("\nRules for Surgut (Лист and Среда):")
print(surgut_rules[['Лист', 'Разделение', 'Среда']])
