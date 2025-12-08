import pandas as pd
import os
import sys

def resource_path(relative_path):
    base_path = os.getcwd()
    path = os.path.join(base_path, relative_path)
    return path

try:
    # Adjust path to where the file is located relative to the root or find it
    file_path = "analytics_ui/Правила названия столбцов.xlsx" 
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        # Try absolute path based on previous ls
        file_path = r"d:\1. Progects\Soft_Analiz_UI_OGPZ_ONGKM\analytics_ui\Правила названия столбцов.xlsx"

    print(f"Reading {file_path}")
    df = pd.read_excel(file_path, engine='openpyxl')
    print("Columns:", df.columns.tolist())
    print("First 5 rows:")
    print(df.head())
    
    print("\nColumn 2 (New Name) unique values:")
    if len(df.columns) > 2:
        print(df.iloc[:, 2].unique())
        
    print("\nChecking Min/Max columns (5 and 6):")
    if len(df.columns) > 6:
        print(df.iloc[:, [2, 5, 6]].head(10))

except Exception as e:
    print(f"Error: {e}")
