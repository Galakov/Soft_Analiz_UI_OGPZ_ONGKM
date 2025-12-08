import pandas as pd
import os
import sys

def resource_path(relative_path):
    base_path = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base_path, relative_path)
    return path

try:
    rules_file = resource_path("analytics_ui/Правила названия столбцов.xlsx")
    print(f"Reading: {rules_file}")
    
    rules_df = pd.read_excel(rules_file, engine='openpyxl')
    print(f"Shape: {rules_df.shape}")
    print(f"Columns: {rules_df.columns.tolist()}")
    
    for i, row in rules_df.iterrows():
        if i > 5: break # Only check first few meaningful rows
        
        print(f"\nRow {i}:")
        # Print first 10 cols
        vals = row.tolist()[:10]
        for idx, v in enumerate(vals):
            print(f"  Col {idx}: {v} (Type: {type(v)})")
            
        if len(row) >= 4:
            new_name = str(row.iloc[2]).strip()
            # Simulation of logic
            units = ""
            if len(row) >= 8 and pd.notna(row.iloc[7]):
                units = str(row.iloc[7]).strip()
                print(f"  Existing units (Col 7): '{units}'")
            else:
                 print(f"  No existing units in Col 7 (Len: {len(row)})")
            
            if not units:
                name_lower = new_name.lower()
                inferred = ""
                if "перепад давления" in name_lower:
                    inferred = "кгс/см2"
                elif "расход" in name_lower:
                    inferred = "тыс. м3/ч"
                elif "температура" in name_lower:
                    inferred = "°C"
                
                if inferred:
                    print(f"  MATCH! '{new_name}' -> Inferred Unit: '{inferred}'")
                else:
                    print(f"  NO MATCH. '{new_name}'")

except Exception as e:
    print(f"Error: {e}")
