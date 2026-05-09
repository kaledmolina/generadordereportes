import pandas as pd

# Load the Excel file
df = pd.read_excel('/workspace/user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Check the "Clasificación" column
print("=== COLUMNA 'Clasificación' (valores únicos) ===")
print(df['Clasificación'].value_counts())

print("\n=== RELACIÓN TÉCNICO vs CLASIFICACIÓN ===")
# Group by técnico principal and clasificación
tech_class = df.groupby(['Técnico Principal', 'Clasificación']).size().unstack(fill_value=0)
print(tech_class)

print("\n=== DATOS COMPLETOS POR TÉCNICO ===")
for tech in df['Técnico Principal'].unique():
    if pd.notna(tech):
        tech_df = df[df['Técnico Principal'] == tech]
        print(f"\n--- {tech} ({len(tech_df)} órdenes) ---")
        class_counts = tech_df['Clasificación'].value_counts()
        for cls, count in class_counts.items():
            print(f"  {cls}: {count}")