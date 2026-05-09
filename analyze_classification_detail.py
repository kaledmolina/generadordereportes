import pandas as pd
import json

# Load the Excel file
df = pd.read_excel('/workspace/user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Check columns
print("=== COLUMNAS EN EL EXCEL ===")
for col in df.columns:
    print(f"  - {col}")

# Check unique values in key columns
print("\n=== VALORES ÚNICOS EN 'Solicitud Suscriptor' ===")
print(df['Solicitud Suscriptor'].value_counts().head(30))

print("\n=== VALORES ÚNICOS EN 'Solución Técnico' ===")
print(df['Solución Técnico'].value_counts().head(50))

# Check classification logic
print("\n=== CLASIFICACIÓN POR 'Solicitud Suscriptor' ===")
solicitud_counts = df['Solicitud Suscriptor'].value_counts()
for val, count in solicitud_counts.items():
    print(f"  {val}: {count}")