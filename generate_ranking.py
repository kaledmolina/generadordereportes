import pandas as pd
import json

# Load the Excel file
df = pd.read_excel('/workspace/user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Get the "Solicitud Suscriptor" values for each technician
print("=== RANKING DE TÉCNICOS CON CLASIFICACIÓN REAL ===")

# Get all unique technicians
tecnicos = df['Técnico Principal'].dropna().unique()

# Create ranking data
ranking_data = []
for tech in tecnicos:
    if tech in ['Sin Asignar', 'Planta externa', 'Soporte Noc']:
        continue  # Skip non-technician entries
    tech_df = df[df['Técnico Principal'] == tech]
    total = len(tech_df)

    # Get counts by Solicitud Suscriptor
    solicitud_counts = tech_df['Solicitud Suscriptor'].value_counts()

    tech_data = {
        'tecnico': tech,
        'total': total,
        'detalle': {}
    }

    for solicitud, count in solicitud_counts.items():
        tech_data['detalle'][solicitud] = int(count)

    ranking_data.append(tech_data)

# Sort by total
ranking_data.sort(key=lambda x: x['total'], reverse=True)

# Print in a simple format
for i, tech in enumerate(ranking_data, 1):
    print(f"\n{i}. {tech['tecnico']} - TOTAL: {tech['total']}")
    for tipo, count in tech['detalle'].items():
        print(f"   {tipo}: {count}")

# Save to JSON for HTML report
with open('/workspace/ranking_data.json', 'w', encoding='utf-8') as f:
    json.dump(ranking_data, f, ensure_ascii=False, indent=2)

print("\n\n=== DATOS GUARDADOS EN ranking_data.json ===")