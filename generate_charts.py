import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import os

# Load the Excel file
df = pd.read_excel('/workspace/user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Create charts directory
os.makedirs('/workspace/charts', exist_ok=True)

# Chart colors
COLORS = {
    'primary': '#1C3557',
    'secondary': '#00B4A6',
    'accent1': '#E63946',
    'accent2': '#FF9F1C',
    'accent3': '#2A9D8F',
    'accent4': '#8338EC',
    'light': '#F0F5F8'
}

plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False

# Chart 1: Orders by Solicitud Suscriptor (Pie)
fig, ax = plt.subplots(figsize=(10, 8))
solicitud_counts = df['Solicitud Suscriptor'].value_counts()
colors = ['#1C3557', '#00B4A6', '#E63946', '#FF9F1C', '#2A9D8F', '#8338EC', '#F4A261', '#264653', '#A8DADC', '#457B9D']
wedges, texts, autotexts = ax.pie(solicitud_counts.values, labels=None, autopct='%1.1f%%',
                                   colors=colors[:len(solicitud_counts)], pctdistance=0.75)
for autotext in autotexts:
    autotext.set_fontsize(9)
ax.legend(solicitud_counts.index, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=9)
ax.set_title('Ordenes por Tipo de Solicitud', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig('/workspace/charts/clasificacion_pie.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 2: Orders by Status (Bar)
fig, ax = plt.subplots(figsize=(10, 6))
estado_counts = df['Estado'].value_counts()
bars = ax.bar(estado_counts.index, estado_counts.values, color=COLORS['primary'])
ax.bar_label(bars, padding=3, fontsize=10)
ax.set_ylabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Ordenes por Estado', fontsize=14, fontweight='bold')
ax.set_xticklabels(estado_counts.index, rotation=30, ha='right')
plt.tight_layout()
plt.savefig('/workspace/charts/estado_bar.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 3: Orders by Zone (Bar)
fig, ax = plt.subplots(figsize=(10, 6))
zona_counts = df['Zona Cliente'].value_counts()
bars = ax.bar(zona_counts.index, zona_counts.values, color=COLORS['secondary'])
ax.bar_label(bars, padding=3, fontsize=10)
ax.set_ylabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Ordenes por Zona', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('/workspace/charts/zona_bar.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 4: Top 15 Neighborhoods (Horizontal Bar)
fig, ax = plt.subplots(figsize=(12, 8))
barrio_counts = df['Barrio'].value_counts().head(15)
bars = ax.barh(barrio_counts.index[::-1], barrio_counts.values[::-1], color=COLORS['primary'])
ax.bar_label(bars, padding=3, fontsize=9)
ax.set_xlabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Top 15 Barrios con Mas Ordenes', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('/workspace/charts/barrios_bar.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 5: Technician Ranking with Solicitud breakdown (Stacked Bar)
fig, ax = plt.subplots(figsize=(14, 8))
tecnicos = df['Técnico Principal'].dropna().unique()
tecnicos = [t for t in tecnicos if t not in ['Sin Asignar', 'Planta externa', 'Soporte Noc']]

# Get unique solicitud types for all technicians
all_solicitudes = df['Solicitud Suscriptor'].unique()

# Create data matrix
tech_solicitud = {}
for tech in tecnicos:
    tech_df = df[df['Técnico Principal'] == tech]
    vc = tech_df['Solicitud Suscriptor'].value_counts()
    tech_solicitud[tech] = dict(zip(vc.index, vc.values))

# Sort technicians by total
tech_totals = {tech: sum(tech_solicitud[tech].values()) for tech in tecnicos}
sorted_tecnicos = sorted(tecnicos, key=lambda x: tech_totals[x], reverse=True)

# Prepare stacked data
solicitud_types = list(all_solicitudes)
bottom = [0] * len(sorted_tecnicos)
colors_stack = plt.cm.tab20.colors[:len(solicitud_types)]

for i, solicitud in enumerate(solicitud_types):
    values = []
    for tech in sorted_tecnicos:
        values.append(tech_solicitud[tech].get(solicitud, 0))
    ax.bar(range(len(sorted_tecnicos)), values, bottom=bottom, label=solicitud, color=colors_stack[i % len(colors_stack)])
    bottom = [b + v for b, v in zip(bottom, values)]

ax.set_xticks(range(len(sorted_tecnicos)))
ax.set_xticklabels(sorted_tecnicos, rotation=45, ha='right')
ax.set_ylabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Ranking de Tecnicos por Tipo de Solicitud', fontsize=14, fontweight='bold')
ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8)
plt.tight_layout()
plt.savefig('/workspace/charts/tecnicos_ranking.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 6: Solicitud Distribution (Doughnut)
fig, ax = plt.subplots(figsize=(10, 8))
solicitud_counts = df['Solicitud Suscriptor'].value_counts()
wedges, texts, autotexts = ax.pie(solicitud_counts.values, labels=None, autopct='%1.1f%%',
                                   colors=colors[:len(solicitud_counts)], pctdistance=0.8,
                                   wedgeprops=dict(width=0.5))
for autotext in autotexts:
    autotext.set_fontsize(9)
ax.legend(solicitud_counts.index, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=9)
ax.set_title('Distribucion por Tipo de Solicitud', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig('/workspace/charts/tipo_doughnut.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 7: Zone vs Solicitud (Stacked)
fig, ax = plt.subplots(figsize=(12, 8))
zonas = df['Zona Cliente'].unique()
solicitud_by_zona = {}
for zona in zonas:
    zona_df = df[df['Zona Cliente'] == zona]
    solicitud_by_zona[zona] = zona_df['Solicitud Suscriptor'].value_counts()

bottom = [0] * len(zonas)
for i, solicitud in enumerate(solicitud_types[:10]):  # Top 10 solicitud types
    values = []
    for zona in zonas:
        vc = solicitud_by_zona[zona]
        if isinstance(vc, pd.Series):
            values.append(vc.get(solicitud, 0))
        else:
            values.append(0)
    ax.bar(range(len(zonas)), values, bottom=bottom, label=solicitud, color=colors_stack[i % len(colors_stack)])
    bottom = [b + v for b, v in zip(bottom, values)]

ax.set_xticks(range(len(zonas)))
ax.set_xticklabels(zonas, rotation=30, ha='right')
ax.set_ylabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Tipo de Solicitud por Zona', fontsize=14, fontweight='bold')
ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8)
plt.tight_layout()
plt.savefig('/workspace/charts/zona_clasificacion.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

# Chart 8: Solicitud Type Bar Chart
fig, ax = plt.subplots(figsize=(12, 6))
solicitud_counts = df['Solicitud Suscriptor'].value_counts()
bars = ax.bar(range(len(solicitud_counts)), solicitud_counts.values, color=COLORS['secondary'])
ax.set_xticks(range(len(solicitud_counts)))
ax.set_xticklabels(solicitud_counts.index, rotation=45, ha='right', fontsize=9)
ax.bar_label(bars, padding=3, fontsize=9)
ax.set_ylabel('Cantidad de Ordenes', fontsize=11)
ax.set_title('Ordenes por Tipo de Solicitud', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('/workspace/charts/solicitud_bar.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

print("=== GRAFICAS GENERADAS ===")
for f in os.listdir('/workspace/charts'):
    print(f"  - {f}")