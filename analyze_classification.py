import pandas as pd
import json
from datetime import datetime
import numpy as np

# Read the Excel file
df = pd.read_excel('user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Convert date columns to datetime
date_cols = ['Fecha Creación', 'Fecha Asignación', 'Fecha Inicio Atención', 'Fecha Fin Atención', 'Fecha Cierre']
for col in date_cols:
    df[col] = pd.to_datetime(df[col], errors='coerce')

# Calculate times
df['Tiempo_Total'] = (df['Fecha Fin Atención'] - df['Fecha Creación']).dt.total_seconds() / 3600

# Clean data
df['Técnico Principal'] = df['Técnico Principal'].fillna('Sin Asignar')
df['Barrio'] = df['Barrio'].fillna('Sin Barrio')
df['Solución Técnico'] = df['Solución Técnico'].fillna('Sin Solución')
df['Solicitud Suscriptor'] = df['Solicitud Suscriptor'].fillna('Sin especificar')

# Classify orders based on solutions and request types
def classify_order(row):
    solution = str(row['Solución Técnico']).upper()
    request = str(row['Solicitud Suscriptor']).upper()
    tiempo = row['Tiempo_Total']

    # Priority classifications (luz roja = high priority/urgent)
    if 'LUZ ROJA' in solution or 'LUZ ROJA' in request:
        return 'LUZ ROJA'
    elif 'GARANTIA' in solution or 'GARANTIA' in request:
        return 'GARANTIA'
    elif tiempo > 72:  # More than 72 hours = priority
        return 'LUZ ROJA'
    elif 'INSTALACION' in solution or 'NUEVO' in request:
        return 'NUEVA INSTALACION'
    elif 'TRASLADO' in solution or 'TRASLADO' in request:
        return 'TRASLADO'
    elif 'CONFIGUR' in solution or 'CAMBIO' in solution:
        return 'CONFIGURACION'
    elif 'MANTENIMIENTO' in solution:
        return 'MANTENIMIENTO'
    else:
        return 'SERVICIO GENERAL'

df['Clasificacion_Orden'] = df.apply(classify_order, axis=1)

# Get classification counts by technician
print("=== CLASIFICACIÓN DE ÓRDENES POR TÉCNICO ===")
tech_classification = df.groupby(['Técnico Principal', 'Clasificacion_Orden']).size().unstack(fill_value=0)
print(tech_classification)

print("\n=== TOTALES POR TÉCNICO ===")
tech_totals = df.groupby('Técnico Principal').size().sort_values(ascending=False)
print(tech_totals)

# Create detailed ranking with classifications
print("\n=== RANKING COMPLETO DE TÉCNICOS ===")
for tech in tech_totals.index:
    tech_data = df[df['Técnico Principal'] == tech]
    total = len(tech_data)
    classifications = tech_data['Clasificacion_Orden'].value_counts()

    print(f"\n{tech}:")
    print(f"  Total órdenes: {total}")
    for clas, count in classifications.items():
        print(f"    {clas}: {count}")

# Graph data for charts
print("\n=== DATA FOR GRAPHS ===")
print("\nOrders by State:")
print(df['Estado'].value_counts())

print("\nOrders by Type:")
print(df['Tipo Orden'].value_counts())

print("\nOrders by Classification:")
print(df['Clasificacion_Orden'].value_counts())

print("\nOrders by Zone:")
print(df['Zona Cliente'].value_counts())

print("\nTop 15 Neighborhoods:")
print(df['Barrio'].value_counts().head(15))

print("\nSolutions frequency:")
print(df['Solución Técnico'].value_counts().head(10))

# Average times by classification
print("\n=== TIEMPOS POR CLASIFICACIÓN ===")
avg_times = df.groupby('Clasificacion_Orden')['Tiempo_Total'].agg(['mean', 'median', 'count']).round(2)
print(avg_times)

# Save classification data
df.to_pickle('orders_with_classification.pkl')
print("\n=== Classification data saved ===")