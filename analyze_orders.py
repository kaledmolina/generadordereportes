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

# Calculate response time (creation to start)
df['Tiempo_Respuesta'] = (df['Fecha Inicio Atención'] - df['Fecha Creación']).dt.total_seconds() / 3600

# Calculate attention time (start to end)
df['Tiempo_Atención'] = (df['Fecha Fin Atención'] - df['Fecha Inicio Atención']).dt.total_seconds() / 3600

# Calculate total time (creation to end)
df['Tiempo_Total'] = (df['Fecha Fin Atención'] - df['Fecha Creación']).dt.total_seconds() / 3600

# Clean NaN values
df['Técnico Principal'] = df['Técnico Principal'].fillna('No Asignado')
df['Técnico Auxiliar'] = df['Técnico Auxiliar'].fillna('Ninguno')
df['Barrio'] = df['Barrio'].fillna('Sin Barrio')

print("=== ESTADÍSTICAS GENERALES ===")
print(f"Total de órdenes: {len(df)}")
print(f"\nEstados de órdenes:")
print(df['Estado'].value_counts())
print(f"\nTipos de orden:")
print(df['Tipo Orden'].value_counts())

print("\n=== ESTADÍSTICAS DE TIEMPO DE ATENCIÓN (horas) ===")
avg_response = df['Tiempo_Respuesta'].mean()
avg_attention = df['Tiempo_Atención'].mean()
avg_total = df['Tiempo_Total'].mean()
print(f"Tiempo promedio de respuesta: {avg_response:.2f} horas")
print(f"Tiempo promedio de atención: {avg_attention:.2f} horas")
print(f"Tiempo promedio total: {avg_total:.2f} horas")

median_response = df['Tiempo_Respuesta'].median()
median_attention = df['Tiempo_Atención'].median()
median_total = df['Tiempo_Total'].median()
print(f"\nMediana de tiempo de respuesta: {median_response:.2f} horas")
print(f"Mediana de tiempo de atención: {median_attention:.2f} horas")
print(f"Mediana de tiempo total: {median_total:.2f} horas")

print("\n=== ESTADÍSTICAS POR TÉCNICO ===")
tech_stats = df.groupby('Técnico Principal').agg({
    'N° Orden': 'count',
    'Tiempo_Atención': 'mean',
    'Tiempo_Total': 'mean'
}).round(2)
tech_stats.columns = ['Total_Ordenes', 'Prom_Tiempo_Atencion', 'Prom_Tiempo_Total']
tech_stats = tech_stats.sort_values('Total_Ordenes', ascending=False)
print(tech_stats)

print("\n=== ESTADÍSTICAS POR BARRIO (Top 15) ===")
barrio_stats = df.groupby('Barrio').agg({
    'N° Orden': 'count',
    'Tiempo_Atención': 'mean',
    'Tiempo_Total': 'mean'
}).round(2)
barrio_stats.columns = ['Total_Ordenes', 'Prom_Tiempo_Atencion', 'Prom_Tiempo_Total']
barrio_stats = barrio_stats.sort_values('Total_Ordenes', ascending=False).head(15)
print(barrio_stats)

print("\n=== CLASIFICACIÓN ===")
print(df['Clasificación'].value_counts())

print("\n=== MEDIO DE SOLICITUD ===")
print(df['Medio Solicitud'].value_counts())

print("\n=== SEDES ===")
print(df['Sede'].value_counts())

print("\n=== ZONAS ===")
print(df['Zona Cliente'].value_counts())

# Save for later use
df.to_pickle('orders_analysis.pkl')
print("\n=== Data saved for PDF generation ===")