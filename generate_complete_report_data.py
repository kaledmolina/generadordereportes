"""
Reporte Completo de Órdenes de Servicio
Incluye:
- Totales correctos (842 órdenes)
- Campos vacíos categorizados como "Otros" o "Sin solución"
- Tiempos de atención (Fecha Creación → Fecha Fin Atención)
- Reportes detallados por técnicos
- Órdenes por barrio
- Análisis completo de clasificación y soluciones
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Leer datos
df = pd.read_excel('user_input_files/reporte-ordenes-2026-04-01-a-2026-04-30 (1).xlsx')

# Total general de órdenes
TOTAL_ORDENES = len(df)
print(f"Total de órdenes en el dataset: {TOTAL_ORDENES}")

# Convertir columnas de fecha
df['Fecha Creación'] = pd.to_datetime(df['Fecha Creación'], errors='coerce')
df['Fecha Fin Atención'] = pd.to_datetime(df['Fecha Fin Atención'], errors='coerce')
df['Fecha Inicio Atención'] = pd.to_datetime(df['Fecha Inicio Atención'], errors='coerce')

# Rellenar valores nulos/vacíos con categorías específicas
df['Barrio'] = df['Barrio'].fillna('OTROS')
df['Barrio'] = df['Barrio'].replace('', 'OTROS')
df['Barrio'] = df['Barrio'].str.strip()
df.loc[df['Barrio'] == '', 'Barrio'] = 'OTROS'

df['Técnico Principal'] = df['Técnico Principal'].fillna('SIN ASIGNAR')
df['Técnico Principal'] = df['Técnico Principal'].replace('', 'SIN ASIGNAR')
df['Técnico Principal'] = df['Técnico Principal'].str.strip()
df.loc[df['Técnico Principal'] == '', 'Técnico Principal'] = 'SIN ASIGNAR'

df['Solución Técnico'] = df['Solución Técnico'].fillna('SIN SOLUCIÓN')
df['Solución Técnico'] = df['Solución Técnico'].replace('', 'SIN SOLUCIÓN')
df['Solución Técnico'] = df['Solución Técnico'].str.strip()
df.loc[df['Solución Técnico'] == '', 'Solución Técnico'] = 'SIN SOLUCIÓN'

df['Clasificación'] = df['Clasificación'].fillna('OTROS')
df['Clasificación'] = df['Clasificación'].replace('', 'OTROS')
df['Clasificación'] = df['Clasificación'].str.strip()
df.loc[df['Clasificación'] == '', 'Clasificación'] = 'OTROS'

df['Tipo Orden'] = df['Tipo Orden'].fillna('OTROS')
df['Tipo Orden'] = df['Tipo Orden'].replace('', 'OTROS')
df['Tipo Orden'] = df['Tipo Orden'].str.strip()
df.loc[df['Tipo Orden'] == '', 'Tipo Orden'] = 'OTROS'

df['Estado'] = df['Estado'].fillna('OTROS')
df['Estado'] = df['Estado'].replace('', 'OTROS')
df['Estado'] = df['Estado'].str.strip()
df.loc[df['Estado'] == '', 'Estado'] = 'OTROS'

# Calcular tiempo de atención (Fecha Creación → Fecha Fin Atención)
df['Tiempo Atención'] = df['Fecha Fin Atención'] - df['Fecha Creación']
df['Tiempo Atención (horas)'] = df['Tiempo Atención'].dt.total_seconds() / 3600
df['Tiempo Atención (horas)'] = df['Tiempo Atención (horas)'].fillna(0)

# Tiempo de atención positivo (solo contar cuando Fin Atención > Creación)
df['Tiempo Atención Positivo'] = df['Tiempo Atención (horas)'].apply(lambda x: x if x > 0 else 0)

# ============ ANÁLISIS POR BARRIO ============
print("\n=== ANÁLISIS POR BARRIO ===")
barrio_counts = df['Barrio'].value_counts().reset_index()
barrio_counts.columns = ['Barrio', 'Cantidad']
barrio_counts['Porcentaje'] = (barrio_counts['Cantidad'] / TOTAL_ORDENES * 100).round(2)
barrio_counts = barrio_counts.sort_values('Cantidad', ascending=False)
print(barrio_counts.to_string())
print(f"\nSuma de barrio_counts: {barrio_counts['Cantidad'].sum()}")

# ============ ANÁLISIS POR TÉCNICO ============
print("\n=== ANÁLISIS POR TÉCNICO ===")
tecnico_counts = df['Técnico Principal'].value_counts().reset_index()
tecnico_counts.columns = ['Técnico', 'Cantidad Órdenes']
tecnico_counts['Porcentaje'] = (tecnico_counts['Cantidad Órdenes'] / TOTAL_ORDENES * 100).round(2)
tecnico_counts = tecnico_counts.sort_values('Cantidad Órdenes', ascending=False)
print(tecnico_counts.to_string())
print(f"\nSuma técnico: {tecnico_counts['Cantidad Órdenes'].sum()}")

# Tiempo promedio por técnico
tiempo_por_tecnico = df.groupby('Técnico Principal').agg({
    'Tiempo Atención Positivo': ['mean', 'count'],
    'N° Orden': 'count'
}).round(2)
tiempo_por_tecnico.columns = ['Tiempo Promedio (horas)', 'Órdenes con Tiempo', 'Total Órdenes']
tiempo_por_tecnico = tiempo_por_tecnico.reset_index()
tiempo_por_tecnico = tiempo_por_tecnico.sort_values('Total Órdenes', ascending=False)
print("\nTiempo por técnico:")
print(tiempo_por_tecnico.to_string())

# ============ ANÁLISIS POR SOLUCIÓN ============
print("\n=== ANÁLISIS POR SOLUCIÓN ===")
solucion_counts = df['Solución Técnico'].value_counts().reset_index()
solucion_counts.columns = ['Solución', 'Cantidad']
solucion_counts['Porcentaje'] = (solucion_counts['Cantidad'] / TOTAL_ORDENES * 100).round(2)
solucion_counts = solucion_counts.sort_values('Cantidad', ascending=False)
print(solucion_counts.head(20).to_string())
print(f"\nSuma soluciones: {solucion_counts['Cantidad'].sum()}")

# ============ ANÁLISIS POR CLASIFICACIÓN ============
print("\n=== ANÁLISIS POR CLASIFICACIÓN ===")
clasificacion_counts = df['Clasificación'].value_counts().reset_index()
clasificacion_counts.columns = ['Clasificación', 'Cantidad']
clasificacion_counts['Porcentaje'] = (clasificacion_counts['Cantidad'] / TOTAL_ORDENES * 100).round(2)
clasificacion_counts = clasificacion_counts.sort_values('Cantidad', ascending=False)
print(clasificacion_counts.to_string())
print(f"\nSuma clasificación: {clasificacion_counts['Cantidad'].sum()}")

# ============ ANÁLISIS POR TIPO DE ORDEN ============
print("\n=== ANÁLISIS POR TIPO DE ORDEN ===")
tipo_counts = df['Tipo Orden'].value_counts().reset_index()
tipo_counts.columns = ['Tipo Orden', 'Cantidad']
tipo_counts['Porcentaje'] = (tipo_counts['Cantidad'] / TOTAL_ORDENES * 100).round(2)
tipo_counts = tipo_counts.sort_values('Cantidad', ascending=False)
print(tipo_counts.to_string())
print(f"\nSuma tipos: {tipo_counts['Cantidad'].sum()}")

# ============ ANÁLISIS POR ESTADO ============
print("\n=== ANÁLISIS POR ESTADO ===")
estado_counts = df['Estado'].value_counts().reset_index()
estado_counts.columns = ['Estado', 'Cantidad']
estado_counts['Porcentaje'] = (estado_counts['Cantidad'] / TOTAL_ORDENES * 100).round(2)
estado_counts = estado_counts.sort_values('Cantidad', ascending=False)
print(estado_counts.to_string())
print(f"\nSuma estados: {estado_counts['Cantidad'].sum()}")

# ============ ESTADÍSTICAS DE TIEMPO ============
print("\n=== ESTADÍSTICAS DE TIEMPO DE ATENCIÓN ===")
tiempo_stats = df['Tiempo Atención Positivo'].describe()
print(tiempo_stats)

# Órdenes con tiempo negativo o cero
ordenes_sin_tiempo = len(df[df['Tiempo Atención Positivo'] <= 0])
ordenes_con_tiempo = len(df[df['Tiempo Atención Positivo'] > 0])
print(f"\nÓrdenes sin tiempo de atención: {ordenes_sin_tiempo}")
print(f"Órdenes con tiempo de atención: {ordenes_con_tiempo}")
print(f"Total: {ordenes_sin_tiempo + ordenes_con_tiempo}")

# Guardar datos procesados
df.to_pickle('ordenes_procesadas.pkl')
barrio_counts.to_pickle('barrio_counts.pkl')
tecnico_counts.to_pickle('tecnico_counts.pkl')
solucion_counts.to_pickle('solucion_counts.pkl')
clasificacion_counts.to_pickle('clasificacion_counts.pkl')
tipo_counts.to_pickle('tipo_counts.pkl')
estado_counts.to_pickle('estado_counts.pkl')
tiempo_por_tecnico.to_pickle('tiempo_por_tecnico.pkl')

print("\n=== ARCHIVOS GUARDADOS ===")
print("Datos procesados guardados exitosamente")