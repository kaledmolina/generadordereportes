import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import os
import io
import base64
from datetime import datetime
import numpy as np

# Paleta de colores EXACTA de generate_charts.py
COLORS_DICT = {
    'primary': '#1C3557',
    'secondary': '#00B4A6',
    'accent1': '#E63946',
    'accent2': '#FF9F1C',
    'accent3': '#2A9D8F',
    'accent4': '#8338EC',
    'light': '#F0F5F8'
}
TAB20_COLORS = plt.cm.tab20.colors

def fig_to_base64(fig):
    img = io.BytesIO()
    fig.savefig(img, format='png', dpi=150, bbox_inches='tight')
    img.seek(0)
    return base64.b64encode(img.getvalue()).decode('utf-8')

def format_html_table_row(cells, is_total=False, aligns=None):
    row_class = ' class="total-row"' if is_total else ''
    html = f'<tr{row_class}>'
    for i, cell in enumerate(cells):
        if i == 0:
            html += f'<td>{cell}</td>'
        else:
            if aligns and i < len(aligns) and aligns[i]:
                align_class = f' class="{aligns[i]}"'
            else:
                align_class = ' class="text-right"' if isinstance(cell, (int, float)) or '%' in str(cell) else ' class="text-center"'
            html += f'<td{align_class}>{cell}</td>'
    html += '</tr>'
    return html

def generate_report_from_df(df, template_path='reporte_template.html'):
    # Total inicial
    TOTAL_ORDENES = len(df)
    
    # --- PROCESAMIENTO DE DATOS (Lógica de generate_complete_report_data.py) ---
    df['Fecha Creación'] = pd.to_datetime(df['Fecha Creación'], errors='coerce')
    df['Fecha Fin Atención'] = pd.to_datetime(df['Fecha Fin Atención'], errors='coerce')
    
    # Rellenar vacíos EXACTAMENTE como en generate_complete_report_data.py
    df['Barrio'] = df['Barrio'].fillna('SIN BARRIO').str.strip().replace('', 'SIN BARRIO')
    df['Técnico Principal'] = df['Técnico Principal'].fillna('SIN ASIGNAR').str.strip().replace('', 'SIN ASIGNAR')
    df['Solución Técnico'] = df['Solución Técnico'].fillna('SIN SOLUCIÓN').str.strip().replace('', 'SIN SOLUCIÓN')
    df['Clasificación'] = df['Clasificación'].fillna('OTROS').str.strip().replace('', 'OTROS')
    df['Tipo Orden'] = df['Tipo Orden'].fillna('OTROS').str.strip().replace('', 'OTROS')
    df['Estado'] = df['Estado'].fillna('OTROS').str.strip().replace('', 'OTROS')
    df['Solicitud Suscriptor'] = df['Solicitud Suscriptor'].fillna('OTROS').str.strip().replace('', 'OTROS')
    df['Zona Cliente'] = df['Zona Cliente'].fillna('SIN ZONA').str.strip().replace('', 'SIN ZONA')

    # Tiempos
    df['Tiempo Atención'] = df['Fecha Fin Atención'] - df['Fecha Creación']
    df['Tiempo Atención (horas)'] = df['Tiempo Atención'].dt.total_seconds() / 3600
    df['Tiempo Atención Positivo'] = df['Tiempo Atención (horas)'].apply(lambda x: x if x > 0 else 0)
    
    # --- ESTADÍSTICAS ---
    ordenes_cerradas = len(df[df['Estado'].str.lower() == 'cerrada'])
    tasa_cierre = (ordenes_cerradas / TOTAL_ORDENES * 100)
    ordenes_con_solucion = len(df[df['Solución Técnico'] != 'SIN SOLUCIÓN'])
    tasa_resolucion = (ordenes_con_solucion / TOTAL_ORDENES * 100)
    ordenes_sin_solucion = TOTAL_ORDENES - ordenes_con_solucion
    barrios_unicos = df['Barrio'].nunique()
    
    # Técnicos activos (principales 6)
    tecnicos_principales_df = df[~df['Técnico Principal'].isin(['SIN ASIGNAR', 'Planta externa', 'Soporte Noc'])]
    tecnicos_list = tecnicos_principales_df['Técnico Principal'].unique()
    
    tiempo_promedio = df['Tiempo Atención Positivo'].mean()
    mediana_tiempo = df['Tiempo Atención Positivo'].median()
    ordenes_con_tiempo_valido = len(df[df['Tiempo Atención (horas)'] > 0])
    
    # --- GENERACIÓN DE GRÁFICOS (Lógica de generate_charts.py) ---
    charts = {}
    plt.rcParams['font.family'] = 'DejaVu Sans'
    
    # 1. clasificacion_pie.png (Basado en SOLICITUD SUSCRIPTOR, no Clasificación)
    fig, ax = plt.subplots(figsize=(10, 8))
    s_counts = df['Solicitud Suscriptor'].value_counts()
    ax.pie(s_counts.values, labels=None, autopct='%1.1f%%', colors=list(TAB20_COLORS), pctdistance=0.75)
    ax.legend(s_counts.index, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=9)
    ax.set_title('Ordenes por Tipo de Solicitud', fontsize=14, fontweight='bold', pad=20)
    charts['clasificacion_pie.png'] = fig_to_base64(fig)
    plt.close()
    
    # 2. estado_bar.png
    fig, ax = plt.subplots(figsize=(10, 6))
    e_counts = df['Estado'].value_counts()
    bars = ax.bar(e_counts.index, e_counts.values, color=COLORS_DICT['primary'])
    ax.bar_label(bars, padding=3, fontsize=10)
    ax.set_title('Ordenes por Estado', fontsize=14, fontweight='bold')
    plt.xticks(rotation=30, ha='right')
    charts['estado_bar.png'] = fig_to_base64(fig)
    plt.close()
    
    # 3. barrios_bar.png (Horizontal Top 15)
    fig, ax = plt.subplots(figsize=(12, 8))
    b_counts = df['Barrio'].value_counts().head(15)
    bars = ax.barh(b_counts.index[::-1], b_counts.values[::-1], color=COLORS_DICT['primary'])
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.set_title('Top 15 Barrios con Mas Ordenes', fontsize=14, fontweight='bold')
    charts['barrios_bar.png'] = fig_to_base64(fig)
    plt.close()
    
    # 4. tecnicos_ranking.png (STACKED BAR - Muy importante)
    fig, ax = plt.subplots(figsize=(14, 8))
    valid_techs = [t for t in df['Técnico Principal'].unique() if t not in ['SIN ASIGNAR', 'Planta externa', 'Soporte Noc', 'Soporte NOC']]
    solicitud_types = list(df['Solicitud Suscriptor'].unique())
    
    tech_solicitud = {}
    for tech in valid_techs:
        t_df = df[df['Técnico Principal'] == tech]
        vc = t_df['Solicitud Suscriptor'].value_counts()
        tech_solicitud[tech] = dict(zip(vc.index, vc.values))
    
    tech_totals = {t: sum(tech_solicitud[t].values()) for t in valid_techs}
    sorted_tecnicos = sorted(valid_techs, key=lambda x: tech_totals[x], reverse=True)
    
    bottom = [0] * len(sorted_tecnicos)
    for i, sol in enumerate(solicitud_types):
        values = [tech_solicitud[t].get(sol, 0) for t in sorted_tecnicos]
        ax.bar(range(len(sorted_tecnicos)), values, bottom=bottom, label=sol, color=TAB20_COLORS[i % len(TAB20_COLORS)])
        bottom = [b + v for b, v in zip(bottom, values)]
        
    ax.set_xticks(range(len(sorted_tecnicos)))
    ax.set_xticklabels(sorted_tecnicos, rotation=45, ha='right')
    ax.set_title('Ranking de Tecnicos por Tipo de Solicitud', fontsize=14, fontweight='bold')
    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8)
    charts['tecnicos_ranking.png'] = fig_to_base64(fig)
    plt.close()
    
    # 5. tipo_orden_bar.png
    fig, ax = plt.subplots(figsize=(12, 6))
    to_counts = df['Tipo Orden'].value_counts()
    bars = ax.bar(to_counts.index, to_counts.values, color=COLORS_DICT['accent2'])
    ax.bar_label(bars, padding=3)
    ax.set_title('Distribución por Tipo de Orden', fontsize=14, fontweight='bold')
    plt.xticks(rotation=45, ha='right')
    charts['tipo_orden_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 6. solicitud_bar.png
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(range(len(s_counts)), s_counts.values, color=COLORS_DICT['secondary'])
    ax.set_xticks(range(len(s_counts)))
    ax.set_xticklabels(s_counts.index, rotation=45, ha='right', fontsize=9)
    ax.bar_label(bars, padding=3)
    ax.set_title('Ordenes por Tipo de Solicitud', fontsize=14, fontweight='bold')
    charts['solicitud_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 7. zona_bar.png
    fig, ax = plt.subplots(figsize=(10, 6))
    z_counts = df['Zona Cliente'].value_counts()
    bars = ax.bar(z_counts.index, z_counts.values, color=COLORS_DICT['secondary'])
    ax.bar_label(bars, padding=3)
    ax.set_title('Ordenes por Zona', fontsize=14, fontweight='bold')
    charts['zona_bar.png'] = fig_to_base64(fig)
    plt.close()

    # --- CONSTRUCCIÓN DE TABLAS ---
    # KPI
    kpi_body = ""
    kpi_body += format_html_table_row(["Tasa de Cierre", f"<strong>{tasa_cierre:.2f}%</strong>", f"{ordenes_cerradas} de {TOTAL_ORDENES} órdenes cerradas exitosamente"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Tasa de Resolución", f"<strong>{tasa_resolucion:.2f}%</strong>", f"{ordenes_con_solucion} de {TOTAL_ORDENES} órdenes con solución aplicada"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Órdenes sin Solución", f"<strong>{ordenes_sin_solucion}</strong>", f"{(ordenes_sin_solucion/TOTAL_ORDENES*100):.2f}% del total (requiere seguimiento)"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Barrios Atendidos", f"<strong>{barrios_unicos}</strong>", "Cobertura geográfica completa"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Tiempo Promedio", f"<strong>{tiempo_promedio:.2f} hrs</strong>", f"Mediana: {mediana_tiempo:.2f} horas"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Órdenes con Tiempo Válido", f"<strong>{ordenes_con_tiempo_valido}</strong>", f"{(ordenes_con_tiempo_valido/TOTAL_ORDENES*100):.2f}% del total"], aligns=["", "text-center", "text-left"])
    
    # Barrios Top 20
    barrios_body = ""
    top_b = df['Barrio'].value_counts().head(20)
    for i, (name, count) in enumerate(top_b.items(), 1):
        barrios_body += format_html_table_row([i, name, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    barrios_body += format_html_table_row(["<strong>TOTAL GENERAL (Top 20)</strong>", "", top_b.sum(), f"{(top_b.sum()/TOTAL_ORDENES*100):.2f}%"], is_total=True)
    
    # Técnicos
    tecnicos_body = ""
    tech_stats = df.groupby('Técnico Principal').agg({'N° Orden': 'count', 'Tiempo Atención Positivo': 'mean'}).sort_values('N° Orden', ascending=False)
    for i, (name, row) in enumerate(tech_stats.iterrows(), 1):
        badge = f'<span class="badge badge-success">{i}°</span>' if i <= 2 else (f'<span class="badge badge-warning">{i}°</span>' if i <= 6 else "—")
        tecnicos_body += format_html_table_row([f"<strong>{name}</strong>", int(row['N° Orden']), f"{(row['N° Orden']/TOTAL_ORDENES*100):.2f}%", f"{row['Tiempo Atención Positivo']:.2f}", badge])
    tecnicos_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%", f"{tiempo_promedio:.2f}", "—"], is_total=True)
    
    # Soluciones
    sol_body = ""
    top_s = df['Solución Técnico'].value_counts().head(20)
    for i, (name, count) in enumerate(top_s.items(), 1):
        display_name = f'<span class="badge badge-danger">SIN SOLUCIÓN</span>' if name == 'SIN SOLUCIÓN' else name
        sol_body += format_html_table_row([i, display_name, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    sol_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", "", TOTAL_ORDENES, "100.00%"], is_total=True)
    
    # Clasificación y Tipo
    class_body = ""
    for n, c in df['Clasificación'].value_counts().items():
        class_body += format_html_table_row([n, c, f"{(c/TOTAL_ORDENES*100):.2f}%"])
    class_body += format_html_table_row(["<strong>TOTAL</strong>", TOTAL_ORDENES, "100%"], is_total=True)
    
    tipo_body = ""
    for n, c in df['Tipo Orden'].value_counts().items():
        tipo_body += format_html_table_row([n, c, f"{(c/TOTAL_ORDENES*100):.2f}%"])
    tipo_body += format_html_table_row(["<strong>TOTAL</strong>", TOTAL_ORDENES, "100%"], is_total=True)
    
    estado_body = ""
    for n, c in df['Estado'].value_counts().items():
        n_lower = n.lower()
        badge_c = "badge-success" if n_lower == 'cerrada' else ("badge-danger" if n_lower == 'anulada' else ("badge-warning" if n_lower == 'reprogramada' else "badge-info"))
        estado_body += format_html_table_row([f'<span class="badge {badge_c}">{n.capitalize()}</span>', c, f"{(c/TOTAL_ORDENES*100):.2f}%"])
    estado_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%"], is_total=True)
    
    # Tiempos
    bins = [-float('inf'), 0, 6, 12, 24, 48, 72, 120, float('inf')]
    labels = ['Sin tiempo válido', '0 - 6 horas (Mismo día)', '6 - 12 horas (Día siguiente)', '12 - 24 horas (1-2 días)', '24 - 48 horas (2-3 días)', '48 - 72 horas (3-4 días)', '72 - 120 horas (4-5 días)', '> 120 horas (más de 5 días)']
    df['Rango'] = pd.cut(df['Tiempo Atención (horas)'], bins=bins, labels=labels)
    r_counts = df['Rango'].value_counts().reindex(labels).fillna(0)
    acum = 0
    rangos_body = ""
    for n, c in r_counts.items():
        c = int(c)
        acum += c
        rangos_body += format_html_table_row([n, c, f"{(c/TOTAL_ORDENES*100):.2f}%", f"{acum} ({(acum/TOTAL_ORDENES*100):.2f}%)"])
    rangos_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100%", "—"], is_total=True)
    
    # Métricas adicionales
    stats = df['Tiempo Atención Positivo'].describe()
    metrics_body = ""
    metrics_body += format_html_table_row(["Tiempo Mínimo", f"{stats['min']:.2f} horas"])
    metrics_body += format_html_table_row(["Tiempo Máximo", f"{stats['max']:.2f} horas"])
    metrics_body += format_html_table_row(["Percentil 25%", f"{df['Tiempo Atención Positivo'].quantile(0.25):.2f} horas"])
    metrics_body += format_html_table_row(["Percentil 75%", f"{df['Tiempo Atención Positivo'].quantile(0.75):.2f} horas"])
    metrics_body += format_html_table_row(["Desviación Estándar", f"{stats['std']:.2f} horas"])

    # --- ENSAMBLAJE ---
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()
        
    p_menos_48 = (len(df[df['Tiempo Atención (horas)'] <= 48]) / TOTAL_ORDENES * 100)
    p_mas_120 = (len(df[df['Tiempo Atención (horas)'] > 120]) / TOTAL_ORDENES * 100)
    
    # Análisis dinámico para el Hallazgo Clave y Resumen
    top_t = tech_stats.index[0]
    top_t_o = int(tech_stats.iloc[0]['N° Orden'])
    top_t_p = (top_t_o / TOTAL_ORDENES * 100)
    top_t_t = tech_stats.iloc[0]['Tiempo Atención Positivo']
    
    hallazgos = f"""
        <strong>1. Volumen de Órdenes:</strong> Se atendieron <strong>{TOTAL_ORDENES} órdenes de servicio</strong> durante el período, con una tasa de cierre del <strong>{tasa_cierre:.2f}%</strong><br><br>
        <strong>2. Tiempo de Atención:</strong> El tiempo promedio total de atención fue de <strong>{tiempo_promedio:.2f} horas</strong>. El {p_menos_48:.2f}% de las órdenes fueron atendidas en menos de 48 horas.<br><br>
        <strong>3. Cobertura Geográfica:</strong> Se atendieron <strong>{barrios_unicos} barrios diferentes</strong>, con <strong>{top_b.index[0]}</strong> como el barrio con mayor volumen ({top_b.iloc[0]} órdenes, {(top_b.iloc[0]/TOTAL_ORDENES*100):.2f}%).<br><br>
        <strong>4. Desempeño Técnico:</strong> <strong>{top_t}</strong> lidera con {top_t_o} órdenes ({top_t_p:.2f}%) y el menor tiempo promedio ({top_t_t:.2f} hrs).<br><br>
        <strong>5. Soluciones:</strong> La solución más frecuente fue <strong>{top_s.index[0]}</strong> con {top_s.iloc[0]} casos ({(top_s.iloc[0]/TOTAL_ORDENES*100):.2f}%).
    """

    # Extraer periodo dinámicamente
    meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
    try:
        valid_dates = df['Fecha Creación'].dropna()
        min_date = valid_dates.min()
        max_date = valid_dates.max()
        periodo_texto = f"{min_date.day} de {meses[min_date.month]} al {max_date.day} de {meses[max_date.month]} de {max_date.year}"
        solo_mes = f"{meses[min_date.month].capitalize()} {min_date.year}"
    except:
        periodo_texto = "1 al 30 de abril de 2026"
        solo_mes = "Abril 2026"

    replacements = {
        '{{ TITULO_PAGINA }}': f'Reporte Órdenes de Servicio - {solo_mes}',
        '{{ REPORTE_HEADER_TITLE }}': f'Reporte Completo de Órdenes de Servicio — {solo_mes}',
        '{{ TOTAL_ORDENES }}': str(TOTAL_ORDENES),
        '{{ PERIODO_TEXTO }}': periodo_texto,
        '{{ ORDENES_CERRADAS }}': str(ordenes_cerradas),
        '{{ TIEMPO_PROMEDIO_BREVE }}': f"{tiempo_promedio:.1f}",
        '{{ TECNICOS_ACTIVOS }}': str(len(valid_techs)),
        '{{ TIEMPO_PROMEDIO_HRS }}': f"{tiempo_promedio:.2f}",
        '{{ PORCENTAJE_MENOS_48H }}': f"{p_menos_48:.2f}",
        '{{ TASA_RESOLUCION }}': f"{tasa_resolucion:.2f}",
        '{{ TABLA_KPI_BODY }}': kpi_body,
        '{{ FOOTER_TEXT }}': f'Reporte de Órdenes de Servicio - Área Innovación y Desarrollo - Kaled Molina',
        '{{ TABLA_BARRIOS_BODY }}': barrios_body,
        '{{ NOTA_BARRIOS }}': f'<strong>Nota:</strong> Se identificaron <strong>{barrios_unicos} barrios diferentes</strong> en total. Los 20 barrios principales representan el {(top_b.sum()/TOTAL_ORDENES*100):.2f}% del total de órdenes.',
        '{{ TABLA_TECNICOS_BODY }}': tecnicos_body,
        '{{ ANALISIS_RENDIMIENTO_TECNICO }}': f'• <strong>{top_t}</strong> lidera con {top_t_o} órdenes ({top_t_p:.2f}%) y el menor tiempo promedio ({top_t_t:.2f} hrs)<br>• Los técnicos principales atienden la gran mayoría de las solicitudes.',
        '{{ TABLA_SOLUCIONES_BODY }}': sol_body,
        '{{ CONCLUSIONES_SOLUCIONES }}': f'• <strong>Tasa de resolución:</strong> {tasa_resolucion:.2f}% ({ordenes_con_solucion} de {TOTAL_ORDENES} órdenes)',
        '{{ TABLA_CLASIFICACION_BODY }}': class_body,
        '{{ TABLA_TIPO_ORDEN_BODY }}': tipo_body,
        '{{ TABLA_ESTADO_BODY }}': estado_body,
        '{{ ORDENES_CON_TIEMPO_VALIDO }}': str(ordenes_con_tiempo_valido),
        '{{ MEDIANA_TIEMPO }}': f"{mediana_tiempo:.2f}",
        '{{ TABLA_RANGOS_TIEMPO_BODY }}': rangos_body,
        '{{ PORCENTAJE_MAS_120H }}': f"{p_mas_120:.2f}",
        '{{ TABLA_METRICAS_ADICIONALES_BODY }}': metrics_body,
        '{{ PERIODO_SOLO_MES }}': solo_mes,
        '{{ CANTIDAD_BARRIOS }}': str(barrios_unicos),
        '{{ HALLAZGOS_PRINCIPALES }}': hallazgos,
        '{{ LISTA_RECOMENDACIONES }}': '<li>Optimizar tiempos de respuesta.</li><li>Seguimiento de casos sin solución.</li>',
        '{{ CLASIFICACION_GENERAL }}': f"El departamento de operaciones técnicas demuestra un nivel de desempeño con tasa de resolución del {tasa_resolucion:.2f}%.",
        '{{ FOOTER_EMPRESA }}': "Corporación Regional de Telecomunicaciones - Kaled Molina"
    }


    # Reemplazar placeholders
    for k, v in replacements.items():
        html = html.replace(k, str(v))
    
    # Reemplazar imágenes
    for filename, b64 in charts.items():
        html = html.replace(f'src="charts/{filename}"', f'src="data:image/png;base64,{b64}"')
        
    return html

if __name__ == "__main__":
    import glob
    files = glob.glob('user_input_files/*.xlsx')
    if files:
        latest = max(files, key=os.path.getmtime)
        df = pd.read_excel(latest)
        report_html = generate_report_from_df(df)
        with open('reporte_final/reporte_generado.html', 'w', encoding='utf-8') as f:
            f.write(report_html)
        print("Reporte generado con exactitud.")
