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

import re
import unicodedata

def clean_text_repr(val):
    if not isinstance(val, str):
        return ""
    val = unicodedata.normalize('NFKD', str(val)).encode('ASCII', 'ignore').decode('utf-8')
    val = re.sub(r'[^a-zA-Z0-9\s]', ' ', val).lower().strip()
    return re.sub(r'\s+', ' ', val)

def normalize_dataframe_columns(df):
    """
    Normaliza los nombres de las columnas del DataFrame para ser tolerante a:
    - Mayúsculas/minúsculas
    - Tildes / acentos ausentes o caracteres con encoding corrupto
    - Espacios extra al inicio/final o entre palabras
    - Variaciones de nombres (ej. 'Fecha Creacion', 'Fecha de Creacion', etc.)
    - Filas de encabezado desplazadas por títulos superiores
    - Columnas opcionales ausentes
    """
    df = df.copy()
    
    # 1. Si los encabezados están en una fila posterior debido a títulos superiores
    def count_known_columns(cols):
        cleaned = [clean_text_repr(str(c)) for c in cols]
        score = 0
        keywords = ['fecha', 'orden', 'estado', 'cliente', 'tecnico', 'solicitud', 'barrio', 'clasificacion', 'creador', 'solucion']
        for c in cleaned:
            for kw in keywords:
                if kw in c:
                    score += 1
                    break
        return score

    if count_known_columns(df.columns) < 2:
        for i in range(min(10, len(df))):
            row_vals = df.iloc[i].tolist()
            if count_known_columns(row_vals) >= 2:
                df.columns = [str(x).strip() if pd.notna(x) else f"Col_{j}" for j, x in enumerate(row_vals)]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
                
    # 2. Mapeo de patrones de columnas canónicas
    COLUMN_MAPPINGS = {
        'Fecha Creación': [
            'fecha creacion', 'fecha de creacion', 'fecha creacion hora', 'f creacion', 'f_creacion',
            'fechacreacion', 'fecha registro', 'fecha creaci n', 'fecha creacin', 'fecha creacion', 'fecha creación'
        ],
        'Fecha Fin Atención': [
            'fecha fin atencion', 'fecha fin de atencion', 'fecha atencion', 'fecha fin', 
            'fecha solucion', 'fecha resolucion', 'fecha fin atenci n', 'fecha cierre', 'fechafin',
            'fecha fin atencion', 'fecha fin atención'
        ],
        'Fecha Asignación': ['fecha asignacion', 'fecha de asignacion', 'f asignacion', 'fecha asignaci n', 'fecha asignacion'],
        'Fecha Inicio Atención': ['fecha inicio atencion', 'fecha inicio', 'fecha inicio atenci n', 'fecha inicio atencion'],
        'Fecha Cierre': ['fecha cierre', 'fecha de cierre'],
        'N° Orden': ['n orden', 'no orden', 'numero orden', 'numero de orden', 'orden', 'id orden', 'id_orden', 'n  orden', 'n orden', 'num orden', 'n de orden'],
        'Barrio': ['barrio', 'barrio cliente', 'barrio suscriptor', 'vecindario', 'sector barrio'],
        'Técnico Principal': ['tecnico principal', 'tecnico', 'nombre tecnico', 'responsable', 'tecnico asignado', 't cnico principal', 'tcnico principal', 'cuadrilla'],
        'Técnico Auxiliar': ['tecnico auxiliar', 'auxiliar', 't cnico auxiliar', 'tcnico auxiliar'],
        'Solución Técnico': ['solucion tecnico', 'solucion', 'solucion tecnica', 'tipo solucion', 'soluci n t cnico', 'solucion dada', 'solucion final'],
        'Clasificación': ['clasificacion', 'tipo clasificacion', 'clasificaci n', 'categoria'],
        'Tipo Orden': ['tipo orden', 'tipo de orden', 'tipo', 'tipo_orden'],
        'Estado': ['estado', 'estado orden', 'estado actual', 'status'],
        'Estado Interno': ['estado interno', 'status interno'],
        'Solicitud Suscriptor': ['solicitud suscriptor', 'solicitud', 'tipo solicitud', 'motivo', 'solicitud cliente', 'problema', 'falla'],
        'Zona Cliente': ['zona cliente', 'zona', 'region', 'sector', 'zona suscriptor'],
        'Creador': ['creador', 'creado por', 'usuario creador', 'agente creador', 'usuario'],
        'Valor Total': ['valor total', 'valor', 'total', 'precio', 'costo', 'valor orden'],
        'Código': ['codigo', 'cod suscriptor', 'codigo suscriptor', 'id suscriptor', 'c digo', 'codigo cliente', 'contrato'],
        'Cliente': ['cliente', 'nombre cliente', 'nombre suscriptor', 'suscriptor', 'usuario cliente'],
        'Dirección': ['direccion', 'direccion cliente', 'direcci n', 'domicilio'],
        'Teléfono': ['telefono', 'telefono cliente', 'celular', 'tel fono', 'movil'],
        'Cédula': ['cedula', 'cedula cliente', 'c dula', 'identificacion', 'documento', 'nit', 'cc'],
        # --- Columnas del formato actualizado ---
        'Tipo de Servicio': ['tipo de servicio', 'tipo servicio', 'tipo de servicio contratado'],
        'Plan Internet': ['plan internet', 'plan de internet', 'plan contratado', 'nombre del plan'],
        'Origen (Enlace/Manual)': ['origen enlace manual', 'origen de la orden', 'origen orden', 'origen'],
        'Aplicado en Enlace': ['aplicado en enlace', 'aplicado enlace', 'aplicado'],
        'Medio Solicitud': ['medio solicitud', 'medio de solicitud', 'canal de solicitud', 'canal'],
        'Prioridad': ['prioridad', 'prioridad de la orden', 'prioridad orden'],
        'Coordenadas': ['coordenadas', 'coordenada', 'ubicacion geografica', 'geolocalizacion'],
        'Novedades NOC': ['novedades noc', 'novedad noc', 'novedades'],
        'Observaciones': ['observaciones', 'observacion', 'notas de la orden', 'comentarios'],
        'Vencimiento (SLA)': ['vencimiento sla', 'vencimiento de sla', 'vencimiento', 'fecha vencimiento'],
        'Fecha Llegada': ['fecha llegada', 'fecha de llegada', 'hora de llegada'],
        'Sede': ['sede', 'sede cliente', 'sede de la orden'],
        'Vendedor': ['vendedor', 'asesor comercial', 'vendedor cliente', 'asesor'],
        'Precinto': ['precinto', 'precintos'],
        'Ubicación Fotos': ['ubicacion fotos', 'ubicacion foto', 'evidencia fotografica', 'fotos de la orden'],
        'MAC Router': ['mac router', 'mac del router'],
        'MAC Bridge': ['mac bridge', 'mac del bridge'],
        'MAC ONT': ['mac ont', 'mac de la ont'],
        'Otros Equipos': ['otros equipos', 'otro equipo', 'equipos adicionales']
    }
    
    # Renombrar columnas encontradas con prioridad en coincidencias exactas
    current_cols = list(df.columns)
    rename_dict = {}
    used_canonicals = set()
    
    # Paso 1: Coincidencias exactas
    for col in current_cols:
        col_str = str(col).strip()
        cleaned = clean_text_repr(col_str)
        
        for canonical, aliases in COLUMN_MAPPINGS.items():
            if canonical in used_canonicals:
                continue
            canonical_clean = clean_text_repr(canonical)
            cleaned_aliases = [clean_text_repr(a) for a in aliases]
            
            if cleaned == canonical_clean or cleaned in cleaned_aliases:
                rename_dict[col] = canonical
                used_canonicals.add(canonical)
                break
                
    # Paso 2: Coincidencias por prefijo / inclusión controlada solo para columnas no asignadas
    for col in current_cols:
        if col in rename_dict:
            continue
        col_str = str(col).strip()
        cleaned = clean_text_repr(col_str)
        
        for canonical, aliases in COLUMN_MAPPINGS.items():
            if canonical in used_canonicals:
                continue
            cleaned_aliases = [clean_text_repr(a) for a in aliases]
            matched = False
            for a in cleaned_aliases:
                # Solo si empieza por el alias o contiene la frase completa
                if len(a) >= 5 and (cleaned.startswith(a) or a == cleaned):
                    matched = True
                    break
            if matched:
                rename_dict[col] = canonical
                used_canonicals.add(canonical)
                break
            
    df = df.rename(columns=rename_dict)
    
    # Descartar columnas duplicadas si las hubiere
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    # Si después del mapeo no encontramos 'Fecha Creación', intentar buscar cualquier columna con fecha o creac
    if 'Fecha Creación' not in df.columns:
        for c in df.columns:
            c_clean = clean_text_repr(str(c))
            if 'creac' in c_clean or ('fecha' in c_clean and 'fin' not in c_clean and 'cierr' not in c_clean and 'asig' not in c_clean and 'inic' not in c_clean):
                df = df.rename(columns={c: 'Fecha Creación'})
                break
                
    if 'Fecha Creación' not in df.columns:
        # Si aún no se encuentra, buscar la primera columna que tenga fechas válidas
        for c in df.columns:
            try:
                converted = pd.to_datetime(df[c], errors='coerce')
                if converted.notna().sum() > len(df) * 0.3:  # más del 30% son fechas válidas
                    df = df.rename(columns={c: 'Fecha Creación'})
                    break
            except Exception:
                continue

    if 'Fecha Creación' not in df.columns:
        cols_str = ", ".join([f"'{c}'" for c in df.columns])
        raise KeyError(
            f"No se encontró la columna de fecha de creación ('Fecha Creación') en el archivo. "
            f"Columnas detectadas en el archivo: [{cols_str}]. "
            f"Por favor asegúrate de que el Excel contenga una columna de fecha."
        )
        
    # Inicializar columnas por defecto si no existen
    defaults = {
        'Fecha Fin Atención': pd.NaT,
        'Barrio': 'SIN BARRIO',
        'Técnico Principal': 'SIN ASIGNAR',
        'Solución Técnico': 'SIN SOLUCIÓN',
        'Clasificación': 'OTROS',
        'Tipo Orden': 'OTROS',
        'Estado': 'OTROS',
        'Solicitud Suscriptor': 'OTROS',
        'Zona Cliente': 'SIN ZONA',
        'Creador': 'SISTEMA/OTROS',
        'Valor Total': 0,
        'N° Orden': range(1, len(df) + 1),
        'Código': df.get('Cliente', range(1, len(df) + 1)),
        'Cliente': 'CLIENTE GENERAL',
        'Dirección': 'SIN DIRECCIÓN',
        'Teléfono': 'SIN TELÉFONO',
        # Columnas nuevas del formato actualizado (vacías si el archivo no las trae)
        'Fecha Asignación': pd.NaT,
        'Fecha Inicio Atención': pd.NaT,
        'Fecha Cierre': pd.NaT,
        'Vencimiento (SLA)': pd.NaT,
        'Fecha Llegada': pd.NaT,
        'Tipo de Servicio': 'SIN DATO',
        'Plan Internet': 'SIN DATO',
        'Origen (Enlace/Manual)': 'SIN DATO',
        'Aplicado en Enlace': 'SIN DATO',
        'Medio Solicitud': 'SIN DATO',
        'Prioridad': 'SIN DATO',
        'Sede': 'SIN DATO',
        'Vendedor': 'SIN DATO',
        'Coordenadas': '',
        'Novedades NOC': '',
        'Observaciones': '',
        'Precinto': '',
        'Ubicación Fotos': '',
        'MAC Router': '',
        'MAC Bridge': '',
        'MAC ONT': '',
        'Otros Equipos': ''
    }
    
    for col_name, default_val in defaults.items():
        if col_name not in df.columns:
            if isinstance(default_val, range):
                df[col_name] = list(default_val)
            else:
                df[col_name] = default_val

    return df

def generate_report_from_df(df, template_path=None):
    if template_path is None:
        template_path = os.path.join(os.path.dirname(__file__), 'reporte_template.html')
    
    # Normalizar columnas de forma robusta
    df = normalize_dataframe_columns(df)
    
    # --- PROCESAMIENTO DE DATOS ---
    df['Fecha Creación'] = pd.to_datetime(df['Fecha Creación'], errors='coerce')
    df['Fecha Fin Atención'] = pd.to_datetime(df['Fecha Fin Atención'], errors='coerce')
    for _dt_col in ['Fecha Asignación', 'Fecha Inicio Atención', 'Fecha Cierre', 'Vencimiento (SLA)', 'Fecha Llegada']:
        if _dt_col in df.columns:
            df[_dt_col] = pd.to_datetime(df[_dt_col], errors='coerce')
    
    # Rellenar vacíos de forma segura
    df['Barrio'] = df['Barrio'].fillna('SIN BARRIO').astype(str).str.strip().replace('', 'SIN BARRIO')
    df['Técnico Principal'] = df['Técnico Principal'].fillna('SIN ASIGNAR').astype(str).str.strip().replace('', 'SIN ASIGNAR')
    df['Solución Técnico'] = df['Solución Técnico'].fillna('SIN SOLUCIÓN').astype(str).str.strip().replace('', 'SIN SOLUCIÓN')
    df['Clasificación'] = df['Clasificación'].fillna('OTROS').astype(str).str.strip().replace('', 'OTROS')
    df['Tipo Orden'] = df['Tipo Orden'].fillna('OTROS').astype(str).str.strip().replace('', 'OTROS')
    df['Estado'] = df['Estado'].fillna('OTROS').astype(str).str.strip().replace('', 'OTROS')
    df['Solicitud Suscriptor'] = df['Solicitud Suscriptor'].fillna('OTROS').astype(str).str.strip().replace('', 'OTROS')
    df['Zona Cliente'] = df['Zona Cliente'].fillna('SIN ZONA').astype(str).str.strip().replace('', 'SIN ZONA')
    df['Creador'] = df['Creador'].fillna('SISTEMA/OTROS').astype(str).str.strip().replace('', 'SISTEMA/OTROS')
    df['Valor Total'] = pd.to_numeric(df['Valor Total'], errors='coerce').fillna(0)

    # Tiempos
    df['Tiempo Atención'] = df['Fecha Fin Atención'] - df['Fecha Creación']
    df['Tiempo Atención (horas)'] = df['Tiempo Atención'].dt.total_seconds() / 3600
    df['Tiempo Atención Positivo'] = df['Tiempo Atención (horas)'].apply(lambda x: x if x > 0 else 0)
    
    # --- COMPARATIVA MULTIMES (NUEVO) ---
    df_full = df.copy()
    
    def get_dias_ejecucion(row):
        estado = str(row['Estado']).lower()
        if estado in ['cerrada', 'ejecutada']:
            if pd.isna(row['Tiempo Atención (horas)']) or row['Tiempo Atención (horas)'] < 0:
                return 'Mal registradas'
            dias = int(row['Tiempo Atención (horas)'] // 24)
            if dias >= 7:
                return '=> 7'
            return str(dias)
        else:
            return 'No aplica'

    df_full['Dias_Ejecucion_Cat'] = df_full.apply(get_dias_ejecucion, axis=1)
    df_full['Mes_Anio'] = df_full['Fecha Creación'].dt.to_period('M')
    
    meses_unicos = sorted([m for m in df_full['Mes_Anio'].unique() if pd.notna(m)])
    meses_comparativa = meses_unicos[-6:] if len(meses_unicos) > 3 else meses_unicos
    
    nombres_meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    
    comparativa_headers = ""
    meses_info = []
    
    for m in meses_comparativa:
        mes_nombre = f"{nombres_meses[m.month - 1].upper()} {m.year if m.year != 2026 else ''}".strip()
        if len(meses_comparativa) <= 3:
            mes_nombre = f"{nombres_meses[m.month - 1].upper()} {m.year}"
        comparativa_headers += f'<th class="text-center" style="background-color: #1C3557; color: white;">NUMERO DE<br>ORDENES<br>{mes_nombre}</th>\n'
        meses_info.append((m, mes_nombre))
        
    comparativa_headers += '<th class="text-center" style="background-color: #1C3557; color: white;"># ORDE/TOTAL<br>ORDE</th>\n<th class="text-center" style="background-color: #1C3557; color: white;">PORCENTAJE<br>ACUMULADO</th>'
    
    categorias = ['0', '1', '2', '3', '4', '5', '6', '=> 7', 'Mal registradas']
    tabla_comparativa_body = ""
    acumulado_ultimo_mes = 0
    ultimo_mes = meses_info[-1][0] if meses_info else None
    
    for cat in categorias:
        row_html = "<tr>\n"
        if cat == '6' or cat == '=> 7':
            row_html = '<tr style="background-color: #fbd4b4;">\n'
        
        texto_cat = cat
        if cat == '6' or cat == '=> 7':
            texto_cat = f'<span style="color: red;">{cat}</span>'
            
        row_html += f'<td class="text-center" style="border: 1px solid #000;">{texto_cat}</td>\n'
        
        count_ultimo_mes = 0
        for m, _ in meses_info:
            df_mes = df_full[df_full['Mes_Anio'] == m]
            count = len(df_mes[df_mes['Dias_Ejecucion_Cat'] == cat])
            row_html += f'<td class="text-center" style="border: 1px solid #000;">{count}</td>\n'
            if m == ultimo_mes:
                count_ultimo_mes = count
                
        total_m = len(df_full[(df_full['Mes_Anio'] == ultimo_mes) & (df_full['Dias_Ejecucion_Cat'] != 'No aplica')])
        if total_m == 0: total_m = 1
        pct_orde = (count_ultimo_mes / total_m * 100)
        acumulado_ultimo_mes += pct_orde
        
        if cat == 'Mal registradas':
            row_html += f'<td class="text-center" style="border: 1px solid #000;">0.00%</td>\n'
            row_html += f'<td class="text-center" style="border: 1px solid #000;">100.00%</td>\n'
        else:
            row_html += f'<td class="text-center" style="border: 1px solid #000;">{pct_orde:.2f}%</td>\n'
            row_html += f'<td class="text-center" style="border: 1px solid #000;">{acumulado_ultimo_mes:.2f}%</td>\n'
        row_html += "</tr>\n"
        tabla_comparativa_body += row_html
        
    # Totales
    row_total = '<tr style="background-color: #5b9bd5; color: white; font-weight: bold;">\n'
    row_total += '<td style="border: 1px solid #000;">TOTAL DE ORDENES</td>\n'
    for m, _ in meses_info:
        tot = len(df_full[df_full['Mes_Anio'] == m])
        row_total += f'<td class="text-center" style="border: 1px solid #000;">{tot}</td>\n'
    row_total += '<td class="text-center" style="border: 1px solid #000;">100.00%</td>\n<td style="border: 1px solid #000;"></td>\n</tr>\n'
    tabla_comparativa_body += row_total
    
    # Total órdenes diarias
    row_diarias = '<tr style="background-color: #5b9bd5; color: white; font-weight: bold;">\n'
    row_diarias += '<td style="border: 1px solid #000;">TOTAL ORDENES DIARIAS</td>\n'
    diarias_por_mes = {}
    for m, _ in meses_info:
        df_mes = df_full[df_full['Mes_Anio'] == m]
        dias_activos = df_mes['Fecha Creación'].dt.date.nunique()
        tot = len(df_mes)
        diaria = round(tot / dias_activos) if dias_activos > 0 else 0
        diarias_por_mes[m] = diaria
        row_diarias += f'<td class="text-center" style="border: 1px solid #000;">{diaria}</td>\n'
    row_diarias += '<td style="border: 1px solid #000;"></td>\n<td style="border: 1px solid #000;"></td>\n</tr>\n'
    tabla_comparativa_body += row_diarias
    
    row_pct = '<tr style="background-color: #5b9bd5; color: white; font-weight: bold;">\n'
    row_pct += '<td style="border: 1px solid #000;">PORCENTAJE</td>\n'
    for m, _ in meses_info:
        tot = len(df_full[df_full['Mes_Anio'] == m])
        diaria = diarias_por_mes[m]
        pct = (diaria / tot * 100) if tot > 0 else 0
        row_pct += f'<td class="text-center" style="border: 1px solid #000;">{pct:.2f}%</td>\n'
    row_pct += '<td style="border: 1px solid #000;"></td>\n<td style="border: 1px solid #000;"></td>\n</tr>\n'
    tabla_comparativa_body += row_pct
    
    row_cuadrilla = '<tr style="background-color: #5b9bd5; color: white; font-weight: bold;">\n'
    row_cuadrilla += '<td style="border: 1px solid #000;">TOTAL ORDENES DIARIAS X CUADRILLA</td>\n'
    for m, _ in meses_info:
        df_mes = df_full[df_full['Mes_Anio'] == m]
        valid_techs_m = [t for t in df_mes['Técnico Principal'].unique() if t not in ['SIN ASIGNAR', 'Planta externa', 'Soporte Noc', 'Soporte NOC']]
        num_cuadrillas = len(valid_techs_m)
        diaria = diarias_por_mes[m]
        x_cuadrilla = round(diaria / num_cuadrillas) if num_cuadrillas > 0 else 0
        row_cuadrilla += f'<td class="text-center" style="border: 1px solid #000;">{x_cuadrilla}</td>\n'
    row_cuadrilla += '<td style="border: 1px solid #000;"></td>\n<td style="border: 1px solid #000;"></td>\n</tr>\n'
    tabla_comparativa_body += row_cuadrilla
    
    # --- FILTRAR PARA EL RESTO DEL REPORTE ---
    if ultimo_mes is not None:
        df = df_full[df_full['Mes_Anio'] == ultimo_mes].copy()
        
    TOTAL_ORDENES = len(df)
    if TOTAL_ORDENES == 0:
        TOTAL_ORDENES = 1

    # --- ANÁLISIS DE LAS NUEVAS COLUMNAS DEL FORMATO ---
    def _text_col(col):
        if col not in df.columns:
            return pd.Series([''] * len(df), index=df.index, dtype=object)
        return df[col].fillna('').astype(str).str.strip()

    EMPTY_VALUES = ['', 'nan', 'none', 'n/a', 'na', 'sin dato', 'sin datos', 'no aplica', 'null', '-']

    def _is_filled(series):
        return ~series.str.lower().isin(EMPTY_VALUES)

    # SLA: cierre real vs vencimiento prometido
    df['_Fecha_Cierre_Efectiva'] = df['Fecha Cierre'].where(
        df['Fecha Cierre'].notna(), df['Fecha Fin Atención']
    )
    df['_SLA_Evaluada'] = df['Vencimiento (SLA)'].notna() & df['_Fecha_Cierre_Efectiva'].notna()
    df['_SLA_OK'] = df['_SLA_Evaluada'] & (df['_Fecha_Cierre_Efectiva'] <= df['Vencimiento (SLA)'])

    sla_evaluadas = int(df['_SLA_Evaluada'].sum())
    sla_cumplidas = int(df['_SLA_OK'].sum())
    sla_incumplidas = sla_evaluadas - sla_cumplidas
    sla_pct = (sla_cumplidas / sla_evaluadas * 100) if sla_evaluadas > 0 else 0
    sla_vencidas_abiertas = int((
        df['Vencimiento (SLA)'].notna()
        & df['_Fecha_Cierre_Efectiva'].isna()
        & (df['Vencimiento (SLA)'] < pd.Timestamp.now())
    ).sum())
    sla_sin_dato = int(df['Vencimiento (SLA)'].isna().sum())
    sla_pct_display = f"{sla_pct:.2f}%" if sla_evaluadas > 0 else "—"

    # Origen de la orden: enlace (automático) vs manual
    origen_txt = _text_col('Origen (Enlace/Manual)').str.lower()
    origen_valido = _is_filled(origen_txt)
    origen_enlace = int((origen_valido & origen_txt.str.contains('enlace')).sum())
    origen_manual = int((origen_valido & origen_txt.str.contains('manual')).sum())
    origen_total = origen_enlace + origen_manual
    pct_origen_enlace = (origen_enlace / origen_total * 100) if origen_total > 0 else 0
    pct_origen_enlace_display = f"{pct_origen_enlace:.2f}%" if origen_total > 0 else "—"

    # Aplicado en enlace
    aplicado_txt = _text_col('Aplicado en Enlace').str.lower()
    aplicado_count = int(aplicado_txt.isin(['si', 'sí', 's', 'yes', '1', 'true', 'x']).sum())
    pct_aplicado_enlace = (aplicado_count / TOTAL_ORDENES * 100)

    # Medio de solicitud (canal de ingreso)
    canal_txt = _text_col('Medio Solicitud')
    canal_norm = canal_txt.where(_is_filled(canal_txt), 'Sin registrar')
    canal_counts = canal_norm.value_counts()
    canal_top = canal_counts.index[0] if len(canal_counts) > 0 else 'Sin dato'
    canal_top_count = int(canal_counts.iloc[0]) if len(canal_counts) > 0 else 0
    canal_top_pct = (canal_top_count / TOTAL_ORDENES * 100)

    # Prioridad
    prio_txt = _text_col('Prioridad')

    def _prio_label(val):
        val = str(val).strip()
        if val.lower() in ['', 'nan', 'sin dato', 'sin prioridad', '0', '0.0']:
            return 'Sin prioridad'
        try:
            num = int(float(val))
        except (TypeError, ValueError):
            return val
        return f'Prioridad {num}'

    prio_norm = prio_txt.map(_prio_label)
    prio_counts = prio_norm.value_counts().sort_index()
    prioridad_alta = int((prio_norm == 'Prioridad 1').sum())
    prioridad_alta_pct = (prioridad_alta / TOTAL_ORDENES * 100)

    # Evidencia fotográfica
    fotos_txt = _text_col('Ubicación Fotos').str.lower()
    con_fotos = int((fotos_txt.ne('') & ~fotos_txt.str.contains('sin foto') & fotos_txt.ne('nan')).sum())
    pct_con_fotos = (con_fotos / TOTAL_ORDENES * 100)

    # Registro de equipos (MAC)
    mac_router = _text_col('MAC Router')
    mac_bridge = _text_col('MAC Bridge')
    mac_ont = _text_col('MAC ONT')
    otros_equipos = _text_col('Otros Equipos')
    ont_reg = int(_is_filled(mac_ont).sum())
    router_reg = int(_is_filled(mac_router).sum())
    bridge_reg = int(_is_filled(mac_bridge).sum())
    pct_ont_reg = (ont_reg / TOTAL_ORDENES * 100)
    pct_router_reg = (router_reg / TOTAL_ORDENES * 100)
    pct_bridge_reg = (bridge_reg / TOTAL_ORDENES * 100)
    otros_equipos_counts = otros_equipos[_is_filled(otros_equipos)].value_counts()

    # Novedades NOC y observaciones
    noc_txt = _text_col('Novedades NOC')
    noc_count = int(_is_filled(noc_txt).sum())
    pct_noc = (noc_count / TOTAL_ORDENES * 100)

    # Tiempo de llegada (asignación -> llegada al sitio)
    if 'Fecha Llegada' in df.columns and 'Fecha Asignación' in df.columns:
        _llegada_h = (df['Fecha Llegada'] - df['Fecha Asignación']).dt.total_seconds() / 3600
        _llegada_valida = _llegada_h[_llegada_h > 0]
        tiempo_llegada_prom = _llegada_valida.mean() if len(_llegada_valida) > 0 else 0
        ordenes_con_llegada = int(len(_llegada_valida))
    else:
        tiempo_llegada_prom = 0
        ordenes_con_llegada = 0

    # Sede, plan, tipo de servicio y vendedor
    sede_counts = _text_col('Sede').where(_is_filled(_text_col('Sede')), 'Sin registrar').value_counts()
    plan_counts = _text_col('Plan Internet').where(_is_filled(_text_col('Plan Internet')), 'Sin dato').value_counts()
    servicio_counts = _text_col('Tipo de Servicio').where(_is_filled(_text_col('Tipo de Servicio')), 'Sin dato').value_counts()
    vendedor_counts = _text_col('Vendedor').where(_is_filled(_text_col('Vendedor')), 'Sin registrar').value_counts()

    # --- MÉTRICAS ADICIONALES (SOLICITADAS) ---
    
    # 1. Órdenes Pendientes
    ordenes_pendientes = len(df[df['Estado'].str.lower() == 'pendiente'])
    # Órdenes no cerradas ni ejecutadas (Backlog general)
    ordenes_backlog = len(df[~df['Estado'].str.lower().isin(['cerrada', 'ejecutada'])])
    
    # 2. Índice de Reincidencia (Garantías) - 30 DÍAS
    id_col = 'Código' if 'Código' in df.columns else ('Cliente' if 'Cliente' in df.columns else 'Dirección')
    df_reincidencia = df.sort_values([id_col, 'Fecha Creación'])
    df_reincidencia['Diff'] = df_reincidencia.groupby(id_col)['Fecha Creación'].diff().dt.days
    reincidencias_df = df_reincidencia[df_reincidencia['Diff'] < 30].copy()
    reincidencias = len(reincidencias_df)
    tasa_reincidencia = (reincidencias / TOTAL_ORDENES * 100) if TOTAL_ORDENES > 0 else 0
    
    # Identificar clientes reincidentes para mostrar su historia completa
    reincident_ids = reincidencias_df[id_col].unique()
    reincidencia_history_body = ""
    # Tomamos los top 15 clientes con más visitas reincidentes
    top_reincident_ids = df[df[id_col].isin(reincident_ids)].groupby(id_col).size().sort_values(ascending=False).head(15).index
    
    for rid in top_reincident_ids:
        client_visits = df[df[id_col] == rid].sort_values('Fecha Creación')
        client_name = str(client_visits.iloc[0]['Cliente'])[:25] if 'Cliente' in df.columns else rid
        
        # Primera visita (referencia)
        v0 = client_visits.iloc[0]
        solicitud0 = str(v0['Solicitud Suscriptor'])[:30] if 'Solicitud Suscriptor' in df.columns else '---'
        solucion0 = str(v0['Solución Técnico'])[:30] if 'Solución Técnico' in df.columns else '---'
        
        reincidencia_history_body += format_html_table_row([
            f"ID: {rid}", 
            f"<strong>{client_name}</strong>", 
            v0['Fecha Creación'].strftime('%d/%m/%Y'), 
            "Iniciador", 
            f"<strong>Rep:</strong> {solicitud0}<br><strong>Sol:</strong> {solucion0}",
            v0['Técnico Principal']
        ])
        
        # Visitas siguientes con intervalos
        for j in range(1, len(client_visits)):
            v_curr = client_visits.iloc[j]
            v_prev = client_visits.iloc[j-1]
            diff = (v_curr['Fecha Creación'] - v_prev['Fecha Creación']).days
            
            badge_class = "badge-danger" if diff < 15 else "badge-warning"
            interval_badge = f'<span class="badge {badge_class}">{diff} días</span>'
            
            solicitud = str(v_curr['Solicitud Suscriptor'])[:30] if 'Solicitud Suscriptor' in df.columns else '---'
            solucion = str(v_curr['Solución Técnico'])[:30] if 'Solución Técnico' in df.columns else '---'
            
            reincidencia_history_body += format_html_table_row([
                "", "", 
                v_curr['Fecha Creación'].strftime('%d/%m/%Y'), 
                interval_badge, 
                f"<strong>Rep:</strong> {solicitud}<br><strong>Sol:</strong> {solucion}",
                v_curr['Técnico Principal']
            ])
    
    if len(reincident_ids) == 0:
        reincidencia_history_body = '<tr><td colspan="6" class="text-center">No se detectaron casos de reincidencia en este periodo.</td></tr>'
    
    # 3. Promedio de Órdenes Diarias por Técnico
    df['Fecha_Dia'] = df['Fecha Creación'].dt.date
    tech_active_days = df.groupby('Técnico Principal')['Fecha_Dia'].nunique()
    tech_total_orders = df.groupby('Técnico Principal').size()
    tech_avg_daily = (tech_total_orders / tech_active_days).fillna(0)
    
    # 4. Tendencia Temporal
    df['Fecha_Creacion_Dia'] = df['Fecha Creación'].dt.date
    # Basado en el campo 'Fecha Fin Atención', sin filtrar por el texto del estado
    df_con_fecha = df[df['Fecha Fin Atención'].notna()].copy()
    trend_created = df.groupby('Fecha_Creacion_Dia').size()
    trend_executed = df_con_fecha.groupby(df_con_fecha['Fecha Fin Atención'].dt.date).size()
    
    min_date = df['Fecha Creación'].min().date()
    max_date = df['Fecha Creación'].max().date()
    
    # Rango completo de fechas consecutivas del mes (mostrando todos los días y los días que faltan)
    all_days = pd.date_range(start=min_date, end=max_date).date
    
    trend_df = pd.DataFrame(index=all_days)
    trend_df['Creadas'] = trend_created
    trend_df['Ejecutadas'] = trend_executed
    trend_df = trend_df.fillna(0)
    
    # --- ESTADÍSTICAS ---
    ordenes_cerradas = len(df[df['Estado'].str.lower() == 'cerrada'])
    ordenes_ejecutadas = len(df[df['Estado'].str.lower() == 'ejecutada'])
    ordenes_reprogramadas = len(df[df['Estado'].str.lower() == 'reprogramada'])
    ordenes_anuladas = len(df[df['Estado'].str.lower() == 'anulada'])
    
    tasa_cierre = (ordenes_cerradas / TOTAL_ORDENES * 100)
    
    # Órdenes con solución (Ejecutadas o Cerradas)
    ordenes_con_solucion = len(df[df['Estado'].str.lower().isin(['cerrada', 'ejecutada'])])
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

    # 5.1 tipo_orden_pie.png (NUEVO)
    fig, ax = plt.subplots(figsize=(8, 6))
    total_to = to_counts.sum()
    
    def autopct_format(pct, allvals):
        absolute = int(round(pct/100.*np.sum(allvals)))
        return f"{absolute}, {pct:.0f}%" if pct > 0 else ""
        
    wedges, texts, autotexts = ax.pie(to_counts.values, autopct=lambda pct: autopct_format(pct, to_counts.values),
                                      textprops=dict(color="#333", fontsize=9, fontweight='bold'), colors=list(TAB20_COLORS), startangle=90, counterclock=False)
    ax.legend(wedges, to_counts.index,
              title="Tipo de orden",
              loc="center left",
              bbox_to_anchor=(1, 0, 0.5, 1),
              fontsize=9)
    ax.set_title("Total", fontweight="bold", fontsize=14)
    plt.tight_layout()
    charts['tipo_orden_pie.png'] = fig_to_base64(fig)
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

    # 8. tendencia_temporal.png
    fig, ax = plt.subplots(figsize=(14, 7))
    ax.plot(trend_df.index, trend_df['Creadas'], marker='o', label='Órdenes Creadas', color=COLORS_DICT['primary'], linewidth=2.5)
    ax.plot(trend_df.index, trend_df['Ejecutadas'], marker='s', label='Órdenes Ejecutadas', color=COLORS_DICT['secondary'], linewidth=2.5)
    
    # Agregar etiquetas de datos (Números sobre los puntos)
    for i, (x, y) in enumerate(zip(trend_df.index, trend_df['Creadas'])):
        if y > 0:
            ax.text(x, y + 0.5, f'{int(y)}', ha='center', va='bottom', fontsize=9, color=COLORS_DICT['primary'], fontweight='bold')
    
    for i, (x, y) in enumerate(zip(trend_df.index, trend_df['Ejecutadas'])):
        if y > 0:
            ax.text(x, y - 1.5, f'{int(y)}', ha='center', va='top', fontsize=9, color=COLORS_DICT['secondary'], fontweight='bold')

    ax.fill_between(trend_df.index, trend_df['Creadas'], alpha=0.05, color=COLORS_DICT['primary'])
    ax.set_title('Tendencia Temporal: Creadas vs Ejecutadas (Detalle Diario)', fontsize=14, fontweight='bold', pad=20)
    ax.legend(loc='upper right')
    ax.grid(True, linestyle='--', alpha=0.3)
    
    # Mostrar explícitamente todos los días del mes en el eje X
    ax.set_xticks(trend_df.index)
    ax.set_xticklabels([d.strftime('%d/%m') for d in trend_df.index], rotation=45, ha='right', fontsize=9)
    plt.tight_layout()
    charts['tendencia_temporal.png'] = fig_to_base64(fig)
    plt.close()

    # 9. productividad_tecnicos.png
    fig, ax = plt.subplots(figsize=(12, 6))
    prod_stats = tech_avg_daily[tech_avg_daily.index.isin(valid_techs)].sort_values(ascending=False)
    bars = ax.bar(prod_stats.index, prod_stats.values, color=COLORS_DICT['accent3'])
    ax.bar_label(bars, fmt='%.1f', padding=3)
    ax.set_title('Promedio de Órdenes Diarias por Técnico (Días Activos)', fontsize=14, fontweight='bold')
    plt.xticks(rotation=45, ha='right')
    charts['productividad_tecnicos.png'] = fig_to_base64(fig)
    plt.close()

    # 10. tendencia_reincidencia.png
    fig, ax = plt.subplots(figsize=(12, 5))
    reincidencia_trend = reincidencias_df.groupby(reincidencias_df['Fecha Creación'].dt.date).size()
    reincidencia_trend = reincidencia_trend.reindex(all_days, fill_value=0)
    ax.plot(reincidencia_trend.index, reincidencia_trend.values, marker='o', linestyle='-', color=COLORS_DICT['accent1'], linewidth=2)
    ax.fill_between(reincidencia_trend.index, reincidencia_trend.values, alpha=0.1, color=COLORS_DICT['accent1'])
    ax.set_title('Monitoreo Diario de Casos de Reincidencia', fontsize=14, fontweight='bold')
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.xticks(rotation=45)
    charts['tendencia_reincidencia.png'] = fig_to_base64(fig)
    plt.close()
    # 11. creadores_bar.png (Horizontal Top 10)
    fig, ax = plt.subplots(figsize=(12, 6))
    c_counts = df['Creador'].value_counts().head(10)
    bars = ax.barh(c_counts.index[::-1], c_counts.values[::-1], color=COLORS_DICT['accent4'])
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.set_title('Top 10 Creadores de Ordenes (Cantidad)', fontsize=14, fontweight='bold')
    charts['creadores_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 12. equipos_cambios_bar.png
    fig, ax = plt.subplots(figsize=(10, 6))
    eq_categories = {
        'Cambio de Equipo\n(Cód. 3)': df['Solución Técnico'].astype(str).str.contains(r'3 CAMBIO DE EQUIPO|cambio de equipo', case=False, na=False).sum(),
        'Equipo Defectuoso\n(Cód. 7)': df['Solución Técnico'].astype(str).str.contains('7 CAMBIO -EQUIPO DEFECTUOSO|equipo defectuoso', case=False, na=False).sum(),
        'Cambio de Bridge/Router\n(Cód. 24)': df['Solución Técnico'].astype(str).str.contains('24 CAMBIO DE BRIDGE|cambio de bridge', case=False, na=False).sum(),
        'Equipo Quemado\n(Cód. 8)': df['Solución Técnico'].astype(str).str.contains('8 CAMBIO -EQUIPO QUEMADO|equipo quemado', case=False, na=False).sum(),
        'Otros\n(Mal Uso/Cámara)': df['Solución Técnico'].astype(str).str.contains('9 CAMBIO- EQUIPO MAL USO USUARIO|25 CAMBIO DE CAMARA', case=False, na=False).sum()
    }
    eq_categories = {k: v for k, v in eq_categories.items() if v > 0}
    eq_categories = dict(sorted(eq_categories.items(), key=lambda item: item[1], reverse=True))
    
    bar_colors = [COLORS_DICT['primary'], COLORS_DICT['secondary'], COLORS_DICT['accent2'], COLORS_DICT['accent3'], COLORS_DICT['accent1']]
    bars = ax.bar(eq_categories.keys(), eq_categories.values(), color=bar_colors[:len(eq_categories)])
    ax.bar_label(bars, padding=3, fontsize=10)
    ax.set_title('Distribución de Cambios de Equipo por Tipo', fontsize=14, fontweight='bold', pad=20)
    plt.xticks(fontsize=9)
    plt.tight_layout()
    charts['equipos_cambios_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 13. origen_pie.png (Enlace vs Manual - NUEVO FORMATO)
    if origen_total > 0:
        fig, ax = plt.subplots(figsize=(7, 5))
        origen_vals = [origen_enlace, origen_manual]
        origen_labels = [f'Enlace ({origen_enlace})', f'Manual ({origen_manual})']
        ax.pie(origen_vals, labels=origen_labels, autopct='%1.1f%%',
               colors=[COLORS_DICT['secondary'], COLORS_DICT['primary']], startangle=90)
        ax.set_title('Origen de las Ordenes', fontsize=14, fontweight='bold')
        plt.tight_layout()
    else:
        fig, ax = plt.subplots(figsize=(7, 5))
        ax.text(0.5, 0.5, 'Sin informacion de origen', ha='center', va='center', fontsize=12)
        ax.axis('off')
        ax.set_title('Origen de las Ordenes', fontsize=14, fontweight='bold')
    charts['origen_pie.png'] = fig_to_base64(fig)
    plt.close()

    # 14. canal_solicitud_bar.png (Medio de solicitud)
    fig, ax = plt.subplots(figsize=(10, 6))
    canal_top10 = canal_counts.head(10)
    bars = ax.barh(canal_top10.index[::-1], canal_top10.values[::-1], color=COLORS_DICT['secondary'])
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.set_title('Ordenes por Medio de Solicitud', fontsize=14, fontweight='bold')
    plt.tight_layout()
    charts['canal_solicitud_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 15. sla_tecnico_bar.png (Cumplimiento de SLA por tecnico)
    sla_tech = df[df['_SLA_Evaluada']].groupby('Técnico Principal').agg(
        Cumplidas=('_SLA_OK', 'sum'), Evaluadas=('_SLA_OK', 'count')
    )
    sla_tech = sla_tech[sla_tech['Evaluadas'] >= 3]
    if len(sla_tech) > 0:
        sla_tech['Pct'] = sla_tech['Cumplidas'] / sla_tech['Evaluadas'] * 100
        sla_tech = sla_tech.sort_values('Pct')
        fig, ax = plt.subplots(figsize=(11, 6))
        colors_sla = [COLORS_DICT['accent1'] if v < 80 else COLORS_DICT['accent3'] for v in sla_tech['Pct']]
        bars = ax.barh(sla_tech.index, sla_tech['Pct'], color=colors_sla)
        ax.bar_label(bars, fmt='%.1f%%', padding=3, fontsize=9)
        ax.axvline(80, color='#999', linestyle='--', linewidth=1)
        ax.set_xlim(0, 110)
        ax.set_title('Cumplimiento de SLA por Tecnico (%)', fontsize=14, fontweight='bold')
        plt.tight_layout()
    else:
        fig, ax = plt.subplots(figsize=(11, 6))
        ax.text(0.5, 0.5, 'Sin datos suficientes de SLA', ha='center', va='center', fontsize=12)
        ax.axis('off')
        ax.set_title('Cumplimiento de SLA por Tecnico', fontsize=14, fontweight='bold')
    charts['sla_tecnico_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 16. prioridad_bar.png
    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.bar(prio_counts.index, prio_counts.values, color=COLORS_DICT['accent2'])
    ax.bar_label(bars, padding=3, fontsize=10)
    ax.set_title('Ordenes por Prioridad', fontsize=14, fontweight='bold')
    plt.xticks(rotation=20, ha='right')
    plt.tight_layout()
    charts['prioridad_bar.png'] = fig_to_base64(fig)
    plt.close()

    # 17. plan_internet_bar.png (Top 10 planes)
    fig, ax = plt.subplots(figsize=(11, 6))
    plan_top10 = plan_counts.head(10)
    bars = ax.barh(plan_top10.index[::-1], plan_top10.values[::-1], color=COLORS_DICT['accent4'])
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.set_title('Top 10 Planes de Internet', fontsize=14, fontweight='bold')
    plt.tight_layout()
    charts['plan_internet_bar.png'] = fig_to_base64(fig)
    plt.close()


    # --- CONSTRUCCIÓN DE TABLAS ---
    # KPI
    kpi_body = ""
    kpi_body += format_html_table_row(["Tasa de Cierre", f"<strong>{tasa_cierre:.2f}%</strong>", f"{ordenes_cerradas} de {TOTAL_ORDENES} órdenes cerradas exitosamente"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Tasa de Resolución", f"<strong>{tasa_resolucion:.2f}%</strong>", f"{ordenes_con_solucion} de {TOTAL_ORDENES} órdenes con solución (Cerradas/Ejecutadas). Nota: Hay {ordenes_anuladas} órdenes anuladas"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Órdenes sin Solución", f"<strong>{ordenes_sin_solucion}</strong>", f"{(ordenes_sin_solucion/TOTAL_ORDENES*100):.2f}% del total (requiere seguimiento)"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Barrios Atendidos", f"<strong>{barrios_unicos}</strong>", "Cobertura geográfica completa"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Tiempo Promedio", f"<strong>{tiempo_promedio:.2f} hrs</strong>", f"Mediana: {mediana_tiempo:.2f} horas"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Órdenes con Tiempo Válido", f"<strong>{ordenes_con_tiempo_valido}</strong>", f"{(ordenes_con_tiempo_valido/TOTAL_ORDENES*100):.2f}% del total"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Pendientes", f"<strong>{ordenes_pendientes}</strong>", f"Órdenes en estado 'Pendiente'"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Backlog Total", f"<strong>{ordenes_backlog}</strong>", f"Órdenes no cerradas ni ejecutadas"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Índice de Reincidencia", f"<strong>{tasa_reincidencia:.2f}%</strong>", f"{reincidencias} casos de reincidencia (<15 días)"], aligns=["", "text-center", "text-left"])
    # KPIs nuevos (columnas del formato actualizado)
    kpi_body += format_html_table_row(["Cumplimiento de SLA", f"<strong>{sla_pct_display}</strong>", f"{sla_cumplidas} de {sla_evaluadas} órdenes cerradas dentro del vencimiento (SLA). {sla_incumplidas} fuera de SLA y {sla_vencidas_abiertas} vencidas sin cerrar"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Origen de la Orden", f"<strong>{pct_origen_enlace_display}</strong>", f"{origen_enlace} órdenes por enlace vs {origen_manual} manuales; {aplicado_count} aplicadas en enlace ({pct_aplicado_enlace:.2f}%)"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Canal Principal", f"<strong>{canal_top}</strong>", f"{canal_top_count} órdenes ({canal_top_pct:.2f}%) ingresadas por este medio"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Prioridad 1", f"<strong>{prioridad_alta}</strong>", f"{prioridad_alta_pct:.2f}% de las órdenes marcadas con la prioridad más alta"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Evidencia Fotográfica", f"<strong>{pct_con_fotos:.2f}%</strong>", f"{con_fotos} de {TOTAL_ORDENES} órdenes con fotografías registradas"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Equipos Registrados (MAC)", f"<strong>{pct_ont_reg:.2f}%</strong>", f"ONT: {ont_reg}, Router: {router_reg}, Bridge: {bridge_reg} órdenes con identificación del equipo"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Novedades NOC", f"<strong>{noc_count}</strong>", f"{pct_noc:.2f}% de las órdenes tienen novedades registradas por el NOC"], aligns=["", "text-center", "text-left"])
    kpi_body += format_html_table_row(["Tiempo de Llegada", f"<strong>{tiempo_llegada_prom:.2f} hrs</strong>", f"Promedio entre asignación y llegada al sitio ({ordenes_con_llegada} órdenes con registro)"], aligns=["", "text-center", "text-left"])
    
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
        avg_daily = tech_avg_daily.get(name, 0)
        tecnicos_body += format_html_table_row([f"<strong>{name}</strong>", int(row['N° Orden']), f"{avg_daily:.1f}", f"{row['Tiempo Atención Positivo']:.2f}", badge])
    tecnicos_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "—", f"{tiempo_promedio:.2f}", "—"], is_total=True)
    
    # Soluciones
    sol_body = ""
    top_s = df['Solución Técnico'].value_counts().head(20)
    for i, (name, count) in enumerate(top_s.items(), 1):
        display_name = f'<span class="badge badge-danger">SIN SOLUCIÓN</span>' if name == 'SIN SOLUCIÓN' else name
        sol_body += format_html_table_row([i, display_name, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    sol_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", "", TOTAL_ORDENES, "100.00%"], is_total=True)

    # Creadores Ranking
    creadores_body = ""
    creator_stats = df.groupby('Creador').agg({'N° Orden': 'count'}).sort_values('N° Orden', ascending=False).head(20)
    for i, (name, row) in enumerate(creator_stats.iterrows(), 1):
        creadores_body += format_html_table_row([i, f"<strong>{name}</strong>", int(row['N° Orden'])])
    creadores_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", "", TOTAL_ORDENES], is_total=True)
    
    # Clasificación y Tipo
    class_body = ""
    for n, c in df['Clasificación'].value_counts().items():
        class_body += format_html_table_row([n, c, f"{(c/TOTAL_ORDENES*100):.2f}%"])
    class_body += format_html_table_row(["<strong>TOTAL</strong>", TOTAL_ORDENES, "100%"], is_total=True)
    
    tipo_body = ""
    for n, c in df['Tipo Orden'].value_counts().items():
        tipo_body += format_html_table_row([n, c, f"{(c/TOTAL_ORDENES*100):.2f}%"])
    tipo_body += format_html_table_row(["<strong>TOTAL</strong>", TOTAL_ORDENES, "100%"], is_total=True)
    
    # Nueva tabla tipo de orden según requerimiento
    tabla_tipo_orden_nueva_body = ""
    for n, c in to_counts.items():
        tabla_tipo_orden_nueva_body += f'<tr><td style="border: 1px solid #ccc; padding: 5px;">{n}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{c}</td></tr>\n'
    tabla_tipo_orden_nueva_body += f'<tr style="font-weight: bold; background-color: #E2EFDA;"><td style="border: 1px solid #ccc; padding: 5px;">Total general</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{total_to}</td></tr>\n'
    
    # Nueva tabla TIPO Y SOLUCION
    tabla_tipo_solucion_body = ""
    # Agrupar por Tipo Orden y contar totales
    tipo_totales = df['Tipo Orden'].value_counts()
    # Ordenar alfabéticamente el tipo de orden, similar a la imagen
    for t_orden in sorted(tipo_totales.index):
        t_count = tipo_totales[t_orden]
        tabla_tipo_solucion_body += f'<tr style="font-weight: bold;"><td style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">&#9634; {t_orden}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">{t_count}</td></tr>\n'
        
        # Obtener las soluciones para este tipo de orden, ordenadas por cantidad descendente
        df_t = df[df['Tipo Orden'] == t_orden]
        soluciones_t = df_t['Solución Técnico'].value_counts()
        for sol, s_count in soluciones_t.items():
            tabla_tipo_solucion_body += f'<tr><td style="border: 1px solid #ccc; padding: 5px; padding-left: 30px;">{sol}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{s_count}</td></tr>\n'
            
    tabla_tipo_solucion_body += f'<tr style="font-weight: bold; background-color: #DDEBF7;"><td style="border: 1px solid #ccc; padding: 5px;">Total general</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{TOTAL_ORDENES}</td></tr>\n'

    # Nueva tabla TECNICO Y SOLUCION
    tabla_tecnico_solucion_body = ""
    # Asegurarnos de que Técnico no tenga nulos (por si acaso) y agrupar
    tecnico_totales = df['Técnico Principal'].value_counts()
    for tecnico in sorted(tecnico_totales.index):
        t_count = tecnico_totales[tecnico]
        tabla_tecnico_solucion_body += f'<tr style="font-weight: bold;"><td style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">&#9634; {tecnico}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">{t_count}</td></tr>\n'
        
        # Obtener soluciones para este técnico, ordenadas por cantidad descendente
        df_tec = df[df['Técnico Principal'] == tecnico].copy()
        df_tec['Solicitud Suscriptor'] = df_tec['Solicitud Suscriptor'].fillna('Sin Solicitud')
        sol_tec_counts = df_tec['Solicitud Suscriptor'].value_counts()
        for sol, s_count in sol_tec_counts.items():
            tabla_tecnico_solucion_body += f'<tr><td style="border: 1px solid #ccc; padding: 5px; padding-left: 30px;">{sol}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{s_count}</td></tr>\n'
            
    tabla_tecnico_solucion_body += f'<tr style="font-weight: bold; background-color: #DDEBF7;"><td style="border: 1px solid #ccc; padding: 5px;">Total general</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{TOTAL_ORDENES}</td></tr>\n'

    # Nueva tabla TIPO ORDEN -> ZONA -> BARRIO
    tabla_tipo_zona_barrio_body = ""
    # Agrupar por Tipo Orden, ordenar alfabéticamente
    for t_orden in sorted(tipo_totales.index):
        t_count = tipo_totales[t_orden]
        tabla_tipo_zona_barrio_body += f'<tr style="font-weight: bold;"><td style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">&#9634; {t_orden}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">{t_count}</td></tr>\n'
        
        df_t = df[df['Tipo Orden'] == t_orden]
        # Para evitar errores con valores nulos, llenarlos con 'Sin Zona'
        df_t_zona = df_t.copy()
        df_t_zona['Zona Cliente'] = df_t_zona['Zona Cliente'].fillna('Sin Zona')
        zona_totales = df_t_zona['Zona Cliente'].value_counts()
        
        for zona in sorted(zona_totales.index):
            z_count = zona_totales[zona]
            tabla_tipo_zona_barrio_body += f'<tr style="font-weight: bold;"><td style="border: 1px solid #ccc; padding: 5px; padding-left: 20px;">&#9634; {zona}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{z_count}</td></tr>\n'
            
            df_z = df_t_zona[df_t_zona['Zona Cliente'] == zona]
            # Para evitar errores con barrios nulos
            df_z_barrio = df_z.copy()
            df_z_barrio['Barrio'] = df_z_barrio['Barrio'].fillna('Sin Barrio')
            barrio_counts = df_z_barrio['Barrio'].value_counts()
            
            for barrio, b_count in barrio_counts.items():
                tabla_tipo_zona_barrio_body += f'<tr style="font-weight: bold; color: #333;"><td style="border: 1px solid #ccc; padding: 5px; padding-left: 40px;">&#9634; {barrio}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{b_count}</td></tr>\n'
                
                # Desglose por Solicitud Suscriptor
                df_b = df_z_barrio[df_z_barrio['Barrio'] == barrio]
                df_b_solicitudes = df_b.copy()
                df_b_solicitudes['Solicitud Suscriptor'] = df_b_solicitudes['Solicitud Suscriptor'].fillna('Sin Solicitud')
                solicitud_counts = df_b_solicitudes['Solicitud Suscriptor'].value_counts()
                
                for solicitud, s_count in solicitud_counts.items():
                    tabla_tipo_zona_barrio_body += f'<tr><td style="border: 1px solid #ccc; padding: 5px; padding-left: 60px; color: #555; font-size: 0.9em;">- {solicitud}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px; color: #555; font-size: 0.9em;">{s_count}</td></tr>\n'
                
    tabla_tipo_zona_barrio_body += f'<tr style="font-weight: bold; background-color: #DDEBF7;"><td style="border: 1px solid #ccc; padding: 5px;">Total general</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{TOTAL_ORDENES}</td></tr>\n'

    # Nueva tabla TIPO ORDEN -> SOLICITUD
    tabla_tipo_solicitud_body = ""
    # Agrupar por Tipo Orden, ordenar alfabéticamente
    for t_orden in sorted(tipo_totales.index):
        t_count = tipo_totales[t_orden]
        tabla_tipo_solicitud_body += f'<tr style="font-weight: bold;"><td style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">&#9634; {t_orden}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;">{t_count}</td></tr>\n'
        
        df_t = df[df['Tipo Orden'] == t_orden]
        # Para evitar errores con valores nulos
        df_t_sol = df_t.copy()
        df_t_sol['Solicitud Suscriptor'] = df_t_sol['Solicitud Suscriptor'].fillna('Sin Solicitud')
        solicitudes_t = df_t_sol['Solicitud Suscriptor'].value_counts()
        
        for sol, s_count in solicitudes_t.items():
            tabla_tipo_solicitud_body += f'<tr><td style="border: 1px solid #ccc; padding: 5px; padding-left: 30px;">{sol}</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{s_count}</td></tr>\n'
            
    tabla_tipo_solicitud_body += f'<tr style="font-weight: bold; background-color: #DDEBF7;"><td style="border: 1px solid #ccc; padding: 5px;">Total general</td><td class="text-right" style="border: 1px solid #ccc; padding: 5px;">{TOTAL_ORDENES}</td></tr>\n'




    
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

    # --- ANÁLISIS DE CAMBIOS DE EQUIPOS ---
    df['Es_Cambio_Equipo'] = df['Solución Técnico'].astype(str).str.contains(
        r'7 CAMBIO -EQUIPO DEFECTUOSO|24 CAMBIO DE BRIDGE|8 CAMBIO -EQUIPO QUEMADO|9 CAMBIO- EQUIPO|25 CAMBIO DE CAMARA|cambio de equipo|cambio de bridge|equipo defectuoso|equipo quemado|cambio de camara',
        case=False, na=False
    )
    
    total_eq_orders = df['Es_Cambio_Equipo'].sum()
    pct_total_eq = (total_eq_orders / TOTAL_ORDENES * 100) if TOTAL_ORDENES > 0 else 0
    
    equipos_cod3_count = df['Solución Técnico'].astype(str).str.contains(r'3 CAMBIO DE EQUIPO|cambio de equipo', case=False, na=False).sum()
    equipos_defectuosos_count = df['Solución Técnico'].astype(str).str.contains('7 CAMBIO -EQUIPO DEFECTUOSO|equipo defectuoso', case=False, na=False).sum()
    equipos_bridge_count = df['Solución Técnico'].astype(str).str.contains('24 CAMBIO DE BRIDGE|cambio de bridge', case=False, na=False).sum()
    equipos_quemados_count = df['Solución Técnico'].astype(str).str.contains('8 CAMBIO -EQUIPO QUEMADO|equipo quemado', case=False, na=False).sum()
    
    # Tabla de Cambios de Equipo por Técnico
    tech_eq = df.groupby('Técnico Principal').agg(
        Total_Ordenes=('N° Orden', 'count'),
        Cambios_Equipos=('Es_Cambio_Equipo', 'sum')
    ).reset_index()
    
    tech_eq = tech_eq[tech_eq['Cambios_Equipos'] > 0].sort_values('Cambios_Equipos', ascending=False)
    
    tabla_equipos_body = ""
    for i, (idx, row) in enumerate(tech_eq.iterrows(), 1):
        name = row['Técnico Principal']
        tot_ord = int(row['Total_Ordenes'])
        eq_chg = int(row['Cambios_Equipos'])
        percentage = (eq_chg / tot_ord * 100) if tot_ord > 0 else 0
        tabla_equipos_body += format_html_table_row([
            i,
            f"<strong>{name}</strong>",
            tot_ord,
            eq_chg,
            f"{percentage:.2f}%"
        ])
    
    tabla_equipos_body += format_html_table_row([
        "<strong>TOTAL GENERAL</strong>",
        "",
        TOTAL_ORDENES,
        int(total_eq_orders),
        f"{pct_total_eq:.2f}%"
    ], is_total=True)

    # --- TABLAS DEL FORMATO ACTUALIZADO ---
    def _sla_pct_txt(evaluadas, cumplidas):
        return f"{(cumplidas / evaluadas * 100):.2f}%" if evaluadas > 0 else "—"

    # Canales de ingreso y su cumplimiento de SLA
    tabla_canales_body = ""
    for canal, count in canal_counts.items():
        mask = canal_norm == canal
        ev = int(df.loc[mask, '_SLA_Evaluada'].sum())
        ok = int(df.loc[mask, '_SLA_OK'].sum())
        tabla_canales_body += format_html_table_row([canal, count, f"{(count/TOTAL_ORDENES*100):.2f}%", _sla_pct_txt(ev, ok)])
    tabla_canales_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%", _sla_pct_txt(sla_evaluadas, sla_cumplidas)], is_total=True)

    # Origen de la orden
    pct_origen_manual = (origen_manual / origen_total * 100) if origen_total > 0 else 0
    if origen_total > 0:
        tabla_origen_body = format_html_table_row(["Enlace (generada en el enlace)", origen_enlace, f"{pct_origen_enlace:.2f}%"])
        tabla_origen_body += format_html_table_row(["Manual / Directa", origen_manual, f"{pct_origen_manual:.2f}%"])
        tabla_origen_body += format_html_table_row(["<strong>TOTAL</strong>", origen_total, "100.00%"], is_total=True)
    else:
        tabla_origen_body = '<tr><td colspan="3" class="text-center">Este archivo no incluye la columna de origen de la orden.</td></tr>'

    # Prioridad
    tabla_prioridad_body = ""
    for prio, count in prio_counts.items():
        tabla_prioridad_body += format_html_table_row([prio, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    tabla_prioridad_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%"], is_total=True)

    # SLA por técnico
    tabla_sla_tecnico_body = ""
    if len(sla_tech) > 0:
        for name, row in sla_tech.sort_values('Pct', ascending=False).iterrows():
            pct_tech = row['Pct']
            badge_class = "badge-success" if pct_tech >= 80 else ("badge-warning" if pct_tech >= 50 else "badge-danger")
            tabla_sla_tecnico_body += format_html_table_row([
                f"<strong>{name}</strong>",
                int(row['Evaluadas']),
                int(row['Cumplidas']),
                int(row['Evaluadas'] - row['Cumplidas']),
                f'<span class="badge {badge_class}">{pct_tech:.2f}%</span>'
            ])
        tabla_sla_tecnico_body += format_html_table_row([
            "<strong>TOTAL GENERAL</strong>", sla_evaluadas, sla_cumplidas, sla_incumplidas, f"{sla_pct:.2f}%"
        ], is_total=True)
    else:
        tabla_sla_tecnico_body = '<tr><td colspan="5" class="text-center">No hay datos de vencimiento (SLA) suficientes para evaluar.</td></tr>'

    # Planes de Internet
    tabla_planes_body = ""
    for plan, count in plan_counts.head(15).items():
        tabla_planes_body += format_html_table_row([plan, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    tabla_planes_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%"], is_total=True)

    # Tipo de servicio
    tabla_servicio_body = ""
    for serv, count in servicio_counts.items():
        tabla_servicio_body += format_html_table_row([serv, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    tabla_servicio_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%"], is_total=True)

    # Sedes
    tabla_sede_body = ""
    for sede, count in sede_counts.items():
        tabla_sede_body += format_html_table_row([sede, count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    tabla_sede_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", TOTAL_ORDENES, "100.00%"], is_total=True)

    # Vendedores
    tabla_vendedores_body = ""
    for i, (vend, count) in enumerate(vendedor_counts.head(15).items(), 1):
        tabla_vendedores_body += format_html_table_row([i, f"<strong>{vend}</strong>", count, f"{(count/TOTAL_ORDENES*100):.2f}%"])
    tabla_vendedores_body += format_html_table_row(["<strong>TOTAL GENERAL</strong>", "", sum(vendedor_counts.head(15).values), f"{(sum(vendedor_counts.head(15).values)/TOTAL_ORDENES*100):.2f}%"], is_total=True)

    # Registro de equipos por MAC
    tabla_equipos_registro_body = ""
    for equipo, reg, pct_reg in [('ONT', ont_reg, pct_ont_reg), ('Router', router_reg, pct_router_reg), ('Bridge', bridge_reg, pct_bridge_reg)]:
        badge_class = "badge-success" if pct_reg >= 80 else ("badge-warning" if pct_reg >= 50 else "badge-danger")
        tabla_equipos_registro_body += format_html_table_row([
            equipo, reg, TOTAL_ORDENES - reg, f'<span class="badge {badge_class}">{pct_reg:.2f}%</span>'
        ])

    # Novedades NOC más frecuentes
    tabla_novedades_body = ""
    novedades_top = noc_txt[_is_filled(noc_txt)].value_counts().head(10)
    if len(novedades_top) > 0:
        for nota, count in novedades_top.items():
            nota_corta = (nota[:110] + '…') if len(nota) > 110 else nota
            tabla_novedades_body += format_html_table_row([nota_corta, count])
    else:
        tabla_novedades_body = '<tr><td colspan="2" class="text-center">No se registraron novedades del NOC en el período.</td></tr>'

    # Otros equipos
    otros_top = otros_equipos_counts.head(10)
    tabla_otros_equipos_body = ""
    if len(otros_top) > 0:
        for eq, count in otros_top.items():
            eq_corta = (eq[:80] + '…') if len(eq) > 80 else eq
            tabla_otros_equipos_body += format_html_table_row([eq_corta, count])
    else:
        tabla_otros_equipos_body = '<tr><td colspan="2" class="text-center">Sin otros equipos registrados.</td></tr>'

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
        <strong>1. Volumen y Tendencia:</strong> Se atendieron <strong>{TOTAL_ORDENES} órdenes</strong> con una tasa de cierre del <strong>{tasa_cierre:.2f}%</strong>. El pico de demanda se observa en el gráfico de tendencia temporal.<br><br>
        <strong>2. Calidad de Servicio:</strong> El índice de reincidencia es del <strong>{tasa_reincidencia:.2f}%</strong> ({reincidencias} casos), lo que indica una buena efectividad en la primera visita.<br><br>
        <strong>3. Pendientes:</strong> Hay <strong>{ordenes_pendientes} órdenes pendientes</strong> en el sistema.<br><br>
        <strong>4. Eficiencia Técnica:</strong> <strong>{top_t}</strong> destaca no solo por volumen sino por un promedio de <strong>{tech_avg_daily.get(top_t, 0):.1f} órdenes diarias</strong>.<br><br>
        <strong>5. Tiempos:</strong> Promedio de <strong>{tiempo_promedio:.2f} horas</strong> por atención.
    """

    hallazgo_principal = f"Se logró una tasa de resolución del {tasa_resolucion:.2f}% (con {ordenes_anuladas} órdenes anuladas y {ordenes_pendientes} órdenes pendientes)."

    hallazgos_gestion = (
        f"• <strong>Cumplimiento de SLA:</strong> {sla_pct:.2f}% de las {sla_evaluadas} órdenes evaluadas se cerraron dentro del vencimiento; "
        f"{sla_incumplidas} quedaron fuera de SLA y {sla_vencidas_abiertas} están vencidas sin cerrar.<br>"
        f"• <strong>Canal de ingreso:</strong> «{canal_top}» concentra {canal_top_pct:.2f}% de las órdenes; "
        f"el {pct_origen_enlace:.2f}% se originó en el enlace y {aplicado_count} órdenes ({pct_aplicado_enlace:.2f}%) quedaron aplicadas en enlace.<br>"
        f"• <strong>Control de equipos:</strong> se registró MAC de ONT en {pct_ont_reg:.2f}% de las órdenes "
        f"({ont_reg} de {TOTAL_ORDENES}); router {pct_router_reg:.2f}% y bridge {pct_bridge_reg:.2f}%.<br>"
        f"• <strong>Evidencias y NOC:</strong> {con_fotos} órdenes ({pct_con_fotos:.2f}%) tienen fotografías y {noc_count} "
        f"({pct_noc:.2f}%) registran novedades del NOC. Tiempo promedio de llegada al sitio: {tiempo_llegada_prom:.2f} horas."
    )

    # Extraer periodo dinámicamente
    import calendar
    meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
    try:
        if ultimo_mes is not None:
            mes = ultimo_mes.month
            anio = ultimo_mes.year
            ultimo_dia = calendar.monthrange(anio, mes)[1]
            periodo_texto = f"1 de {meses[mes]} al {ultimo_dia} de {meses[mes]} de {anio}"
            solo_mes = f"{meses[mes].capitalize()} {anio}"
        else:
            valid_dates = df['Fecha Creación'].dropna()
            min_date = valid_dates.min()
            max_date = valid_dates.max()
            periodo_texto = f"{min_date.day} de {meses[min_date.month]} al {max_date.day} de {meses[max_date.month]} de {max_date.year}"
            solo_mes = f"{meses[min_date.month].capitalize()} {min_date.year}"
    except:
        periodo_texto = "1 al 31 de mayo de 2026"
        solo_mes = "Mayo 2026"

    replacements = {
        '{{ EQUIPOS_CAMBIOS_COUNT }}': str(total_eq_orders),
        '{{ PORCENTAJE_CAMBIOS_EQUIPOS }}': f"{pct_total_eq:.2f}",
        '{{ EQUIPOS_COD3_COUNT }}': str(equipos_cod3_count),
        '{{ EQUIPOS_DEFECTUOSOS_COUNT }}': str(equipos_defectuosos_count),
        '{{ EQUIPOS_QUEMADOS_COUNT }}': str(equipos_quemados_count),
        '{{ EQUIPOS_BRIDGE_COUNT }}': str(equipos_bridge_count),
        '{{ TABLA_EQUIPOS_BODY }}': tabla_equipos_body,
        '{{ TITULO_PAGINA }}': f'Reporte Órdenes de Servicio - {solo_mes}',
        '{{ REPORTE_HEADER_TITLE }}': f'Reporte Completo de Órdenes de Servicio — {solo_mes}',
        '{{ TOTAL_ORDENES }}': str(TOTAL_ORDENES),
        '{{ PERIODO_TEXTO }}': periodo_texto,
        '{{ ORDENES_CERRADAS }}': str(ordenes_cerradas),
        '{{ ORDENES_EJECUTADAS }}': str(ordenes_ejecutadas),
        '{{ ORDENES_REPROGRAMADAS }}': str(ordenes_reprogramadas),
        '{{ ORDENES_ANULADAS }}': str(ordenes_anuladas),
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
        '{{ CONCLUSIONES_SOLUCIONES }}': f'• <strong>Tasa de resolución:</strong> {tasa_resolucion:.2f}% ({ordenes_con_solucion} de {TOTAL_ORDENES} órdenes. Nota: {ordenes_anuladas} órdenes anuladas)',
        '{{ TABLA_CLASIFICACION_BODY }}': class_body,
        '{{ TABLA_TIPO_ORDEN_BODY }}': tipo_body,
        '{{ TABLA_TIPO_ORDEN_NUEVA_BODY }}': tabla_tipo_orden_nueva_body,
        '{{ TABLA_TIPO_SOLUCION_BODY }}': tabla_tipo_solucion_body,
        '{{ TABLA_TECNICO_SOLUCION_BODY }}': tabla_tecnico_solucion_body,
        '{{ TABLA_TIPO_ZONA_BARRIO_BODY }}': tabla_tipo_zona_barrio_body,
        '{{ TABLA_TIPO_SOLICITUD_BODY }}': tabla_tipo_solicitud_body,
        '{{ TABLA_ESTADO_BODY }}': estado_body,
        '{{ ORDENES_CON_TIEMPO_VALIDO }}': str(ordenes_con_tiempo_valido),
        '{{ MEDIANA_TIEMPO }}': f"{mediana_tiempo:.2f}",
        '{{ TABLA_RANGOS_TIEMPO_BODY }}': rangos_body,
        '{{ PORCENTAJE_MAS_120H }}': f"{p_mas_120:.2f}",
        '{{ TABLA_METRICAS_ADICIONALES_BODY }}': metrics_body,
        '{{ PERIODO_SOLO_MES }}': solo_mes,
        '{{ CANTIDAD_BARRIOS }}': str(barrios_unicos),
        '{{ HALLAZGOS_PRINCIPALES }}': hallazgos,
        '{{ HALLAZGO_ESTRELLA }}': hallazgo_principal,
        '{{ BACKLOG_TOTAL }}': str(ordenes_backlog),
        '{{ PENDIENTES_COUNT }}': str(ordenes_pendientes),
        '{{ TASA_REINCIDENCIA }}': f"{tasa_reincidencia:.2f}%",
        '{{ REINCIDENCIAS_COUNT }}': str(reincidencias),
        '{{ TABLA_REINCIDENCIA_BODY }}': reincidencia_history_body,
        '{{ TABLA_CREADORES_BODY }}': creadores_body,
        '{{ LISTA_RECOMENDACIONES }}': '<li>Reducir el backlog de órdenes pendientes.</li><li>Investigar casos de reincidencia para mejora de procesos.</li>',
        '{{ CLASIFICACION_GENERAL }}': f"El departamento demuestra un nivel de desempeño sólido con tasa de resolución del {tasa_resolucion:.2f}% (considerando {ordenes_anuladas} órdenes anuladas).",
        '{{ FOOTER_EMPRESA }}': "Corporación Regional de Telecomunicaciones - Kaled Molina",
        '{{ COMPARATIVA_HEADERS }}': comparativa_headers,
        '{{ TABLA_COMPARATIVA_BODY }}': tabla_comparativa_body,
        # Nuevas columnas del formato actualizado
        '{{ SLA_CUMPLIMIENTO }}': sla_pct_display,
        '{{ SLA_CUMPLIDAS }}': str(sla_cumplidas),
        '{{ SLA_INCUMPLIDAS }}': str(sla_incumplidas),
        '{{ SLA_EVALUADAS }}': str(sla_evaluadas),
        '{{ SLA_VENCIDAS_ABIERTAS }}': str(sla_vencidas_abiertas),
        '{{ ORIGEN_ENLACE_COUNT }}': str(origen_enlace),
        '{{ ORIGEN_MANUAL_COUNT }}': str(origen_manual),
        '{{ ORIGEN_ENLACE_PCT }}': pct_origen_enlace_display,
        '{{ APLICADO_ENLACE_COUNT }}': str(aplicado_count),
        '{{ APLICADO_ENLACE_PCT }}': f"{pct_aplicado_enlace:.2f}",
        '{{ CANAL_TOP }}': canal_top,
        '{{ CANAL_TOP_PCT }}': f"{canal_top_pct:.2f}",
        '{{ PRIORIDAD_ALTA_COUNT }}': str(prioridad_alta),
        '{{ PRIORIDAD_ALTA_PCT }}': f"{prioridad_alta_pct:.2f}",
        '{{ ORDENES_CON_FOTOS }}': str(con_fotos),
        '{{ PORCENTAJE_FOTOS }}': f"{pct_con_fotos:.2f}",
        '{{ PORCENTAJE_ONT }}': f"{pct_ont_reg:.2f}",
        '{{ NOVEDADES_NOC_COUNT }}': str(noc_count),
        '{{ NOVEDADES_NOC_PCT }}': f"{pct_noc:.2f}",
        '{{ TIEMPO_LLEGADA_PROM }}': f"{tiempo_llegada_prom:.2f}",
        '{{ TABLA_CANALES_BODY }}': tabla_canales_body,
        '{{ TABLA_ORIGEN_BODY }}': tabla_origen_body,
        '{{ TABLA_PRIORIDAD_BODY }}': tabla_prioridad_body,
        '{{ TABLA_SLA_TECNICO_BODY }}': tabla_sla_tecnico_body,
        '{{ TABLA_PLANES_BODY }}': tabla_planes_body,
        '{{ TABLA_SERVICIO_BODY }}': tabla_servicio_body,
        '{{ TABLA_SEDE_BODY }}': tabla_sede_body,
        '{{ TABLA_VENDEDORES_BODY }}': tabla_vendedores_body,
        '{{ TABLA_EQUIPOS_REGISTRO_BODY }}': tabla_equipos_registro_body,
        '{{ TABLA_NOVEDADES_BODY }}': tabla_novedades_body,
        '{{ TABLA_OTROS_EQUIPOS_BODY }}': tabla_otros_equipos_body,
        '{{ HALLAZGOS_GESTION }}': hallazgos_gestion
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
