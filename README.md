# Generador de Reportes Automático - Órdenes de Servicio

Este proyecto es una herramienta automatizada diseñada para transformar datos brutos de órdenes de servicio (en formato Excel) en reportes profesionales de alta calidad en formato HTML. El reporte incluye análisis de tiempos, productividad técnica, índices de reincidencia y un dashboard ejecutivo optimizado para móviles.

---

## 📋 Requisitos del Sistema

Para ejecutar este generador, asegúrate de tener instalado:

1.  **Python 3.8 o superior**
2.  **Bibliotecas de Python necesarias:**
    *   `pandas`: Para el procesamiento y análisis de datos.
    *   `matplotlib`: Para la generación de gráficos.
    *   `openpyxl`: Para la lectura de archivos Excel (.xlsx).
    *   `jinja2` (Opcional, si se usa para plantillas avanzadas).

### Instalación de Dependencias
Ejecuta el siguiente comando en tu terminal para instalar todo lo necesario:
```bash
pip install pandas matplotlib openpyxl
```

---

## 🚀 Guía de Despliegue Local

1.  **Clonar o Descargar el Proyecto:**
    Asegúrate de tener la siguiente estructura de carpetas:
    ```text
    /package
    ├── user_input_files/          # Aquí se colocan los archivos Excel de entrada
    ├── reporte_final/             # Aquí se generará el reporte final
    ├── generar_reporte_automatico.py  # Script principal
    └── reporte_template.html      # Plantilla base del reporte
    ```

2.  **Preparar los Datos:**
    Exporta el reporte de órdenes desde tu sistema y guárdalo en la carpeta `user_input_files/`. El script procesará automáticamente el archivo más reciente que encuentre en esa carpeta.

---

## 📖 Guía de Uso

### 1. Generación del Reporte
Simplemente ejecuta el script principal desde la raíz de la carpeta `package`:
```bash
python generar_reporte_automatico.py
```

### 2. Ver Resultados
Una vez que el script termine (verás el mensaje "Reporte generado con exactitud"), abre la carpeta `reporte_final/` y visualiza el archivo:
*   `reporte_generado.html`: Puedes abrirlo con cualquier navegador (Chrome, Edge, Safari).

### 3. Compartir
El reporte está diseñado para ser **auto-contenido**. Todas las imágenes y gráficos están embebidos directamente en el archivo HTML, lo que significa que puedes enviarlo por correo o WhatsApp y se verá perfectamente sin necesidad de archivos adicionales.

---

## 📊 Métricas Clave Incluidas

*   **Dashboard Ejecutivo (One-Pager):** Resumen ultra-rápido de KPIs críticos (Total, Tasa de Resolución, Backlog).
*   **Índice de Reincidencia (30 días):** Monitoreo de clientes que requieren más de una visita en un mes, con historial detallado y soluciones aplicadas.
*   **Tendencia Temporal:** Gráfico detallado de órdenes creadas vs. ejecutadas día por día.
*   **Productividad Técnica:** Análisis de órdenes diarias promedio por técnico (normalizado por días activos).
*   **Análisis de Tiempos:** Distribución de tiempos de respuesta y cumplimiento de SLAs (48h / 5 días).

---

## 🛠️ Personalización

*   **Colores y Logos:** Puedes modificar los colores base en el diccionario `COLORS_DICT` dentro del script de Python para ajustarlos a la marca de la empresa.
*   **Plantilla:** El diseño visual se puede ajustar editando el archivo `reporte_template.html` (CSS y Estructura).

---
**Desarrollado para:** Área de Innovación y Desarrollo - Kaled Molina
