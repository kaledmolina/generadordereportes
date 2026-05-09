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

### 2. Ver Resultados (Opción Rápida)
Abre la carpeta `reporte_final/` y visualiza el archivo `reporte_generado.html` directamente en tu navegador.

### 3. Visualización mediante Servidor Local (Puerto 5000)
Si prefieres visualizar el reporte a través de un servidor local (simulando un entorno web), puedes ejecutar:

```bash

python -m http.server 5000
```

Luego, abre tu navegador e ingresa a la siguiente URL:
👉 **[http://127.0.0.1:5000/reporte_final/reporte_generado.html](http://127.0.0.1:5000/reporte_final/reporte_generado.html)**

---

---

## 🌐 Guía de Visualización en Navegador

El reporte está optimizado para una experiencia interactiva y profesional:

1.  **Navegación**: El reporte está dividido en páginas lógicas (Dashboard, Tiempos, Técnicos, Reincidencia). Puedes desplazarte hacia abajo para ver todas las secciones.
2.  **Vista Móvil**: El "Resumen Ejecutivo" (primera página) está diseñado específicamente para ser consultado desde un celular. Es ideal para tomar una captura de pantalla y enviarla por WhatsApp para una actualización rápida.
3.  **Exportar a PDF**:
    *   Presiona `Ctrl + P` (Windows) o `Cmd + P` (Mac).
    *   Selecciona **"Guardar como PDF"** como destino.
    *   **IMPORTANTE**: En la configuración de impresión, asegúrate de activar la opción **"Gráficos de fondo"** (Background Graphics) para que los colores y degradados se vean correctamente.
    *   El reporte está pre-configurado para ajustarse perfectamente al tamaño **A4**.
4.  **Interactividad**: Los gráficos son de alta resolución y se adaptan al ancho de tu pantalla para una visualización clara de las etiquetas de datos.

---

## 🌐 Guía de Despliegue en la Web (Hosting)

Si deseas que el reporte sea accesible a través de una URL pública o privada para tu equipo, sigue estas opciones:

### 1. GitHub Pages (Recomendado y Gratis)
Es la forma más rápida de tener el reporte en la web:
1.  Sube el archivo `reporte_generado.html` a un repositorio de GitHub.
2.  Ve a **Settings** > **Pages**.
3.  En "Branch", selecciona `main` y la carpeta root `/`.
4.  GitHub te dará una URL (ej. `https://usuario.github.io/proyecto/reporte_generado.html`) para compartir.

### 2. Vercel o Netlify (Despliegue Rápido)
1.  Arrastra y suelta la carpeta `reporte_final/` en el panel de control de Vercel o Netlify.
2.  Se generará un enlace público automáticamente.

### 3. Servidor Interno
*   Simplemente copia el archivo `reporte_generado.html` a la carpeta pública de tu servidor (Apache/Nginx).
*   Al ser un archivo **único y estático**, no requiere bases de datos ni configuración de servidor adicional para ser visualizado.

> [!IMPORTANT]
> **Privacidad**: Recuerda que el reporte contiene datos de clientes. Si lo subes a la web, asegúrate de que el repositorio sea **privado** o usa servicios que permitan proteger el acceso con contraseña si es información sensible.

---

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
