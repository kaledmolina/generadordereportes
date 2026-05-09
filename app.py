from flask import Flask, render_template, request, send_file, Response
import pandas as pd
import io
import os
from generar_reporte_automatico import generate_report_from_df

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64MB limit

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'file' not in request.files:
        return "No hay archivo", 400
    
    file = request.files['file']
    if file.filename == '':
        return "Nombre de archivo vacío", 400
    
    if file and file.filename.endswith('.xlsx'):
        try:
            # Leer el excel directamente de la memoria
            df = pd.read_excel(file)
            
            # Generar el reporte
            report_html = generate_report_from_df(df)
            
            # Retornar el HTML para previsualización
            return report_html
        except Exception as e:
            return f"Error procesando el archivo: {str(e)}", 500
    
    return "Formato de archivo no soportado (debe ser .xlsx)", 400

@app.route('/download', methods=['POST'])
def download():
    # El HTML se envía desde el cliente para descargar el mismo que se previsualizó
    html_content = request.form.get('html_content')
    if not html_content:
        return "Contenido vacío", 400
    
    # Crear un buffer en memoria
    buffer = io.BytesIO()
    buffer.write(html_content.encode('utf-8'))
    buffer.seek(0)
    
    return send_file(
        buffer,
        as_attachment=True,
        download_name='reporte_ordenes.html',
        mimetype='text/html'
    )

if __name__ == '__main__':
    # Asegurarse de que las carpetas necesarias existen
    os.makedirs('templates', exist_ok=True)
    app.run(debug=True, port=5000)
