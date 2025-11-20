from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
import os
from rpa_engine import procesar_pedidos

app = Flask(__name__)
app.secret_key = os.environ.get('SESSION_SECRET', 'natural-conexion-rpa-2025')

UPLOAD_FOLDER = 'data'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    if 'archivo' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    file = request.files['archivo']
    
    if file.filename == '':
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            resultados = procesar_pedidos(filepath)
            
            archivos_generados = []
            archivos_disponibles = [
                'PedidosValidados.xlsx',
                'ErroresRPA.xlsx',
                'PedidosRegistradosSAG.xlsx',
                'ReporteProcesados.xlsx',
                'DashboardRPA.xlsx',
                'LogCorreos.txt',
                'graficos_dashboard.png'
            ]
            
            for archivo in archivos_disponibles:
                ruta_completa = os.path.join(app.config['OUTPUT_FOLDER'], archivo)
                if os.path.exists(ruta_completa):
                    archivos_generados.append(archivo)
            
            return render_template('resultado.html', 
                                 estadisticas=resultados, 
                                 archivos=archivos_generados)
        
        except Exception as e:
            flash(f'Error al procesar el archivo: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    else:
        flash('Tipo de archivo no permitido. Solo se permiten archivos .xlsx, .xls o .csv', 'error')
        return redirect(url_for('index'))

@app.route('/descargar/<filename>')
def descargar(filename):
    try:
        return send_file(
            os.path.join(app.config['OUTPUT_FOLDER'], filename),
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        flash(f'Error al descargar el archivo: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/limpiar')
def limpiar():
    try:
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
        flash('Archivos limpiados exitosamente', 'success')
    except Exception as e:
        flash(f'Error al limpiar archivos: {str(e)}', 'error')
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
