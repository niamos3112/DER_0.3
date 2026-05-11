from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import json
import os

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)
DATA_FILE = 'data.json'

def leer_datos():
    if not os.path.exists(DATA_FILE):
        return {"linkPostulacion": "#", "directorio": [], "eventos": [], "anexos": []}
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def guardar_datos(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# ==========================================
# RUTAS DE LAS PÁGINAS WEB (FRONTEND)
# ==========================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/innovacion-emprendedores')
def emprendedores():
    return render_template('innovacionEmprendedores.html')

@app.route('/innovacion-empresarial')
def empresarial():
    return render_template('innovacionEmpresarial.html')

# ==========================================
# RUTAS DE LA API (DATOS)
# ==========================================
@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify(leer_datos())

@app.route('/api/data', methods=['POST'])
def update_data():
    nuevo_data = request.json
    guardar_datos(nuevo_data)
    return jsonify({"mensaje": "Datos actualizados"}), 200

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)