from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import json
import os

# Le indicamos a Flask dónde están las carpetas
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

# --- NUEVA RUTA PARA SERVIR EL FRONTEND ---
@app.route('/')
def index():
    return render_template('index.html')

# --- RUTAS DE LA API ---
@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify(leer_datos())

@app.route('/api/data', methods=['POST'])
def update_data():
    nuevo_data = request.json
    guardar_datos(nuevo_data)
    return jsonify({"mensaje": "Datos actualizados correctamente"}), 200

if __name__ == '__main__':
    # Render asigna un puerto automáticamente a través de variables de entorno
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)