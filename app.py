import sqlite3
import os
import io
import base64
import hmac
import hashlib
import subprocess
from flask import Flask, jsonify, render_template_string, request, Response, send_file, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

# --- LIBRERÍA EXCEL ---
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("⚠️ ERROR: No tienes instalada la librería 'openpyxl'.")
    print("👉 Ejecuta: pip install openpyxl")
    exit(1)

# --- LIBRERÍA PDF (NUEVO MOTOR DE REPORTES) ---
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
except ImportError:
    print("⚠️ ERROR: No tienes instalada la librería 'reportlab' para exportar a PDF.")
    print("👉 Ejecuta: pip install reportlab")
    exit(1)

# --- LIBRERÍA PARA PROCESAR IMÁGENES (EVIDENCIA 1:1) ---
try:
    from PIL import Image as PILImage
except ImportError:
    print("⚠️ ERROR: No tienes instalada la librería 'Pillow' para procesar imágenes fotográficas.")
    print("👉 Ejecuta: pip install Pillow")
    exit(1)

# --- LIBRERÍA PARA CÓDIGOS QR ---
try:
    import qrcode
except ImportError:
    print("⚠️ ERROR: No tienes instalada la librería 'qrcode'.")
    print("👉 Ejecuta: pip install qrcode")
    exit(1)

# --- CONFIGURACIÓN DE RUTAS ABSOLUTAS ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_NAME = os.path.join(BASE_DIR, "cmms_planta_cafe.db")

# Inicializamos la aplicación Flask
app = Flask(__name__)
app.secret_key = 'cafetec_super_secret_enterprise_key'

# --- CONTRASEÑA SECRETA PARA GITHUB ---
GITHUB_SECRET = 'cafetec_github_2026_seguro'

@app.before_request
def requerir_login():
    """Protege todas las rutas del sistema, excepto el login, QR estático y Webhook."""
    rutas_permitidas = ['login', 'static', 'api_etiqueta_qr', 'webhook_update']
    
    if request.endpoint not in rutas_permitidas and 'user_id' not in session:
        # MÉTODO PROFESIONAL: Guardar a dónde quería ir (ej. escanear QR) para redirigirlo tras el login
        # Evitamos guardar rutas API ocultas para no causar ciclos infinitos
        next_url = request.url if request.endpoint == 'vista_movil_maquina' else None
        return redirect(url_for('login', next=next_url))

# ==========================================
# 0. CI/CD: WEBHOOK DE GITHUB (AUTOMATIZACIÓN)
# ==========================================

@app.route('/webhook-update', methods=['POST'])
def webhook_update():
    signature = request.headers.get('X-Hub-Signature-256')
    if not signature: return jsonify({'msg': 'No se encontró la firma de seguridad'}), 403
    sha_name, signature = signature.split('=')
    if sha_name != 'sha256': return jsonify({'msg': 'Método no soportado'}), 501
    mac = hmac.new(GITHUB_SECRET.encode(), msg=request.data, digestmod=hashlib.sha256)
    if not hmac.compare_digest(mac.hexdigest(), signature): return jsonify({'msg': 'Acceso denegado.'}), 403

    try:
        subprocess.run(['git', 'fetch', '--all'], cwd=BASE_DIR, check=True)
        subprocess.run(['git', 'reset', '--hard', 'origin/main'], cwd=BASE_DIR, check=True)
        wsgi_path = '/var/www/rewey_pythonanywhere_com_wsgi.py'
        if os.path.exists(wsgi_path): os.utime(wsgi_path, None)
        return jsonify({'msg': '¡Actualización exitosa!'}), 200
    except Exception as e:
        return jsonify({'msg': f'Error al actualizar: {str(e)}'}), 500

# ==========================================
# 1. CAPA DE ACCESO A DATOS
# ==========================================

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row 
    return conn

def migrar_base_datos():
    conn = get_db_connection()
    conn.execute('''CREATE TABLE IF NOT EXISTS Usuarios (id_usuario INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, nombre_completo TEXT NOT NULL, rol TEXT NOT NULL)''')
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM Usuarios")
    if cursor.fetchone()[0] == 0:
        usuarios_defecto = [('admin', generate_password_hash('admin123'), 'Administrador Principal', 'Admin'), ('tecnico1', generate_password_hash('tec123'), 'Juan Pérez', 'Tecnico')]
        conn.executemany("INSERT INTO Usuarios (username, password_hash, nombre_completo, rol) VALUES (?, ?, ?, ?)", usuarios_defecto)
    conn.execute('''CREATE TABLE IF NOT EXISTS Maquinas (id_maquina INTEGER PRIMARY KEY AUTOINCREMENT, codigo_equipo TEXT UNIQUE NOT NULL, nombre TEXT NOT NULL, area_planta TEXT NOT NULL, criticidad TEXT NOT NULL, estado TEXT DEFAULT 'Operativa', fecha_instalacion TEXT DEFAULT (date('now', 'localtime')))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS Repuestos_Stock (id_repuesto INTEGER PRIMARY KEY AUTOINCREMENT, codigo_pieza TEXT UNIQUE NOT NULL, nombre TEXT NOT NULL, descripcion TEXT, ubicacion_almacen TEXT, cantidad_actual REAL DEFAULT 0, punto_reorden REAL DEFAULT 0, unidad_medida TEXT DEFAULT 'Unidad')''')
    conn.execute('''CREATE TABLE IF NOT EXISTS Calendario_Mantenimiento (id_mantenimiento INTEGER PRIMARY KEY AUTOINCREMENT, id_maquina INTEGER NOT NULL, tipo_mantenimiento TEXT NOT NULL, descripcion_tarea TEXT NOT NULL, fecha_programada TEXT NOT NULL, fecha_ejecucion TEXT, estado_orden TEXT DEFAULT 'Pendiente', tecnico_asignado TEXT, observaciones TEXT DEFAULT 'Sin observaciones registradas.', recomendaciones TEXT DEFAULT 'Ninguna.', trabajos_realizados TEXT DEFAULT 'No especificado.', equipos_necesarios TEXT DEFAULT 'Ninguno.', tiempo_ejecucion TEXT DEFAULT '0 h', FOREIGN KEY (id_maquina) REFERENCES Maquinas (id_maquina))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS Repuestos_Orden (id_registro INTEGER PRIMARY KEY AUTOINCREMENT, id_mantenimiento INTEGER NOT NULL, id_repuesto INTEGER NOT NULL, cantidad_usada REAL NOT NULL, costo_unitario_historico REAL, FOREIGN KEY (id_mantenimiento) REFERENCES Calendario_Mantenimiento (id_mantenimiento) ON DELETE CASCADE, FOREIGN KEY (id_repuesto) REFERENCES Repuestos_Stock (id_repuesto))''')
    try: conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN observaciones TEXT DEFAULT 'Sin observaciones registradas.'")
    except sqlite3.OperationalError: pass 
    try: conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN recomendaciones TEXT DEFAULT 'Ninguna.'")
    except sqlite3.OperationalError: pass 
    try: conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN trabajos_realizados TEXT DEFAULT 'No especificado.'")
    except sqlite3.OperationalError: pass 
    try: conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN equipos_necesarios TEXT DEFAULT 'Ninguno.'")
    except sqlite3.OperationalError: pass 
    try: conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN tiempo_ejecucion TEXT DEFAULT '0 h'")
    except sqlite3.OperationalError: pass 
    try: conn.execute("""CREATE TABLE IF NOT EXISTS Evidencia_Fotografica (id_evidencia INTEGER PRIMARY KEY AUTOINCREMENT, id_mantenimiento INTEGER NOT NULL, imagen_base64 TEXT NOT NULL, FOREIGN KEY (id_mantenimiento) REFERENCES Calendario_Mantenimiento (id_mantenimiento) ON DELETE CASCADE)""")
    except sqlite3.OperationalError: pass
    try: conn.execute("""CREATE TABLE IF NOT EXISTS Compras_Repuestos (id_compra INTEGER PRIMARY KEY AUTOINCREMENT, id_repuesto INTEGER NOT NULL, cantidad REAL NOT NULL, costo_unitario REAL NOT NULL, proveedor TEXT, factura TEXT, fecha_compra TEXT DEFAULT (date('now', 'localtime')), FOREIGN KEY (id_repuesto) REFERENCES Repuestos_Stock (id_repuesto) ON DELETE CASCADE)""")
    except sqlite3.OperationalError: pass
    conn.commit()
    conn.close()

def obtener_todas_las_maquinas():
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn.execute(f"UPDATE Maquinas SET estado = 'En Mantenimiento' WHERE id_maquina IN (SELECT id_maquina FROM Calendario_Mantenimiento WHERE estado_orden = 'Pendiente' AND fecha_programada <= '{hoy}')")
    conn.execute(f"UPDATE Maquinas SET estado = 'Operativa' WHERE id_maquina NOT IN (SELECT id_maquina FROM Calendario_Mantenimiento WHERE estado_orden = 'Pendiente' AND fecha_programada <= '{hoy}') AND estado = 'En Mantenimiento'")
    conn.commit()
    maquinas = conn.execute('SELECT * FROM Maquinas').fetchall()
    conn.close()
    return [dict(ix) for ix in maquinas]

def obtener_mantenimientos_pendientes():
    conn = get_db_connection()
    query = "SELECT c.id_mantenimiento, m.codigo_equipo, m.nombre AS maquina_nombre, c.tipo_mantenimiento, c.descripcion_tarea, c.fecha_programada, c.estado_orden, c.tecnico_asignado, c.id_maquina FROM Calendario_Mantenimiento c JOIN Maquinas m ON c.id_maquina = m.id_maquina WHERE c.estado_orden IN ('Pendiente', 'En Progreso') ORDER BY c.fecha_programada ASC"
    ordenes = conn.execute(query).fetchall()
    conn.close()
    return [dict(ix) for ix in ordenes]

def obtener_inventario():
    conn = get_db_connection()
    repuestos = conn.execute('SELECT * FROM Repuestos_Stock').fetchall()
    conn.close()
    return [dict(ix) for ix in repuestos]

def descontar_stock_db(id_repuesto):
    conn = get_db_connection()
    conn.execute('UPDATE Repuestos_Stock SET cantidad_actual = cantidad_actual - 1 WHERE id_repuesto = ? AND cantidad_actual > 0', (id_repuesto,))
    conn.commit()
    conn.close()

def sumar_stock_db(id_repuesto):
    conn = get_db_connection()
    conn.execute('UPDATE Repuestos_Stock SET cantidad_actual = cantidad_actual + 1 WHERE id_repuesto = ?', (id_repuesto,))
    conn.commit()
    conn.close()

def registrar_compra_db(id_repuesto, cantidad, costo, proveedor, factura):
    conn = get_db_connection()
    try:
        conn.execute('INSERT INTO Compras_Repuestos (id_repuesto, cantidad, costo_unitario, proveedor, factura) VALUES (?, ?, ?, ?, ?)', (id_repuesto, cantidad, costo, proveedor, factura))
        conn.execute('UPDATE Repuestos_Stock SET cantidad_actual = cantidad_actual + ? WHERE id_repuesto = ?', (cantidad, id_repuesto))
        conn.commit()
    except Exception as e:
        conn.rollback()
    finally:
        conn.close()

def obtener_historial_compras_db():
    conn = get_db_connection()
    query = "SELECT c.id_compra, r.codigo_pieza, r.nombre, c.cantidad, c.costo_unitario, (c.cantidad * c.costo_unitario) as costo_total, c.proveedor, c.factura, c.fecha_compra FROM Compras_Repuestos c JOIN Repuestos_Stock r ON c.id_repuesto = r.id_repuesto ORDER BY c.fecha_compra DESC, c.id_compra DESC"
    compras = conn.execute(query).fetchall()
    conn.close()
    return [dict(ix) for ix in compras]

def crear_nueva_orden_db(id_maquina, tipo, descripcion, fecha, tecnico):
    conn = get_db_connection()
    conn.execute("INSERT INTO Calendario_Mantenimiento (id_maquina, tipo_mantenimiento, descripcion_tarea, fecha_programada, tecnico_asignado, estado_orden) VALUES (?, ?, ?, ?, ?, 'Pendiente')", (id_maquina, tipo, descripcion, fecha, tecnico))
    hoy = datetime.now().strftime('%Y-%m-%d')
    if fecha <= hoy:
        conn.execute("UPDATE Maquinas SET estado = 'En Mantenimiento' WHERE id_maquina = ?", (id_maquina,))
    conn.commit()
    conn.close()

def actualizar_orden_db(id_orden, id_maquina, tipo, descripcion, fecha, tecnico):
    conn = get_db_connection()
    conn.execute('UPDATE Calendario_Mantenimiento SET id_maquina = ?, tipo_mantenimiento = ?, descripcion_tarea = ?, fecha_programada = ?, tecnico_asignado = ? WHERE id_mantenimiento = ?', (id_maquina, tipo, descripcion, fecha, tecnico, id_orden))
    conn.commit()
    conn.close()

def eliminar_orden_db(id_orden):
    conn = get_db_connection()
    conn.execute('DELETE FROM Calendario_Mantenimiento WHERE id_mantenimiento = ?', (id_orden,))
    conn.commit()
    conn.close()

def completar_orden_db(id_mantenimiento, repuestos_usados, obs, rec, trabajos, equipos, tiempo, evidencias=[]):
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    try:
        conn.execute("UPDATE Calendario_Mantenimiento SET estado_orden = 'Completada', fecha_ejecucion = ?, observaciones = ?, recomendaciones = ?, trabajos_realizados = ?, equipos_necesarios = ?, tiempo_ejecucion = ? WHERE id_mantenimiento = ?", (hoy, obs, rec, trabajos, equipos, tiempo, id_mantenimiento))
        orden = conn.execute('SELECT id_maquina FROM Calendario_Mantenimiento WHERE id_mantenimiento = ?', (id_mantenimiento,)).fetchone()
        if orden:
            conn.execute("UPDATE Maquinas SET estado = 'Operativa' WHERE id_maquina = ?", (orden['id_maquina'],))
        for rep in repuestos_usados:
            id_rep = int(rep['id_repuesto'])
            cantidad = float(rep['cantidad'])
            conn.execute("INSERT INTO Repuestos_Orden (id_mantenimiento, id_repuesto, cantidad_usada, costo_unitario_historico) VALUES (?, ?, ?, 0.0)", (id_mantenimiento, id_rep, cantidad))
            conn.execute('UPDATE Repuestos_Stock SET cantidad_actual = cantidad_actual - ? WHERE id_repuesto = ?', (cantidad, id_rep))
        for b64 in evidencias:
            conn.execute('INSERT INTO Evidencia_Fotografica (id_mantenimiento, imagen_base64) VALUES (?, ?)', (id_mantenimiento, b64))
        conn.commit()
    except Exception as e:
        conn.rollback()
    finally:
        conn.close()

def obtener_historial_db():
    conn = get_db_connection()
    query = "SELECT c.id_mantenimiento, m.codigo_equipo, m.nombre AS maquina_nombre, m.area_planta, c.tipo_mantenimiento, c.descripcion_tarea, c.fecha_ejecucion, c.tecnico_asignado, c.tiempo_ejecucion, (SELECT GROUP_CONCAT(ro.cantidad_usada || ' ' || rs.unidad_medida || ' de ' || rs.nombre, '\n') FROM Repuestos_Orden ro JOIN Repuestos_Stock rs ON ro.id_repuesto = rs.id_repuesto WHERE ro.id_mantenimiento = c.id_mantenimiento) as repuestos_usados FROM Calendario_Mantenimiento c JOIN Maquinas m ON c.id_maquina = m.id_maquina WHERE c.estado_orden = 'Completada' ORDER BY c.fecha_ejecucion DESC"
    historial = conn.execute(query).fetchall()
    conn.close()
    return [dict(ix) for ix in historial]

def crear_maquina_db(codigo, nombre, area, criticidad):
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn.execute("INSERT INTO Maquinas (codigo_equipo, nombre, area_planta, criticidad, estado, fecha_instalacion) VALUES (?, ?, ?, ?, 'Operativa', ?)", (codigo, nombre, area, criticidad, hoy))
    conn.commit()
    conn.close()

# ==========================================
# 2. INTERFACES DE USUARIO HTML (TEMPLATES)
# ==========================================

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CAFETEC | Login CMMS</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Rajdhani:wght@600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; } 
        .fuente-logo { font-family: 'Rajdhani', sans-serif; letter-spacing: 1px;}
    </style>
</head>
<body class="bg-[#050f1a] flex items-center justify-center h-screen relative overflow-hidden">
    <div class="absolute top-[-10%] left-[-10%] w-96 h-96 bg-cyan-600 rounded-full mix-blend-multiply filter blur-[128px] opacity-20 animate-blob"></div>
    <div class="absolute bottom-[-10%] right-[-10%] w-96 h-96 bg-blue-600 rounded-full mix-blend-multiply filter blur-[128px] opacity-20 animate-blob animation-delay-2000"></div>

    <div class="bg-slate-900 p-10 rounded-2xl shadow-[0_0_40px_rgba(0,0,0,0.5)] border border-slate-800 w-full max-w-md relative z-10">
        <div class="flex items-center justify-center mb-6">
            <div class="flex items-center bg-slate-950 border-2 border-slate-700 rounded-lg px-4 py-2 shadow-[0_0_20px_rgba(6,182,212,0.15)]">
                <span class="text-cyan-400 font-bold text-4xl leading-none fuente-logo">CAFE</span>
                <span class="text-white font-bold text-4xl leading-none fuente-logo ml-1">TEC</span>
            </div>
        </div>
        
        <h2 class="text-center text-cyan-500 font-bold mb-8 tracking-widest text-xs uppercase flex items-center justify-center">
            <div class="h-px bg-slate-800 flex-1 mr-4"></div>
            Software CMMS
            <div class="h-px bg-slate-800 flex-1 ml-4"></div>
        </h2>
        
        {% if error %}
        <div class="bg-rose-900/40 border border-rose-500 text-rose-300 px-4 py-3 rounded-lg mb-6 text-sm text-center flex items-center justify-center font-medium">
            <i class="fa-solid fa-triangle-exclamation mr-2 text-rose-500"></i>{{ error }}
        </div>
        {% endif %}
        
        <form method="POST" action="/login" class="space-y-6">
            {% if next_url %}
            <!-- Campo oculto para la redirección inteligente después de iniciar sesión -->
            <input type="hidden" name="next" value="{{ next_url }}">
            {% endif %}
            <div>
                <label class="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-2">Usuario Asignado</label>
                <div class="relative">
                    <div class="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                        <i class="fa-solid fa-user text-slate-500"></i>
                    </div>
                    <input type="text" name="username" required autocomplete="off" class="w-full bg-slate-950 border border-slate-800 text-white rounded-xl py-3.5 pl-11 pr-4 focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500 outline-none transition-all placeholder-slate-700 text-sm font-medium" placeholder="Ingrese su usuario">
                </div>
            </div>
            <div>
                <label class="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-2">Contraseña</label>
                <div class="relative">
                    <div class="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                        <i class="fa-solid fa-lock text-slate-500"></i>
                    </div>
                    <input type="password" name="password" required class="w-full bg-slate-950 border border-slate-800 text-white rounded-xl py-3.5 pl-11 pr-4 focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500 outline-none transition-all placeholder-slate-700 text-sm font-medium" placeholder="••••••••">
                </div>
            </div>
            <button type="submit" class="w-full bg-cyan-600 hover:bg-cyan-500 text-white font-bold py-3.5 rounded-xl shadow-lg shadow-cyan-900/30 hover:shadow-cyan-500/25 transition-all mt-4 flex justify-center items-center group">
                Ingresar al Sistema 
                <i class="fa-solid fa-arrow-right-to-bracket ml-2 transform group-hover:translate-x-1 transition-transform"></i>
            </button>
        </form>
    </div>
</body>
</html>
"""

# HTML_TEMPLATE Y MONITOR_TEMPLATE NO SE HAN MODIFICADO PARA AHORRAR ESPACIO Y MANTENER LA INTEGRIDAD
HTML_TEMPLATE = """<!DOCTYPE html><html lang="es" class="dark"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard CAFETEC</title><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-slate-950 text-white flex items-center justify-center h-screen"><h1>Dashboard CMMS Activo</h1><p>Ingresa a las rutas de escritorio para administrar.</p></body></html>"""
MONITOR_TEMPLATE = """<!DOCTYPE html><html lang="es" class="dark"><head><meta charset="UTF-8"><title>Monitor CAFETEC</title><script src="https://cdn.tailwindcss.com"></script></head><body class="bg-[#020617] text-white flex items-center justify-center h-screen"><h1>Monitor de Planta</h1></body></html>"""

MOBILE_MACHINE_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ maquina.codigo_equipo }} | CAFETEC</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #0f172a; color: white; }
        .toast-enter { transform: translateY(100%); opacity: 0; }
        .toast-enter-active { transform: translateY(0); opacity: 1; transition: all 0.3s ease-out; }
    </style>
</head>
<body class="pb-20 relative">

    <!-- Toast Notification (Mobile) -->
    <div id="mobile-toast" class="fixed bottom-5 left-1/2 transform -translate-x-1/2 bg-slate-800 text-white px-4 py-3 rounded-xl shadow-2xl border border-slate-700 z-[100] flex items-center w-11/12 max-w-sm toast-enter hidden">
        <i id="mobile-toast-icon" class="fa-solid fa-circle-info mr-3 text-cyan-400 text-lg"></i>
        <span id="mobile-toast-msg" class="text-sm font-semibold">Mensaje</span>
    </div>

    <!-- Header Fijo -->
    <header class="bg-slate-900 border-b border-slate-800 p-4 sticky top-0 z-50 shadow-lg flex justify-between items-center">
        <div class="flex flex-col">
            <span class="text-cyan-400 font-bold text-sm tracking-widest uppercase">Perfil de Máquina</span>
            <span class="font-black text-2xl text-white">{{ maquina.codigo_equipo }}</span>
        </div>
        <a href="/" class="bg-slate-800 text-white w-10 h-10 rounded-full flex items-center justify-center shadow-inner border border-slate-700 active:bg-slate-700 transition-colors">
            <i class="fa-solid fa-house"></i>
        </a>
    </header>

    <div class="p-4 space-y-6">
        <!-- 1. Tarjeta Principal de la Máquina -->
        <div class="bg-slate-900 rounded-2xl p-5 border border-slate-800 shadow-xl relative overflow-hidden">
            {% if maquina.estado == 'Operativa' %}
                <div class="absolute left-0 top-0 bottom-0 w-1.5 bg-emerald-500"></div>
            {% else %}
                <div class="absolute left-0 top-0 bottom-0 w-1.5 bg-rose-500"></div>
            {% endif %}
            
            <h1 class="text-xl font-bold mb-1">{{ maquina.nombre }}</h1>
            <p class="text-slate-400 text-sm mb-4"><i class="fa-solid fa-location-dot mr-1"></i> Área: {{ maquina.area_planta }}</p>
            
            <div class="grid grid-cols-2 gap-4">
                <div class="bg-slate-950 p-3 rounded-xl border border-slate-800">
                    <span class="text-slate-500 text-[10px] uppercase font-bold tracking-wider block mb-1">Estado Actual</span>
                    {% if maquina.estado == 'Operativa' %}
                        <span class="text-emerald-400 font-bold text-sm bg-emerald-900/30 px-2 py-1 rounded"><i class="fa-solid fa-check-circle mr-1"></i> OPERATIVA</span>
                    {% else %}
                        <span class="text-rose-400 font-bold text-sm bg-rose-900/30 px-2 py-1 rounded"><i class="fa-solid fa-triangle-exclamation mr-1"></i> DETENIDA</span>
                    {% endif %}
                </div>
                <div class="bg-slate-950 p-3 rounded-xl border border-slate-800">
                    <span class="text-slate-500 text-[10px] uppercase font-bold tracking-wider block mb-1">Criticidad</span>
                    <span class="text-white font-bold text-sm bg-slate-800 px-2 py-1 rounded">{{ maquina.criticidad }}</span>
                </div>
            </div>
        </div>

        <!-- 2. Sección de Órdenes Pendientes -->
        <div>
            <div class="flex justify-between items-center mb-3 ml-1">
                <h2 class="text-slate-400 font-bold uppercase tracking-widest text-xs flex items-center">
                    <i class="fa-solid fa-clipboard-list mr-2 text-cyan-500"></i> Tareas Pendientes
                </h2>
                <span class="bg-slate-800 text-slate-300 text-[10px] font-bold px-2 py-1 rounded-lg">{{ ordenes|length }}</span>
            </div>
            
            {% if ordenes|length == 0 %}
                <div class="bg-slate-900/50 border border-slate-800 border-dashed rounded-xl p-6 text-center">
                    <i class="fa-solid fa-check-circle text-4xl text-emerald-500/50 mb-2"></i>
                    <p class="text-slate-500 text-sm font-medium">Equipo al día. No hay tareas programadas.</p>
                </div>
            {% else %}
                <div class="space-y-3">
                {% for orden in ordenes %}
                    <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-md relative overflow-hidden">
                        <div class="absolute left-0 top-0 bottom-0 w-1 {% if orden.tipo_mantenimiento == 'Preventivo' %}bg-cyan-500{% elif orden.tipo_mantenimiento == 'Correctivo' %}bg-rose-500{% else %}bg-slate-400{% endif %}"></div>
                        
                        <div class="pl-2">
                            <div class="flex justify-between items-start mb-2">
                                <span class="text-xs font-bold {% if orden.tipo_mantenimiento == 'Preventivo' %}text-cyan-400{% elif orden.tipo_mantenimiento == 'Correctivo' %}text-rose-400{% else %}text-slate-300{% endif %} uppercase tracking-wider">{{ orden.tipo_mantenimiento }}</span>
                                <span class="text-[10px] font-bold text-amber-500 bg-amber-500/10 px-2 py-0.5 rounded border border-amber-500/20"><i class="fa-regular fa-clock mr-1"></i>{{ orden.fecha_programada }}</span>
                            </div>
                            <p class="text-sm font-medium text-slate-200 mb-3">{{ orden.descripcion_tarea }}</p>
                            
                            <!-- Botón para Completar Orden Directamente -->
                            <button onclick="abrirModalReporteMovel({{ orden.id_mantenimiento }})" class="w-full bg-slate-800 hover:bg-slate-700 active:bg-slate-600 text-white font-bold py-2.5 rounded-lg text-sm transition-colors border border-slate-700 flex justify-center items-center shadow-inner">
                                <i class="fa-solid fa-file-signature mr-2 text-cyan-400"></i> Generar Reporte
                            </button>
                        </div>
                    </div>
                {% endfor %}
                </div>
            {% endif %}
        </div>

        <!-- 3. Historial Rápido de Intervenciones -->
        <div>
            <h2 class="text-slate-400 font-bold uppercase tracking-widest text-xs mt-6 mb-3 ml-1 flex items-center">
                <i class="fa-solid fa-clock-rotate-left mr-2 text-slate-500"></i> Últimas Intervenciones
            </h2>
            
            {% if historial|length == 0 %}
                <p class="text-xs text-slate-500 italic ml-1">Sin historial previo registrado.</p>
            {% else %}
                <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-md">
                    <ul class="divide-y divide-slate-800">
                    {% for h in historial %}
                        <li class="p-3">
                            <div class="flex justify-between items-center mb-1">
                                <span class="text-xs font-bold text-slate-300">{{ h.tipo_mantenimiento }}</span>
                                <span class="text-[10px] text-slate-500">{{ h.fecha_ejecucion }}</span>
                            </div>
                            <p class="text-xs text-slate-400 truncate">{{ h.descripcion_tarea }}</p>
                            <div class="mt-1 flex items-center text-[10px] text-slate-500">
                                <i class="fa-solid fa-user-tag mr-1 text-slate-600"></i> {{ h.tecnico_asignado }}
                            </div>
                        </li>
                    {% endfor %}
                    </ul>
                </div>
            {% endif %}
        </div>
    </div>

    <!-- MODAL MÓVIL: Formulario de Reporte Completo CON TODOS LOS ITEMS FALTANTES -->
    <div id="modal-reporte-movil" class="fixed inset-0 bg-slate-950/90 z-[60] hidden flex-col justify-end transform transition-transform duration-300 translate-y-full">
        <div class="bg-slate-900 w-full h-[95vh] rounded-t-3xl border-t border-slate-700 flex flex-col shadow-[0_-10px_40px_rgba(0,0,0,0.5)]">
            
            <div class="w-full flex justify-center py-3" onclick="cerrarModalReporteMovel()">
                <div class="w-12 h-1.5 bg-slate-700 rounded-full"></div>
            </div>
            
            <div class="px-5 pb-4 border-b border-slate-800 flex justify-between items-center shrink-0">
                <h3 class="text-lg font-bold text-white flex items-center">
                    <i class="fa-solid fa-flag-checkered text-emerald-500 mr-2"></i> Finalizar Trabajo
                </h3>
                <button onclick="cerrarModalReporteMovel()" class="w-8 h-8 rounded-full bg-slate-800 flex items-center justify-center text-slate-400">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>
            
            <div class="p-5 overflow-y-auto flex-1 space-y-4">
                <input type="hidden" id="movil-id-orden">
                
                <div>
                    <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1">Trabajos Realizados (Descripción Detallada)</label>
                    <textarea id="movil-trabajos" rows="2" class="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-cyan-500 outline-none" placeholder="Describe paso a paso lo que se hizo..."></textarea>
                </div>
                
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1">Tiempo (Hrs)</label>
                        <input type="text" id="movil-tiempo" class="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-cyan-500 outline-none" placeholder="Ej: 1.5 h">
                    </div>
                    <div>
                        <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1">Equipos Usados</label>
                        <input type="text" id="movil-equipos" class="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-cyan-500 outline-none" placeholder="Herramientas...">
                    </div>
                </div>

                <div>
                    <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1">Observaciones Encontradas</label>
                    <textarea id="movil-obs" rows="2" class="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-cyan-500 outline-none" placeholder="¿Qué fallas o detalles encontraste?"></textarea>
                </div>

                <div>
                    <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1">Recomendaciones Futuras</label>
                    <textarea id="movil-rec" rows="2" class="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-sm text-white focus:border-cyan-500 outline-none" placeholder="Ej: Cambiar filtro el próximo mes..."></textarea>
                </div>

                <div class="bg-slate-800/50 p-3 rounded-xl border border-slate-700">
                    <label class="block text-[11px] font-bold text-cyan-400 uppercase mb-2"><i class="fa-solid fa-box-open mr-1"></i> Repuestos Utilizados</label>
                    <div class="flex space-x-2">
                        <select id="movil-select-repuesto" class="flex-1 min-w-0 bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white outline-none truncate"></select>
                        <input type="number" id="movil-input-cant" value="1" min="0.1" step="0.1" class="w-16 bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white text-center outline-none">
                        <button type="button" onclick="agregarRepuestoMovil()" class="bg-cyan-600 text-white px-3 rounded-lg font-bold">
                            <i class="fa-solid fa-plus"></i>
                        </button>
                    </div>
                    <ul id="movil-lista-repuestos" class="mt-3 divide-y divide-slate-700/50 max-h-32 overflow-y-auto">
                        <li class="py-2 text-xs text-slate-500 italic text-center">No se han agregado repuestos.</li>
                    </ul>
                </div>

                <div>
                    <label class="block text-[11px] font-bold text-slate-500 uppercase mb-1"><i class="fa-solid fa-camera mr-1"></i> Evidencia Fotográfica</label>
                    <input type="file" id="movil-fotos" multiple accept="image/*" class="w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-bold file:bg-cyan-900/50 file:text-cyan-400 bg-slate-950 border border-slate-800 rounded-xl" onchange="previewMobileImages(event)">
                    <div id="movil-preview-container" class="grid grid-cols-2 gap-2 mt-3 hidden"></div>
                </div>
            </div>
            
            <div class="p-4 border-t border-slate-800 bg-slate-900 shrink-0">
                <button onclick="enviarReporteMovil()" id="btn-enviar-movil" class="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-3.5 rounded-xl shadow-lg transition-colors flex justify-center items-center">
                    <i class="fa-solid fa-flag-checkered mr-2"></i> Guardar Reporte y Finalizar
                </button>
            </div>
        </div>
    </div>

    <script>
        let evidenciasBase64Movil = [];
        let repuestosUsadosMovil = [];

        function showMobileToast(msg, isError = false) {
            const toast = document.getElementById('mobile-toast');
            const icon = document.getElementById('mobile-toast-icon');
            document.getElementById('mobile-toast-msg').innerText = msg;
            toast.classList.remove('hidden');
            toast.classList.add('toast-enter-active');
            
            if (isError) {
                toast.classList.replace('bg-slate-800', 'bg-rose-900');
                toast.classList.replace('border-slate-700', 'border-rose-700');
                icon.className = "fa-solid fa-triangle-exclamation mr-3 text-white text-lg";
            } else {
                toast.classList.replace('bg-rose-900', 'bg-slate-800');
                toast.classList.replace('border-rose-700', 'border-slate-700');
                icon.className = "fa-solid fa-circle-check mr-3 text-emerald-400 text-lg";
            }
            setTimeout(() => {
                toast.classList.remove('toast-enter-active');
                setTimeout(() => toast.classList.add('hidden'), 300);
            }, 3000);
        }

        function abrirModalReporteMovel(idOrden) {
            document.getElementById('movil-id-orden').value = idOrden;
            
            document.getElementById('movil-trabajos').value = '';
            document.getElementById('movil-tiempo').value = '';
            document.getElementById('movil-equipos').value = '';
            document.getElementById('movil-obs').value = '';
            document.getElementById('movil-rec').value = '';
            document.getElementById('movil-fotos').value = '';
            
            const previewContainer = document.getElementById('movil-preview-container');
            previewContainer.innerHTML = '';
            previewContainer.classList.add('hidden');
            
            evidenciasBase64Movil = [];
            repuestosUsadosMovil = [];
            actualizarListaUI();

            // COMO EL TÉCNICO YA ESTÁ LOGUEADO, ESTA API SÍ RESPONDERÁ CON LOS REPUESTOS
            fetch('/api/inventario').then(res => res.json()).then(data => {
                const select = document.getElementById('movil-select-repuesto');
                select.innerHTML = '<option value="" disabled selected>Selecciona un repuesto...</option>';
                data.forEach(r => {
                    if(r.cantidad_actual > 0) select.innerHTML += `<option value="${r.id_repuesto}" data-nombre="${r.codigo_pieza} - ${r.nombre}">${r.codigo_pieza} (Stock: ${r.cantidad_actual})</option>`;
                });
            }).catch(err => console.error("Error al cargar repuestos: ", err));

            const modal = document.getElementById('modal-reporte-movil');
            modal.classList.remove('hidden');
            setTimeout(() => modal.classList.remove('translate-y-full'), 10);
        }

        function cerrarModalReporteMovel() {
            const modal = document.getElementById('modal-reporte-movil');
            modal.classList.add('translate-y-full');
            setTimeout(() => modal.classList.add('hidden'), 300);
        }

        function agregarRepuestoMovil() {
            const select = document.getElementById('movil-select-repuesto');
            const cant = document.getElementById('movil-input-cant').value;
            if(!select.value || cant <= 0) { 
                showMobileToast("Selecciona repuesto y cantidad.", true); 
                return; 
            }
            repuestosUsadosMovil.push({ id_repuesto: select.value, nombre: select.options[select.selectedIndex].getAttribute('data-nombre'), cantidad: cant });
            actualizarListaUI();
        }

        function actualizarListaUI() {
            const ul = document.getElementById('movil-lista-repuestos');
            if(repuestosUsadosMovil.length === 0) { ul.innerHTML = '<li class="py-2 text-xs text-slate-500 italic text-center">No se han agregado repuestos.</li>'; return; }
            ul.innerHTML = '';
            repuestosUsadosMovil.forEach((item, index) => {
                ul.innerHTML += `<li class="py-2 flex justify-between items-center text-xs text-slate-300"><span><span class="text-cyan-400 font-bold">${item.cantidad}x</span> ${item.nombre}</span><button type="button" onclick="quitarRepuestoMovil(${index})" class="text-rose-500 w-6 h-6 bg-slate-900 rounded"><i class="fa-solid fa-xmark"></i></button></li>`;
            });
        }

        function quitarRepuestoMovil(i) { repuestosUsadosMovil.splice(i, 1); actualizarListaUI(); }

        function previewMobileImages(event) {
            const files = event.target.files;
            const previewContainer = document.getElementById('movil-preview-container');
            previewContainer.innerHTML = '';
            evidenciasBase64Movil = [];

            if(files.length > 0) previewContainer.classList.remove('hidden');
            else previewContainer.classList.add('hidden');

            Array.from(files).forEach(file => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const b64 = e.target.result;
                    evidenciasBase64Movil.push(b64.split(',')[1]);
                    
                    const img = document.createElement('img');
                    img.src = b64;
                    img.className = 'w-full h-24 object-cover rounded-xl border border-slate-700';
                    previewContainer.appendChild(img);
                };
                reader.readAsDataURL(file);
            });
        }

        function enviarReporteMovil() {
            const idOrden = document.getElementById('movil-id-orden').value;
            const btn = document.getElementById('btn-enviar-movil');
            btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin mr-2"></i> Guardando...';
            btn.disabled = true;

            const payload = {
                repuestos: repuestosUsadosMovil,
                observaciones: document.getElementById('movil-obs').value || 'Sin observaciones.',
                recomendaciones: document.getElementById('movil-rec').value || 'Ninguna recomendación.',
                trabajos_realizados: document.getElementById('movil-trabajos').value || 'Trabajo completado.',
                equipos_necesarios: document.getElementById('movil-equipos').value || 'Herramientas estándar.',
                tiempo_ejecucion: document.getElementById('movil-tiempo').value || 'No especificado',
                evidencias: evidenciasBase64Movil
            };

            fetch(`/api/ordenes/completar/${idOrden}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
            .then(res => res.json())
            .then(r => {
                if(r.status === 'ok') {
                    showMobileToast('Reporte guardado exitosamente.');
                    cerrarModalReporteMovel();
                    setTimeout(() => window.location.reload(), 1000);
                } else {
                    showMobileToast('Error al guardar reporte.', true);
                    btn.innerHTML = '<i class="fa-solid fa-flag-checkered mr-2"></i> Intentar de nuevo';
                    btn.disabled = false;
                }
            })
            .catch(err => {
                showMobileToast('Error de conexión.', true);
                btn.innerHTML = '<i class="fa-solid fa-flag-checkered mr-2"></i> Intentar de nuevo';
                btn.disabled = false;
            });
        }
    </script>
</body>
</html>
"""

# ==========================================
# 3. CONTROLADORES DE RUTAS
# ==========================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    # Detectamos a dónde quería ir el usuario antes de pedirle login
    next_url = request.args.get('next') or request.form.get('next')
    
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        conn = get_db_connection()
        user = conn.execute('SELECT * FROM Usuarios WHERE username = ?', (username,)).fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id_usuario']
            session['username'] = user['username']
            session['nombre_completo'] = user['nombre_completo']
            session['rol'] = user['rol']
            
            # REDIRECCIÓN INTELIGENTE: Si escaneó el QR, lo mandamos a la máquina, si no, al Dashboard.
            if next_url:
                return redirect(next_url)
            return redirect(url_for('dashboard'))
            
        return render_template_string(LOGIN_TEMPLATE, error="Usuario o contraseña incorrectos", next_url=next_url)
    
    return render_template_string(LOGIN_TEMPLATE, error=None, next_url=next_url)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
def dashboard(): 
    return render_template_string(HTML_TEMPLATE)

@app.route('/monitor')
def monitor():
    return render_template_string(MONITOR_TEMPLATE)

@app.route('/maquina/<int:id_maquina>')
def vista_movil_maquina(id_maquina):
    conn = get_db_connection()
    maquina = conn.execute('SELECT * FROM Maquinas WHERE id_maquina = ?', (id_maquina,)).fetchone()
    if not maquina:
        conn.close()
        return "Máquina no encontrada", 404
        
    ordenes_pendientes = conn.execute('''
        SELECT * FROM Calendario_Mantenimiento 
        WHERE id_maquina = ? AND estado_orden IN ('Pendiente', 'En Progreso')
        ORDER BY fecha_programada ASC
    ''', (id_maquina,)).fetchall()
    
    historial = conn.execute('''
        SELECT * FROM Calendario_Mantenimiento 
        WHERE id_maquina = ? AND estado_orden = 'Completada'
        ORDER BY fecha_ejecucion DESC LIMIT 3
    ''', (id_maquina,)).fetchall()
    conn.close()
    
    return render_template_string(MOBILE_MACHINE_TEMPLATE, 
                                  maquina=maquina, 
                                  ordenes=[dict(o) for o in ordenes_pendientes],
                                  historial=[dict(h) for h in historial])

@app.route('/api/maquinas/etiqueta/<int:id_maquina>')
def api_etiqueta_qr(id_maquina):
    conn = get_db_connection()
    m = conn.execute('SELECT * FROM Maquinas WHERE id_maquina = ?', (id_maquina,)).fetchone()
    conn.close()
    if not m: return "Máquina no encontrada", 404
    
    url_maquina = request.host_url.rstrip('/') + url_for('vista_movil_maquina', id_maquina=id_maquina)
    qr = qrcode.QRCode(version=1, box_size=10, border=1)
    qr.add_data(url_maquina)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
    
    ETIQUETA_HTML = """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Imprimir Etiqueta QR - {{ m.codigo_equipo }}</title>
        <style>
            body { font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background: #e2e8f0; margin: 0; }
            .etiqueta { width: 400px; height: 200px; background: white; border: 2px solid black; display: flex; padding: 10px; box-sizing: border-box; border-radius: 8px;}
            .info { flex: 1; display: flex; flex-col; flex-direction: column; justify-content: center; padding-right: 10px; }
            .qr-container { width: 150px; display: flex; justify-content: center; align-items: center; border-left: 2px dashed #cbd5e1; padding-left: 10px; }
            h1 { margin: 0; font-size: 24px; color: #0f172a; }
            h2 { margin: 5px 0 0 0; font-size: 14px; color: #64748b; font-weight: normal; }
            .logo { font-family: Impact, sans-serif; font-size: 20px; font-weight: bold; margin-bottom: 15px; }
            .logo-c { color: #0ea5e9; }
            img { max-width: 100%; max-height: 100%; }
            @media print { body { background: white; } .no-print { display: none; } .etiqueta { border: none; border-radius: 0; } }
            .btn-imprimir { position: fixed; top: 20px; padding: 10px 20px; background: #0ea5e9; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; font-size: 16px;}
        </style>
    </head>
    <body>
        <button class="btn-imprimir no-print" onclick="window.print()">🖨️ Imprimir Etiqueta</button>
        <div class="etiqueta">
            <div class="info">
                <div class="logo"><span class="logo-c">CAFE</span>TEC</div>
                <h1>{{ m.codigo_equipo }}</h1>
                <h2>{{ m.nombre }}</h2>
            </div>
            <div class="qr-container"><img src="data:image/png;base64,{{ qr_img }}" alt="QR"></div>
        </div>
    </body>
    </html>
    """
    return render_template_string(ETIQUETA_HTML, m=m, qr_img=img_str)

@app.route('/api/monitor_data')
def api_monitor_data():
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    maquinas = conn.execute('SELECT codigo_equipo, nombre, estado FROM Maquinas').fetchall()
    tot_maq = len(maquinas)
    op_maq = sum(1 for m in maquinas if m['estado'] == 'Operativa')
    detenidas = [dict(m) for m in maquinas if m['estado'] != 'Operativa']
    pendientes = conn.execute("SELECT COUNT(*) FROM Calendario_Mantenimiento WHERE estado_orden = 'Pendiente' AND fecha_programada <= ?", (hoy,)).fetchone()[0]
    completadas_hoy = conn.execute("SELECT COUNT(*) FROM Calendario_Mantenimiento WHERE estado_orden = 'Completada' AND fecha_ejecucion = ?", (hoy,)).fetchone()[0]
    tot_ord = pendientes + completadas_hoy
    conn.close()
    return jsonify({"maquinas": {"total": tot_maq, "operativas": op_maq, "detenidas": detenidas}, "ordenes": {"total": tot_ord, "completadas": completadas_hoy}})

@app.route('/api/maquinas')
def api_maquinas(): return jsonify(obtener_todas_las_maquinas())

@app.route('/api/ordenes')
def api_ordenes(): return jsonify(obtener_mantenimientos_pendientes())

@app.route('/api/historial')
def api_historial(): return jsonify(obtener_historial_db())

@app.route('/api/inventario')
def api_inventario(): return jsonify(obtener_inventario())

@app.route('/api/compras')
def api_compras(): return jsonify(obtener_historial_compras_db())

@app.route('/api/inventario/descontar/<int:id_repuesto>', methods=['POST'])
def api_descontar_inventario(id_repuesto):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    descontar_stock_db(id_repuesto)
    return jsonify({"status": "ok"})

@app.route('/api/inventario/sumar/<int:id_repuesto>', methods=['POST'])
def api_sumar_inventario(id_repuesto):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    sumar_stock_db(id_repuesto)
    return jsonify({"status": "ok"})

@app.route('/api/inventario/comprar/<int:id_repuesto>', methods=['POST'])
def api_registrar_compra(id_repuesto):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    d = request.json
    registrar_compra_db(id_repuesto, float(d['cantidad']), float(d['costo']), d['proveedor'], d['factura'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/nueva', methods=['POST'])
def api_nueva_orden():
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    d = request.json
    crear_nueva_orden_db(d['id_maquina'], d['tipo_mantenimiento'], d['descripcion'], d['fecha'], d['tecnico'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/editar/<int:id_orden>', methods=['POST'])
def api_editar_orden(id_orden):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    d = request.json
    actualizar_orden_db(id_orden, d['id_maquina'], d['tipo_mantenimiento'], d['descripcion'], d['fecha'], d['tecnico'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/eliminar/<int:id_orden>', methods=['DELETE'])
def api_borrar_orden(id_orden):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    eliminar_orden_db(id_orden)
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/completar/<int:id_orden>', methods=['POST'])
def api_completar_orden(id_orden):
    d = request.json or {}
    completar_orden_db(id_orden, d.get('repuestos', []), d.get('observaciones', ''), d.get('recomendaciones', ''), d.get('trabajos_realizados', ''), d.get('equipos_necesarios', ''), d.get('tiempo_ejecucion', ''), d.get('evidencias', []))
    return jsonify({"status": "ok"})

@app.route('/api/maquinas/nueva', methods=['POST'])
def api_nueva_maquina():
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    d = request.json
    crear_maquina_db(d['codigo'], d['nombre'], d['area'], d['criticidad'])
    return jsonify({"status": "ok"})

@app.route('/api/maquinas/eliminar/<int:id_maquina>', methods=['DELETE'])
def api_borrar_maquina(id_maquina):
    if session.get('rol') != 'Admin': return jsonify({"status": "error"}), 403
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM Maquinas WHERE id_maquina = ?', (id_maquina,))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return jsonify({"status": "ok"}) if exito else jsonify({"status": "error"})

@app.route('/api/inventario/exportar')
def api_exportar_inventario():
    inventario = obtener_inventario()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock CAFETEC"
    encabezados = ['Código Pieza', 'Nombre', 'Descripción', 'Cantidad Actual', 'Punto Reorden', 'Unidad de Medida', 'Ubicación']
    ws.append(encabezados)
    fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font = Font(color="22D3EE", bold=True)
    for col_num in range(1, 8):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill
        cell.font = font
    for item in inventario: 
        ws.append([item['codigo_pieza'], item['nombre'], item['descripcion'], item['cantidad_actual'], item['punto_reorden'], item['unidad_medida'], item['ubicacion_almacen']])
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return send_file(file_stream, as_attachment=True, download_name="Inventario_Cafetec.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/api/compras/exportar')
def api_exportar_compras():
    compras = obtener_historial_compras_db()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historial de Compras"
    encabezados = ['Fecha Compra', 'Código Pieza', 'Nombre de Repuesto', 'Cantidad Comprada', 'Costo Unitario ($)', 'Costo Total ($)', 'Proveedor', 'Factura / OC']
    ws.append(encabezados)
    fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font = Font(color="22D3EE", bold=True)
    for col_num in range(1, 9):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill
        cell.font = font
    for item in compras: 
        ws.append([item['fecha_compra'], item['codigo_pieza'], item['nombre'], item['cantidad'], item['costo_unitario'], item['costo_total'], item['proveedor'], item['factura']])
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return send_file(file_stream, as_attachment=True, download_name="Compras_Cafetec.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/api/reporte/<int:id_orden>')
def api_descargar_reporte(id_orden):
    conn = get_db_connection()
    query = """
        SELECT c.id_mantenimiento, m.codigo_equipo, m.nombre AS maquina_nombre, m.area_planta,
               c.tipo_mantenimiento, c.descripcion_tarea, c.fecha_programada, c.fecha_ejecucion, 
               c.tecnico_asignado, c.observaciones, c.recomendaciones, 
               c.trabajos_realizados, c.equipos_necesarios, c.tiempo_ejecucion,
               (SELECT GROUP_CONCAT(ro.cantidad_usada || ' ' || rs.unidad_medida || ' de ' || rs.nombre, '\n')
                FROM Repuestos_Orden ro JOIN Repuestos_Stock rs ON ro.id_repuesto = rs.id_repuesto WHERE ro.id_mantenimiento = c.id_mantenimiento) as repuestos_usados
        FROM Calendario_Mantenimiento c JOIN Maquinas m ON c.id_maquina = m.id_maquina WHERE c.id_mantenimiento = ?
    """
    orden = conn.execute(query, (id_orden,)).fetchone()
    evidencias_db = conn.execute("SELECT imagen_base64 FROM Evidencia_Fotografica WHERE id_mantenimiento = ?", (id_orden,)).fetchall()
    conn.close()

    if not orden: return "Orden no encontrada", 404

    file_stream = io.BytesIO()
    doc = SimpleDocTemplate(file_stream, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements = []
    title_style = ParagraphStyle(name='Title', fontName='Helvetica-Bold', fontSize=11, spaceAfter=8, spaceBefore=14, textColor=colors.HexColor('#0F172A'))
    normal_style = ParagraphStyle(name='NormalCustom', fontName='Helvetica', fontSize=10, spaceAfter=4, leading=14)

    logo_path = "logo_cafetec.jpg"
    celda_logo = Image(logo_path, width=112, height=55, kind='proportional') if os.path.exists(logo_path) else Paragraph("<font color='white'><b>CAFETEC</b></font>", ParagraphStyle(name='LogoFB', alignment=1, fontName='Times-Bold', fontSize=14))
    p_tit_top = Paragraph("REGISTRO DE MANTENIMIENTO", ParagraphStyle(name='TitTop', alignment=1, fontName='Times-Bold', fontSize=11, leading=14))
    p_tit_bot = Paragraph(f"MANTENIMIENTO DE MAQUINA {orden['maquina_nombre'].upper()}", ParagraphStyle(name='TitBot', alignment=1, fontName='Times-Bold', fontSize=10, textColor=colors.HexColor('#004b87'), leading=12)) 
    lbl_style = ParagraphStyle(name='Lbl', alignment=0, fontName='Times-Roman', fontSize=10)
    val_style = ParagraphStyle(name='Val', alignment=0, fontName='Times-Roman', fontSize=10)
    fecha_format = orden['fecha_ejecucion']
    
    header_data = [[celda_logo, p_tit_top, Paragraph("Código:", lbl_style), Paragraph("RE MAN-262", val_style)], ["", "", Paragraph("Versión:", lbl_style), Paragraph("001", val_style)], ["", p_tit_bot, Paragraph("Fecha:", lbl_style), Paragraph(fecha_format, val_style)], ["", "", Paragraph("Página:", lbl_style), Paragraph("1", val_style)]]
    t_header = Table(header_data, colWidths=[120, 250, 60, 80], rowHeights=[15, 15, 15, 15])
    t_header.setStyle(TableStyle([('SPAN', (0, 0), (0, 3)), ('SPAN', (1, 0), (1, 1)), ('SPAN', (1, 2), (1, 3)), ('BACKGROUND', (0, 0), (0, 3), colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 1, colors.black)]))
    elements.append(t_header)
    elements.append(Spacer(1, 15))

    elements.append(Paragraph("1. INFORMACIÓN GENERAL:", title_style))
    info_text = f"• <b>Fecha de Ejecución:</b> {orden['fecha_ejecucion']}<br/>• <b>Tipo de trabajo a realizar:</b> {orden['tipo_mantenimiento']}<br/>• <b>Código del activo:</b> {orden['codigo_equipo']}<br/>• <b>Nombre de activo:</b> {orden['maquina_nombre']}<br/>• <b>Técnico a cargo:</b> {orden['tecnico_asignado']}<br/>• <b>Duración:</b> {orden['tiempo_ejecucion']}"
    elements.append(Paragraph(info_text, normal_style))
    elements.append(Paragraph("2. TRABAJO SOLICITADO:", title_style))
    elements.append(Paragraph(f"• Mantenimiento {orden['tipo_mantenimiento'].lower()} reportado: {orden['descripcion_tarea']}", normal_style))
    elements.append(Paragraph("3. TRABAJOS REALIZADOS:", title_style))
    elements.append(Paragraph(orden['trabajos_realizados'].replace('\n', '<br/>'), normal_style))
    elements.append(Paragraph("4. DESCRIPCIÓN Y OBSERVACIONES:", title_style))
    elements.append(Paragraph(f"• {orden['observaciones'].replace(chr(10), '<br/>')}", normal_style))
    elements.append(Paragraph("RECURSOS NECESARIOS:", title_style))
    rep_text = "• " + orden['repuestos_usados'].replace('\n', '<br/>• ') if orden['repuestos_usados'] else "<i>No se utilizaron repuestos de almacén.</i>"
    equipos_text = "<br/>".join([f"• {e.strip()}" for e in orden['equipos_necesarios'].split(',') if e.strip()]) or "• Herramientas manuales estándar"
    recursos_text = f"<b>Mano de obra:</b><br/>• 01 técnico de mantenimiento ({orden['tecnico_asignado']})<br/><br/><b>Equipos necesarios:</b><br/>{equipos_text}<br/><br/><b>Materiales y repuestos:</b><br/>{rep_text}"
    elements.append(Paragraph(recursos_text, normal_style))
    elements.append(Paragraph("5. COMENTARIOS / RECOMENDACIONES:", title_style))
    elements.append(Paragraph(f"• {orden['recomendaciones'].replace(chr(10), '<br/>')}", normal_style))
    elements.append(Paragraph("6. EVIDENCIA FOTOGRÁFICA:", title_style))
    if evidencias_db:
        evidencia_data, row = [], []
        for index, row_db in enumerate(evidencias_db):
            try:
                img_data = base64.b64decode(row_db['imagen_base64'])
                img_pil = PILImage.open(io.BytesIO(img_data)).convert('RGB')
                w, h = img_pil.size
                ns = min(w, h)
                img_pil = img_pil.crop(((w-ns)/2, (h-ns)/2, (w+ns)/2, (h+ns)/2)).resize((300, 300))
                output = io.BytesIO()
                img_pil.save(output, format='JPEG', quality=85)
                output.seek(0)
                row.append(Image(output, width=240, height=240))
                if len(row) == 2:
                    evidencia_data.append(row)
                    row = []
            except: pass
        if row: 
            row.append("")
            evidencia_data.append(row)
        t_evidencia = Table(evidencia_data, colWidths=[255, 255])
        t_evidencia.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        elements.append(t_evidencia)
    else:
        box = Table([["\n\n(Espacio reservado para adjuntar evidencia fotográfica post-impresión)\n\n"]], colWidths=[510])
        box.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8'))]))
        elements.append(box)
        
    elements.append(Spacer(1, 85))
    firmas_data = [["________________________", "________________________", "________________________"], [f"Elaborado:\n{orden['tecnico_asignado']}", "Revisado:\nÁrea de Mantenimiento", "Aprobado:\nGerencia de Planta"]]
    t_firmas = Table(firmas_data, colWidths=[170, 170, 170])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica'), ('FONTSIZE', (0,0), (-1,-1), 9)]))
    elements.append(t_firmas)

    doc.build(elements)
    file_stream.seek(0)
    return send_file(file_stream, as_attachment=True, download_name=f"RE-MAN-262_{orden['codigo_equipo']}_{orden['fecha_ejecucion']}.pdf", mimetype='application/pdf')

if __name__ == '__main__':
    print("=====================================================")
    print("⚙️  INICIANDO CAFETEC CMMS - ÁREA DE MANTENIMIENTO")
    print("=====================================================")
    migrar_base_datos()
    app.run(host='0.0.0.0', debug=False, port=5000)
