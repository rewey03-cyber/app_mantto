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

# --- CONFIGURACIÓN DE RUTAS ABSOLUTAS (PARA EL SERVIDOR EN LA NUBE) ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_NAME = os.path.join(BASE_DIR, "cmms_planta_cafe.db")

# Inicializamos la aplicación Flask
app = Flask(__name__)
# Llave secreta necesaria para mantener las sesiones de usuario seguras
app.secret_key = 'cafetec_super_secret_enterprise_key'

# --- CONTRASEÑA SECRETA PARA GITHUB (WEBHOOK) ---
GITHUB_SECRET = 'cafetec_github_2026_seguro'

@app.before_request
def requerir_login():
    """Protege todas las rutas del sistema, excepto el login, QR, vista móvil y Webhook de GitHub."""
    rutas_permitidas = ['login', 'static', 'api_etiqueta_qr', 'vista_movil_maquina', 'webhook_update']
    if request.endpoint not in rutas_permitidas and 'user_id' not in session:
        return redirect(url_for('login'))

# ==========================================
# 0. CI/CD: WEBHOOK DE GITHUB (AUTOMATIZACIÓN)
# ==========================================

@app.route('/webhook-update', methods=['POST'])
def webhook_update():
    """Endpoint automático para actualizar la app en PythonAnywhere al hacer push a GitHub."""
    signature = request.headers.get('X-Hub-Signature-256')
    if not signature:
        return jsonify({'msg': 'Firma de seguridad ausente'}), 403
        
    sha_name, signature = signature.split('=')
    if sha_name != 'sha256':
        return jsonify({'msg': 'Algoritmo no soportado'}), 501

    mac = hmac.new(GITHUB_SECRET.encode(), msg=request.data, digestmod=hashlib.sha256)
    if not hmac.compare_digest(mac.hexdigest(), signature):
        return jsonify({'msg': 'Firma secreta inválida'}), 403

    try:
        subprocess.run(['git', 'fetch', '--all'], cwd=BASE_DIR, check=True)
        subprocess.run(['git', 'reset', '--hard', 'origin/main'], cwd=BASE_DIR, check=True)
        
        # Tocar el archivo WSGI para reiniciar automáticamente el servidor en PythonAnywhere
        wsgi_path = '/var/www/rewey_pythonanywhere_com_wsgi.py'
        if os.path.exists(wsgi_path):
            os.utime(wsgi_path, None)

        return jsonify({'msg': '¡Actualización y reinicio automático completados!'}), 200
    except Exception as e:
        return jsonify({'msg': f'Error en despliegue: {str(e)}'}), 500

# ==========================================
# 1. CAPA DE ACCESO A DATOS (CRUD BÁSICO Y AVANZADO)
# ==========================================

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row 
    return conn

def migrar_base_datos():
    conn = get_db_connection()
    
    # 1. Crear Tabla de Usuarios
    conn.execute('''
        CREATE TABLE IF NOT EXISTS Usuarios (
            id_usuario INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            nombre_completo TEXT NOT NULL,
            rol TEXT NOT NULL
        )
    ''')

    # 2. Inyectar Usuarios por Defecto (Si la tabla está vacía)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM Usuarios")
    if cursor.fetchone()[0] == 0:
        usuarios_defecto = [
            ('admin', generate_password_hash('admin123'), 'Administrador Principal', 'Admin'),
            ('tecnico1', generate_password_hash('tec123'), 'Juan Pérez (Téc. Mecánico)', 'Tecnico'),
            ('tecnico2', generate_password_hash('tec123'), 'Carlos Gomez (Téc. Eléctrico)', 'Tecnico'),
            ('tecnico3', generate_password_hash('tec123'), 'Ana Silva (Téc. Instrumentación)', 'Tecnico')
        ]
        conn.executemany("INSERT INTO Usuarios (username, password_hash, nombre_completo, rol) VALUES (?, ?, ?, ?)", usuarios_defecto)

    # 3. Tablas de Operación Normal
    conn.execute('''
        CREATE TABLE IF NOT EXISTS Maquinas (
            id_maquina INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_equipo TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            area_planta TEXT NOT NULL,
            criticidad TEXT NOT NULL,
            estado TEXT DEFAULT 'Operativa',
            fecha_instalacion TEXT DEFAULT (date('now', 'localtime'))
        )
    ''')

    conn.execute('''
        CREATE TABLE IF NOT EXISTS Repuestos_Stock (
            id_repuesto INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_pieza TEXT UNIQUE NOT NULL,
            nombre TEXT NOT NULL,
            descripcion TEXT,
            ubicacion_almacen TEXT,
            cantidad_actual REAL DEFAULT 0,
            punto_reorden REAL DEFAULT 0,
            unidad_medida TEXT DEFAULT 'Unidad'
        )
    ''')
    
    conn.execute('''
        CREATE TABLE IF NOT EXISTS Calendario_Mantenimiento (
            id_mantenimiento INTEGER PRIMARY KEY AUTOINCREMENT,
            id_maquina INTEGER NOT NULL,
            tipo_mantenimiento TEXT NOT NULL,
            descripcion_tarea TEXT NOT NULL,
            fecha_programada TEXT NOT NULL,
            fecha_ejecucion TEXT,
            estado_orden TEXT DEFAULT 'Pendiente',
            tecnico_asignado TEXT,
            observaciones TEXT DEFAULT 'Sin observaciones registradas.',
            recomendaciones TEXT DEFAULT 'Ninguna.',
            trabajos_realizados TEXT DEFAULT 'No especificado.',
            equipos_necesarios TEXT DEFAULT 'Ninguno.',
            tiempo_ejecucion TEXT DEFAULT '0 h',
            FOREIGN KEY (id_maquina) REFERENCES Maquinas (id_maquina)
        )
    ''')

    conn.execute('''
        CREATE TABLE IF NOT EXISTS Repuestos_Orden (
            id_registro INTEGER PRIMARY KEY AUTOINCREMENT,
            id_mantenimiento INTEGER NOT NULL,
            id_repuesto INTEGER NOT NULL,
            cantidad_usada REAL NOT NULL,
            costo_unitario_historico REAL,
            FOREIGN KEY (id_mantenimiento) REFERENCES Calendario_Mantenimiento (id_mantenimiento) ON DELETE CASCADE,
            FOREIGN KEY (id_repuesto) REFERENCES Repuestos_Stock (id_repuesto)
        )
    ''')

    try:
        conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN observaciones TEXT DEFAULT 'Sin observaciones registradas.'")
        conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN recomendaciones TEXT DEFAULT 'Ninguna.'")
        conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN trabajos_realizados TEXT DEFAULT 'No especificado.'")
        conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN equipos_necesarios TEXT DEFAULT 'Ninguno.'")
        conn.execute("ALTER TABLE Calendario_Mantenimiento ADD COLUMN tiempo_ejecucion TEXT DEFAULT '0 h'")
    except sqlite3.OperationalError:
        pass 
        
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS Evidencia_Fotografica (
                id_evidencia INTEGER PRIMARY KEY AUTOINCREMENT,
                id_mantenimiento INTEGER NOT NULL,
                imagen_base64 TEXT NOT NULL,
                FOREIGN KEY (id_mantenimiento) REFERENCES Calendario_Mantenimiento (id_mantenimiento) ON DELETE CASCADE
            )
        """)
    except sqlite3.OperationalError:
        pass
        
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS Compras_Repuestos (
                id_compra INTEGER PRIMARY KEY AUTOINCREMENT,
                id_repuesto INTEGER NOT NULL,
                cantidad REAL NOT NULL,
                costo_unitario REAL NOT NULL,
                proveedor TEXT,
                factura TEXT,
                fecha_compra TEXT DEFAULT (date('now', 'localtime')),
                FOREIGN KEY (id_repuesto) REFERENCES Repuestos_Stock (id_repuesto) ON DELETE CASCADE
            )
        """)
    except sqlite3.OperationalError:
        pass
        
    conn.commit()
    conn.close()

def obtener_todas_las_maquinas():
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    
    conn.execute(f"""
        UPDATE Maquinas 
        SET estado = 'En Mantenimiento' 
        WHERE id_maquina IN (
            SELECT id_maquina FROM Calendario_Mantenimiento 
            WHERE estado_orden = 'Pendiente' AND fecha_programada <= '{hoy}'
        )
    """)
    
    conn.execute(f"""
        UPDATE Maquinas 
        SET estado = 'Operativa' 
        WHERE id_maquina NOT IN (
            SELECT id_maquina FROM Calendario_Mantenimiento 
            WHERE estado_orden = 'Pendiente' AND fecha_programada <= '{hoy}'
        ) AND estado = 'En Mantenimiento'
    """)
    conn.commit()

    maquinas = conn.execute('SELECT * FROM Maquinas').fetchall()
    conn.close()
    return [dict(ix) for ix in maquinas]

def obtener_mantenimientos_pendientes():
    conn = get_db_connection()
    query = """
        SELECT c.id_mantenimiento, m.codigo_equipo, m.nombre AS maquina_nombre, 
               c.tipo_mantenimiento, c.descripcion_tarea, c.fecha_programada, 
               c.estado_orden, c.tecnico_asignado, c.id_maquina
        FROM Calendario_Mantenimiento c
        JOIN Maquinas m ON c.id_maquina = m.id_maquina
        WHERE c.estado_orden IN ('Pendiente', 'En Progreso')
        ORDER BY c.fecha_programada ASC
    """
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
        conn.execute('''
            INSERT INTO Compras_Repuestos (id_repuesto, cantidad, costo_unitario, proveedor, factura)
            VALUES (?, ?, ?, ?, ?)
        ''', (id_repuesto, cantidad, costo, proveedor, factura))
        
        conn.execute('''
            UPDATE Repuestos_Stock 
            SET cantidad_actual = cantidad_actual + ? 
            WHERE id_repuesto = ?
        ''', (cantidad, id_repuesto))
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"Error en BD al registrar compra: {e}")
    finally:
        conn.close()

def obtener_historial_compras_db():
    conn = get_db_connection()
    query = """
        SELECT c.id_compra, r.codigo_pieza, r.nombre, c.cantidad, c.costo_unitario,
               (c.cantidad * c.costo_unitario) as costo_total,
               c.proveedor, c.factura, c.fecha_compra
        FROM Compras_Repuestos c
        JOIN Repuestos_Stock r ON c.id_repuesto = r.id_repuesto
        ORDER BY c.fecha_compra DESC, c.id_compra DESC
    """
    compras = conn.execute(query).fetchall()
    conn.close()
    return [dict(ix) for ix in compras]

def crear_nueva_orden_db(id_maquina, tipo, descripcion, fecha, tecnico):
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO Calendario_Mantenimiento 
        (id_maquina, tipo_mantenimiento, descripcion_tarea, fecha_programada, tecnico_asignado, estado_orden)
        VALUES (?, ?, ?, ?, ?, 'Pendiente')
    ''', (id_maquina, tipo, descripcion, fecha, tecnico))
    
    hoy = datetime.now().strftime('%Y-%m-%d')
    if fecha <= hoy:
        conn.execute("UPDATE Maquinas SET estado = 'En Mantenimiento' WHERE id_maquina = ?", (id_maquina,))
        
    conn.commit()
    conn.close()

def actualizar_orden_db(id_orden, id_maquina, tipo, descripcion, fecha, tecnico):
    conn = get_db_connection()
    conn.execute('''
        UPDATE Calendario_Mantenimiento 
        SET id_maquina = ?, tipo_mantenimiento = ?, descripcion_tarea = ?, fecha_programada = ?, tecnico_asignado = ?
        WHERE id_mantenimiento = ?
    ''', (id_maquina, tipo, descripcion, fecha, tecnico, id_orden))
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
        conn.execute('''
            UPDATE Calendario_Mantenimiento 
            SET estado_orden = 'Completada', 
                fecha_ejecucion = ?,
                observaciones = ?,
                recomendaciones = ?,
                trabajos_realizados = ?,
                equipos_necesarios = ?,
                tiempo_ejecucion = ?
            WHERE id_mantenimiento = ?
        ''', (hoy, obs, rec, trabajos, equipos, tiempo, id_mantenimiento))
        
        orden = conn.execute('SELECT id_maquina FROM Calendario_Mantenimiento WHERE id_mantenimiento = ?', (id_mantenimiento,)).fetchone()
        if orden:
            conn.execute("UPDATE Maquinas SET estado = 'Operativa' WHERE id_maquina = ?", (orden['id_maquina'],))
            
        for rep in repuestos_usados:
            id_rep = int(rep['id_repuesto'])
            cantidad = float(rep['cantidad'])
            
            conn.execute('''
                INSERT INTO Repuestos_Orden (id_mantenimiento, id_repuesto, cantidad_usada, costo_unitario_historico)
                VALUES (?, ?, ?, 0.0)
            ''', (id_mantenimiento, id_rep, cantidad))
            
            conn.execute('''
                UPDATE Repuestos_Stock 
                SET cantidad_actual = cantidad_actual - ? 
                WHERE id_repuesto = ?
            ''', (cantidad, id_rep))
            
        for b64 in evidencias:
            conn.execute('''
                INSERT INTO Evidencia_Fotografica (id_mantenimiento, imagen_base64)
                VALUES (?, ?)
            ''', (id_mantenimiento, b64))
            
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"Error en BD al completar orden: {e}")
    finally:
        conn.close()

def obtener_historial_db():
    conn = get_db_connection()
    query = """
        SELECT c.id_mantenimiento, m.codigo_equipo, m.nombre AS maquina_nombre, m.area_planta,
               c.tipo_mantenimiento, c.descripcion_tarea, c.fecha_ejecucion, 
               c.tecnico_asignado, c.tiempo_ejecucion,
               (SELECT GROUP_CONCAT(ro.cantidad_usada || ' ' || rs.unidad_medida || ' de ' || rs.nombre, '\n')
                FROM Repuestos_Orden ro
                JOIN Repuestos_Stock rs ON ro.id_repuesto = rs.id_repuesto
                WHERE ro.id_mantenimiento = c.id_mantenimiento) as repuestos_usados
        FROM Calendario_Mantenimiento c
        JOIN Maquinas m ON c.id_maquina = m.id_maquina
        WHERE c.estado_orden = 'Completada'
        ORDER BY c.fecha_ejecucion DESC
    """
    historial = conn.execute(query).fetchall()
    conn.close()
    return [dict(ix) for ix in historial]

def crear_maquina_db(codigo, nombre, area, criticidad):
    conn = get_db_connection()
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn.execute('''
        INSERT INTO Maquinas (codigo_equipo, nombre, area_planta, criticidad, estado, fecha_instalacion)
        VALUES (?, ?, ?, ?, 'Operativa', ?)
    ''', (codigo, nombre, area, criticidad, hoy))
    conn.commit()
    conn.close()

# ==========================================
# 2. INTERFAZ DE USUARIO (HTML + TAILWIND CSS)
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

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cafetec - Sistema de Mantenimiento</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = { darkMode: 'class', theme: { extend: {} } }
    </script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Rajdhani:wght@600;700&display=swap" rel="stylesheet">
    
    <style>
        body { font-family: 'Inter', sans-serif; }
        .fuente-logo { font-family: 'Rajdhani', sans-serif; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .dark ::-webkit-scrollbar-thumb { background: #475569; }
        ::-webkit-scrollbar-thumb:hover { background: #22d3ee; }
        input:focus, select:focus, textarea:focus { outline: none; box-shadow: 0 0 0 2px rgba(34, 211, 238, 0.3); border-color: #22d3ee; }
        input[type="number"]::-webkit-inner-spin-button, input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        input[type="number"] { -moz-appearance: textfield; }
    </style>
</head>
<body class="bg-slate-50 dark:bg-slate-950 antialiased text-slate-800 dark:text-slate-200 relative transition-colors duration-300">

    <script>
        const CURRENT_USER = "{{ session.get('nombre_completo', 'Usuario') }}";
        const CURRENT_ROLE = "{{ session.get('rol', 'Tecnico') }}";
    </script>

    <div id="toast-notificacion" class="fixed top-5 right-5 bg-slate-900 dark:bg-slate-800 text-white px-6 py-3 rounded-lg shadow-2xl transform translate-x-[150%] transition-transform duration-300 z-[100] border-l-4 border-cyan-500 font-bold flex items-center">
        <i id="toast-icon" class="fa-solid fa-circle-info mr-3 text-cyan-400 text-lg"></i>
        <span id="toast-msg">Mensaje</span>
    </div>

    <!-- NAVBAR CAFETEC CORPORATIVO -->
    <nav class="bg-[#050f1a] text-white shadow-xl border-b-[3px] border-cyan-500 relative z-20">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-20">
                <div class="flex items-center select-none cursor-default fuente-logo mt-1.5">
                    <div class="border-[3px] border-white px-2 flex items-center justify-center leading-none mr-1.5 bg-transparent h-10">
                        <span class="text-cyan-400 text-4xl font-bold tracking-wide leading-none">CAFE</span>
                    </div>
                    <div class="leading-none mr-4 h-10 flex items-center">
                        <span class="text-white text-4xl font-bold tracking-wide leading-none">TEC</span>
                    </div>
                    <span class="ml-2 text-[10px] font-sans font-bold uppercase tracking-widest text-slate-400 border-l border-slate-700 pl-4 py-1 flex items-center h-8 mt-1 hidden sm:flex">
                        <i class="fa-solid fa-screwdriver-wrench text-cyan-400 mr-1.5"></i> CMMS
                    </span>
                </div>
                
                <div class="flex items-center space-x-4">
                    <a href="/monitor" target="_blank" class="hidden sm:flex text-slate-300 hover:text-cyan-400 transition-colors bg-slate-800 hover:bg-slate-700 px-3 py-2 rounded-lg items-center shadow-inner font-bold text-xs" title="Abrir Monitor de Planta">
                        <i class="fa-solid fa-display mr-2"></i> Monitor
                    </a>

                    <button onclick="toggleTheme()" class="text-slate-300 hover:text-cyan-400 transition-colors bg-slate-800 hover:bg-slate-700 p-2.5 rounded-full w-10 h-10 flex justify-center items-center shadow-inner" title="Cambiar Tema">
                        <i id="theme-icon" class="fa-solid fa-sun"></i>
                    </button>

                    <div class="text-sm font-medium bg-slate-900 pl-4 pr-2 py-1.5 rounded-full border border-slate-800 shadow-inner flex items-center text-slate-200">
                        <div class="flex flex-col text-right mr-3 leading-tight hidden sm:flex">
                            <span class="font-bold text-white text-[11px]">{{ session.get('nombre_completo', 'Usuario') }}</span>
                            <span class="text-[9px] text-cyan-400 uppercase tracking-widest font-bold">{{ session.get('rol', 'Técnico') }}</span>
                        </div>
                        <div class="w-8 h-8 rounded-full bg-cyan-600 flex justify-center items-center mr-3 shadow-md border border-cyan-400">
                            <i class="fa-solid fa-user-astronaut text-white"></i>
                        </div>
                        <div class="h-6 w-px bg-slate-700 mr-2"></div>
                        <a href="/logout" class="w-8 h-8 rounded-full bg-rose-600/20 text-rose-500 hover:bg-rose-500 hover:text-white flex justify-center items-center transition-colors" title="Cerrar Sesión">
                            <i class="fa-solid fa-right-from-bracket"></i>
                        </a>
                    </div>
                </div>
            </div>
        </div>
    </nav>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        <!-- Tarjetas de Resumen (KPIs) -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 p-5 relative overflow-hidden group hover:shadow-md transition-all">
                <div class="absolute left-0 top-0 bottom-0 w-1 bg-cyan-500"></div>
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider mb-0.5">Equipos Registrados</p>
                        <p class="text-2xl font-black text-slate-800 dark:text-slate-100 leading-none" id="kpi-equipos">0</p>
                    </div>
                    <div class="w-10 h-10 rounded-full bg-sky-50 dark:bg-sky-900/30 border border-sky-100 dark:border-sky-800 flex items-center justify-center text-cyan-500 group-hover:bg-cyan-500 group-hover:text-white transition-colors">
                        <i class="fa-solid fa-industry text-sm"></i>
                    </div>
                </div>
            </div>
            
            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 p-5 relative overflow-hidden group hover:shadow-md transition-all">
                <div class="absolute left-0 top-0 bottom-0 w-1 bg-rose-500"></div>
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider mb-0.5">Órdenes Pendientes</p>
                        <p class="text-2xl font-black text-slate-800 dark:text-slate-100 leading-none" id="kpi-ordenes">0</p>
                    </div>
                    <div class="w-10 h-10 rounded-full bg-rose-50 dark:bg-rose-900/30 border border-rose-100 dark:border-rose-800 flex items-center justify-center text-rose-500 group-hover:bg-rose-500 group-hover:text-white transition-colors">
                        <i class="fa-solid fa-triangle-exclamation text-sm"></i>
                    </div>
                </div>
            </div>

            <div id="kpi-planta-card" class="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 p-5 relative overflow-hidden group hover:shadow-md transition-all">
                <div id="kpi-planta-border" class="absolute left-0 top-0 bottom-0 w-1 bg-emerald-500 transition-colors duration-300"></div>
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider mb-0.5">Estado Planta</p>
                        <p id="kpi-planta-text" class="text-2xl font-black text-slate-800 dark:text-slate-100 leading-none">Operativa</p>
                    </div>
                    <div id="kpi-planta-icon-bg" class="w-10 h-10 rounded-full bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-100 dark:border-emerald-800 flex items-center justify-center text-emerald-600 transition-colors duration-300 group-hover:bg-emerald-500 group-hover:text-white">
                        <i id="kpi-planta-icon" class="fa-solid fa-check-double text-sm"></i>
                    </div>
                </div>
            </div>
        </div>

        <!-- Gráficos de Estado -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 p-5">
                <h3 class="text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-3 flex items-center">
                    <i class="fa-solid fa-chart-pie text-cyan-600 mr-2"></i> Salud de Equipos
                </h3>
                <div class="relative h-48 w-full flex justify-center">
                    <canvas id="graficoMaquinas"></canvas>
                </div>
            </div>

            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 p-5">
                <h3 class="text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-3 flex items-center">
                    <i class="fa-solid fa-chart-bar text-cyan-600 mr-2"></i> Carga de Trabajo Activa
                </h3>
                <div class="relative h-48 w-full flex justify-center">
                    <canvas id="graficoOrdenes"></canvas>
                </div>
            </div>
        </div>

        <!-- TABLAS PRINCIPALES -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
            <!-- 1. Tabla de Órdenes -->
            <div class="bg-white dark:bg-slate-900 shadow-sm rounded-xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col h-auto">
                <div class="px-5 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50 flex justify-between items-center">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center">
                        <i class="fa-solid fa-wrench text-cyan-500 mr-2"></i> Órdenes de Trabajo
                    </h3>
                    <div class="flex space-x-2">
                        <button onclick="abrirModalHistorial()" class="bg-slate-900 dark:bg-slate-700 hover:bg-slate-800 dark:hover:bg-slate-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm flex items-center">
                            <i class="fa-solid fa-clock-rotate-left mr-1.5"></i> Historial
                        </button>
                        <button id="btn-add-orden" onclick="abrirModalOrden()" class="bg-cyan-500 hover:bg-cyan-400 text-slate-900 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm">
                            <i class="fa-solid fa-plus mr-1"></i> Nueva Orden
                        </button>
                    </div>
                </div>
                <div class="overflow-y-auto max-h-80 flex-1 bg-white dark:bg-slate-900">
                    <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700">
                        <thead class="bg-slate-900 sticky top-0 z-10 shadow-sm border-b-2 border-cyan-500">
                            <tr>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Equipo</th>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Tarea</th>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Estado</th>
                            </tr>
                        </thead>
                        <tbody id="tabla-ordenes" class="divide-y divide-slate-100 dark:divide-slate-800"></tbody>
                    </table>
                </div>
            </div>

            <!-- 2. Tabla de Máquinas -->
            <div class="bg-white dark:bg-slate-900 shadow-sm rounded-xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col h-auto">
                <div class="px-5 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50 flex justify-between items-center">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center">
                        <i class="fa-solid fa-gears text-cyan-500 mr-2"></i> Estado de Equipos
                    </h3>
                    <div class="flex items-center space-x-2">
                        <div class="relative">
                            <input type="text" id="buscador-maquinas" onkeyup="filtrarMaquinas()" placeholder="Buscar equipo..." class="pl-8 pr-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded-lg text-xs focus:ring-cyan-500 focus:border-cyan-500 shadow-sm w-40 bg-white dark:bg-slate-800 dark:text-white transition-shadow">
                            <i class="fa-solid fa-magnifying-glass absolute left-2.5 top-2 text-slate-400 text-xs"></i>
                        </div>
                        <button id="btn-add-maquina" onclick="abrirModalMaquina()" class="bg-slate-900 dark:bg-slate-700 hover:bg-slate-800 dark:hover:bg-slate-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm">
                            <i class="fa-solid fa-plus mr-1"></i> Agregar
                        </button>
                    </div>
                </div>
                <div class="overflow-y-auto max-h-80 flex-1 bg-white dark:bg-slate-900">
                    <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700">
                        <thead class="bg-slate-900 sticky top-0 z-10 shadow-sm border-b-2 border-cyan-500">
                            <tr>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Código</th>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Nombre</th>
                                <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Estado</th>
                            </tr>
                        </thead>
                        <tbody id="tabla-maquinas" class="divide-y divide-slate-100 dark:divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- 3. Inventario y Almacén -->
        <div class="bg-white dark:bg-slate-900 shadow-sm rounded-xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col mb-8">
            <div class="px-5 py-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50 flex justify-between items-center">
                <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center">
                    <i class="fa-solid fa-boxes-stacked text-cyan-500 mr-2"></i> Inventario de Repuestos y Almacén
                </h3>
                <div class="flex items-center space-x-2">
                    <div class="relative">
                        <input type="text" id="buscador-inventario" onkeyup="filtrarInventario()" placeholder="Buscar código o pieza..." class="pl-8 pr-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded-lg text-xs focus:ring-cyan-500 focus:border-cyan-500 shadow-sm w-56 bg-white dark:bg-slate-800 dark:text-white transition-shadow">
                        <i class="fa-solid fa-magnifying-glass absolute left-2.5 top-2 text-slate-400 text-xs"></i>
                    </div>
                    <button id="btn-historial-compras" onclick="abrirModalCompras()" class="bg-slate-900 dark:bg-slate-700 hover:bg-slate-800 dark:hover:bg-slate-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm flex items-center">
                        <i class="fa-solid fa-receipt mr-1.5"></i> Compras
                    </button>
                    <!-- EXPORTAR A EXCEL BOTÓN -->
                    <button onclick="descargarInventarioExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-sm flex items-center">
                        <i class="fa-solid fa-file-excel mr-1.5"></i> Exportar
                    </button>
                </div>
            </div>
            
            <div class="overflow-y-auto max-h-[400px] flex-1 bg-white dark:bg-slate-900">
                <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700">
                    <thead class="bg-slate-900 sticky top-0 z-10 shadow-sm border-b-2 border-cyan-500">
                        <tr>
                            <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Código / Pieza</th>
                            <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Ubicación</th>
                            <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Stock Actual</th>
                            <th class="px-5 py-3 text-left text-[10px] font-bold text-cyan-400 uppercase tracking-wider">Alerta</th>
                        </tr>
                    </thead>
                    <tbody id="tabla-inventario" class="divide-y divide-slate-100 dark:divide-slate-800"></tbody>
                </table>
            </div>
        </div>

        <!-- MODAL: Programar/Editar Orden -->
        <div id="modal-orden" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-center backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-slate-900 p-8 rounded-xl shadow-2xl w-full max-w-md relative border border-slate-700 modal-content" onclick="event.stopPropagation()">
                <h2 id="modal-orden-title" class="text-xl font-bold mb-5 text-white flex items-center">
                    <i class="fa-solid fa-clipboard-list text-cyan-400 mr-2"></i> Programar Orden
                </h2>
                <form id="form-nueva-orden" onsubmit="guardarOrden(event)">
                    <input type="hidden" id="orden-id">
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Máquina</label>
                        <select id="select-maquina" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white truncate"></select>
                    </div>
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Tipo</label>
                        <select id="select-tipo" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                            <option value="Preventivo">Preventivo</option>
                            <option value="Correctivo">Correctivo</option>
                            <option value="Predictivo">Predictivo</option>
                        </select>
                    </div>
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Descripción</label>
                        <input type="text" id="input-descripcion" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                    </div>
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Fecha</label>
                        <input type="date" id="input-fecha" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white [color-scheme:dark]">
                    </div>
                    <div class="mb-6">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Técnico</label>
                        <input type="text" id="input-tecnico" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                    </div>
                    <div class="flex justify-end space-x-3 pt-4 border-t border-slate-700">
                        <button type="button" onclick="cerrarModalOrden()" class="px-5 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-semibold transition-colors">Cancelar</button>
                        <button type="submit" class="px-5 py-2 bg-cyan-500 text-slate-900 rounded-lg hover:bg-cyan-400 font-bold transition-all">Guardar Orden</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: Registrar Máquina -->
        <div id="modal-maquina" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-center backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-slate-900 p-8 rounded-xl shadow-2xl w-full max-w-md relative border border-slate-700 modal-content" onclick="event.stopPropagation()">
                <h2 class="text-xl font-bold mb-5 text-white flex items-center">
                    <i class="fa-solid fa-server text-cyan-400 mr-2"></i> Registrar Equipo
                </h2>
                <form id="form-nueva-maquina" onsubmit="guardarMaquina(event)">
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Código</label>
                        <input type="text" id="input-maq-codigo" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                    </div>
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Nombre y Modelo</label>
                        <input type="text" id="input-maq-nombre" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                    </div>
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Área de la Planta</label>
                        <select id="select-maq-area" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                            <option value="Recepcion">Recepción</option>
                            <option value="Tostado">Tostado</option>
                            <option value="Molienda">Molienda</option>
                            <option value="Empaque">Empaque</option>
                            <option value="General">Servicios Generales</option>
                        </select>
                    </div>
                    <div class="mb-6">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Criticidad</label>
                        <select id="select-maq-criticidad" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white">
                            <option value="Alta">Alta</option>
                            <option value="Media">Media</option>
                            <option value="Baja">Baja</option>
                        </select>
                    </div>
                    <div class="flex justify-end space-x-3 pt-4 border-t border-slate-700">
                        <button type="button" onclick="cerrarModalMaquina()" class="px-5 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-semibold transition-colors">Cancelar</button>
                        <button type="submit" class="px-5 py-2 bg-cyan-500 text-slate-900 rounded-lg hover:bg-cyan-400 font-bold transition-all">Guardar Equipo</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: Ingreso de Mercadería (Compras) -->
        <div id="modal-compra" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-center backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-slate-900 p-8 rounded-xl shadow-2xl w-full max-w-md relative border border-slate-700 modal-content" onclick="event.stopPropagation()">
                <h2 class="text-xl font-bold mb-2 text-white flex items-center">
                    <i class="fa-solid fa-truck-ramp-box text-cyan-400 mr-2"></i> Ingreso de Mercadería
                </h2>
                <p id="compra-subtitulo" class="text-sm text-slate-400 mb-5 pb-3 border-b border-slate-800 font-medium">Repuesto Seleccionado</p>
                
                <form id="form-registro-compra" onsubmit="guardarCompra(event)">
                    <input type="hidden" id="compra-id-repuesto">
                    
                    <div class="grid grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Cantidad Recibida</label>
                            <input type="number" id="compra-cantidad" required min="0.1" step="0.1" class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white" placeholder="Ej: 10">
                        </div>
                        <div>
                            <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Costo Unitario ($)</label>
                            <input type="number" id="compra-costo" required min="0" step="0.01" class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white" placeholder="Ej: 15.50">
                        </div>
                    </div>
                    
                    <div class="mb-4">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">Proveedor / Tienda</label>
                        <input type="text" id="compra-proveedor" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white" placeholder="Ej: Ferretería Industrial">
                    </div>
                    
                    <div class="mb-6">
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-1">N° Factura / Boleta / OC</label>
                        <input type="text" id="compra-factura" required class="w-full border border-slate-700 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-800 text-white" placeholder="Ej: F001-4589">
                    </div>

                    <div class="flex justify-end space-x-3 pt-4 border-t border-slate-700">
                        <button type="button" onclick="cerrarModalCompra()" class="px-5 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-semibold transition-colors">Cancelar</button>
                        <button type="submit" class="px-5 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-500 font-bold transition-all flex items-center">
                            <i class="fa-solid fa-check mr-2"></i> Registrar Stock
                        </button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: Completar Orden y Generar Reporte -->
        <div id="modal-completar" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-start pt-10 pb-10 backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-slate-900 p-8 rounded-xl shadow-2xl w-full max-w-2xl relative border border-slate-700 my-auto modal-content" onclick="event.stopPropagation()">
                <h2 class="text-xl font-bold mb-5 text-white flex items-center border-b border-slate-800 pb-3">
                    <i class="fa-solid fa-file-signature text-cyan-400 mr-2"></i> Generar Reporte Técnico
                </h2>
                
                <input type="hidden" id="completar-id-orden">
                
                <div class="mb-4">
                    <label class="block text-xs font-bold text-slate-400 uppercase mb-2">Trabajos Realizados (Descripción Detallada)</label>
                    <textarea id="input-trabajos" rows="2" class="w-full border border-slate-700 rounded-lg p-3 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-950 text-white placeholder-slate-600 resize-none" placeholder="Describe paso a paso lo que se hizo durante el mantenimiento..."></textarea>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-2">Equipos/Herramientas Necesarios</label>
                        <input type="text" id="input-equipos" class="w-full border border-slate-700 rounded-lg p-3 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-950 text-white placeholder-slate-600" placeholder="Ej: Multímetro, hidrolavadora, andamios..">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-2">Tiempo de Ejecución</label>
                        <input type="text" id="input-tiempo" class="w-full border border-slate-700 rounded-lg p-3 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-950 text-white placeholder-slate-600" placeholder="Ej: 2 horas, 45 min...">
                    </div>
                </div>
                
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4 pt-4 border-t border-slate-800">
                    <div>
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-2">Observaciones Encontradas</label>
                        <textarea id="input-obs" rows="2" class="w-full border border-slate-700 rounded-lg p-3 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-950 text-white placeholder-slate-600 resize-none" placeholder="¿Qué fallas o detalles encontraste?"></textarea>
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-400 uppercase mb-2">Recomendaciones Futuras</label>
                        <textarea id="input-rec" rows="2" class="w-full border border-slate-700 rounded-lg p-3 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-950 text-white placeholder-slate-600 resize-none" placeholder="Ej: Cambiar filtro el próximo mes..."></textarea>
                    </div>
                </div>

                <div class="mb-4 bg-slate-800/50 p-4 rounded-lg border border-slate-700">
                    <label class="block text-xs font-bold text-cyan-400 uppercase mb-2"><i class="fa-solid fa-box-open mr-1"></i> Repuestos Utilizados</label>
                    <div class="flex space-x-2">
                        <select id="select-uso-repuesto" class="flex-1 min-w-0 flex-shrink-0 border border-slate-700 rounded-lg p-2 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-900 text-white truncate"></select>
                        <input type="number" id="input-uso-cant" value="1" min="0.1" step="0.1" class="w-20 border border-slate-700 rounded-lg p-2 text-sm focus:ring-2 focus:ring-cyan-500 bg-slate-900 text-white" placeholder="Cant.">
                        <button type="button" onclick="agregarRepuestoALista()" class="bg-cyan-600 hover:bg-cyan-500 text-white px-4 rounded-lg font-bold transition-colors">
                            <i class="fa-solid fa-plus"></i>
                        </button>
                    </div>
                </div>

                <div class="mb-6">
                    <div class="bg-slate-950 border border-slate-800 rounded-lg max-h-32 overflow-y-auto p-2">
                        <ul id="lista-repuestos-usados" class="divide-y divide-slate-800">
                            <li class="py-2 text-sm text-slate-500 italic text-center">No se han agregado repuestos.</li>
                        </ul>
                    </div>
                </div>

                <div class="mb-4 pt-4 border-t border-slate-800">
                    <label class="block text-xs font-bold text-slate-400 uppercase mb-2"><i class="fa-solid fa-camera mr-1"></i> Evidencia Fotográfica</label>
                    <input type="file" id="input-fotos" multiple accept="image/*" class="w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-bold file:bg-cyan-900/50 file:text-cyan-400 hover:file:bg-cyan-900 transition-colors bg-slate-950 border border-slate-800 rounded-lg cursor-pointer" onchange="procesarImagenes(event)">
                    <div id="preview-fotos" class="grid grid-cols-2 md:grid-cols-4 gap-3 mt-3"></div>
                </div>

                <div class="flex justify-end space-x-3 pt-4 border-t border-slate-800">
                    <button type="button" onclick="cerrarModalCompletar()" class="px-5 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-semibold transition-colors">Cancelar</button>
                    <button type="button" onclick="confirmarCompletarOrden()" class="px-5 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-500 font-bold transition-all flex items-center">
                        <i class="fa-solid fa-flag-checkered mr-2"></i> Guardar Reporte y Finalizar
                    </button>
                </div>
            </div>
        </div>

        <!-- MODAL: Historial de Auditoría (PDF) -->
        <div id="modal-historial" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-center backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-2xl w-full max-w-6xl relative border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col max-h-[85vh] modal-content" onclick="event.stopPropagation()">
                <div class="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-900 flex justify-between items-center">
                    <h2 class="text-xl font-bold text-white flex items-center">
                        <i class="fa-solid fa-clipboard-check text-cyan-400 mr-2"></i> Historial de Auditoría
                    </h2>
                    <button onclick="cerrarModalHistorial()" class="text-slate-400 hover:text-white transition-colors">
                        <i class="fa-solid fa-xmark text-2xl"></i>
                    </button>
                </div>
                
                <div class="overflow-y-auto flex-1 bg-slate-50 dark:bg-slate-900 p-4">
                    <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700 border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden">
                        <thead class="bg-slate-900 sticky top-0 z-10 border-b-2 border-cyan-500">
                            <tr>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Fecha</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Equipo</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Tipo</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Tiempo Ejecución</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Técnico</th>
                                <th class="px-4 py-3 text-center text-xs font-bold text-cyan-400 uppercase tracking-wider">Reporte</th>
                            </tr>
                        </thead>
                        <tbody id="tabla-historial" class="bg-white dark:bg-slate-900 divide-y divide-slate-100 dark:divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- MODAL: Libro Mayor de Compras -->
        <div id="modal-historial-compras" class="hidden fixed inset-0 bg-slate-950/80 overflow-y-auto h-full w-full z-50 flex justify-center items-center backdrop-blur-sm transition-opacity modal-container">
            <div class="bg-white dark:bg-slate-900 rounded-xl shadow-2xl w-full max-w-6xl relative border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col max-h-[85vh] modal-content" onclick="event.stopPropagation()">
                <div class="px-6 py-4 border-b border-slate-200 dark:border-slate-700 bg-slate-900 flex justify-between items-center">
                    <h2 class="text-xl font-bold text-white flex items-center">
                        <i class="fa-solid fa-file-invoice-dollar text-cyan-400 mr-2"></i> Registro de Ingresos y Compras
                    </h2>
                    <div class="flex items-center space-x-4">
                        <button onclick="descargarComprasExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-lg text-sm font-bold shadow-sm flex items-center transition-colors">
                            <i class="fa-solid fa-file-excel mr-2"></i> Exportar a Excel
                        </button>
                        <button onclick="cerrarModalCompras()" class="text-slate-400 hover:text-white transition-colors">
                            <i class="fa-solid fa-xmark text-2xl"></i>
                        </button>
                    </div>
                </div>
                
                <div class="overflow-y-auto flex-1 bg-slate-50 dark:bg-slate-900 p-4">
                    <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700 border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden">
                        <thead class="bg-slate-900 sticky top-0 z-10 border-b-2 border-cyan-500">
                            <tr>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Fecha</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Repuesto</th>
                                <th class="px-4 py-3 text-center text-xs font-bold text-cyan-400 uppercase tracking-wider">Cantidad</th>
                                <th class="px-4 py-3 text-right text-xs font-bold text-cyan-400 uppercase tracking-wider">Costo Unit.</th>
                                <th class="px-4 py-3 text-right text-xs font-bold text-cyan-400 uppercase tracking-wider">Total</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider pl-6">Proveedor</th>
                                <th class="px-4 py-3 text-left text-xs font-bold text-cyan-400 uppercase tracking-wider">Factura/OC</th>
                            </tr>
                        </thead>
                        <tbody id="tabla-historial-compras" class="bg-white dark:bg-slate-900 divide-y divide-slate-100 dark:divide-slate-800"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- BOTÓN FLOTANTE: Alertas -->
        <button id="btn-alertas-flotante" onclick="abrirModalAlertas(event)" class="hidden fixed bottom-8 right-8 z-50 bg-rose-600 hover:bg-rose-500 text-white rounded-full p-4 shadow-2xl transition-transform hover:scale-110 group">
            <span class="absolute top-0 right-0 -mt-1 -mr-1 flex h-4 w-4">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-4 w-4 bg-rose-500 border-2 border-white"></span>
            </span>
            <i class="fa-solid fa-bell text-2xl"></i>
            <span class="absolute bottom-full right-0 mb-2 w-max px-3 py-1 bg-slate-900 text-white text-xs font-bold rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none">
                Alertas de Stock
            </span>
        </button>

        <!-- PANEL FLOTANTE: Repuestos a Reabastecer -->
        <div id="modal-alertas-stock" class="hidden fixed bottom-24 right-8 bg-[#0f172a] p-5 rounded-xl shadow-2xl w-96 border border-slate-700 z-50 modal-content" onclick="event.stopPropagation()">
            <div class="flex justify-between items-center mb-4 border-b border-slate-800 pb-3">
                <h2 class="text-base font-bold text-white flex items-center">
                    <i class="fa-solid fa-triangle-exclamation text-rose-500 mr-2"></i> A Reabastecer
                </h2>
                <button onclick="cerrarModalAlertas()" class="text-slate-400 hover:text-white transition-colors">
                    <i class="fa-solid fa-xmark text-xl"></i>
                </button>
            </div>
            <div class="overflow-y-auto max-h-64 pr-2">
                <ul id="lista-alertas-stock" class="space-y-2"></ul>
            </div>
        </div>

    </main>

    <script>
        // --- LOGICA DE TEMA (DARK/LIGHT) ---
        function toggleTheme() {
            const html = document.documentElement;
            const isDark = html.classList.contains('dark');
            const icon = document.getElementById('theme-icon');
            
            if (isDark) {
                html.classList.remove('dark');
                localStorage.setItem('theme', 'light');
                icon.className = 'fa-solid fa-moon';
                actualizarGraficosTema(false);
            } else {
                html.classList.add('dark');
                localStorage.setItem('theme', 'dark');
                icon.className = 'fa-solid fa-sun text-amber-400';
                actualizarGraficosTema(true);
            }
        }

        function applySavedTheme() {
            const html = document.documentElement;
            const icon = document.getElementById('theme-icon');
            if (localStorage.theme === 'dark' || (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
                html.classList.add('dark');
                if(icon) icon.className = 'fa-solid fa-sun text-amber-400';
            } else {
                html.classList.remove('dark');
                if(icon) icon.className = 'fa-solid fa-moon';
            }
        }
        applySavedTheme();

        // --- SISTEMA DE NOTIFICACIONES ---
        function mostrarNotificacion(mensaje, tipo='info') {
            const toast = document.getElementById('toast-notificacion');
            const icon = document.getElementById('toast-icon');
            document.getElementById('toast-msg').innerText = mensaje;
            
            if(tipo === 'error') {
                toast.className = "fixed top-5 right-5 bg-rose-600 text-white px-6 py-3 rounded-lg shadow-2xl transform z-[100] font-bold flex items-center transition-transform duration-300";
                icon.className = "fa-solid fa-triangle-exclamation mr-3 text-white text-lg";
            } else {
                toast.className = "fixed top-5 right-5 bg-slate-900 dark:bg-slate-800 text-white px-6 py-3 rounded-lg shadow-2xl transform z-[100] border-l-4 border-cyan-500 font-bold flex items-center transition-transform duration-300";
                icon.className = "fa-solid fa-circle-check mr-3 text-cyan-400 text-lg";
            }

            toast.style.transform = "translateX(0)";
            setTimeout(() => { toast.style.transform = "translateX(150%)"; }, 3500);
        }

        function getBadge(status) {
            let color = 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700';
            if (status === 'Operativa' || status === 'Completada') color = 'bg-emerald-50 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-400 border-emerald-200 dark:border-emerald-800/50';
            if (status === 'En Mantenimiento' || status === 'En Progreso') color = 'bg-amber-50 dark:bg-amber-900/30 text-amber-700 dark:text-amber-400 border-amber-200 dark:border-amber-800/50';
            if (status === 'Fuera de Servicio' || status === 'Pendiente') color = 'bg-rose-50 dark:bg-rose-900/30 text-rose-700 dark:text-rose-400 border-rose-200 dark:border-rose-800/50';
            return `<span class="px-2.5 py-1 inline-flex text-[10px] leading-5 font-bold rounded-md border ${color} shadow-sm uppercase tracking-wide">${status}</span>`;
        }

        let chartMaquinas = null;
        let chartOrdenes = null;

        // Autorización en Interfaz (Ocultar botones si es Técnico)
        if (CURRENT_ROLE === 'Tecnico') {
            document.getElementById('btn-add-maquina')?.classList.add('hidden');
            document.getElementById('btn-historial-compras')?.classList.add('hidden');
            document.getElementById('btn-add-orden')?.classList.add('hidden');
        }

        // Cargar Máquinas
        fetch('/api/maquinas')
            .then(response => response.json())
            .then(data => {
                document.getElementById('kpi-equipos').innerText = data.length;
                let html = '';
                let maquinasDetenidas = 0;
                let maquinasOperativas = 0;
                
                data.forEach(maquina => {
                    if (maquina.estado !== 'Operativa') maquinasDetenidas++;
                    else maquinasOperativas++;

                    let btnImprimirQR = CURRENT_ROLE === 'Admin' ? `<a href="/api/maquinas/etiqueta/${maquina.id_maquina}" target="_blank" class="text-cyan-500 hover:text-cyan-400 transition-colors ml-2 opacity-0 group-hover:opacity-100" title="Imprimir QR"><i class="fa-solid fa-qrcode"></i></a>` : '';
                    let btnEliminar = CURRENT_ROLE === 'Admin' ? `<button onclick="eliminarMaquina(${maquina.id_maquina})" class="text-slate-300 dark:text-slate-600 hover:text-rose-600 dark:hover:text-rose-500 transition-colors ml-2 opacity-0 group-hover:opacity-100" title="Eliminar equipo"><i class="fa-solid fa-trash-can"></i></button>` : '';

                    html += `
                        <tr class="hover:bg-cyan-50 dark:hover:bg-slate-800/50 transition-colors group">
                            <td class="px-5 py-3 whitespace-nowrap text-[11px] font-bold text-slate-800 dark:text-slate-200">${maquina.codigo_equipo}</td>
                            <td class="px-5 py-3 whitespace-nowrap text-[11px] text-slate-600 dark:text-slate-400">
                                <div class="font-bold text-slate-800 dark:text-slate-200 uppercase">${maquina.nombre}</div>
                                <span class="text-[10px] text-slate-400 dark:text-slate-500 font-medium">Área: ${maquina.area_planta}</span>
                            </td>
                            <td class="px-5 py-3 whitespace-nowrap">
                                <div class="flex items-center justify-between">
                                    ${getBadge(maquina.estado)}
                                    <div class="flex items-center">
                                        ${btnImprimirQR}
                                        ${btnEliminar}
                                    </div>
                                </div>
                            </td>
                        </tr>
                    `;
                });
                document.getElementById('tabla-maquinas').innerHTML = html;

                const border = document.getElementById('kpi-planta-border');
                const iconBg = document.getElementById('kpi-planta-icon-bg');
                const icon = document.getElementById('kpi-planta-icon');
                const text = document.getElementById('kpi-planta-text');

                if (data.length === 0) {
                    text.innerText = "Sin Datos";
                } else if (maquinasDetenidas === 0) {
                    text.innerText = "Operativa";
                    border.className = "absolute left-0 top-0 bottom-0 w-1 bg-emerald-500 transition-colors duration-300";
                    iconBg.className = "w-10 h-10 rounded-full bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-100 dark:border-emerald-800 flex items-center justify-center text-emerald-600 dark:text-emerald-400 transition-colors duration-300 group-hover:bg-emerald-500 group-hover:text-white";
                    icon.className = "fa-solid fa-check-double text-sm";
                } else if (maquinasDetenidas > 0 && maquinasDetenidas < data.length) {
                    text.innerText = "Op. Parcial";
                    border.className = "absolute left-0 top-0 bottom-0 w-1 bg-amber-500 transition-colors duration-300";
                    iconBg.className = "w-10 h-10 rounded-full bg-amber-50 dark:bg-amber-900/30 border border-amber-100 dark:border-amber-800 flex items-center justify-center text-amber-600 dark:text-amber-400 transition-colors duration-300 group-hover:bg-amber-500 group-hover:text-white";
                    icon.className = "fa-solid fa-triangle-exclamation text-sm";
                } else {
                    text.innerText = "Parada General";
                    border.className = "absolute left-0 top-0 bottom-0 w-1 bg-rose-500 transition-colors duration-300";
                    iconBg.className = "w-10 h-10 rounded-full bg-rose-50 dark:bg-rose-900/30 border border-rose-100 dark:border-rose-800 flex items-center justify-center text-rose-600 dark:text-rose-400 transition-colors duration-300 group-hover:bg-rose-500 group-hover:text-white";
                    icon.className = "fa-solid fa-circle-xmark text-sm";
                }

                const isDark = document.documentElement.classList.contains('dark');
                const darkColor = isDark ? '#334155' : '#1e293b';
                
                const ctxMaquinas = document.getElementById('graficoMaquinas').getContext('2d');
                chartMaquinas = new Chart(ctxMaquinas, {
                    type: 'doughnut',
                    data: {
                        labels: ['Operativas', 'En Mantenimiento'],
                        datasets: [{ data: [maquinasOperativas, maquinasDetenidas], backgroundColor: ['#22d3ee', darkColor], borderWidth: 0 }]
                    },
                    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, cutout: '75%' }
                });
            });

        // Cargar Órdenes
        fetch('/api/ordenes')
            .then(response => response.json())
            .then(data => {
                document.getElementById('kpi-ordenes').innerText = data.length;
                let html = '';
                let prev = 0, corr = 0, pred = 0;

                if(data.length === 0) {
                    html = '<tr><td colspan="3" class="px-5 py-6 text-center text-xs text-slate-500 dark:text-slate-400 italic font-semibold">No hay órdenes pendientes. ¡Buen trabajo!</td></tr>';
                } else {
                    data.forEach(orden => {
                        if(orden.tipo_mantenimiento === 'Preventivo') prev++;
                        if(orden.tipo_mantenimiento === 'Correctivo') corr++;
                        if(orden.tipo_mantenimiento === 'Predictivo') pred++;

                        let btnsAdmin = '';
                        if (CURRENT_ROLE === 'Admin') {
                            btnsAdmin = `
                                <button onclick="editarOrden(${orden.id_mantenimiento}, ${orden.id_maquina}, '${orden.tipo_mantenimiento}', '${orden.descripcion_tarea}', '${orden.fecha_programada}', '${orden.tecnico_asignado}')" class="text-amber-500 hover:text-amber-400 transition-colors ml-2 opacity-0 group-hover:opacity-100" title="Editar Orden"><i class="fa-solid fa-pen-to-square"></i></button>
                                <button onclick="eliminarOrden(${orden.id_mantenimiento})" class="text-rose-600 hover:text-rose-500 transition-colors ml-2 opacity-0 group-hover:opacity-100" title="Eliminar Orden"><i class="fa-solid fa-trash-can"></i></button>
                            `;
                        }

                        html += `
                            <tr class="hover:bg-cyan-50 dark:hover:bg-slate-800/50 transition-colors group">
                                <td class="px-5 py-3 whitespace-nowrap text-[11px] font-bold text-slate-800 dark:text-slate-200">${orden.codigo_equipo}</td>
                                <td class="px-5 py-3 text-[11px] text-slate-600 dark:text-slate-400">
                                    <div class="font-bold text-slate-800 dark:text-slate-200">${orden.tipo_mantenimiento}</div>
                                    <div class="truncate w-40 text-slate-500 dark:text-slate-400" title="${orden.descripcion_tarea}">${orden.descripcion_tarea}</div>
                                    <div class="text-[10px] text-cyan-600 dark:text-cyan-400 font-semibold mt-0.5"><i class="fa-regular fa-calendar mr-1"></i> ${orden.fecha_programada}</div>
                                </td>
                                <td class="px-5 py-3 whitespace-nowrap">
                                    <div class="flex items-center justify-between">
                                        ${getBadge(orden.estado_orden)}
                                        <div class="flex items-center">
                                            ${btnsAdmin}
                                            <button onclick="abrirModalCompletar(${orden.id_mantenimiento})" class="text-slate-400 dark:text-slate-500 hover:text-cyan-500 dark:hover:text-cyan-400 transition-colors ml-3 opacity-0 group-hover:opacity-100" title="Marcar finalizada y generar reporte">
                                                <i class="fa-solid fa-file-signature text-xl"></i>
                                            </button>
                                        </div>
                                    </div>
                                </td>
                            </tr>
                        `;
                    });
                }
                document.getElementById('tabla-ordenes').innerHTML = html;

                const isDark = document.documentElement.classList.contains('dark');
                const darkGrid = isDark ? '#334155' : '#e2e8f0';
                const darkText = isDark ? '#94a3b8' : '#64748b';

                const ctxOrdenes = document.getElementById('graficoOrdenes').getContext('2d');
                chartOrdenes = new Chart(ctxOrdenes, {
                    type: 'bar',
                    data: {
                        labels: ['Preventivo', 'Correctivo', 'Predictivo'],
                        datasets: [{ label: 'Órdenes Activas', data: [prev, corr, pred], backgroundColor: ['#22d3ee', '#e11d48', '#64748b'], borderRadius: 3 }]
                    },
                    options: { 
                        responsive: true, maintainAspectRatio: false, 
                        plugins: { legend: { display: false } }, 
                        scales: { 
                            y: { beginAtZero: true, grid: { color: darkGrid }, ticks: { stepSize: 1, precision: 0, font: {size: 10}, color: darkText } }, 
                            x: { grid: { display: false }, ticks: {font: {size: 10}, color: darkText} } 
                        } 
                    }
                });
            });

        // Cargar Inventario
        fetch('/api/inventario')
            .then(response => response.json())
            .then(data => {
                let html = '';
                let htmlAlertas = '';
                let itemsCriticos = 0;

                data.forEach(item => {
                    let estadoStock = '<span class="px-2.5 py-1 inline-flex text-[10px] leading-5 font-bold rounded-md bg-emerald-50 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800/50 uppercase tracking-wide">Óptimo</span>';
                    let textClass = "text-slate-800 dark:text-slate-200";
                    
                    if (item.cantidad_actual <= item.punto_reorden) {
                        estadoStock = '<span class="px-2.5 py-1 inline-flex text-[10px] leading-5 font-bold rounded-md bg-rose-50 dark:bg-rose-900/30 text-rose-700 dark:text-rose-400 border border-rose-200 dark:border-rose-800/50 uppercase tracking-wide">Reabastecer</span>';
                        textClass = "text-rose-600 dark:text-rose-400 font-bold";
                        itemsCriticos++;
                        
                        htmlAlertas += `
                            <li class="p-3 bg-slate-900 border border-slate-800 rounded-lg flex justify-between items-center transition-colors shadow-sm">
                                <div>
                                    <span class="text-rose-400 font-bold block text-sm">${item.codigo_pieza}</span>
                                    <span class="text-slate-400 text-xs w-48 truncate block">${item.nombre}</span>
                                </div>
                                <div class="text-right flex flex-col items-end">
                                    <div class="flex items-baseline space-x-1">
                                        <span class="text-rose-500 text-lg font-bold">${item.cantidad_actual}</span> 
                                        <span class="text-slate-500 text-[10px] font-semibold">${item.unidad_medida}</span>
                                    </div>
                                    <div class="text-[9px] text-slate-500 mt-0.5 uppercase tracking-wide bg-slate-800 px-1.5 py-0.5 rounded">Mín: ${item.punto_reorden}</div>
                                </div>
                            </li>
                        `;
                    }

                    let botonesAccionHtml = '';
                    if(CURRENT_ROLE === 'Admin') {
                        botonesAccionHtml = `
                            <div class="flex items-center space-x-1">
                                <button onclick="descontarStock(${item.id_repuesto})" class="bg-white dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 w-6 h-6 rounded border border-slate-200 dark:border-slate-700 shadow-sm flex items-center justify-center transition-colors" title="Ajuste Rápido (-)"><i class="fa-solid fa-minus text-[10px]"></i></button>
                                <button onclick="sumarStock(${item.id_repuesto})" class="bg-white dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 w-6 h-6 rounded border border-slate-200 dark:border-slate-700 shadow-sm flex items-center justify-center transition-colors" title="Ajuste Rápido (+)"><i class="fa-solid fa-plus text-[10px]"></i></button>
                                <div class="w-px h-3 bg-slate-300 dark:bg-slate-700 mx-1"></div>
                                <button onclick="abrirModalCompra(${item.id_repuesto}, '${item.codigo_pieza}', '${item.nombre}')" class="bg-cyan-50 dark:bg-cyan-900/30 hover:bg-cyan-100 dark:hover:bg-cyan-900/60 text-cyan-600 dark:text-cyan-400 w-6 h-6 rounded border border-cyan-200 dark:border-cyan-800/50 shadow-sm flex items-center justify-center transition-colors" title="Registrar Ingreso (Compra)"><i class="fa-solid fa-file-invoice-dollar text-[10px]"></i></button>
                            </div>
                        `;
                    }

                    html += `
                        <tr class="hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
                            <td class="px-5 py-3 whitespace-nowrap">
                                <div class="text-[11px] font-bold text-slate-800 dark:text-slate-200">${item.codigo_pieza}</div>
                                <div class="text-[11px] text-slate-500 dark:text-slate-400">${item.nombre}</div>
                            </td>
                            <td class="px-5 py-3 whitespace-nowrap text-[11px] text-slate-600 dark:text-slate-400 font-medium">
                                <i class="fa-solid fa-location-dot mr-1 text-slate-400 dark:text-slate-500"></i> ${item.ubicacion_almacen}
                            </td>
                            <td class="px-5 py-3 whitespace-nowrap text-[11px]">
                                <div class="flex items-center space-x-3">
                                    <div class="w-16">
                                        <span class="${textClass} text-lg">${item.cantidad_actual}</span> <span class="text-slate-400 dark:text-slate-500 text-[10px] font-semibold">${item.unidad_medida}</span>
                                        <div class="text-[9px] text-slate-400 dark:text-slate-500 font-medium uppercase mt-0.5">Mínimo: ${item.punto_reorden}</div>
                                    </div>
                                    ${botonesAccionHtml}
                                </div>
                            </td>
                            <td class="px-5 py-3 whitespace-nowrap">${estadoStock}</td>
                        </tr>
                    `;
                });
                document.getElementById('tabla-inventario').innerHTML = html;
                
                const btnFlotante = document.getElementById('btn-alertas-flotante');
                const listaAlertas = document.getElementById('lista-alertas-stock');
                
                if (itemsCriticos > 0) {
                    btnFlotante.classList.remove('hidden');
                    listaAlertas.innerHTML = htmlAlertas;
                } else {
                    btnFlotante.classList.add('hidden');
                }
            });

        function actualizarGraficosTema(isDark) {
            const darkColor = isDark ? '#334155' : '#1e293b';
            const darkGrid = isDark ? '#334155' : '#e2e8f0';
            const darkText = isDark ? '#94a3b8' : '#64748b';

            if(chartMaquinas) {
                chartMaquinas.data.datasets[0].backgroundColor[1] = darkColor;
                chartMaquinas.update();
            }
            if(chartOrdenes) {
                chartOrdenes.options.scales.x.ticks.color = darkText;
                chartOrdenes.options.scales.y.ticks.color = darkText;
                chartOrdenes.options.scales.y.grid.color = darkGrid;
                chartOrdenes.update();
            }
        }

        function descontarStock(id) { fetch(`/api/inventario/descontar/${id}`, {method:'POST'}).then(r=>r.json()).then(d=>{if(d.status==='ok')window.location.reload();}); }
        function sumarStock(id) { fetch(`/api/inventario/sumar/${id}`, {method:'POST'}).then(r=>r.json()).then(d=>{if(d.status==='ok')window.location.reload();}); }
        function filtrarInventario() { const txt = document.getElementById('buscador-inventario').value.toLowerCase(); document.querySelectorAll('#tabla-inventario tr').forEach(row => { row.style.display = row.innerText.toLowerCase().includes(txt) ? '' : 'none'; }); }
        function filtrarMaquinas() { const txt = document.getElementById('buscador-maquinas').value.toLowerCase(); document.querySelectorAll('#tabla-maquinas tr').forEach(row => { row.style.display = row.innerText.toLowerCase().includes(txt) ? '' : 'none'; }); }

        // MODAL COMPRAS
        function abrirModalCompra(id_repuesto, codigo, nombre) {
            document.getElementById('compra-id-repuesto').value = id_repuesto;
            document.getElementById('compra-subtitulo').innerText = `${codigo} - ${nombre}`;
            document.getElementById('form-registro-compra').reset();
            document.getElementById('modal-compra').classList.remove('hidden');
        }
        function cerrarModalCompra() { document.getElementById('modal-compra').classList.add('hidden'); }
        function guardarCompra(e) {
            e.preventDefault();
            const id = document.getElementById('compra-id-repuesto').value;
            fetch(`/api/inventario/comprar/${id}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    cantidad: document.getElementById('compra-cantidad').value,
                    costo: document.getElementById('compra-costo').value,
                    proveedor: document.getElementById('compra-proveedor').value,
                    factura: document.getElementById('compra-factura').value
                })
            }).then(r => r.json()).then(res => { if(res.status === 'ok') window.location.reload(); });
        }
        
        function abrirModalCompras() {
            fetch('/api/compras').then(res => res.json()).then(data => {
                let html = '';
                if(data.length === 0) {
                    html = '<tr><td colspan="7" class="px-6 py-8 text-center text-sm text-slate-500 dark:text-slate-400 italic">No hay registros de compras.</td></tr>';
                } else {
                    data.forEach(c => {
                        html += `
                            <tr class="hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
                                <td class="px-4 py-3 whitespace-nowrap text-xs text-slate-600 dark:text-slate-400 font-semibold"><i class="fa-regular fa-calendar text-emerald-500 mr-1"></i> ${c.fecha_compra}</td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs"><span class="font-bold text-slate-800 dark:text-slate-200">${c.codigo_pieza}</span><br><span class="text-[10px] text-slate-500 dark:text-slate-400">${c.nombre}</span></td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs font-bold text-slate-700 dark:text-slate-300 text-center">${c.cantidad}</td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs text-slate-600 dark:text-slate-400 text-right">$${c.costo_unitario.toFixed(2)}</td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs font-bold text-emerald-600 dark:text-emerald-400 text-right">$${c.costo_total.toFixed(2)}</td>
                                <td class="px-4 py-3 text-xs text-slate-600 dark:text-slate-400 pl-6">${c.proveedor}</td>
                                <td class="px-4 py-3 whitespace-nowrap"><span class="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 px-2 py-0.5 rounded text-[10px] font-bold border border-slate-200 dark:border-slate-700">${c.factura}</span></td>
                            </tr>
                        `;
                    });
                }
                document.getElementById('tabla-historial-compras').innerHTML = html;
                document.getElementById('modal-historial-compras').classList.remove('hidden');
            });
        }
        function cerrarModalCompras() { document.getElementById('modal-historial-compras').classList.add('hidden'); }

        // MODAL ÓRDENES
        function abrirModalOrden() { 
            document.getElementById('modal-orden-title').innerHTML = '<i class="fa-solid fa-clipboard-list text-cyan-400 mr-2"></i> Programar Orden';
            document.getElementById('form-nueva-orden').reset();
            document.getElementById('orden-id').value = '';
            document.getElementById('modal-orden').classList.remove('hidden'); 
            
            fetch('/api/maquinas').then(r => r.json()).then(data => {
                let s = document.getElementById('select-maquina'); 
                s.innerHTML = '<option value="" disabled selected>Seleccione...</option>'; 
                data.forEach(m => {
                    s.innerHTML += `<option value="${m.id_maquina}" title="${m.codigo_equipo} - ${m.nombre}">${m.codigo_equipo} - ${m.nombre}</option>`;
                });
            });
        }

        function editarOrden(id_orden, id_maquina, tipo, desc, fecha, tecnico) {
            document.getElementById('modal-orden-title').innerHTML = '<i class="fa-solid fa-pen-to-square text-amber-500 mr-2"></i> Editar Orden';
            document.getElementById('orden-id').value = id_orden;
            document.getElementById('select-tipo').value = tipo;
            document.getElementById('input-descripcion').value = desc;
            document.getElementById('input-fecha').value = fecha;
            document.getElementById('input-tecnico').value = tecnico;
            
            fetch('/api/maquinas').then(r => r.json()).then(data => {
                let s = document.getElementById('select-maquina'); 
                s.innerHTML = ''; 
                data.forEach(m => {
                    let sel = m.id_maquina === id_maquina ? 'selected' : '';
                    s.innerHTML += `<option value="${m.id_maquina}" ${sel}>${m.codigo_equipo} - ${m.nombre}</option>`;
                });
            });
            document.getElementById('modal-orden').classList.remove('hidden');
        }

        function eliminarOrden(id) {
            fetch(`/api/ordenes/eliminar/${id}`, {method:'DELETE'}).then(r=>r.json()).then(res=>{
                if(res.status==='ok') window.location.reload(); 
            });
        }

        function cerrarModalOrden() { 
            document.getElementById('modal-orden').classList.add('hidden'); 
        }

        function guardarOrden(e) { 
            e.preventDefault(); 
            const id_orden = document.getElementById('orden-id').value;
            const payload = {
                id_maquina: document.getElementById('select-maquina').value, 
                tipo_mantenimiento: document.getElementById('select-tipo').value, 
                descripcion: document.getElementById('input-descripcion').value, 
                fecha: document.getElementById('input-fecha').value, 
                tecnico: document.getElementById('input-tecnico').value
            };

            const endpoint = id_orden ? `/api/ordenes/editar/${id_orden}` : '/api/ordenes/nueva';
            
            fetch(endpoint, {
                method: 'POST',
                headers: {'Content-Type':'application/json'},
                body: JSON.stringify(payload)
            }).then(r=>r.json()).then(res=>{
                if(res.status==='ok') window.location.reload();
            }); 
        }

        function abrirModalMaquina() { document.getElementById('modal-maquina').classList.remove('hidden'); }
        function cerrarModalMaquina() { document.getElementById('modal-maquina').classList.add('hidden'); document.getElementById('form-nueva-maquina').reset(); }
        function guardarMaquina(e) { e.preventDefault(); fetch('/api/maquinas/nueva',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({codigo:document.getElementById('input-maq-codigo').value, nombre:document.getElementById('input-maq-nombre').value, area:document.getElementById('select-maq-area').value, criticidad:document.getElementById('select-maq-criticidad').value})}).then(r=>r.json()).then(res=>{if(res.status==='ok') window.location.reload();}); }
        
        function eliminarMaquina(id) { 
            fetch(`/api/maquinas/eliminar/${id}`,{method:'DELETE'}).then(r=>r.json()).then(res=>{
                if(res.status==='ok') window.location.reload(); 
                else mostrarNotificacion('No se puede eliminar: Protegido por auditoría.', 'error');
            }); 
        }

        // LÓGICA AVANZADA: COMPLETAR ORDEN Y REPORTE
        let repuestosUsadosTemp = [];
        let evidenciasBase64 = [];
        const modalCompletar = document.getElementById('modal-completar');

        function abrirModalCompletar(id_orden) {
            document.getElementById('completar-id-orden').value = id_orden;
            
            document.getElementById('input-obs').value = '';
            document.getElementById('input-rec').value = '';
            document.getElementById('input-trabajos').value = '';
            document.getElementById('input-equipos').value = '';
            document.getElementById('input-tiempo').value = '';
            
            // Limpiar área fotográfica
            document.getElementById('input-fotos').value = '';
            document.getElementById('preview-fotos').innerHTML = '';
            
            repuestosUsadosTemp = [];
            evidenciasBase64 = [];
            actualizarListaUI();
            
            fetch('/api/inventario').then(res => res.json()).then(data => {
                const select = document.getElementById('select-uso-repuesto');
                select.innerHTML = '<option value="" disabled selected>Selecciona un repuesto...</option>';
                data.forEach(r => {
                    if(r.cantidad_actual > 0) select.innerHTML += `<option value="${r.id_repuesto}" data-nombre="${r.codigo_pieza} - ${r.nombre}">${r.codigo_pieza} - ${r.nombre} (Stock: ${r.cantidad_actual})</option>`;
                });
            });
            modalCompletar.classList.remove('hidden');
        }

        function cerrarModalCompletar() { modalCompletar.classList.add('hidden'); }

        function agregarRepuestoALista() {
            const select = document.getElementById('select-uso-repuesto');
            const cant = document.getElementById('input-uso-cant').value;
            if(!select.value || cant <= 0) { 
                mostrarNotificacion("Selecciona un repuesto y cantidad válida.", 'error'); 
                return; 
            }
            repuestosUsadosTemp.push({ id_repuesto: select.value, nombre: select.options[select.selectedIndex].getAttribute('data-nombre'), cantidad: cant });
            actualizarListaUI();
        }

        function actualizarListaUI() {
            const ul = document.getElementById('lista-repuestos-usados');
            if(repuestosUsadosTemp.length === 0) { ul.innerHTML = '<li class="py-2 text-sm text-slate-500 italic text-center">No se han agregado repuestos.</li>'; return; }
            ul.innerHTML = '';
            repuestosUsadosTemp.forEach((item, index) => {
                ul.innerHTML += `<li class="py-2 flex justify-between items-center text-sm text-slate-300"><span><span class="text-cyan-400 font-bold">${item.cantidad}x</span> ${item.nombre}</span><button type="button" onclick="quitarRepuesto(${index})" class="text-rose-500 hover:text-rose-400"><i class="fa-solid fa-xmark"></i></button></li>`;
            });
        }
        
        function quitarRepuesto(i) { repuestosUsadosTemp.splice(i, 1); actualizarListaUI(); }

        function procesarImagenes(event) {
            const files = event.target.files;
            const previewContainer = document.getElementById('preview-fotos');
            previewContainer.innerHTML = '';
            evidenciasBase64 = [];

            Array.from(files).forEach(file => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const b64 = e.target.result;
                    const base64Data = b64.split(',')[1];
                    evidenciasBase64.push(base64Data);

                    const img = document.createElement('img');
                    img.src = b64;
                    img.className = 'w-full h-24 object-cover rounded-lg border border-slate-700 shadow-sm';
                    previewContainer.appendChild(img);
                };
                reader.readAsDataURL(file);
            });
        }

        function confirmarCompletarOrden() {
            const id_orden = document.getElementById('completar-id-orden').value;
            const obs = document.getElementById('input-obs').value || 'Sin observaciones adicionales.';
            const rec = document.getElementById('input-rec').value || 'Ninguna recomendación específica.';
            const trabajos = document.getElementById('input-trabajos').value || 'Trabajos de rutina.';
            const equipos = document.getElementById('input-equipos').value || 'Herramientas manuales.';
            const tiempo = document.getElementById('input-tiempo').value || 'No especificado.';
            
            fetch(`/api/ordenes/completar/${id_orden}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    repuestos: repuestosUsadosTemp, 
                    observaciones: obs, 
                    recomendaciones: rec, 
                    trabajos_realizados: trabajos, 
                    equipos_necesarios: equipos, 
                    tiempo_ejecucion: tiempo,
                    evidencias: evidenciasBase64
                })
            }).then(res => res.json()).then(r => { if(r.status === 'ok') window.location.reload(); });
        }

        // HISTORIAL DE AUDITORÍA (PDF)
        function abrirModalHistorial() {
            fetch('/api/historial').then(res => res.json()).then(data => {
                let html = '';
                if(data.length === 0) {
                    html = '<tr><td colspan="6" class="px-6 py-8 text-center text-sm text-slate-500 dark:text-slate-400 italic">El historial está vacío.</td></tr>';
                } else {
                    data.forEach(h => {
                        html += `
                            <tr class="hover:bg-cyan-50 dark:hover:bg-slate-800/50 transition-colors">
                                <td class="px-4 py-3 whitespace-nowrap text-xs text-slate-600 dark:text-slate-400 font-semibold"><i class="fa-regular fa-calendar-check text-emerald-500 mr-1"></i> ${h.fecha_ejecucion}</td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs font-bold text-slate-800 dark:text-slate-200">${h.codigo_equipo}</td>
                                <td class="px-4 py-3 whitespace-nowrap"><span class="bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 px-2 py-0.5 rounded text-[10px] font-bold border border-slate-200 dark:border-slate-700">${h.tipo_mantenimiento}</span></td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs text-slate-600 dark:text-slate-400">${h.tiempo_ejecucion}</td>
                                <td class="px-4 py-3 whitespace-nowrap text-xs text-slate-700 dark:text-slate-300 font-medium"><i class="fa-solid fa-user-tag text-cyan-500 mr-1"></i> ${h.tecnico_asignado}</td>
                                <td class="px-4 py-3 text-center whitespace-nowrap flex justify-center">
                                    <button onclick="descargarReportePDF(${h.id_mantenimiento})" class="bg-cyan-600 hover:bg-cyan-500 text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow-sm transition-all flex items-center">
                                        <i class="fa-solid fa-file-pdf mr-1.5"></i> .PDF
                                    </button>
                                </td>
                            </tr>
                        `;
                    });
                }
                document.getElementById('tabla-historial').innerHTML = html;
                document.getElementById('modal-historial').classList.remove('hidden');
            });
        }
        function cerrarModalHistorial() { document.getElementById('modal-historial').classList.add('hidden'); }
        
        function descargarReportePDF(id) { 
            mostrarNotificacion("Generando Reporte PDF oficial...");
            window.location.href = `/api/reporte/${id}`; 
        }

        function abrirModalAlertas(e) { 
            if(e) e.stopPropagation();
            document.getElementById('modal-alertas-stock').classList.remove('hidden'); 
        }
        function cerrarModalAlertas() { document.getElementById('modal-alertas-stock').classList.add('hidden'); }
        
        function descargarInventarioExcel() { 
            mostrarNotificacion("Descargando libro de Excel...");
            window.location.href = "/api/inventario/exportar"; 
        }
        function descargarComprasExcel() { 
            mostrarNotificacion("Descargando libro contable de Compras...");
            window.location.href = "/api/compras/exportar"; 
        }

        window.addEventListener('click', function(e) {
            if (e.target.classList.contains('modal-container')) {
                e.target.classList.add('hidden');
            }
            const panelAlertas = document.getElementById('modal-alertas-stock');
            const btnAlertas = document.getElementById('btn-alertas-flotante');
            if (panelAlertas && !panelAlertas.classList.contains('hidden') && !panelAlertas.contains(e.target) && e.target !== btnAlertas && !btnAlertas.contains(e.target)) {
                cerrarModalAlertas();
            }
        });

        window.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' || e.key === 'Esc') {
                document.querySelectorAll('.modal-container').forEach(m => m.classList.add('hidden'));
                const mAlertas = document.getElementById('modal-alertas-stock');
                if (mAlertas && !mAlertas.classList.contains('hidden')) cerrarModalAlertas();
            }
        });
    </script>
</body>
</html>
"""

MONITOR_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitor de Planta | CAFETEC</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700&family=Rajdhani:wght@700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #020617; color: white; overflow: hidden; }
        .fuente-logo { font-family: 'Rajdhani', sans-serif; }
    </style>
</head>
<body class="flex flex-col h-screen p-4 gap-4">
    <!-- HEADER -->
    <header class="flex justify-between items-center bg-slate-900 p-4 rounded-2xl border border-slate-800 shadow-2xl shrink-0">
        <div class="flex items-center">
            <div class="flex items-center bg-slate-950 border-2 border-slate-700 rounded-lg px-3 py-1 shadow-[0_0_15px_rgba(6,182,212,0.15)]">
                <span class="text-cyan-400 font-bold text-4xl fuente-logo tracking-wider">CAFE</span>
                <span class="text-white font-bold text-4xl fuente-logo tracking-wider ml-1">TEC</span>
            </div>
            <span class="ml-6 pl-6 border-l-2 border-slate-700 text-slate-400 uppercase tracking-widest font-bold text-lg flex items-center">
                <i class="fa-solid fa-tower-broadcast text-rose-500 animate-pulse mr-3"></i> Monitor en Tiempo Real
            </span>
        </div>
        <div class="text-right flex flex-col items-end">
            <div id="reloj" class="text-4xl font-black text-white tracking-widest font-mono drop-shadow-md">00:00:00</div>
            <div id="fecha" class="text-cyan-400 text-sm font-bold uppercase mt-1">---</div>
        </div>
    </header>

    <!-- MAIN CONTENT -->
    <main class="flex-1 grid grid-cols-1 lg:grid-cols-3 gap-4 min-h-0">
        <!-- COL 1: KPIs Principales -->
        <div class="flex flex-col gap-4 min-h-0">
            <div class="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-2xl flex-1 flex flex-col justify-center items-center text-center relative overflow-hidden">
                <div id="bg-estado" class="absolute inset-0 opacity-10 bg-emerald-500 transition-colors duration-1000"></div>
                <div class="absolute top-4 left-4 flex items-center text-slate-500 text-xs font-bold uppercase tracking-widest">
                    <i class="fa-solid fa-circle-dot mr-2"></i> Estatus Planta
                </div>
                <div id="estado-texto" class="text-[3.5rem] leading-none font-black text-emerald-400 z-10 transition-colors duration-1000 mt-4">OPERATIVA</div>
            </div>
            
            <div class="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-2xl flex-1 flex flex-col justify-center">
                <h2 class="text-slate-400 font-bold uppercase tracking-widest mb-6 text-sm"><i class="fa-solid fa-list-check mr-2"></i> Progreso de Órdenes (Hoy)</h2>
                <div class="flex justify-between items-end mb-4">
                    <div>
                        <span class="text-[3.5rem] leading-none font-black text-white" id="ordenes-progreso">0/0</span>
                        <span class="text-slate-500 font-bold ml-2 uppercase text-sm">Completadas</span>
                    </div>
                    <span class="text-cyan-400 font-black text-4xl" id="ordenes-porcentaje">0%</span>
                </div>
                <div class="w-full bg-slate-950 rounded-full h-6 border border-slate-800 overflow-hidden shadow-inner">
                    <div class="bg-cyan-500 h-full rounded-full transition-all duration-1000 relative overflow-hidden" id="barra-progreso" style="width: 0%">
                        <div class="absolute inset-0 bg-white/20 w-full h-full" style="background-image: linear-gradient(45deg,rgba(255,255,255,.15) 25%,transparent 25%,transparent 50%,rgba(255,255,255,.15) 50%,rgba(255,255,255,.15) 75%,transparent 75%,transparent); background-size: 1rem 1rem;"></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- COL 2: Gráfico de Disponibilidad -->
        <div class="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-2xl flex flex-col items-center relative overflow-hidden min-h-0">
            <div class="absolute top-0 w-full h-1.5 bg-gradient-to-r from-cyan-500 to-blue-500"></div>
            <h2 class="text-slate-400 font-bold uppercase tracking-widest mb-4 w-full text-left text-sm flex items-center shrink-0">
                <i class="fa-solid fa-chart-pie mr-2"></i> Disponibilidad de Equipos
            </h2>
            <div class="flex-1 w-full relative flex items-center justify-center min-h-0">
                <div class="relative w-full h-full max-h-64 flex justify-center items-center">
                    <canvas id="chartDisponibilidad"></canvas>
                    <div class="absolute inset-0 flex items-center justify-center pointer-events-none flex-col mt-4">
                        <span class="text-6xl font-black text-white drop-shadow-lg" id="disp-porcentaje">100%</span>
                        <span class="text-sm text-emerald-400 font-bold uppercase tracking-widest mt-1">Online</span>
                    </div>
                </div>
            </div>
            <div class="w-full mt-4 flex justify-center space-x-6 shrink-0">
                <div class="flex items-center"><span class="w-4 h-4 rounded-full bg-emerald-500 mr-2 shadow-[0_0_10px_rgba(16,185,129,0.5)]"></span><span class="text-slate-300 font-bold uppercase text-xs tracking-wider">Operativas</span></div>
                <div class="flex items-center"><span class="w-4 h-4 rounded-full bg-rose-600 mr-2 shadow-[0_0_10px_rgba(225,29,72,0.5)]"></span><span class="text-slate-300 font-bold uppercase text-xs tracking-wider">En Mantenimiento</span></div>
            </div>
        </div>

        <!-- COL 3: Equipos Intervenidos -->
        <div class="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-2xl flex flex-col relative overflow-hidden min-h-0">
            <div class="absolute top-0 w-full h-1.5 bg-gradient-to-r from-rose-500 to-orange-500"></div>
            <h2 class="text-slate-400 font-bold uppercase tracking-widest mb-4 flex justify-between items-center text-sm shrink-0">
                <span><i class="fa-solid fa-triangle-exclamation mr-2"></i> Equipos Detenidos</span>
                <span class="bg-rose-600 text-white px-3 py-1 rounded-lg text-lg shadow-[0_0_15px_rgba(225,29,72,0.5)]" id="badge-detenidos">0</span>
            </h2>
            <div class="flex-1 overflow-y-auto pr-2 space-y-3 min-h-0" id="lista-detenidos">
                <!-- Se inyecta por JS -->
                <div class="h-full flex flex-col items-center justify-center text-slate-500 italic opacity-50">
                    <i class="fa-solid fa-check-circle text-6xl mb-4 text-emerald-500"></i>
                    <p class="font-bold text-lg">Todos los equipos operativos</p>
                </div>
            </div>
        </div>
    </main>

    <script>
        function actualizarReloj() {
            const ahora = new Date();
            document.getElementById('reloj').textContent = ahora.toLocaleTimeString('es-ES', { hour12: false });
            const opciones = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
            let fechaStr = ahora.toLocaleDateString('es-ES', opciones);
            document.getElementById('fecha').textContent = fechaStr;
        }
        setInterval(actualizarReloj, 1000);
        actualizarReloj();

        let chartDisp = null;

        function actualizarDatos() {
            fetch('/api/monitor_data')
                .then(r => r.json())
                .then(d => {
                    const totOrd = d.ordenes.total;
                    const compOrd = d.ordenes.completadas;
                    document.getElementById('ordenes-progreso').textContent = `${compOrd}/${totOrd}`;
                    let pctOrd = 100;
                    if(totOrd > 0) pctOrd = Math.round((compOrd / totOrd) * 100);
                    
                    const countUp = (id, target) => {
                        document.getElementById(id).textContent = target + '%';
                    };
                    countUp('ordenes-porcentaje', pctOrd);
                    document.getElementById('barra-progreso').style.width = `${pctOrd}%`;

                    const totMaq = d.maquinas.total;
                    const opMaq = d.maquinas.operativas;
                    const detMaq = d.maquinas.detenidas;
                    
                    let estadoGlobal = 'OPERATIVA';
                    let colorEstado = 'emerald';
                    if(detMaq.length > 0 && detMaq.length < totMaq) { estadoGlobal = 'OP. PARCIAL'; colorEstado = 'amber'; }
                    else if(detMaq.length === totMaq && totMaq > 0) { estadoGlobal = 'PARADA GENERAL'; colorEstado = 'rose'; }
                    
                    const bgEstado = document.getElementById('bg-estado');
                    const txtEstado = document.getElementById('estado-texto');
                    
                    bgEstado.className = `absolute inset-0 opacity-10 bg-${colorEstado}-500 transition-colors duration-1000`;
                    txtEstado.className = `text-[3.5rem] leading-none font-black text-${colorEstado}-400 z-10 transition-colors duration-1000 mt-4`;
                    txtEstado.textContent = estadoGlobal;

                    let pctDisp = 100;
                    if(totMaq > 0) pctDisp = Math.round((opMaq / totMaq) * 100);
                    countUp('disp-porcentaje', pctDisp);

                    if(!chartDisp) {
                        const ctx = document.getElementById('chartDisponibilidad').getContext('2d');
                        chartDisp = new Chart(ctx, {
                            type: 'doughnut',
                            data: { labels: ['Operativas', 'Detenidas'], datasets: [{ data: [opMaq, detMaq.length], backgroundColor: ['#10B981', '#E11D48'], borderWidth: 0 }] },
                            options: { responsive: true, maintainAspectRatio: false, cutout: '80%', plugins: { legend: { display: false }, tooltip: { enabled: false } }, animation: { animateScale: true } }
                        });
                    } else {
                        chartDisp.data.datasets[0].data = [opMaq, detMaq.length];
                        chartDisp.update();
                    }

                    document.getElementById('badge-detenidos').textContent = detMaq.length;
                    const lista = document.getElementById('lista-detenidos');
                    if(detMaq.length === 0) {
                        lista.innerHTML = `
                            <div class="h-full flex flex-col items-center justify-center text-slate-500 italic opacity-50 transition-opacity duration-500">
                                <i class="fa-solid fa-check-circle text-6xl mb-4 text-emerald-500"></i>
                                <p class="font-bold text-lg">Todos los equipos operativos</p>
                            </div>
                        `;
                    } else {
                        lista.innerHTML = detMaq.map(m => `
                            <div class="bg-slate-950 border-l-4 border-rose-500 rounded-lg p-4 shadow-md flex items-center justify-between transform transition-all duration-300 hover:scale-105">
                                <div>
                                    <div class="font-black text-rose-400 text-lg tracking-wider">${m.codigo_equipo}</div>
                                    <div class="text-slate-300 font-medium truncate w-48 text-sm">${m.nombre}</div>
                                </div>
                                <div class="w-10 h-10 rounded-full bg-rose-900/30 flex items-center justify-center text-rose-500 animate-pulse">
                                    <i class="fa-solid fa-wrench"></i>
                                </div>
                            </div>
                        `).join('');
                    }
                })
                .catch(err => console.error("Error en monitor:", err));
        }

        setInterval(actualizarDatos, 5000);
        actualizarDatos();
    </script>
</body>
</html>
"""

MOBILE_MACHINE_TEMPLATE = """
<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ maquina.codigo_equipo }} | CAFETEC</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Rajdhani:wght@600;700&display=swap" rel="stylesheet">
    <style>body { font-family: 'Inter', sans-serif; background-color: #0f172a; color: white; }</style>
</head>
<body class="pb-20">
    <header class="bg-slate-900 border-b border-slate-800 p-4 sticky top-0 z-50 shadow-lg flex justify-between items-center">
        <div class="flex flex-col">
            <span class="text-cyan-400 font-bold text-sm tracking-widest uppercase">Perfil de Máquina</span>
            <span class="font-black text-2xl text-white">{{ maquina.codigo_equipo }}</span>
        </div>
        <a href="/" class="bg-slate-800 text-white w-10 h-10 rounded-full flex items-center justify-center shadow-inner">
            <i class="fa-solid fa-house"></i>
        </a>
    </header>

    <div class="p-4 space-y-4">
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
                    <span class="text-slate-500 text-[10px] uppercase font-bold tracking-wider block">Estado Actual</span>
                    {% if maquina.estado == 'Operativa' %}
                        <span class="text-emerald-400 font-bold text-sm"><i class="fa-solid fa-check-circle mr-1"></i> OPERATIVA</span>
                    {% else %}
                        <span class="text-rose-400 font-bold text-sm"><i class="fa-solid fa-triangle-exclamation mr-1"></i> EN MANTENIMIENTO</span>
                    {% endif %}
                </div>
                <div class="bg-slate-950 p-3 rounded-xl border border-slate-800">
                    <span class="text-slate-500 text-[10px] uppercase font-bold tracking-wider block">Criticidad</span>
                    <span class="text-white font-bold text-sm">{{ maquina.criticidad }}</span>
                </div>
            </div>
        </div>

        <h2 class="text-slate-400 font-bold uppercase tracking-widest text-xs mt-6 mb-2 ml-1">Acciones Rápidas</h2>
        <div class="grid grid-cols-2 gap-3">
            <a href="/" class="bg-cyan-600 hover:bg-cyan-500 text-white rounded-xl p-4 font-bold flex flex-col items-center justify-center text-center shadow-lg transition-colors">
                <i class="fa-solid fa-clipboard-list text-2xl mb-2"></i>
                <span class="text-sm">Ver Órdenes</span>
            </a>
            <a href="/" class="bg-slate-800 hover:bg-slate-700 text-white rounded-xl p-4 font-bold flex flex-col items-center justify-center text-center shadow-lg border border-slate-700 transition-colors">
                <i class="fa-solid fa-boxes-stacked text-2xl mb-2 text-cyan-400"></i>
                <span class="text-sm">Repuestos</span>
            </a>
        </div>
    </div>
</body>
</html>
"""

# ==========================================
# 3. CONTROLADORES DE RUTAS (ENDPOINTS FLASK)
# ==========================================

@app.route('/login', methods=['GET', 'POST'])
def login():
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
            return redirect(url_for('dashboard'))
            
        return render_template_string(LOGIN_TEMPLATE, error="Usuario o contraseña incorrectos")
        
    return render_template_string(LOGIN_TEMPLATE, error=None)

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
    conn.close()
    if not maquina: return "Máquina no encontrada", 404
    return render_template_string(MOBILE_MACHINE_TEMPLATE, maquina=maquina)

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
            @media print {
                body { background: white; }
                .no-print { display: none; }
                .etiqueta { border: none; border-radius: 0; }
            }
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
            <div class="qr-container">
                <img src="data:image/png;base64,{{ qr_img }}" alt="QR Code">
            </div>
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
    comp_ord = completadas_hoy
    conn.close()
    
    return jsonify({
        "maquinas": {"total": tot_maq, "operativas": op_maq, "detenidas": detenidas},
        "ordenes": {"total": tot_ord, "completadas": comp_ord}
    })

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
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    descontar_stock_db(id_repuesto)
    return jsonify({"status": "ok"})

@app.route('/api/inventario/sumar/<int:id_repuesto>', methods=['POST'])
def api_sumar_inventario(id_repuesto):
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    sumar_stock_db(id_repuesto)
    return jsonify({"status": "ok"})

@app.route('/api/inventario/comprar/<int:id_repuesto>', methods=['POST'])
def api_registrar_compra(id_repuesto):
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    d = request.json
    registrar_compra_db(id_repuesto, float(d['cantidad']), float(d['costo']), d['proveedor'], d['factura'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/nueva', methods=['POST'])
def api_nueva_orden():
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    d = request.json
    crear_nueva_orden_db(d['id_maquina'], d['tipo_mantenimiento'], d['descripcion'], d['fecha'], d['tecnico'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/editar/<int:id_orden>', methods=['POST'])
def api_editar_orden(id_orden):
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    d = request.json
    actualizar_orden_db(id_orden, d['id_maquina'], d['tipo_mantenimiento'], d['descripcion'], d['fecha'], d['tecnico'])
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/eliminar/<int:id_orden>', methods=['DELETE'])
def api_borrar_orden(id_orden):
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    eliminar_orden_db(id_orden)
    return jsonify({"status": "ok"})

@app.route('/api/ordenes/completar/<int:id_orden>', methods=['POST'])
def api_completar_orden(id_orden):
    d = request.json or {}
    repuestos = d.get('repuestos', [])
    obs = d.get('observaciones', 'Sin observaciones.')
    rec = d.get('recomendaciones', 'Ninguna.')
    trabajos = d.get('trabajos_realizados', 'No especificado.')
    equipos = d.get('equipos_necesarios', 'Ninguno.')
    tiempo = d.get('tiempo_ejecucion', '0 h')
    evidencias = d.get('evidencias', [])
    
    completar_orden_db(id_orden, repuestos, obs, rec, trabajos, equipos, tiempo, evidencias)
    return jsonify({"status": "ok"})

@app.route('/api/maquinas/nueva', methods=['POST'])
def api_nueva_maquina():
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
    d = request.json
    crear_maquina_db(d['codigo'], d['nombre'], d['area'], d['criticidad'])
    return jsonify({"status": "ok"})

@app.route('/api/maquinas/eliminar/<int:id_maquina>', methods=['DELETE'])
def api_borrar_maquina(id_maquina):
    if session.get('rol') != 'Admin': return jsonify({"status": "error", "msg": "No autorizado"}), 403
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
    for col_num, _ in enumerate(encabezados, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill
        cell.font = font

    for item in inventario: 
        ws.append([item['codigo_pieza'], item['nombre'], item['descripcion'], item['cantidad_actual'], item['punto_reorden'], item['unidad_medida'], item['ubicacion_almacen']])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(cell.value)
            except: pass
        ws.column_dimensions[column].width = max_length + 2

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
    for col_num, _ in enumerate(encabezados, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill
        cell.font = font

    for item in compras: 
        ws.append([item['fecha_compra'], item['codigo_pieza'], item['nombre'], item['cantidad'], item['costo_unitario'], item['costo_total'], item['proveedor'], item['factura']])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(cell.value)
            except: pass
        ws.column_dimensions[column].width = max_length + 2

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
                FROM Repuestos_Orden ro
                JOIN Repuestos_Stock rs ON ro.id_repuesto = rs.id_repuesto
                WHERE ro.id_mantenimiento = c.id_mantenimiento) as repuestos_usados
        FROM Calendario_Mantenimiento c
        JOIN Maquinas m ON c.id_maquina = m.id_maquina
        WHERE c.id_mantenimiento = ?
    """
    orden = conn.execute(query, (id_orden,)).fetchone()
    
    evidencias_db = conn.execute("SELECT imagen_base64 FROM Evidencia_Fotografica WHERE id_mantenimiento = ?", (id_orden,)).fetchall()
    conn.close()

    if not orden: return "Orden no encontrada", 404

    file_stream = io.BytesIO()
    doc = SimpleDocTemplate(file_stream, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements = []
    
    title_style = ParagraphStyle(
        name='Title',
        fontName='Helvetica-Bold',
        fontSize=11,
        spaceAfter=8,
        spaceBefore=14,
        textColor=colors.HexColor('#0F172A')
    )
    normal_style = ParagraphStyle(
        name='NormalCustom',
        fontName='Helvetica',
        fontSize=10,
        spaceAfter=4,
        leading=14
    )

    logo_path = os.path.join(BASE_DIR, "logo_cafetec.jpg")
    if os.path.exists(logo_path):
        celda_logo = Image(logo_path, width=112, height=55, kind='proportional')
    else:
        celda_logo = Paragraph("<font color='white'><b>CAFETEC</b></font>", ParagraphStyle(name='LogoFB', alignment=1, fontName='Times-Bold', fontSize=14))

    p_tit_top = Paragraph("REGISTRO DE MANTENIMIENTO", ParagraphStyle(name='TitTop', alignment=1, fontName='Times-Bold', fontSize=11, leading=14))
    p_tit_bot = Paragraph(f"MANTENIMIENTO DE MAQUINA {orden['maquina_nombre'].upper()}", ParagraphStyle(name='TitBot', alignment=1, fontName='Times-Bold', fontSize=10, textColor=colors.HexColor('#004b87'), leading=12)) 
    
    lbl_style = ParagraphStyle(name='Lbl', alignment=0, fontName='Times-Roman', fontSize=10)
    val_style = ParagraphStyle(name='Val', alignment=0, fontName='Times-Roman', fontSize=10)

    fecha_format = orden['fecha_ejecucion']
    try:
        f_obj = datetime.strptime(fecha_format, '%Y-%m-%d')
        fecha_format = f_obj.strftime('%d/%m/%y')
    except:
        pass

    header_data = [
        [celda_logo, p_tit_top, Paragraph("Código:", lbl_style), Paragraph("RE MAN-262", val_style)],
        ["", "", Paragraph("Versión:", lbl_style), Paragraph("001", val_style)],
        ["", p_tit_bot, Paragraph("Fecha:", lbl_style), Paragraph(fecha_format, val_style)],
        ["", "", Paragraph("Página:", lbl_style), Paragraph("1", val_style)]
    ]
    
    t_header = Table(header_data, colWidths=[120, 250, 60, 80], rowHeights=[15, 15, 15, 15])
    t_header.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 3)),
        ('SPAN', (1, 0), (1, 1)),
        ('SPAN', (1, 2), (1, 3)),
        ('BACKGROUND', (0, 0), (0, 3), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 15))

    elements.append(Paragraph("1. INFORMACIÓN GENERAL:", title_style))
    info_text = f"""
    • <b>Fecha de Ejecución:</b> {orden['fecha_ejecucion']}<br/>
    • <b>Tipo de trabajo a realizar:</b> {orden['tipo_mantenimiento']}<br/>
    • <b>Código del activo:</b> {orden['codigo_equipo']}<br/>
    • <b>Nombre de activo:</b> {orden['maquina_nombre']}<br/>
    • <b>Técnico a cargo:</b> {orden['tecnico_asignado']}<br/>
    • <b>Duración:</b> {orden['tiempo_ejecucion']}
    """
    elements.append(Paragraph(info_text, normal_style))

    elements.append(Paragraph("2. TRABAJO SOLICITADO:", title_style))
    elements.append(Paragraph(f"• Mantenimiento {orden['tipo_mantenimiento'].lower()} reportado: {orden['descripcion_tarea']}", normal_style))

    elements.append(Paragraph("3. TRABAJOS REALIZADOS:", title_style))
    trabajos = orden['trabajos_realizados'].replace('\n', '<br/>')
    elements.append(Paragraph(trabajos, normal_style))

    elements.append(Paragraph("4. DESCRIPCIÓN Y OBSERVACIONES:", title_style))
    observaciones = orden['observaciones'].replace('\n', '<br/>')
    elements.append(Paragraph(f"• {observaciones}", normal_style))

    elements.append(Paragraph("RECURSOS NECESARIOS:", title_style))
    rep_text = orden['repuestos_usados'].replace('\n', '<br/>• ') if orden['repuestos_usados'] else "<i>No se utilizaron repuestos de almacén.</i>"
    if orden['repuestos_usados']:
        rep_text = "• " + rep_text
        
    equipos_list = orden['equipos_necesarios'].split(',')
    equipos_text = "<br/>".join([f"• {e.strip()}" for e in equipos_list if e.strip()])
    if not equipos_text: equipos_text = "• Herramientas manuales estándar"

    recursos_text = f"""
    <b>Mano de obra:</b><br/>• 01 técnico de mantenimiento ({orden['tecnico_asignado']})<br/><br/>
    <b>Equipos necesarios:</b><br/>{equipos_text}<br/><br/>
    <b>Materiales y repuestos:</b><br/>{rep_text}
    """
    elements.append(Paragraph(recursos_text, normal_style))

    elements.append(Paragraph("5. COMENTARIOS / RECOMENDACIONES:", title_style))
    recomendaciones = orden['recomendaciones'].replace('\n', '<br/>')
    elements.append(Paragraph(f"• {recomendaciones}", normal_style))

    elements.append(Paragraph("6. EVIDENCIA FOTOGRÁFICA:", title_style))
    
    if evidencias_db:
        try:
            resample_mode = PILImage.Resampling.LANCZOS
        except AttributeError:
            resample_mode = PILImage.LANCZOS

        evidencia_data = []
        row = []
        
        for index, row_db in enumerate(evidencias_db):
            try:
                img_data = base64.b64decode(row_db['imagen_base64'])
                img_pil = PILImage.open(io.BytesIO(img_data))
                if img_pil.mode in ('RGBA', 'P'):
                    img_pil = img_pil.convert('RGB')
                
                width, height = img_pil.size
                new_size = min(width, height)
                left = (width - new_size) / 2
                top = (height - new_size) / 2
                right = (width + new_size) / 2
                bottom = (height + new_size) / 2
                img_pil = img_pil.crop((left, top, right, bottom))
                
                img_pil = img_pil.resize((300, 300), resample_mode)
                
                output = io.BytesIO()
                img_pil.save(output, format='JPEG', quality=85)
                output.seek(0)
                
                rl_img = Image(output, width=240, height=240)
                row.append(rl_img)
                
                if len(row) == 2:
                    evidencia_data.append(row)
                    row = []
            except Exception as e:
                print(f"Error procesando imagen para el reporte: {e}")
                
        if row: 
            row.append("")
            evidencia_data.append(row)
            
        t_evidencia = Table(evidencia_data, colWidths=[255, 255])
        t_evidencia.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(t_evidencia)
    else:
        box = Table([["\n\n(Espacio reservado para adjuntar evidencia fotográfica post-impresión)\n\n"]], colWidths=[510])
        box.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#94a3b8')),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#64748b')),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Oblique'),
        ]))
        elements.append(box)
        
    elements.append(Spacer(1, 85))

    firmas_data = [
        ["________________________", "________________________", "________________________"],
        [f"Elaborado:\n{orden['tecnico_asignado']}", "Revisado:\nÁrea de Mantenimiento", "Aprobado:\nGerencia de Planta"]
    ]
    t_firmas = Table(firmas_data, colWidths=[170, 170, 170])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elements.append(t_firmas)

    doc.build(elements)
    file_stream.seek(0)
    
    nombre_archivo = f"RE-MAN-262_{orden['codigo_equipo']}_{orden['fecha_ejecucion']}.pdf"
    return send_file(file_stream, as_attachment=True, download_name=nombre_archivo, mimetype='application/pdf')

if __name__ == '__main__':
    print("=====================================================")
    print("⚙️  INICIANDO CAFETEC CMMS - ÁREA DE MANTENIMIENTO")
    print("=====================================================")
    migrar_base_datos()
    app.run(host='0.0.0.0', debug=False, port=5000)
