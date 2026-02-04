from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
import pandas as pd
import io
import qrcode
import matplotlib.pyplot as plt
import os
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

#Configuración de la app flask
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'clave_secreta_por_defecto')

# Variable global para almacenar mensaje de limpieza de duplicados
mensaje_limpieza_global = None

# # Lista de programas para docentes y estudiantes
# PROGRAMAS = [
#     {"value": "actividad-fisica-deporte", "label": "Actividad Física y Deporte (Presencial)"},
#     {"value": "administracion-empresas-distancia", "label": "Administración de Empresas (Distancia)"},
#     {"value": "administracion-empresas-presencial", "label": "Administración de Empresas (Presencial)"},
#     {"value": "arquitectura-presencial", "label": "Arquitectura (Presencial)"},
#     {"value": "ciclo-prof-tecnologos-sistemas", "label": "Ciclo de Profesionalización para Tecnólogos en Sistemas"},
#     {"value": "comunicacion-social", "label": "Comunicación Social (Presencial)"},
#     {"value": "contaduria-publica-presencial", "label": "Contaduría Pública (Presencial)"},
#     {"value": "contaduria-publica-distancia", "label": "Contaduría Pública (Distancia)"},
#     {"value": "derecho-presencial", "label": "Derecho (Presencial)"},
#     {"value": "desarrollo-familiar-distancia", "label": "Desarrollo Familiar (Distancia)"},
#     {"value": "desarrollo-familiar-presencial", "label": "Desarrollo Familiar (Presencial)"},
#     {"value": "diseno-grafico-presencial", "label": "Diseño Gráfico (Presencial)"},
#     {"value": "economia-presencial", "label": "Economía (Presencial)"},
#     {"value": "filosofia-presencial", "label": "Filosofía (Presencial)"},
#     {"value": "gastronomia-presencial", "label": "Gastronomía (Presencial)"},
#     {"value": "ingenieria-sistemas-presencial", "label": "Ingeniería de Sistemas (Presencial)"},
#     {"value": "ingenieria-industrial-presencial", "label": "Ingeniería Industrial (Presencial)"},
#     {"value": "ingenieria-civil-presencial", "label": "Ingeniería Civil (Presencial)"},
#     {"value": "lic-ingles-presencial", "label": "Licenciatura en Educación Básica - Inglés (Presencial)"},
#     {"value": "lic-matematicas", "label": "Licenciatura en Educación Básica - Matemáticas"},
#     {"value": "lic-tecnologia", "label": "Licenciatura en Educación Básica - Tecnología e Informática"},
#     {"value": "lic-educacion-infantil", "label": "Licenciatura en Educación Infantil (Presencial)"},
#     {"value": "lic-educacion-preescolar", "label": "Licenciatura en Educación Preescolar (Presencial - Distancia)"},
#     {"value": "lic-filosofia", "label": "Licenciatura en Filosofía (Presencial - Distancia)"},
#     {"value": "lic-ingles", "label": "Licenciatura en Inglés (Presencial)"},
#     {"value": "lic-lengua-castellana", "label": "Licenciatura en Lengua Castellana (Distancia)"},
#     {"value": "lic-lenguas-extranjeras", "label": "Licenciatura en Lenguas Extranjeras con énfasis en Inglés (Presencial)"},
#     {"value": "lic-pedagogia-reeducativa", "label": "Licenciatura en Pedagogía Reeducativa (Distancia)"},
#     {"value": "lic-teologia", "label": "Licenciatura en Teología (Presencial - Distancia)"},
#     {"value": "negocios-internacionales", "label": "Negocios Internacionales (Presencial)"},
#     {"value": "psicologia-presencial", "label": "Psicología (Presencial)"},
#     {"value": "psicologia-distancia", "label": "Psicología (Distancia)"},
#     {"value": "publicidad-presencial", "label": "Publicidad (Presencial)"},
#     {"value": "tecnologia-desarrollo-software", "label": "Tecnología en Desarrollo de Software (Presencial)"},
#     {"value": "teologia-presencial", "label": "Teología (Presencial)"},
#     {"value": "trabajo-social", "label": "Trabajo Social (Distancia)"}
# ]

# def migrate_programs_to_db():
#     """Migrar programas de la lista PROGRAMAS a la base de datos"""
#     with get_db_connection() as conn:
#         cursor = conn.cursor()
#         for programa in PROGRAMAS:
#             try:
#                 cursor.execute("""
#                     INSERT OR IGNORE INTO programas (nombre, activo)
#                     VALUES (?, 1)
#                 """, (programa['label'],))
#             except:
#                 pass
#         conn.commit()


# --- Funciones auxiliares ---
def get_db_connection():
    """Función auxiliar para obtener conexión a la base de datos"""
    conn = sqlite3.connect("biblioteca.db")
    conn.row_factory = sqlite3.Row
    return conn

def get_programas_list():
    """Obtener lista de todos los programas desde la base de datos"""
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nombre FROM programas WHERE activo = 1 ORDER BY nombre ASC")
            programas = cursor.fetchall()
            return [{"value": p["nombre"], "label": p["nombre"]} for p in programas]
    except:
        return []

def get_programas_map():
    """Obtener diccionario de mapeo nombre -> nombre (por compatibilidad)"""
    programas_map = {}
    for programa in get_programas_list():
        programas_map[programa['value']] = programa['label']
    return programas_map

def get_ventana_anos(num_anos=5):
    """
    Obtiene una ventana deslizante de años basada en el año actual.
    Por defecto retorna los últimos 5 años.
    
    Ejemplo en 2024: [2020, 2021, 2022, 2023, 2024]
    Ejemplo en 2025: [2021, 2022, 2023, 2024, 2025]
    
    Args:
        num_anos (int): Número de años a incluir en la ventana (por defecto 5)
    
    Returns:
        tuple: (lista de años, año_inicio, año_fin)
    """
    from datetime import datetime
    ano_actual = datetime.now().year
    ano_inicio = ano_actual - (num_anos - 1)
    anos = list(range(ano_inicio, ano_actual + 1))
    
    return anos, ano_inicio, ano_actual

def limpiar_datos_antiguos(anos_a_mantener=5):
    """
    Limpia datos de asistencias que estén fuera de la ventana de años especificada.
    Esta función puede ejecutarse manualmente o programarse para ejecutarse periódicamente.
    
    Args:
        anos_a_mantener (int): Número de años de datos a mantener
    
    Returns:
        dict: Información sobre los registros eliminados
    """
    _, ano_inicio, _ = get_ventana_anos(anos_a_mantener)
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # Contar registros a eliminar
            cursor.execute("""
                SELECT COUNT(*) as total
                FROM asistencias
                WHERE CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) < ?
            """, (ano_inicio,))
            
            total_a_eliminar = cursor.fetchone()['total']
            
            if total_a_eliminar > 0:
                # Eliminar registros antiguos
                cursor.execute("""
                    DELETE FROM asistencias
                    WHERE CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) < ?
                """, (ano_inicio,))
                
                conn.commit()
                
                return {
                    'success': True,
                    'registros_eliminados': total_a_eliminar,
                    'ano_limite': ano_inicio,
                    'mensaje': f'Se eliminaron {total_a_eliminar} registros anteriores al año {ano_inicio}'
                }
            else:
                return {
                    'success': True,
                    'registros_eliminados': 0,
                    'ano_limite': ano_inicio,
                    'mensaje': 'No hay registros antiguos para eliminar'
                }
                
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'mensaje': f'Error al limpiar datos antiguos: {str(e)}'
        }

# --- Inicialización de la base de datos ---
def init_db():
    """
    Inicializa la base de datos y retorna mensaje de limpieza de duplicados si aplica
    """
    mensaje_limpieza = None
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Tabla de usuarios
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tabla de programas académicos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS programas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                activo INTEGER DEFAULT 1,
                fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                fecha_modificacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # # Migrar programas existentes a la base de datos
        # migrate_programs_to_db()

        # Tabla de asistencias (capacitaciones)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS asistencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_evento TEXT NOT NULL,
                dictado_por TEXT NOT NULL,
                docente TEXT NOT NULL,
                programa_docente TEXT NOT NULL,
                numero_identificacion TEXT NOT NULL,
                nombre_completo TEXT NOT NULL,
                programa_estudiante TEXT NOT NULL,
                modalidad TEXT NOT NULL,
                tipo_asistente TEXT NOT NULL,
                sede TEXT NOT NULL,
                fecha_evento TEXT NOT NULL,
                fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Verificar columna fecha_evento en asistencias
        cursor.execute("PRAGMA table_info(asistencias)")
        columnas = [columna[1] for columna in cursor.fetchall()]
        
        if 'fecha_evento' not in columnas:
            cursor.execute("ALTER TABLE asistencias ADD COLUMN fecha_evento TEXT")
            cursor.execute("""
                UPDATE asistencias 
                SET fecha_evento = DATE(fecha_registro) 
                WHERE fecha_evento IS NULL
            """)
        
        # Crear índice UNIQUE para prevenir duplicados
        # Un registro es duplicado si tiene la misma identificación, evento y fecha
        try:
            # Primero, verificar si hay duplicados existentes
            cursor.execute("""
                SELECT COUNT(*) as count FROM (
                    SELECT numero_identificacion, nombre_evento, fecha_evento, COUNT(*) as cantidad
                    FROM asistencias
                    GROUP BY numero_identificacion, nombre_evento, fecha_evento
                    HAVING COUNT(*) > 1
                )
            """)
            duplicados_count = cursor.fetchone()[0]
            
            if duplicados_count > 0:
                print(f"⚠️ Se encontraron {duplicados_count} grupos con duplicados. Limpiando...")
                
                # Eliminar duplicados manteniendo solo el registro más antiguo (menor ID)
                cursor.execute("""
                    DELETE FROM asistencias
                    WHERE id NOT IN (
                        SELECT MIN(id)
                        FROM asistencias
                        GROUP BY numero_identificacion, nombre_evento, fecha_evento
                    )
                """)
                eliminados = cursor.rowcount
                conn.commit()
                print(f"✅ Se eliminaron {eliminados} registros duplicados")
                
                # Guardar mensaje para mostrarlo a los usuarios
                mensaje_limpieza = f"Se detectaron y eliminaron {eliminados} registros duplicados en la base de datos. El sistema ahora está protegido contra duplicados."
            
            # Ahora crear el índice UNIQUE
            cursor.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_asistencias_unique 
                ON asistencias(numero_identificacion, nombre_evento, fecha_evento)
            """)
            print("✅ Índice UNIQUE creado para prevenir duplicados")
            
        except sqlite3.OperationalError as e:
            # El índice ya existe o hay otro error
            print(f"ℹ️ Índice UNIQUE: {str(e)}")
            pass
        # Tabla de inversiones institucionales
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inversiones_institucionales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                año INTEGER NOT NULL,
                monto_libros REAL NOT NULL DEFAULT 0,
                monto_revistas REAL NOT NULL DEFAULT 0,
                monto_bases_datos REAL NOT NULL DEFAULT 0,
                total REAL GENERATED ALWAYS AS (monto_libros + monto_revistas + monto_bases_datos) STORED,
                observaciones TEXT,
                fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(año)
            )
        """)
        
        # Tabla de inversiones por programa
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inversiones_programas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                año INTEGER NOT NULL,
                programa TEXT NOT NULL,
                libros_titulos INTEGER NOT NULL DEFAULT 0,
                libros_volumenes INTEGER NOT NULL DEFAULT 0,
                libros_valor REAL NOT NULL DEFAULT 0,
                revistas_titulos INTEGER NOT NULL DEFAULT 0,
                revistas_valor REAL NOT NULL DEFAULT 0,
                donaciones_titulos INTEGER NOT NULL DEFAULT 0,
                donaciones_volumenes INTEGER NOT NULL DEFAULT 0,
                donaciones_trabajos_grado INTEGER NOT NULL DEFAULT 0,
                observaciones TEXT,
                fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(año, programa)
            )
        """)
        
        # Tabla de evaluaciones de capacitaciones
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS evaluaciones_capacitaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                asistencia_id INTEGER NOT NULL,
                calidad_contenido INTEGER NOT NULL CHECK(calidad_contenido BETWEEN 1 AND 5),
                actualidad_contenidos INTEGER NOT NULL CHECK(actualidad_contenidos BETWEEN 1 AND 5),
                intensidad_horaria INTEGER NOT NULL CHECK(intensidad_horaria BETWEEN 1 AND 5),
                dominio_tema INTEGER NOT NULL CHECK(dominio_tema BETWEEN 1 AND 5),
                metodologia INTEGER NOT NULL CHECK(metodologia BETWEEN 1 AND 5),
                ayudas_didacticas INTEGER NOT NULL CHECK(ayudas_didacticas BETWEEN 1 AND 5),
                lenguaje_comprensible INTEGER NOT NULL CHECK(lenguaje_comprensible BETWEEN 1 AND 5),
                manejo_grupo INTEGER NOT NULL CHECK(manejo_grupo BETWEEN 1 AND 5),
                solucion_inquietudes INTEGER NOT NULL CHECK(solucion_inquietudes BETWEEN 1 AND 5),
                puntualidad INTEGER NOT NULL CHECK(puntualidad BETWEEN 1 AND 5),
                promedio REAL GENERATED ALWAYS AS (
                    (calidad_contenido + actualidad_contenidos + intensidad_horaria + 
                     dominio_tema + metodologia + ayudas_didacticas + lenguaje_comprensible + 
                     manejo_grupo + solucion_inquietudes + puntualidad) / 10.0
                ) STORED,
                fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (asistencia_id) REFERENCES asistencias(id),
                UNIQUE(asistencia_id)
            )
        """)
        
        conn.commit()
    print("Base de datos inicializada.")
    return mensaje_limpieza

# --- Rutas ---
@app.route("/")
def home():
    return redirect(url_for("login"))

# Login
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("username", "").strip()
        clave = request.form.get("password", "")
        
        if not usuario or not clave:
            return render_template("login.html", error="Por favor completa todos los campos")
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM usuarios WHERE username=?", (usuario,))
            user = cursor.fetchone()
            
        if user and check_password_hash(user['password'], clave):
            session["usuario"] = usuario
            return redirect(url_for("dashboard"))  # ← CAMBIO AQUÍ: redirige a dashboard
        else:
            return render_template("login.html", error="Usuario o contraseña incorrectos")
    
    return render_template("login.html")

# Dashboard/Home
@app.route("/dashboard")
def dashboard():
    if "usuario" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html")

# Registro de usuario
@app.route("/registro", methods=["GET", "POST"])
def registro():
    if request.method == "POST":
        usuario = request.form.get("username", "").strip()
        clave = request.form.get("password", "")
        
        if not usuario or not clave:
            return render_template("registro.html", error="Por favor completa todos los campos")
        
        if len(usuario) < 7:
            return render_template("registro.html", error="El usuario debe tener al menos 7 caracteres")
        
        if len(clave) < 8:
            return render_template("registro.html", error="La contraseña debe tener al menos 8 caracteres")
        
        try:
            password_hash = generate_password_hash(clave)
            
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("INSERT INTO usuarios (username, password) VALUES (?, ?)", 
                             (usuario, password_hash))
                conn.commit()
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            return render_template("registro.html", error="El usuario ya existe")
    
    return render_template("registro.html")

# Formulario de asistencia
@app.route("/formulario", methods=["GET", "POST"])
def formulario():
    es_acceso_publico = request.args.get('publico') == '1'
    
    if request.method == "POST":
        campos_requeridos = [
            "nombre_evento", "dictado_por", "docente", "programa_docente",
            "numero_identificacion", "nombre_completo", "programa_estudiante",
            "modalidad", "tipo_asistente", "sede"
        ]
        
        datos = {}
        for campo in campos_requeridos:
            valor = request.form.get(campo, "").strip()
            if not valor:
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                return render_template("formulario.html", 
                                     programas=get_programas_list(), 
                                     error=f"El campo '{campo.replace('_', ' ').title()}' es requerido",
                                     form_data=request.form,
                                     es_acceso_publico=es_acceso_publico,
                                     fecha_actual=fecha_actual)
            datos[campo] = valor
        
        try:
            # Obtener la fecha del evento del formulario (o usar fecha actual si no se trae)
            fecha_evento = request.form.get("fecha_evento", "").strip()
            if not fecha_evento:
                fecha_evento = datetime.now().strftime("%Y-%m-%d")
            
            # Validar formato de fecha
            try:
                datetime.strptime(fecha_evento, "%Y-%m-%d")
            except ValueError:
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                return render_template("formulario.html", 
                                     programas=get_programas_list(), 
                                     error="El formato de la fecha del evento no es válido",
                                     form_data=request.form,
                                     es_acceso_publico=es_acceso_publico,
                                     fecha_actual=fecha_actual)
            
            with get_db_connection() as conn:
                cursor = conn.cursor()
                # Verificar si ya existe un registro con la misma cédula, evento y fecha
                cursor.execute("""
                    SELECT COUNT(*) FROM asistencias 
                    WHERE numero_identificacion = ? 
                    AND nombre_evento = ? 
                    AND fecha_evento = ?
                """, (datos['numero_identificacion'], datos['nombre_evento'], fecha_evento))
                
                count = cursor.fetchone()[0]
                
                if count > 0:
                    # Formatear la fecha para mostrarla al usuario
                    fecha_formateada = datetime.strptime(fecha_evento, "%Y-%m-%d").strftime("%d/%m/%Y")
                    fecha_actual = datetime.now().strftime("%Y-%m-%d")
                    return render_template("formulario.html", 
                                         programas=get_programas_list(), 
                                         error=f"La identificación {datos['numero_identificacion']} ya está registrada para el evento '{datos['nombre_evento']}' el día {fecha_formateada}. No puede registrarse dos veces para el mismo evento en la misma fecha.",
                                         form_data=request.form,
                                         es_acceso_publico=es_acceso_publico,
                                         fecha_actual=fecha_actual)
                
                # Insertar el nuevo registro
                cursor.execute("""
                    INSERT INTO asistencias (
                        nombre_evento, dictado_por, docente, programa_docente,
                        numero_identificacion, nombre_completo, programa_estudiante,
                        modalidad, tipo_asistente, sede, fecha_evento
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (datos["nombre_evento"], datos["dictado_por"], datos["docente"], 
                      datos["programa_docente"], datos["numero_identificacion"], 
                      datos["nombre_completo"], datos["programa_estudiante"],
                      datos["modalidad"], datos["tipo_asistente"], datos["sede"], 
                      fecha_evento))
                conn.commit()
            
            # Obtener el ID de la asistencia recién creada
            asistencia_id = cursor.lastrowid
            
            if es_acceso_publico:
                return redirect(url_for('formulario_evaluacion', asistencia_id=asistencia_id, publico='1'))
            else:
                return redirect(url_for('formulario_evaluacion', asistencia_id=asistencia_id))
            
        except Exception as e:
            fecha_actual = datetime.now().strftime("%Y-%m-%d")
            return render_template("formulario.html", 
                                 programas=get_programas_list(), 
                                 error=f"Error al registrar asistencia: {str(e)}",
                                 form_data=request.form,
                                 es_acceso_publico=es_acceso_publico,
                                 fecha_actual=fecha_actual)
    
    # Obtener la fecha actual por defecto
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    return render_template("formulario.html", 
                          programas=get_programas_list(),
                          es_acceso_publico=es_acceso_publico,
                          fecha_actual=fecha_actual)

@app.route("/formulario/success")
def formulario_success():
    es_acceso_publico = request.args.get('publico') == '1'
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    return render_template("formulario.html", 
                          programas=get_programas_list(),
                          registro_exitoso=True,
                          es_acceso_publico=es_acceso_publico,
                          fecha_actual=fecha_actual)

@app.route("/registro-publico")
def registro_publico():
    return redirect(url_for('formulario', publico='1'))

@app.route("/qr_formulario")
def qr_formulario():
    try:
        url = request.url_root + "registro-publico"
        
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10, 
            border=4
        )
        qr.add_data(url)
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        
        buffer = io.BytesIO()
        img.save(buffer, "PNG")
        buffer.seek(0)
        
        return send_file(buffer, 
                        mimetype="image/png",
                        as_attachment=False,
                        download_name="qr_formulario.png")
    except Exception as e:
        return f"Error generando QR: {str(e)}", 500

# Función para convertir los nombres de los programas con guiones a sin guiones
# Función para convertir los nombres de los programas con guiones a sin guiones
def convertir_programas_para_vista(datos):
    programas_map = get_programas_map()
    datos_convertidos = []
    for fila in datos:
        fila_convertida = list(fila)
        
        if len(fila_convertida) > 3 and fila_convertida[3]:
            fila_convertida[3] = programas_map.get(fila_convertida[3], fila_convertida[3])
        
        if len(fila_convertida) > 6 and fila_convertida[6]:
            fila_convertida[6] = programas_map.get(fila_convertida[6], fila_convertida[6])
        
        datos_convertidos.append(tuple(fila_convertida))
    
    return datos_convertidos



@app.route("/panel")
def panel():
    global mensaje_limpieza_global
    
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    filtro = request.args.get("filtro", "").strip()
    error = request.args.get("error", "").strip()
    success = request.args.get("success", "").strip()
    
    # Si hay mensaje de limpieza pendiente, agregarlo al success
    if mensaje_limpieza_global and not success:
        success = f"⚠️ {mensaje_limpieza_global}"
        mensaje_limpieza_global = None  # Limpiar para que solo se muestre una vez
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            if filtro:
                cursor.execute("""
                    SELECT nombre_evento, dictado_por, docente, programa_docente,
                           numero_identificacion, nombre_completo, programa_estudiante,
                           modalidad, tipo_asistente, sede
                    FROM asistencias
                    WHERE nombre_evento LIKE ? 
                       OR dictado_por LIKE ?
                       OR docente LIKE ?
                       OR programa_docente LIKE ?
                       OR numero_identificacion LIKE ?
                       OR nombre_completo LIKE ?
                       OR programa_estudiante LIKE ?
                       OR modalidad LIKE ?
                       OR tipo_asistente LIKE ?
                       OR sede LIKE ?
                    ORDER BY fecha_evento DESC, nombre_evento
                """, tuple([f"%{filtro}%" for _ in range(10)]))
            else:
                cursor.execute("""
                    SELECT nombre_evento, dictado_por, docente, programa_docente,
                           numero_identificacion, nombre_completo, programa_estudiante,
                           modalidad, tipo_asistente, sede
                    FROM asistencias
                    ORDER BY fecha_evento DESC, nombre_evento
                """)
            
            datos = cursor.fetchall()
            
        datos_con_nombres_completos = convertir_programas_para_vista(datos)
        return render_template("panel.html", datos=datos_con_nombres_completos, filtro=filtro, error=error if error else None, success=success if success else None)
    
    except Exception as e:
        return render_template("panel.html", datos=[], error=f"Error cargando datos: {str(e)}")

@app.route("/panel/cargar_excel", methods=["POST"])
def panel_cargar_excel():
    """Ruta para cargar datos desde archivo Excel al panel de asistencias"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        # Verificar que se haya enviado un archivo
        if 'file' not in request.files:
            return redirect(url_for('panel', error='No se seleccionó ningún archivo'))
        
        file = request.files['file']
        
        if file.filename == '':
            return redirect(url_for('panel', error='No se seleccionó ningún archivo'))
        
        # Verificar que sea un archivo Excel
        if not file.filename.endswith(('.xlsx', '.xls')):
            return redirect(url_for('panel', error='El archivo debe ser un Excel (.xlsx o .xls)'))
        
        # Leer el archivo Excel
        df = pd.read_excel(file)
        
        # Validar que el Excel tenga las columnas necesarias
        columnas_requeridas = ['nombre_evento', 'dictado_por', 'docente', 'programa_docente',
                               'numero_identificacion', 'nombre_completo', 'programa_estudiante',
                               'modalidad', 'tipo_asistente', 'sede', 'fecha_evento']
        
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        if columnas_faltantes:
            return redirect(url_for('panel', error=f'Faltan las columnas: {", ".join(columnas_faltantes)}'))
        
        # Insertar los datos en la base de datos
        registros_insertados = 0
        registros_duplicados = 0
        errores = []
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            for index, row in df.iterrows():
                try:
                    # Intentar insertar el registro
                    cursor.execute("""
                        INSERT INTO asistencias (
                            nombre_evento, dictado_por, docente, programa_docente,
                            numero_identificacion, nombre_completo, programa_estudiante,
                            modalidad, tipo_asistente, sede, fecha_evento
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row['nombre_evento']),
                        str(row['dictado_por']),
                        str(row['docente']),
                        str(row['programa_docente']),
                        str(row['numero_identificacion']),
                        str(row['nombre_completo']),
                        str(row['programa_estudiante']),
                        str(row['modalidad']),
                        str(row['tipo_asistente']),
                        str(row['sede']),
                        str(row['fecha_evento'])
                    ))
                    registros_insertados += 1
                except sqlite3.IntegrityError as e:
                    # Registro duplicado detectado por el índice UNIQUE
                    registros_duplicados += 1
                    # Guardar información del duplicado (opcional, para debugging)
                    if registros_duplicados <= 5:  # Solo guardar los primeros 5 para no saturar
                        errores.append(f"Fila {index + 2}: {row['nombre_completo']} - {row['nombre_evento']} - {row['fecha_evento']}")
                    continue
                except Exception as e:
                    # Otros errores (datos inválidos, etc.)
                    print(f"Error insertando fila {index + 2}: {str(e)}")
                    errores.append(f"Fila {index + 2}: Error - {str(e)}")
                    continue
            
            conn.commit()
        
        # Preparar mensaje de éxito
        mensaje = f'Se insertaron {registros_insertados} registros correctamente.'
        if registros_duplicados > 0:
            mensaje += f' ⚠️ Se omitieron {registros_duplicados} registros duplicados.'
        if len(errores) > 0 and registros_duplicados <= 5:
            mensaje += f' Duplicados: {", ".join(errores)}'
        
        return redirect(url_for('panel', success=mensaje))
        
    except Exception as e:
        print(f"Error cargando Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return redirect(url_for('panel', error=f'Error procesando el archivo: {str(e)}'))

@app.route("/exportar")
def exportar():
    if "usuario" not in session:
        return redirect(url_for("login"))

    try:
        filtros = [request.args.get(f"col{i}", "").strip() for i in range(10)]
        columnas = ["nombre_evento","dictado_por","docente","programa_docente","numero_identificacion",
                    "nombre_completo","programa_estudiante","modalidad","tipo_asistente","sede"]

        query = "SELECT * FROM asistencias"
        condiciones = []
        params = []

        for i, val in enumerate(filtros):
            if val:
                condiciones.append(f"{columnas[i]} LIKE ?")
                params.append(f"%{val}%")

        if condiciones:
            query += " WHERE " + " AND ".join(condiciones)
        
        query += " ORDER BY fecha_evento DESC, nombre_evento"

        with get_db_connection() as conn:
            df = pd.read_sql_query(query, conn, params=params)

        if df.empty:
            return "No hay datos para exportar", 400

        programas_map = get_programas_map()
        
        def convertir_programa(nombre_programa):
            return programas_map.get(nombre_programa, nombre_programa)

        if 'programa_docente' in df.columns:
            df['programa_docente'] = df['programa_docente'].apply(convertir_programa)
        
        if 'programa_estudiante' in df.columns:
            df['programa_estudiante'] = df['programa_estudiante'].apply(convertir_programa)

        if 'id' in df.columns:
            df = df.drop('id', axis=1)

        df.rename(columns={
            'nombre_evento': 'Nombre del Evento',
            'dictado_por': 'Dictado Por',
            'docente': 'Docente Acompañante',
            'programa_docente': 'Programa del Docente',
            'numero_identificacion': 'Número de Identificación',
            'nombre_completo': 'Nombre Completo del Estudiante',
            'programa_estudiante': 'Programa del Estudiante',
            'modalidad': 'Modalidad',
            'tipo_asistente': 'Tipo de Asistente',
            'sede': 'Sede',
            'fecha_evento': 'Fecha del Evento',
            'fecha_registro': 'Fecha de Registro'
        }, inplace=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Asistencias")

        output.seek(0)

        return send_file(output,
                         as_attachment=True,
                         download_name="asistencias.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return f"Error exportando datos: {str(e)}", 500

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ========================================
# MÓDULO DE INVERSIONES
# ========================================

@app.route("/inversiones")
def inversiones():
    """Página principal del módulo de inversiones con selector de sub-módulos"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    return render_template("inversiones_home.html")

# --- Sub-módulo 1: Inversiones Institucionales ---
@app.route("/inversiones/institucional")
def inversiones_institucional():
    """Panel de inversiones institucionales"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT año, monto_libros, monto_revistas, monto_bases_datos, total, observaciones
                FROM inversiones_institucionales
                ORDER BY año DESC
            """)
            datos = cursor.fetchall()
        
        return render_template("inversiones_institucional.html", datos=datos)
    except Exception as e:
        return render_template("inversiones_institucional.html", datos=[], error=f"Error cargando datos: {str(e)}")

@app.route("/inversiones/institucional/registrar", methods=["GET", "POST"])
def inversiones_institucional_registrar():
    """Formulario para registrar inversión institucional"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    if request.method == "POST":
        try:
            año = request.form.get("año", "").strip()
            monto_libros = request.form.get("monto_libros", "0").strip()
            monto_revistas = request.form.get("monto_revistas", "0").strip()
            monto_bases_datos = request.form.get("monto_bases_datos", "0").strip()
            observaciones = request.form.get("observaciones", "").strip()
            
            # Validaciones
            if not año:
                return render_template("inversiones_institucional_form.html", 
                                     error="El año es requerido", 
                                     form_data=request.form)
            
            # Convertir a números
            año = int(año)
            monto_libros = float(monto_libros.replace(",", ""))
            monto_revistas = float(monto_revistas.replace(",", ""))
            monto_bases_datos = float(monto_bases_datos.replace(",", ""))
            
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO inversiones_institucionales 
                    (año, monto_libros, monto_revistas, monto_bases_datos, observaciones)
                    VALUES (?, ?, ?, ?, ?)
                """, (año, monto_libros, monto_revistas, monto_bases_datos, observaciones))
                conn.commit()
            
            return redirect(url_for("inversiones_institucional"))
            
        except sqlite3.IntegrityError:
            return render_template("inversiones_institucional_form.html",
                                 error=f"Ya existe un registro para el año {año}",
                                 form_data=request.form)
        except Exception as e:
            return render_template("inversiones_institucional_form.html",
                                 error=f"Error al registrar: {str(e)}",
                                 form_data=request.form)
    
    return render_template("inversiones_institucional_form.html")

# --- Sub-módulo 2: Inversiones por Programa ---
@app.route("/inversiones/programas")
def inversiones_programas():
    """Panel de inversiones por programa"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT año, programa, 
                       libros_titulos, libros_volumenes, libros_valor,
                       revistas_titulos, revistas_valor,
                       donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                       observaciones
                FROM inversiones_programas
                ORDER BY año DESC, programa
            """)
            datos = cursor.fetchall()
        
        datos_con_nombres = convertir_programas_para_vista(datos)
        return render_template("inversiones_programas.html", datos=datos_con_nombres)
    except Exception as e:
        return render_template("inversiones_programas.html", datos=[], error=f"Error cargando datos: {str(e)}")

@app.route("/inversiones/programas/registrar", methods=["GET", "POST"])
def inversiones_programas_registrar():
    """Formulario para registrar inversión por programa"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    if request.method == "POST":
        try:
            año = request.form.get("año", "").strip()
            programa = request.form.get("programa", "").strip()
            
            # Libros
            libros_titulos = int(request.form.get("libros_titulos", "0").strip())
            libros_volumenes = int(request.form.get("libros_volumenes", "0").strip())
            libros_valor = float(request.form.get("libros_valor", "0").strip().replace(",", ""))
            
            # Revistas
            revistas_titulos = int(request.form.get("revistas_titulos", "0").strip())
            revistas_valor = float(request.form.get("revistas_valor", "0").strip().replace(",", ""))
            
            # Donaciones
            donaciones_titulos = int(request.form.get("donaciones_titulos", "0").strip())
            donaciones_volumenes = int(request.form.get("donaciones_volumenes", "0").strip())
            donaciones_trabajos_grado = int(request.form.get("donaciones_trabajos_grado", "0").strip())
            
            observaciones = request.form.get("observaciones", "").strip()
            
            # Validaciones
            if not año or not programa:
                return render_template("inversiones_programas_form.html",
                                     programas=get_programas_list(),
                                     error="El año y el programa son requeridos",
                                     form_data=request.form)
            
            año = int(año)
            
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO inversiones_programas 
                    (año, programa, libros_titulos, libros_volumenes, libros_valor,
                     revistas_titulos, revistas_valor,
                     donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                     observaciones)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (año, programa, libros_titulos, libros_volumenes, libros_valor,
                      revistas_titulos, revistas_valor,
                      donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                      observaciones))
                conn.commit()
            
            return redirect(url_for("inversiones_programas"))
            
        except sqlite3.IntegrityError:
            return render_template("inversiones_programas_form.html",
                                 programas=get_programas_list(),
                                 error=f"Ya existe un registro para {programa} en el año {año}",
                                 form_data=request.form)
        except Exception as e:
            return render_template("inversiones_programas_form.html",
                                 programas=get_programas_list(),
                                 error=f"Error al registrar: {str(e)}",
                                 form_data=request.form)
    
    return render_template("inversiones_programas_form.html", programas=get_programas_list())

# Agregar esta función al app.py después de la línea 699

@app.route("/estadisticas")
def estadisticas():
    if "usuario" not in session:
        return redirect(url_for("login"))

    try:
        # ========== NUEVA FUNCIONALIDAD: VENTANA DE 5 AÑOS ==========
        # Obtener la ventana de años (últimos 5 años)
        anos_ventana, ano_inicio_ventana, ano_fin_ventana = get_ventana_anos(5)
        
        # Obtener filtros de la URL
        evento_filtro = request.args.get('evento', '')
        programa_filtro = request.args.get('programa', '')
        fecha_inicio = request.args.get('fecha_inicio', '')
        fecha_fin = request.args.get('fecha_fin', '')
        
        with get_db_connection() as conn:
            # Query base con filtros opcionales
            where_clauses = []
            params = []
            
            # ========== FILTRO AUTOMÁTICO POR VENTANA DE AÑOS ==========
            # Siempre filtrar por la ventana de 5 años, a menos que el usuario especifique fechas
            if not fecha_inicio and not fecha_fin:
                # Aplicar filtro de ventana de años automáticamente
                where_clauses.append("CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) >= ?")
                params.append(ano_inicio_ventana)
                where_clauses.append("CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) <= ?")
                params.append(ano_fin_ventana)
            
            if evento_filtro:
                where_clauses.append("nombre_evento = ?")
                params.append(evento_filtro)
            if programa_filtro:
                where_clauses.append("programa_estudiante = ?")
                params.append(programa_filtro)
            if fecha_inicio:
                where_clauses.append("fecha_evento >= ?")
                params.append(fecha_inicio)
            if fecha_fin:
                where_clauses.append("fecha_evento <= ?")
                params.append(fecha_fin)
            
            where_sql = " WHERE " + " AND ".join(where_clauses) if where_clauses else ""
            
            # 1. Total de asistencias
            df_total = pd.read_sql_query(f"""
                SELECT COUNT(*) as total_asistencias
                FROM asistencias {where_sql}
            """, conn, params=params)
            
            # 2. Total de eventos únicos
            df_eventos_unicos = pd.read_sql_query(f"""
                SELECT COUNT(DISTINCT nombre_evento) as total_eventos
                FROM asistencias {where_sql}
            """, conn, params=params)
            
            # 3. Total de programas únicos
            df_programas_unicos = pd.read_sql_query(f"""
                SELECT COUNT(DISTINCT programa_estudiante) as total_programas
                FROM asistencias {where_sql}
            """, conn, params=params)
            
            # 4. Promedio de evaluaciones
            df_promedio_eval = pd.read_sql_query(f"""
                SELECT AVG(promedio) as promedio_general
                FROM evaluaciones_capacitaciones e
                INNER JOIN asistencias a ON e.asistencia_id = a.id
                {where_sql}
            """, conn, params=params)
            
            # 5. Datos de eventos (asistencias por evento)
            df_eventos = pd.read_sql_query(f"""
                SELECT nombre_evento, COUNT(*) as total_asistencias
                FROM asistencias {where_sql}
                GROUP BY nombre_evento 
                ORDER BY total_asistencias DESC
                LIMIT 15
            """, conn, params=params)
            
            # 6. Datos de programas (asistencias por programa)
            df_programas = pd.read_sql_query(f"""
                SELECT programa_estudiante, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY programa_estudiante 
                ORDER BY total DESC
                LIMIT 15
            """, conn, params=params)
            
            # 7. **NUEVO: Análisis cruzado Programa x Evento**
            df_cruzado = pd.read_sql_query(f"""
                SELECT 
                    nombre_evento,
                    programa_estudiante,
                    COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY nombre_evento, programa_estudiante
                ORDER BY nombre_evento, total DESC
            """, conn, params=params)
            
            # 8. Tendencia mensual
            df_mensual = pd.read_sql_query(f"""
                SELECT 
                    SUBSTR(fecha_evento, 1, 7) as mes,
                    COUNT(*) as total
                FROM asistencias
                WHERE fecha_evento IS NOT NULL 
                    AND fecha_evento != ''
                    {' AND ' + ' AND '.join(where_clauses) if where_clauses else ''}
                GROUP BY mes
                ORDER BY mes
            """, conn, params=params if where_clauses else [])
            
            # 9. Top 5 programas por evento
            query_top = f"""
                SELECT * FROM (
                    SELECT 
                        nombre_evento,
                        programa_estudiante,
                        COUNT(*) as total,
                        ROW_NUMBER() OVER (PARTITION BY nombre_evento ORDER BY COUNT(*) DESC) as ranking
                    FROM asistencias {where_sql}
                    GROUP BY nombre_evento, programa_estudiante
                )
                WHERE ranking <= 5
                ORDER BY nombre_evento, ranking
            """
            df_top_por_evento = pd.read_sql_query(query_top, conn, params=params)
            
            # 10. Obtener listas únicas para filtros
            df_eventos_lista = pd.read_sql_query("SELECT DISTINCT nombre_evento FROM asistencias ORDER BY nombre_evento", conn)
            # Obtener programas únicos directamente de asistencias
            df_programas_lista = pd.read_sql_query("SELECT DISTINCT programa_estudiante as nombre FROM asistencias ORDER BY programa_estudiante", conn)
            
            # 11. Tipo de asistentes
            df_tipo_asistente = pd.read_sql_query(f"""
                SELECT tipo_asistente, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY tipo_asistente
                ORDER BY total DESC
            """, conn, params=params)
            
            # 12. Modalidad
            df_modalidad = pd.read_sql_query(f"""
                SELECT modalidad, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY modalidad
                ORDER BY total DESC
            """, conn, params=params)

        # Verificar si hay datos
        if df_eventos.empty:
            return render_template("estadisticas_avanzadas.html", 
                                 mensaje="No hay datos disponibles para mostrar estadísticas",
                                 ventana_anos={
                                     'anos': anos_ventana,
                                     'ano_inicio': ano_inicio_ventana,
                                     'ano_fin': ano_fin_ventana
                                 },
                                 filtros={
                                     'evento': evento_filtro,
                                     'programa': programa_filtro,
                                     'fecha_inicio': fecha_inicio,
                                     'fecha_fin': fecha_fin,
                                     'eventos_lista': [],
                                     'programas_lista': []
                                 })

        # Procesar datos para el template
        total_asistencias = int(df_total['total_asistencias'].iloc[0]) if not df_total.empty else 0
        total_eventos = int(df_eventos_unicos['total_eventos'].iloc[0]) if not df_eventos_unicos.empty else 0
        total_programas = int(df_programas_unicos['total_programas'].iloc[0]) if not df_programas_unicos.empty else 0
        promedio_evaluaciones = float(df_promedio_eval['promedio_general'].iloc[0]) if not df_promedio_eval.empty and pd.notna(df_promedio_eval['promedio_general'].iloc[0]) else 0
        
        eventos_labels = df_eventos['nombre_evento'].tolist()
        eventos_valores = [int(x) for x in df_eventos['total_asistencias'].tolist()]
        
        programa_labels = df_programas['programa_estudiante'].tolist()
        programa_valores = [int(x) for x in df_programas['total'].tolist()]
        
        # Procesar datos cruzados para matriz
        matriz_cruzada = {}
        for _, row in df_cruzado.iterrows():
            evento = row['nombre_evento']
            programa = row['programa_estudiante']
            total = int(row['total'])
            
            if evento not in matriz_cruzada:
                matriz_cruzada[evento] = {}
            matriz_cruzada[evento][programa] = total
        
        # Procesar tendencia mensual
        meses_labels = df_mensual['mes'].tolist() if not df_mensual.empty else []
        meses_valores = [int(x) for x in df_mensual['total'].tolist()] if not df_mensual.empty else []
        
        # Procesar top por evento (solo top 5)
        top_por_evento = {}
        for _, row in df_top_por_evento.iterrows():
            evento = row['nombre_evento']
            if evento not in top_por_evento:
                top_por_evento[evento] = []
            top_por_evento[evento].append({
                'programa': row['programa_estudiante'],
                'total': int(row['total']),
                'ranking': int(row['ranking'])
            })
        
        # Datos de tipo de asistente
        tipo_asistente_labels = df_tipo_asistente['tipo_asistente'].tolist()
        tipo_asistente_valores = [int(x) for x in df_tipo_asistente['total'].tolist()]
        
        # Datos de modalidad
        modalidad_labels = df_modalidad['modalidad'].tolist()
        modalidad_valores = [int(x) for x in df_modalidad['total'].tolist()]
        
        return render_template("estadisticas_avanzadas.html",
                             # Resumen general
                             total_asistencias=total_asistencias,
                             total_eventos=total_eventos,
                             total_programas=total_programas,
                             promedio_evaluaciones=round(promedio_evaluaciones, 2),
                             
                             # Gráficas principales
                             eventos_labels=eventos_labels,
                             eventos_valores=eventos_valores,
                             programa_labels=programa_labels,
                             programa_valores=programa_valores,
                             
                             # Análisis cruzado
                             matriz_cruzada=matriz_cruzada,
                             top_por_evento=top_por_evento,
                             
                             # Tendencias
                             meses_labels=meses_labels,
                             meses_valores=meses_valores,
                             
                             # Distribuciones
                             tipo_asistente_labels=tipo_asistente_labels,
                             tipo_asistente_valores=tipo_asistente_valores,
                             modalidad_labels=modalidad_labels,
                             modalidad_valores=modalidad_valores,
                             
                             # Ventana de años
                             ventana_anos={
                                 'anos': anos_ventana,
                                 'ano_inicio': ano_inicio_ventana,
                                 'ano_fin': ano_fin_ventana
                             },
                             
                             # Filtros
                             filtros={
                                 'evento': evento_filtro,
                                 'programa': programa_filtro,
                                 'fecha_inicio': fecha_inicio,
                                 'fecha_fin': fecha_fin,
                                 'eventos_lista': df_eventos_lista['nombre_evento'].tolist(),
                                 'programas_lista': df_programas_lista['nombre'].tolist()
                             })
    
    except Exception as e:
        print(f"Error en estadísticas: {str(e)}")
        import traceback
        traceback.print_exc()
        anos_ventana, ano_inicio_ventana, ano_fin_ventana = get_ventana_anos(5)
        return render_template("estadisticas_avanzadas.html", 
                             error=f"Error cargando estadísticas: {str(e)}",
                             ventana_anos={
                                 'anos': anos_ventana,
                                 'ano_inicio': ano_inicio_ventana,
                                 'ano_fin': ano_fin_ventana
                             },
                             filtros={
                                 'evento': '',
                                 'programa': '',
                                 'fecha_inicio': '',
                                 'fecha_fin': '',
                                 'eventos_lista': [],
                                 'programas_lista': []
                             })

# Nueva ruta para el formulario de evaluación
@app.route("/formulario/evaluacion/<int:asistencia_id>", methods=["GET", "POST"])
def formulario_evaluacion(asistencia_id):
    es_acceso_publico = request.args.get('publico') == '1'
    
    if request.method == "POST":
        try:
            # Validar que todos los campos de evaluación estén presentes
            campos_evaluacion = [
                "calidad_contenido", "actualidad_contenidos", "intensidad_horaria",
                "dominio_tema", "metodologia", "ayudas_didacticas",
                "lenguaje_comprensible", "manejo_grupo", "solucion_inquietudes",
                "puntualidad"
            ]
            
            evaluacion = {}
            for campo in campos_evaluacion:
                valor = request.form.get(campo, "").strip()
                if not valor or not valor.isdigit() or int(valor) < 1 or int(valor) > 5:
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("SELECT * FROM asistencias WHERE id = ?", (asistencia_id,))
                        asistencia = cursor.fetchone()
                    
                    return render_template("formulario_evaluacion.html",
                                         asistencia=asistencia,
                                         error=f"El campo '{campo.replace('_', ' ').title()}' debe ser un valor entre 1 y 5",
                                         form_data=request.form,
                                         es_acceso_publico=es_acceso_publico)
                evaluacion[campo] = int(valor)
            
            # Insertar la evaluación
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO evaluaciones_capacitaciones (
                        asistencia_id, calidad_contenido, actualidad_contenidos,
                        intensidad_horaria, dominio_tema, metodologia, ayudas_didacticas,
                        lenguaje_comprensible, manejo_grupo, solucion_inquietudes, puntualidad
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (asistencia_id, evaluacion["calidad_contenido"],
                      evaluacion["actualidad_contenidos"], evaluacion["intensidad_horaria"],
                      evaluacion["dominio_tema"], evaluacion["metodologia"],
                      evaluacion["ayudas_didacticas"], evaluacion["lenguaje_comprensible"],
                      evaluacion["manejo_grupo"], evaluacion["solucion_inquietudes"],
                      evaluacion["puntualidad"]))
                conn.commit()
            
            # Redirigir a página de éxito
            if es_acceso_publico:
                return redirect(url_for('evaluacion_success', publico='1'))
            else:
                return redirect(url_for('evaluacion_success'))
                
        except sqlite3.IntegrityError:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM asistencias WHERE id = ?", (asistencia_id,))
                asistencia = cursor.fetchone()
            
            return render_template("formulario_evaluacion.html",
                                 asistencia=asistencia,
                                 error="Ya existe una evaluación para esta asistencia",
                                 form_data=request.form,
                                 es_acceso_publico=es_acceso_publico)
        except Exception as e:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM asistencias WHERE id = ?", (asistencia_id,))
                asistencia = cursor.fetchone()
            
            return render_template("formulario_evaluacion.html",
                                 asistencia=asistencia,
                                 error=f"Error al registrar evaluación: {str(e)}",
                                 form_data=request.form,
                                 es_acceso_publico=es_acceso_publico)
    
    # GET request
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM asistencias WHERE id = ?", (asistencia_id,))
            asistencia = cursor.fetchone()
            
            if not asistencia:
                return redirect(url_for('formulario'))
        
        return render_template("formulario_evaluacion.html",
                             asistencia=asistencia,
                             es_acceso_publico=es_acceso_publico)
    except Exception as e:
        return redirect(url_for('formulario'))

@app.route("/evaluacion/success")
def evaluacion_success():
    es_acceso_publico = request.args.get('publico') == '1'
    return render_template("evaluacion_success.html",
                          es_acceso_publico=es_acceso_publico)

# --- Rutas para gestión de programas académicos ---
@app.route("/programas")
def gestion_programas():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, nombre, activo, 
                       fecha_creacion, fecha_modificacion
                FROM programas
                ORDER BY nombre ASC
            """)
            programas = cursor.fetchall()
        
        return render_template("gestion_programas.html", programas=programas)
    except Exception as e:
        return render_template("gestion_programas.html", 
                             error=f"Error al cargar programas: {str(e)}",
                             programas=[])

@app.route("/api/programas/activos")
def get_programas_activos():
    """API para obtener solo programas activos (para el formulario)"""
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT nombre
                FROM programas
                WHERE activo = 1
                ORDER BY nombre ASC
            """)
            programas = cursor.fetchall()
        
        return {
            "success": True,
            "programas": [{"value": p["nombre"], "label": p["nombre"]} for p in programas]
        }
    except Exception as e:
        return {"success": False, "error": str(e)}, 500

@app.route("/api/programas/toggle/<int:programa_id>", methods=["POST"])
def toggle_programa(programa_id):
    """Habilitar o deshabilitar un programa"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            # Obtener estado actual
            cursor.execute("SELECT activo FROM programas WHERE id = ?", (programa_id,))
            programa = cursor.fetchone()
            
            if not programa:
                return {"success": False, "error": "Programa no encontrado"}, 404
            
            # Cambiar estado
            nuevo_estado = 0 if programa["activo"] == 1 else 1
            cursor.execute("""
                UPDATE programas 
                SET activo = ?,
                    fecha_modificacion = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (nuevo_estado, programa_id))
            conn.commit()
        
        return {
            "success": True,
            "nuevo_estado": nuevo_estado,
            "mensaje": "Programa habilitado" if nuevo_estado == 1 else "Programa deshabilitado"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}, 500

@app.route("/api/programas/agregar", methods=["POST"])
def agregar_programa():
    """Agregar un nuevo programa académico"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        data = request.get_json()
        nombre = data.get("nombre", "").strip()
        
        if not nombre:
            return {"success": False, "error": "El nombre del programa es requerido"}, 400
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO programas (nombre, activo)
                VALUES (?, 1)
            """, (nombre,))
            conn.commit()
            programa_id = cursor.lastrowid
        
        return {
            "success": True,
            "mensaje": "Programa agregado exitosamente",
            "programa_id": programa_id
        }
    except sqlite3.IntegrityError:
        return {"success": False, "error": "Ya existe un programa con ese nombre"}, 400
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
        return {"success": False, "error": str(e)}, 500

@app.route("/api/programas/eliminar/<int:programa_id>", methods=["DELETE"])
def eliminar_programa(programa_id):
    """Eliminar un programa académico (solo si no tiene registros asociados)"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            
            # Verificar si tiene asistencias registradas
            cursor.execute("""
                SELECT COUNT(*) as count FROM asistencias 
                WHERE programa_estudiante IN (
                    SELECT nombre FROM programas WHERE id = ?
                ) OR programa_docente IN (
                    SELECT nombre FROM programas WHERE id = ?
                )
            """, (programa_id, programa_id))
            
            count = cursor.fetchone()["count"]
            
            if count > 0:
                return {
                    "success": False, 
                    "error": f"No se puede eliminar. El programa tiene {count} registro(s) asociado(s). Considere deshabilitarlo en su lugar."
                }, 400
            
            # Eliminar programa
            cursor.execute("DELETE FROM programas WHERE id = ?", (programa_id,))
            conn.commit()
        
        return {
            "success": True,
            "mensaje": "Programa eliminado exitosamente"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}, 500




@app.route("/admin/limpiar_datos_antiguos")
def limpiar_datos_route():
    """
    Ruta administrativa para limpiar datos antiguos fuera de la ventana de 5 años.
    Solo accesible por usuarios autenticados.
    """
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        resultado = limpiar_datos_antiguos(5)
        
        if resultado['success']:
            return render_template("panel.html",
                                 success=resultado['mensaje'],
                                 datos=[])
        else:
            return render_template("panel.html",
                                 error=resultado['mensaje'],
                                 datos=[])
    except Exception as e:
        return render_template("panel.html",
                             error=f"Error al limpiar datos: {str(e)}",
                             datos=[])


# ============================================
# MANEJADORES DE ERRORES MEJORADOS
# ============================================

@app.errorhandler(404)
def error_404(e):
    """Página no encontrada"""
    return render_template('error.html',
        error_code='Error 404',
        error_title='Página no encontrada',
        error_message='Lo sentimos, la página que buscas no existe o ha sido movida a otra ubicación.',
        error_details='La ruta solicitada no está disponible'
    ), 404


@app.errorhandler(403)
def error_403(e):
    """Acceso prohibido"""
    return render_template('error.html',
        error_code='Error 403',
        error_title='Acceso denegado',
        error_message='No tienes los permisos necesarios para acceder a esta sección.',
        error_details='Se requiere autenticación válida'
    ), 403


@app.errorhandler(401)
def error_401(e):
    """No autorizado"""
    return render_template('error.html',
        error_code='Error 401',
        error_title='Sesión no válida',
        error_message='Tu sesión ha expirado o no has iniciado sesión. Por favor, inicia sesión nuevamente.',
        error_details='Autenticación requerida'
    ), 401


@app.errorhandler(500)
def error_500(e):
    """Error interno del servidor"""
    return render_template('error.html',
        error_code='Error 500',
        error_title='Error interno del servidor',
        error_message='Ha ocurrido un error inesperado. Nuestro equipo ha sido notificado y está trabajando en solucionarlo.',
        error_details='Error interno del sistema' if not app.debug else str(e)
    ), 500


@app.errorhandler(405)
def error_405(e):
    """Método no permitido"""
    return render_template('error.html',
        error_code='Error 405',
        error_title='Método no permitido',
        error_message='El método HTTP usado no está permitido para esta ruta.',
        error_details='Verifica que estés usando el método correcto (GET, POST, etc.)'
    ), 405


@app.errorhandler(400)
def error_400(e):
    """Solicitud incorrecta"""
    return render_template('error.html',
        error_code='Error 400',
        error_title='Solicitud incorrecta',
        error_message='Los datos enviados no son válidos. Por favor verifica la información.',
        error_details=str(e) if app.debug else 'Datos de solicitud inválidos'
    ), 400


if __name__ == "__main__":
    mensaje_limpieza_duplicados = init_db()
    if mensaje_limpieza_duplicados:
        mensaje_limpieza_global = mensaje_limpieza_duplicados
        print(f"ℹ️ IMPORTANTE: {mensaje_limpieza_duplicados}")
    app.run(debug=True)