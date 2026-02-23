from flask import Flask, flash, render_template, request, redirect, url_for, session, send_file
import pandas as pd
import io
import qrcode
import matplotlib.pyplot as plt
import os
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import unicodedata

# --- CONFIGURACIÓN DUAL: SQLite (desarrollo) y PostgreSQL (producción) ---
DATABASE_URL = os.environ.get('DATABASE_URL')  # Render proporciona esta variable

if DATABASE_URL:
    # PostgreSQL en producción
    import psycopg2
    from psycopg2.extras import RealDictCursor
    USE_POSTGRES = True
else:
    # SQLite en desarrollo
    import sqlite3
    USE_POSTGRES = False

#Configuración de la app flask
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'clave_secreta_por_defecto')

# Variable global para almacenar mensaje de limpieza de duplicados
mensaje_limpieza_global = None

# --- Funciones auxiliares para base de datos dual ---
def get_db_connection():
    """Función auxiliar para obtener conexión a la base de datos (SQLite o PostgreSQL)"""
    if USE_POSTGRES:
        # PostgreSQL en producción
        # IMPORTANTE: Render puede dar URL con postgres:// que debe ser postgresql://
        db_url = DATABASE_URL
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        
        conn = psycopg2.connect(db_url, sslmode='require')
        return conn
    else:
        # SQLite en desarrollo
        conn = sqlite3.connect("biblioteca.db")
        conn.row_factory = sqlite3.Row
        return conn

def get_cursor(conn):
    """Obtener cursor con el formato correcto según la base de datos"""
    if USE_POSTGRES:
        return conn.cursor(cursor_factory=RealDictCursor)
    else:
        return conn.cursor()

def adapt_query(query):
    """Adapta la query según la base de datos"""
    if USE_POSTGRES:
        # Reemplazar ? con %s para PostgreSQL
        query = query.replace('?', '%s')
        # Reemplazar SUBSTR con SUBSTRING
        query = query.replace('SUBSTR(', 'SUBSTRING(')
    return query

def normalize_text(text):
    """
    Normaliza un texto removiendo acentos y convirtiendo a minúsculas.
    Usado para búsquedas insensibles a mayúsculas y tildes.
    
    Ejemplo: "Ingeniería de Sistemas" -> "ingenieria de sistemas"
    """
    if not text:
        return ''
    # Convertir a minúsculas
    text = str(text).lower()
    # Remover acentos usando unicodedata
    # NFD = Canonical Decomposition (separa letras de sus acentos)
    text = unicodedata.normalize('NFD', text)
    # Filtrar solo caracteres ASCII (elimina las marcas diacríticas)
    text = text.encode('ascii', 'ignore').decode('utf-8')
    return text

def get_programas_list():
    """Obtener lista de todos los programas desde la base de datos"""
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("SELECT nombre FROM programas WHERE activo = 1 ORDER BY nombre ASC")
            cursor.execute(query)
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

def asistencia_to_tuple(asistencia):
    """
    Convierte un registro de asistencia (dict o Row) a tupla
    Para compatibilidad con plantillas que acceden por índice
    """
    if not asistencia:
        return None
    
    if isinstance(asistencia, dict):
        # PostgreSQL con RealDictCursor - extraer en orden
        # Convertir datetime a string si es necesario
        fecha_evento = asistencia.get('fecha_evento', '')
        if fecha_evento and hasattr(fecha_evento, 'strftime'):
            fecha_evento = fecha_evento.strftime('%Y-%m-%d')
        
        fecha_registro = asistencia.get('fecha_registro', '')
        if fecha_registro and hasattr(fecha_registro, 'strftime'):
            fecha_registro = fecha_registro.strftime('%Y-%m-%d %H:%M:%S')
        
        return (
            asistencia.get('id', ''),
            asistencia.get('nombre_evento', ''),
            asistencia.get('dictado_por', ''),
            asistencia.get('docente', ''),
            asistencia.get('programa_docente', ''),
            asistencia.get('numero_identificacion', ''),
            asistencia.get('nombre_completo', ''),
            asistencia.get('programa_estudiante', ''),
            asistencia.get('modalidad', ''),
            asistencia.get('tipo_asistente', ''),
            asistencia.get('sede', ''),
            fecha_evento,
            asistencia.get('hora_inicio', ''),
            asistencia.get('hora_fin', ''),
            fecha_registro
        )
    else:
        # SQLite - ya es tupla o Row compatible
        return tuple(asistencia)

def programas_to_tuples(programas):
    """
    Convierte una lista de programas (dict o Row) a lista de tuplas
    Para compatibilidad con plantillas que acceden por índice
    """
    if not programas:
        return []
    
    result = []
    for programa in programas:
        if isinstance(programa, dict):
            # PostgreSQL con RealDictCursor - extraer en orden
            # Convertir datetime a string para compatibilidad
            fecha_creacion = programa.get('fecha_creacion', '')
            if fecha_creacion and hasattr(fecha_creacion, 'strftime'):
                fecha_creacion = fecha_creacion.strftime('%Y-%m-%d %H:%M:%S')
            
            fecha_modificacion = programa.get('fecha_modificacion', '')
            if fecha_modificacion and hasattr(fecha_modificacion, 'strftime'):
                fecha_modificacion = fecha_modificacion.strftime('%Y-%m-%d %H:%M:%S')
            
            result.append((
                programa.get('id', ''),
                programa.get('nombre', ''),
                programa.get('activo', 0),
                fecha_creacion,
                fecha_modificacion
            ))
        else:
            # SQLite - ya es tupla o Row compatible
            result.append(tuple(programa))
    
    return result

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
            cursor = get_cursor(conn)
            
            # Contar registros a eliminar
            query_count = adapt_query("""
                SELECT COUNT(*) as total
                FROM asistencias
                WHERE CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) < ?
            """)
            cursor.execute(query_count, (ano_inicio,))
            
            result = cursor.fetchone()
            total_a_eliminar = result['total']
            
            if total_a_eliminar > 0:
                # Eliminar registros antiguos
                query_delete = adapt_query("""
                    DELETE FROM asistencias
                    WHERE CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) < ?
                """)
                cursor.execute(query_delete, (ano_inicio,))
                
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
    
    # Definir tipos de datos según la base de datos
    if DATABASE_URL:
        # PostgreSQL
        SERIAL = "SERIAL"
        AUTOINCREMENT = ""
        INTEGER = "INTEGER"
        TEXT = "TEXT"
        TIMESTAMP_DEFAULT = "TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
    else:
        # SQLite
        SERIAL = "INTEGER"
        AUTOINCREMENT = "AUTOINCREMENT"
        INTEGER = "INTEGER"
        TEXT = "TEXT"
        TIMESTAMP_DEFAULT = "DATETIME DEFAULT CURRENT_TIMESTAMP"
    
    
    # CRÍTICO: Usar conexión explícita, no context manager
    conn = get_db_connection()
    cursor = get_cursor(conn)
    
    try:
        # Tabla de usuarios
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS usuarios (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                username {TEXT} UNIQUE NOT NULL,
                password {TEXT} NOT NULL,
                created_at {TIMESTAMP_DEFAULT}
            )
        """)

        # Tabla de programas académicos
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS programas (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                nombre {TEXT} UNIQUE NOT NULL,
                activo {INTEGER} DEFAULT 1,
                fecha_creacion {TIMESTAMP_DEFAULT},
                fecha_modificacion {TIMESTAMP_DEFAULT}
            )
        """)

        # Tabla de modalidades
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS modalidades (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                nombre {TEXT} NOT NULL UNIQUE,
                activo {INTEGER} DEFAULT 1,
                fecha_creacion {TIMESTAMP_DEFAULT},
                fecha_modificacion {TIMESTAMP_DEFAULT}
            )
        """)
        
        # Insertar modalidades por defecto (solo si no existen)
        modalidades_default = ['Presencial', 'A Distancia', 'Virtual']
        for modalidad in modalidades_default:
            try:
                if USE_POSTGRES:
                    cursor.execute("""
                        INSERT INTO modalidades (nombre, activo) 
                        VALUES (%s, 1)
                        ON CONFLICT (nombre) DO NOTHING
                    """, (modalidad,))
                else:
                    cursor.execute("""
                        INSERT OR IGNORE INTO modalidades (nombre, activo) 
                        VALUES (?, 1)
                    """, (modalidad,))
            except:
                pass
        

        # Tabla de asistencias (capacitaciones)
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS asistencias (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                nombre_evento {TEXT} NOT NULL,
                dictado_por {TEXT} NOT NULL,
                docente {TEXT} NOT NULL,
                programa_docente {TEXT} NOT NULL,
                numero_identificacion {TEXT} NOT NULL,
                nombre_completo {TEXT} NOT NULL,
                programa_estudiante {TEXT} NOT NULL,
                modalidad {TEXT} NOT NULL,
                tipo_asistente {TEXT} NOT NULL,
                sede {TEXT} NOT NULL,
                fecha_evento {TEXT} NOT NULL,
                hora_inicio {TEXT},
                hora_fin {TEXT},
                fecha_registro {TIMESTAMP_DEFAULT}
            )
        """)
        
        # Verificar columna fecha_evento en asistencias (solo para SQLite)
        if not USE_POSTGRES:
            cursor.execute("PRAGMA table_info(asistencias)")
            columnas = [columna[1] for columna in cursor.fetchall()]
            
            if 'fecha_evento' not in columnas:
                cursor.execute("ALTER TABLE asistencias ADD COLUMN fecha_evento TEXT")
                cursor.execute("""
                    UPDATE asistencias 
                    SET fecha_evento = DATE(fecha_registro) 
                    WHERE fecha_evento IS NULL
                """)
            
            # Migración: agregar hora_inicio y hora_fin si no existen
            if 'hora_inicio' not in columnas:
                cursor.execute("ALTER TABLE asistencias ADD COLUMN hora_inicio TEXT")
            if 'hora_fin' not in columnas:
                cursor.execute("ALTER TABLE asistencias ADD COLUMN hora_fin TEXT")
        else:
            # PostgreSQL: agregar columnas si no existen
            cursor.execute("""
                ALTER TABLE asistencias ADD COLUMN IF NOT EXISTS hora_inicio TEXT
            """)
            cursor.execute("""
                ALTER TABLE asistencias ADD COLUMN IF NOT EXISTS hora_fin TEXT
            """)
        
        # Crear índice UNIQUE para prevenir duplicados
        try:
            # Verificar si hay duplicados existentes
            if USE_POSTGRES:
                query_duplicados = """
                    SELECT COUNT(*) as count FROM (
                        SELECT numero_identificacion, nombre_evento, fecha_evento, COUNT(*) as cantidad
                        FROM asistencias
                        GROUP BY numero_identificacion, nombre_evento, fecha_evento
                        HAVING COUNT(*) > 1
                    ) AS duplicados
                """
            else:
                query_duplicados = """
                    SELECT COUNT(*) as count FROM (
                        SELECT numero_identificacion, nombre_evento, fecha_evento, COUNT(*) as cantidad
                        FROM asistencias
                        GROUP BY numero_identificacion, nombre_evento, fecha_evento
                        HAVING COUNT(*) > 1
                    )
                """
            
            cursor.execute(query_duplicados)
            result = cursor.fetchone()
            duplicados_count = result['count']
            
            if duplicados_count > 0:
                
                # Eliminar duplicados manteniendo solo el registro más antiguo (menor ID)
                query_delete = adapt_query("""
                    DELETE FROM asistencias
                    WHERE id NOT IN (
                        SELECT MIN(id)
                        FROM asistencias
                        GROUP BY numero_identificacion, nombre_evento, fecha_evento
                    )
                """)
                cursor.execute(query_delete)
                eliminados = cursor.rowcount
                
                # Guardar mensaje para mostrarlo a los usuarios
                mensaje_limpieza = f"Se detectaron y eliminaron {eliminados} registros duplicados."
            
            # Ahora crear el índice UNIQUE
            cursor.execute("""
                CREATE UNIQUE INDEX IF NOT EXISTS idx_asistencias_unique 
                ON asistencias(numero_identificacion, nombre_evento, fecha_evento)
            """)
            
        except Exception as e:
            # El índice ya existe o hay otro error
            pass
        
        # Tabla de inversiones institucionales
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS inversiones_institucionales (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                año {INTEGER} NOT NULL,
                monto_libros REAL NOT NULL DEFAULT 0,
                monto_revistas REAL NOT NULL DEFAULT 0,
                monto_bases_datos REAL NOT NULL DEFAULT 0,
                total REAL GENERATED ALWAYS AS (monto_libros + monto_revistas + monto_bases_datos) STORED,
                observaciones {TEXT},
                fecha_registro {TIMESTAMP_DEFAULT},
                UNIQUE(año)
            )
        """)
        
        # Tabla de inversiones por programa
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS inversiones_programas (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                año {INTEGER} NOT NULL,
                programa {TEXT} NOT NULL,
                libros_titulos {INTEGER} NOT NULL DEFAULT 0,
                libros_volumenes {INTEGER} NOT NULL DEFAULT 0,
                libros_valor REAL NOT NULL DEFAULT 0,
                revistas_titulos {INTEGER} NOT NULL DEFAULT 0,
                revistas_valor REAL NOT NULL DEFAULT 0,
                donaciones_titulos {INTEGER} NOT NULL DEFAULT 0,
                donaciones_volumenes {INTEGER} NOT NULL DEFAULT 0,
                donaciones_trabajos_grado {INTEGER} NOT NULL DEFAULT 0,
                observaciones {TEXT},
                fecha_registro {TIMESTAMP_DEFAULT},
                UNIQUE(año, programa)
            )
        """)
        
        # Tabla de evaluaciones de capacitaciones
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS evaluaciones_capacitaciones (
                id {SERIAL} PRIMARY KEY {AUTOINCREMENT},
                asistencia_id {INTEGER} NOT NULL,
                calidad_contenido {INTEGER} NOT NULL CHECK(calidad_contenido BETWEEN 1 AND 5),
                metodologia {INTEGER} NOT NULL CHECK(metodologia BETWEEN 1 AND 5),
                lenguaje_comprensible {INTEGER} NOT NULL CHECK(lenguaje_comprensible BETWEEN 1 AND 5),
                manejo_grupo {INTEGER} NOT NULL CHECK(manejo_grupo BETWEEN 1 AND 5),
                solucion_inquietudes {INTEGER} NOT NULL CHECK(solucion_inquietudes BETWEEN 1 AND 5),
                comentarios {TEXT},
                promedio REAL GENERATED ALWAYS AS (
                    (calidad_contenido + metodologia + lenguaje_comprensible + 
                     manejo_grupo + solucion_inquietudes) / 5.0
                ) STORED,
                fecha_registro {TIMESTAMP_DEFAULT},
                FOREIGN KEY (asistencia_id) REFERENCES asistencias(id),
                UNIQUE(asistencia_id)
            )
        """)
        
        # Migración: Actualizar estructura de evaluaciones para nueva versión
        try:
            if USE_POSTGRES:
                # PostgreSQL: Agregar columna comentarios si no existe
                cursor.execute("""
                    ALTER TABLE evaluaciones_capacitaciones 
                    ADD COLUMN IF NOT EXISTS comentarios TEXT
                """)
                
                # Verificar si existen columnas antiguas
                cursor.execute("""
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name = 'evaluaciones_capacitaciones'
                """)
                columnas_existentes = [row[0] if isinstance(row, tuple) else row['column_name'] for row in cursor.fetchall()]
                
                # Eliminar columnas que ya no se usan (si existen)
                columnas_a_eliminar = [
                    'actualidad_contenidos', 'intensidad_horaria', 'dominio_tema',
                    'ayudas_didacticas', 'puntualidad'
                ]
                
                for columna in columnas_a_eliminar:
                    if columna in columnas_existentes:
                        cursor.execute(f"""
                            ALTER TABLE evaluaciones_capacitaciones 
                            DROP COLUMN IF EXISTS {columna} CASCADE
                        """)
                
                # Recrear el promedio con la fórmula correcta si es necesario
                if 'promedio' in columnas_existentes:
                    cursor.execute("""
                        ALTER TABLE evaluaciones_capacitaciones 
                        DROP COLUMN IF EXISTS promedio CASCADE
                    """)
                    cursor.execute("""
                        ALTER TABLE evaluaciones_capacitaciones 
                        ADD COLUMN promedio REAL GENERATED ALWAYS AS (
                            (calidad_contenido + metodologia + lenguaje_comprensible + 
                             manejo_grupo + solucion_inquietudes) / 5.0
                        ) STORED
                    """)
                
            else:
                # SQLite: Solo agregar comentarios si no existe
                cursor.execute("PRAGMA table_info(evaluaciones_capacitaciones)")
                columnas = [columna[1] for columna in cursor.fetchall()]
                
                if 'comentarios' not in columnas:
                    cursor.execute("ALTER TABLE evaluaciones_capacitaciones ADD COLUMN comentarios TEXT")
        except Exception as e:
            # No hacer raise para permitir que la app continúe
            pass
        
        # CRÍTICO: Commit explícito
        conn.commit()
        
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        cursor.close()
        conn.close()
    
    db_type = "PostgreSQL" if USE_POSTGRES else "SQLite"
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
            cursor = get_cursor(conn)
            query = adapt_query("SELECT * FROM usuarios WHERE username=?")
            cursor.execute(query, (usuario,))
            user = cursor.fetchone()
            
        if user and check_password_hash(user['password'], clave):
            session["usuario"] = usuario
            return redirect(url_for("dashboard"))
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
                cursor = get_cursor(conn)
                query = adapt_query("INSERT INTO usuarios (username, password) VALUES (?, ?)")
                cursor.execute(query, (usuario, password_hash))
                conn.commit()
            return redirect(url_for("login"))
        except Exception as e:
            if "UNIQUE" in str(e) or "unique" in str(e):
                return render_template("registro.html", error="El usuario ya existe")
            return render_template("registro.html", error=f"Error: {str(e)}")
    
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
            
            # Obtener horas de inicio y fin de la capacitación
            hora_inicio = request.form.get("hora_inicio", "").strip()
            hora_fin = request.form.get("hora_fin", "").strip()
            
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
                cursor = get_cursor(conn)
                # Verificar si ya existe un registro con la misma cédula, evento y fecha
                query_check = adapt_query("""
                    SELECT COUNT(*) as count FROM asistencias 
                    WHERE numero_identificacion = ? 
                    AND nombre_evento = ? 
                    AND fecha_evento = ?
                """)
                cursor.execute(query_check, (datos['numero_identificacion'], datos['nombre_evento'], fecha_evento))
                
                result = cursor.fetchone()
                count = result['count']
                
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
                if USE_POSTGRES:
                    # PostgreSQL: usar RETURNING id
                    query_insert = """
                        INSERT INTO asistencias (
                            nombre_evento, dictado_por, docente, programa_docente,
                            numero_identificacion, nombre_completo, programa_estudiante,
                            modalidad, tipo_asistente, sede, fecha_evento,
                            hora_inicio, hora_fin
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        RETURNING id
                    """
                    cursor.execute(query_insert, (datos["nombre_evento"], datos["dictado_por"], datos["docente"], 
                          datos["programa_docente"], datos["numero_identificacion"], 
                          datos["nombre_completo"], datos["programa_estudiante"],
                          datos["modalidad"], datos["tipo_asistente"], datos["sede"], 
                          fecha_evento, hora_inicio or None, hora_fin or None))
                    asistencia_id = cursor.fetchone()['id']
                else:
                    # SQLite: usar lastrowid
                    query_insert = """
                        INSERT INTO asistencias (
                            nombre_evento, dictado_por, docente, programa_docente,
                            numero_identificacion, nombre_completo, programa_estudiante,
                            modalidad, tipo_asistente, sede, fecha_evento,
                            hora_inicio, hora_fin
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """
                    cursor.execute(query_insert, (datos["nombre_evento"], datos["dictado_por"], datos["docente"], 
                          datos["programa_docente"], datos["numero_identificacion"], 
                          datos["nombre_completo"], datos["programa_estudiante"],
                          datos["modalidad"], datos["tipo_asistente"], datos["sede"], 
                          fecha_evento, hora_inicio or None, hora_fin or None))
                    asistencia_id = cursor.lastrowid
                
                conn.commit()
            
            # Solo redirigir a evaluación si NO es "Visita de grupos"
            nombre_evento = datos["nombre_evento"]
            
            if nombre_evento != "Visita de Grupos":
                # Es una capacitación/taller - debe evaluarse
                if es_acceso_publico:
                    return redirect(url_for('formulario_evaluacion', asistencia_id=asistencia_id, publico='1'))
                else:
                    return redirect(url_for('formulario_evaluacion', asistencia_id=asistencia_id))
            else:
                # Es "Visita de grupos" - ir directo a página de éxito sin evaluación
                if es_acceso_publico:
                    return redirect(url_for('formulario_success', publico='1'))
                else:
                    return redirect(url_for('formulario_success'))
            
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

def convertir_programas_para_vista(datos):
    programas_map = get_programas_map()
    datos_convertidos = []
    
    for fila in datos:
        # Convertir a lista de valores
        # Si es un diccionario (PostgreSQL con RealDictCursor), extraer valores
        # Si es una Row de SQLite, convertir a lista
        if isinstance(fila, dict):
            # PostgreSQL - Extraer valores en el orden correcto con fecha_evento al final
            fila_convertida = [
                fila.get('nombre_evento', ''),
                fila.get('dictado_por', ''),
                fila.get('docente', ''),
                fila.get('programa_docente', ''),
                fila.get('numero_identificacion', ''),
                fila.get('nombre_completo', ''),
                fila.get('programa_estudiante', ''),
                fila.get('modalidad', ''),
                fila.get('tipo_asistente', ''),
                fila.get('sede', ''),
                fila.get('fecha_evento', '')
            ]
        else:
            # SQLite - Convertir Row a lista
            fila_convertida = list(fila)
        
        # Convertir nombres de programas (las posiciones están en el lugar original)
        # programa_docente está en posición 3
        if len(fila_convertida) > 3 and fila_convertida[3]:
            fila_convertida[3] = programas_map.get(fila_convertida[3], fila_convertida[3])
        
        # programa_estudiante está en posición 6
        if len(fila_convertida) > 6 and fila_convertida[6]:
            fila_convertida[6] = programas_map.get(fila_convertida[6], fila_convertida[6])
        
        datos_convertidos.append(tuple(fila_convertida))
    
    return datos_convertidos

@app.route("/panel")
def panel():
    global mensaje_limpieza_global
    
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    error = request.args.get("error", "").strip()
    success = request.args.get("success", "").strip()
    
    # Si hay mensaje de limpieza pendiente, agregarlo al success
    if mensaje_limpieza_global and not success:
        success = f"⚠️ {mensaje_limpieza_global}"
        mensaje_limpieza_global = None  # Limpiar para que solo se muestre una vez
    
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("SELECT COUNT(*) as total FROM asistencias")
            cursor.execute(query)
            result = cursor.fetchone()
            total_registros = result['total'] if isinstance(result, dict) else result[0]
        
        return render_template("panel.html",
                               total_registros=total_registros,
                               error=error if error else None,
                               success=success if success else None)
    
    except Exception as e:
        return render_template("panel.html", total_registros=0, error=f"Error cargando datos: {str(e)}")


@app.route("/api/asistencias")
def api_asistencias():
    """
    API para DataTables server-side processing.
    Recibe parámetros de DataTables (draw, start, length, search, order, columns)
    y devuelve el JSON que DataTables espera.
    
    INCLUYE: Búsqueda insensible a mayúsculas y acentos usando normalize_text()
    """
    if "usuario" not in session:
        return {"error": "No autorizado"}, 401

    try:
        # Parámetros estándar de DataTables
        draw        = int(request.args.get('draw', 1))
        start       = int(request.args.get('start', 0))
        length      = int(request.args.get('length', 25))
        search_val  = request.args.get('search[value]', '').strip()

        # Columnas en el mismo orden que el SELECT y la tabla HTML
        col_names = [
            'nombre_evento', 'dictado_por', 'docente', 'programa_docente',
            'numero_identificacion', 'nombre_completo', 'programa_estudiante',
            'modalidad', 'tipo_asistente', 'sede', 'fecha_evento',
            'hora_inicio', 'hora_fin'
        ]

        # Columna y dirección de ordenamiento
        order_col_idx = int(request.args.get('order[0][column]', 10))
        order_dir     = request.args.get('order[0][dir]', 'desc')
        order_col     = col_names[order_col_idx] if 0 <= order_col_idx < len(col_names) else 'fecha_evento'
        if order_dir not in ('asc', 'desc'):
            order_dir = 'desc'

        # Filtros individuales por columna
        col_filters = []
        for i in range(len(col_names)):
            val = request.args.get(f'columns[{i}][search][value]', '').strip()
            col_filters.append(val)

        # Construir WHERE con búsqueda normalizada
        conditions = []
        params     = []

        # Búsqueda global (normalizada)
        if search_val:
            search_normalized = normalize_text(search_val)
            if USE_POSTGRES:
                # PostgreSQL: usar unaccent si está disponible, sino usar LOWER
                sub = ' OR '.join([f"LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE({c}, 'á','a'), 'é','e'), 'í','i'), 'ó','o'), 'ú','u')) LIKE %s" for c in col_names])
            else:
                # SQLite: crear una condición con LOWER para cada columna
                sub = ' OR '.join([f"LOWER({c}) LIKE ?" for c in col_names])
            conditions.append(f"({sub})")
            params += [f"%{search_normalized}%" for _ in col_names]

        # Filtros por columna (normalizados)
        for i, val in enumerate(col_filters):
            if val:
                val_normalized = normalize_text(val)
                if USE_POSTGRES:
                    conditions.append(f"LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE({col_names[i]}, 'á','a'), 'é','e'), 'í','i'), 'ó','o'), 'ú','u')) LIKE %s")
                else:
                    conditions.append(f"LOWER({col_names[i]}) LIKE ?")
                params.append(f"%{val_normalized}%")

        where_clause = ('WHERE ' + ' AND '.join(conditions)) if conditions else ''

        with get_db_connection() as conn:
            cursor = get_cursor(conn)

            # Total sin filtros
            cursor.execute(adapt_query("SELECT COUNT(*) as total FROM asistencias"))
            total_records = cursor.fetchone()
            total_records = total_records['total'] if isinstance(total_records, dict) else total_records[0]

            # Total con filtros aplicados
            count_query = adapt_query(f"SELECT COUNT(*) as total FROM asistencias {where_clause}")
            cursor.execute(count_query, params)
            filtered_records = cursor.fetchone()
            filtered_records = filtered_records['total'] if isinstance(filtered_records, dict) else filtered_records[0]

            # Datos paginados (incluir id para el botón de edición)
            data_query = adapt_query(f"""
                SELECT id, nombre_evento, dictado_por, docente, programa_docente,
                       numero_identificacion, nombre_completo, programa_estudiante,
                       modalidad, tipo_asistente, sede, fecha_evento,
                       hora_inicio, hora_fin
                FROM asistencias
                {where_clause}
                ORDER BY {order_col} {order_dir}
                LIMIT ? OFFSET ?
            """)
            cursor.execute(data_query, params + [length, start])
            rows = cursor.fetchall()

        # Convertir programas a nombres completos
        programas_map = get_programas_map()
        data = []
        for row in rows:
            if isinstance(row, dict):
                fila = [
                    row.get('nombre_evento', '') or '',
                    row.get('dictado_por', '') or '',
                    row.get('docente', '') or '',
                    programas_map.get(row.get('programa_docente', ''), row.get('programa_docente', '') or ''),
                    row.get('numero_identificacion', '') or '',
                    row.get('nombre_completo', '') or '',
                    programas_map.get(row.get('programa_estudiante', ''), row.get('programa_estudiante', '') or ''),
                    row.get('modalidad', '') or '',
                    row.get('tipo_asistente', '') or '',
                    row.get('sede', '') or '',
                    row.get('fecha_evento', '') or '',
                    row.get('hora_inicio', '') or '',
                    row.get('hora_fin', '') or '',
                    row.get('id', '')     # índice 13 — usado por el botón de editar
                ]
            else:
                fila = list(row)
                # row: id(0), nombre_evento(1)...hora_fin(13)
                # Reordenar: [1..13, 0] para que data[0]=nombre_evento y data[13]=id
                fila = [
                    fila[1] or '',   # nombre_evento
                    fila[2] or '',   # dictado_por
                    fila[3] or '',   # docente
                    programas_map.get(fila[4] or '', fila[4] or ''),  # programa_docente
                    fila[5] or '',   # numero_identificacion
                    fila[6] or '',   # nombre_completo
                    programas_map.get(fila[7] or '', fila[7] or ''),  # programa_estudiante
                    fila[8] or '',   # modalidad
                    fila[9] or '',   # tipo_asistente
                    fila[10] or '',  # sede
                    fila[11] or '',  # fecha_evento
                    fila[12] or '',  # hora_inicio
                    fila[13] or '',  # hora_fin
                    fila[0]          # id (índice 13 en la respuesta)
                ]
            data.append(fila)

        return {
            "draw":            draw,
            "recordsTotal":    total_records,
            "recordsFiltered": filtered_records,
            "data":            data
        }

    except Exception as e:
        return {"draw": 1, "recordsTotal": 0, "recordsFiltered": 0, "data": [], "error": str(e)}, 500



# ============================================
# API EDICIÓN DE REGISTRO INDIVIDUAL
# ============================================

@app.route("/api/asistencia/<int:asistencia_id>", methods=["GET"])
def api_get_asistencia(asistencia_id):
    """Devuelve los datos de un registro de asistencia por su ID."""
    if "usuario" not in session:
        return {"success": False, "error": "No autorizado"}, 401
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("""
                SELECT id, nombre_evento, dictado_por, docente, programa_docente,
                       numero_identificacion, nombre_completo, programa_estudiante,
                       modalidad, tipo_asistente, sede, fecha_evento,
                       hora_inicio, hora_fin
                FROM asistencias WHERE id = ?
            """)
            cursor.execute(query, (asistencia_id,))
            row = cursor.fetchone()

        if not row:
            return {"success": False, "error": "Registro no encontrado"}, 404

        if isinstance(row, dict):
            fecha = row.get('fecha_evento', '')
            if fecha and hasattr(fecha, 'strftime'):
                fecha = fecha.strftime('%Y-%m-%d')
            registro = {
                'id':                    row.get('id'),
                'nombre_evento':         row.get('nombre_evento', ''),
                'dictado_por':           row.get('dictado_por', ''),
                'docente':               row.get('docente', ''),
                'programa_docente':      row.get('programa_docente', ''),
                'numero_identificacion': row.get('numero_identificacion', ''),
                'nombre_completo':       row.get('nombre_completo', ''),
                'programa_estudiante':   row.get('programa_estudiante', ''),
                'modalidad':             row.get('modalidad', ''),
                'tipo_asistente':        row.get('tipo_asistente', ''),
                'sede':                  row.get('sede', ''),
                'fecha_evento':          fecha,
                'hora_inicio':           row.get('hora_inicio', '') or '',
                'hora_fin':              row.get('hora_fin', '') or '',
            }
        else:
            row = list(row)
            fecha = row[11]
            if fecha and hasattr(fecha, 'strftime'):
                fecha = fecha.strftime('%Y-%m-%d')
            registro = {
                'id':                    row[0],
                'nombre_evento':         row[1] or '',
                'dictado_por':           row[2] or '',
                'docente':               row[3] or '',
                'programa_docente':      row[4] or '',
                'numero_identificacion': row[5] or '',
                'nombre_completo':       row[6] or '',
                'programa_estudiante':   row[7] or '',
                'modalidad':             row[8] or '',
                'tipo_asistente':        row[9] or '',
                'sede':                  row[10] or '',
                'fecha_evento':          fecha or '',
                'hora_inicio':           row[12] or '',
                'hora_fin':              row[13] or '',
            }

        return {"success": True, "registro": registro}

    except Exception as e:
        return {"success": False, "error": str(e)}, 500


@app.route("/api/asistencia/<int:asistencia_id>", methods=["PUT"])
def api_put_asistencia(asistencia_id):
    """Actualiza un registro de asistencia por su ID."""
    if "usuario" not in session:
        return {"success": False, "error": "No autorizado"}, 401

    from flask import json as flask_json
    try:
        datos = request.get_json(force=True)
        if not datos:
            return {"success": False, "error": "No se recibieron datos"}, 400

        campos_requeridos = [
            'nombre_evento', 'dictado_por', 'docente', 'programa_docente',
            'numero_identificacion', 'nombre_completo', 'programa_estudiante',
            'modalidad', 'tipo_asistente', 'sede', 'fecha_evento'
        ]
        for campo in campos_requeridos:
            if not datos.get(campo, '').strip():
                return {"success": False, "error": f"El campo '{campo}' es requerido"}, 400

        # Validar formato de fecha
        try:
            datetime.strptime(datos['fecha_evento'], "%Y-%m-%d")
        except ValueError:
            return {"success": False, "error": "Formato de fecha inválido (use YYYY-MM-DD)"}, 400

        hora_inicio = datos.get('hora_inicio', '') or None
        hora_fin    = datos.get('hora_fin', '') or None

        # Validar horas si ambas se proporcionan
        if hora_inicio and hora_fin and hora_fin <= hora_inicio:
            return {"success": False, "error": "La hora de fin debe ser posterior a la hora de inicio"}, 400

        with get_db_connection() as conn:
            cursor = get_cursor(conn)

            # Verificar que el registro existe
            cursor.execute(adapt_query("SELECT id FROM asistencias WHERE id = ?"), (asistencia_id,))
            if not cursor.fetchone():
                return {"success": False, "error": "Registro no encontrado"}, 404

            query = adapt_query("""
                UPDATE asistencias SET
                    nombre_evento          = ?,
                    dictado_por            = ?,
                    docente                = ?,
                    programa_docente       = ?,
                    numero_identificacion  = ?,
                    nombre_completo        = ?,
                    programa_estudiante    = ?,
                    modalidad              = ?,
                    tipo_asistente         = ?,
                    sede                   = ?,
                    fecha_evento           = ?,
                    hora_inicio            = ?,
                    hora_fin               = ?
                WHERE id = ?
            """)
            cursor.execute(query, (
                datos['nombre_evento'].strip(),
                datos['dictado_por'].strip(),
                datos['docente'].strip(),
                datos['programa_docente'].strip(),
                datos['numero_identificacion'].strip(),
                datos['nombre_completo'].strip(),
                datos['programa_estudiante'].strip(),
                datos['modalidad'].strip(),
                datos['tipo_asistente'].strip(),
                datos['sede'].strip(),
                datos['fecha_evento'].strip(),
                hora_inicio,
                hora_fin,
                asistencia_id
            ))
            conn.commit()

        return {"success": True, "message": "Registro actualizado correctamente"}

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "error": str(e)}, 500

@app.route("/api/stats/asistencias")
def api_stats_asistencias():
    """Devuelve conteos únicos de eventos, programas y sedes para las tarjetas del panel."""
    if "usuario" not in session:
        return {"success": False, "error": "No autorizado"}, 401
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)

            cursor.execute("SELECT COUNT(DISTINCT nombre_evento) as total FROM asistencias")
            eventos = cursor.fetchone()
            eventos = eventos['total'] if isinstance(eventos, dict) else eventos[0]

            cursor.execute("SELECT COUNT(DISTINCT programa_estudiante) as total FROM asistencias")
            programas = cursor.fetchone()
            programas = programas['total'] if isinstance(programas, dict) else programas[0]

            cursor.execute("SELECT COUNT(DISTINCT sede) as total FROM asistencias")
            sedes = cursor.fetchone()
            sedes = sedes['total'] if isinstance(sedes, dict) else sedes[0]

        return {"success": True, "eventos": eventos, "programas": programas, "sedes": sedes}
    except Exception as e:
        return {"success": False, "error": str(e)}, 500

@app.route("/panel/cargar_excel", methods=["POST"])
def panel_cargar_excel():
    """Ruta para cargar datos desde archivo Excel al panel de asistencias"""
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        # Verificar que se haya enviado un archivo
        if 'file' not in request.files:
            flash('texto', 'danger')
            return redirect(url_for('panel'))

        
        file = request.files['file']
        
        if file.filename == '':
            flash('texto', 'danger')
            return redirect(url_for('panel'))

        
        # Verificar que sea un archivo Excel
        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('texto', 'danger')
            return redirect(url_for('panel'))

        
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
            cursor = get_cursor(conn)
            
            for index, row in df.iterrows():
                try:
                    # Intentar insertar el registro
                    query_insert = adapt_query("""
                        INSERT INTO asistencias (
                            nombre_evento, dictado_por, docente, programa_docente,
                            numero_identificacion, nombre_completo, programa_estudiante,
                            modalidad, tipo_asistente, sede, fecha_evento
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """)
                    cursor.execute(query_insert, (
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
                except Exception as e:
                    if "UNIQUE" in str(e) or "unique" in str(e):
                        # Registro duplicado detectado por el índice UNIQUE
                        registros_duplicados += 1
                        # Guardar información del duplicado (opcional, para debugging)
                        if registros_duplicados <= 5:  # Solo guardar los primeros 5 para no saturar
                            errores.append(f"Fila {index + 2}: {row['nombre_completo']} - {row['nombre_evento']} - {row['fecha_evento']}")
                    else:
                        # Otros errores (datos inválidos, etc.)
                        errores.append(f"Fila {index + 2}: Error - {str(e)}")
                    continue
            
            conn.commit()
        
        # Preparar mensaje de éxito
        mensaje = f'Se insertaron {registros_insertados} registros correctamente.'
        if registros_duplicados > 0:
            mensaje += f' ⚠️ Se omitieron {registros_duplicados} registros duplicados.'
        if len(errores) > 0 and registros_duplicados <= 5:
            mensaje += f' Duplicados: {", ".join(errores)}'
        
        flash(mensaje, 'success')
        return redirect(url_for('panel'))

        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return redirect(url_for('panel', error=f'Error procesando el archivo: {str(e)}'))

@app.route("/exportar")
def exportar():
    if "usuario" not in session:
        return redirect(url_for("login"))

    try:
        # Obtener filtros de columnas (13 columnas)
        col_filters = [request.args.get(f"col{i}", "").strip() for i in range(13)]
        
        # Obtener búsqueda global
        global_search = request.args.get("global_search", "").strip()
        
        # Obtener ordenamiento
        order_column = request.args.get("order_column", "10")  # Por defecto columna 10 (fecha_evento)
        order_dir = request.args.get("order_dir", "desc")       # Por defecto descendente
        
        # Nombres de columnas en el mismo orden que la tabla
        col_names = [
            'nombre_evento', 'dictado_por', 'docente', 'programa_docente',
            'numero_identificacion', 'nombre_completo', 'programa_estudiante',
            'modalidad', 'tipo_asistente', 'sede', 'fecha_evento',
            'hora_inicio', 'hora_fin'
        ]
        
        # Validar índice de columna de ordenamiento
        try:
            order_col_idx = int(order_column)
            if 0 <= order_col_idx < len(col_names):
                order_col = col_names[order_col_idx]
            else:
                order_col = 'fecha_evento'
        except:
            order_col = 'fecha_evento'
        
        # Validar dirección de ordenamiento
        if order_dir not in ('asc', 'desc'):
            order_dir = 'desc'

        # Construir condiciones WHERE con normalización (igual que en api_asistencias)
        conditions = []
        params = []

        # Búsqueda global (normalizada)
        if global_search:
            global_normalized = normalize_text(global_search)
            if USE_POSTGRES:
                sub = ' OR '.join([f"LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE({c}, 'á','a'), 'é','e'), 'í','i'), 'ó','o'), 'ú','u')) LIKE %s" for c in col_names])
            else:
                sub = ' OR '.join([f"LOWER(COALESCE({c},'')) LIKE ?" for c in col_names])
            conditions.append(f"({sub})")
            params += [f"%{global_normalized}%" for _ in col_names]

        # Filtros por columna (normalizados)
        for i, val in enumerate(col_filters):
            if val:
                val_normalized = normalize_text(val)
                if USE_POSTGRES:
                    conditions.append(f"LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE({col_names[i]}, 'á','a'), 'é','e'), 'í','i'), 'ó','o'), 'ú','u')) LIKE %s")
                else:
                    conditions.append(f"LOWER(COALESCE({col_names[i]},'')) LIKE ?")
                params.append(f"%{val_normalized}%")

        # Construir query
        query = "SELECT * FROM asistencias"
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        query += f" ORDER BY {order_col} {order_dir}"
        
        # Adaptar query
        query = adapt_query(query)

        with get_db_connection() as conn:
            df = pd.read_sql_query(query, conn, params=params)

        if df.empty:
            return "No hay datos para exportar con los filtros aplicados", 400

        # Convertir programas a nombres completos
        programas_map = get_programas_map()
        
        def convertir_programa(nombre_programa):
            return programas_map.get(nombre_programa, nombre_programa)

        if 'programa_docente' in df.columns:
            df['programa_docente'] = df['programa_docente'].apply(convertir_programa)
        
        if 'programa_estudiante' in df.columns:
            df['programa_estudiante'] = df['programa_estudiante'].apply(convertir_programa)

        if 'id' in df.columns:
            df = df.drop('id', axis=1)

        # Eliminar columnas de hora y fecha de registro del Excel
        cols_excluir = ['hora_inicio', 'hora_fin', 'fecha_registro']
        df = df.drop(columns=[c for c in cols_excluir if c in df.columns])

        # Renombrar columnas para el Excel
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
            'hora_inicio': 'Hora de Inicio',
            'hora_fin': 'Hora de Fin',
            'fecha_registro': 'Fecha de Registro'
        }, inplace=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Asistencias")

        output.seek(0)

        return send_file(output,
                         as_attachment=True,
                         download_name="asistencias_filtradas.xlsx",
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
            cursor = get_cursor(conn)
            query = adapt_query("""
                SELECT año, monto_libros, monto_revistas, monto_bases_datos, total, observaciones
                FROM inversiones_institucionales
                ORDER BY año DESC
            """)
            cursor.execute(query)
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
                cursor = get_cursor(conn)
                query = adapt_query("""
                    INSERT INTO inversiones_institucionales 
                    (año, monto_libros, monto_revistas, monto_bases_datos, observaciones)
                    VALUES (?, ?, ?, ?, ?)
                """)
                cursor.execute(query, (año, monto_libros, monto_revistas, monto_bases_datos, observaciones))
                conn.commit()
            
            return redirect(url_for("inversiones_institucional"))
            
        except Exception as e:
            if "UNIQUE" in str(e) or "unique" in str(e):
                return render_template("inversiones_institucional_form.html",
                                     error=f"Ya existe un registro para el año {año}",
                                     form_data=request.form)
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
            cursor = get_cursor(conn)
            query = adapt_query("""
                SELECT año, programa, 
                       libros_titulos, libros_volumenes, libros_valor,
                       revistas_titulos, revistas_valor,
                       donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                       observaciones
                FROM inversiones_programas
                ORDER BY año DESC, programa
            """)
            cursor.execute(query)
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
                cursor = get_cursor(conn)
                query = adapt_query("""
                    INSERT INTO inversiones_programas 
                    (año, programa, libros_titulos, libros_volumenes, libros_valor,
                     revistas_titulos, revistas_valor,
                     donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                     observaciones)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """)
                cursor.execute(query, (año, programa, libros_titulos, libros_volumenes, libros_valor,
                      revistas_titulos, revistas_valor,
                      donaciones_titulos, donaciones_volumenes, donaciones_trabajos_grado,
                      observaciones))
                conn.commit()
            
            return redirect(url_for("inversiones_programas"))
            
        except Exception as e:
            if "UNIQUE" in str(e) or "unique" in str(e):
                return render_template("inversiones_programas_form.html",
                                     programas=get_programas_list(),
                                     error=f"Ya existe un registro para {programa} en el año {año}",
                                     form_data=request.form)
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
        # Obtener filtros de la URL
        evento_filtro = request.args.get('evento', '')
        programa_filtro = request.args.get('programa', '')
        fecha_inicio = request.args.get('fecha_inicio', '')
        fecha_fin = request.args.get('fecha_fin', '')
        
        with get_db_connection() as conn:
            # Obtener años reales con datos (máximo últimos 5)
            query_anos = adapt_query("""
                SELECT DISTINCT CAST(SUBSTR(fecha_evento, 1, 4) AS INTEGER) as ano
                FROM asistencias
                WHERE fecha_evento IS NOT NULL AND fecha_evento != ''
                ORDER BY ano DESC
                LIMIT 5
            """)
            df_anos = pd.read_sql_query(query_anos, conn)
            anos_con_datos = sorted(df_anos['ano'].tolist()) if not df_anos.empty else []

            # Query base con filtros opcionales
            where_clauses = []
            params = []
            
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
            
            # Adaptar queries
            # 1. Total de asistencias
            query1 = adapt_query(f"SELECT COUNT(*) as total_asistencias FROM asistencias {where_sql}")
            df_total = pd.read_sql_query(query1, conn, params=params)
            
            # 2. Total de eventos únicos
            query2 = adapt_query(f"SELECT COUNT(DISTINCT nombre_evento) as total_eventos FROM asistencias {where_sql}")
            df_eventos_unicos = pd.read_sql_query(query2, conn, params=params)
            
            # 3. Total de programas únicos
            query3 = adapt_query(f"SELECT COUNT(DISTINCT programa_estudiante) as total_programas FROM asistencias {where_sql}")
            df_programas_unicos = pd.read_sql_query(query3, conn, params=params)
            
            # 4. Promedio de evaluaciones
            query4 = adapt_query(f"""
                SELECT AVG(promedio) as promedio_general
                FROM evaluaciones_capacitaciones e
                INNER JOIN asistencias a ON e.asistencia_id = a.id
                {where_sql}
            """)
            df_promedio_eval = pd.read_sql_query(query4, conn, params=params)
            
            # 5. Datos de eventos (asistencias por evento)
            query5 = adapt_query(f"""
                SELECT nombre_evento, COUNT(*) as total_asistencias
                FROM asistencias {where_sql}
                GROUP BY nombre_evento 
                ORDER BY total_asistencias DESC
                LIMIT 15
            """)
            df_eventos = pd.read_sql_query(query5, conn, params=params)
            
            # 6. Datos de programas (asistencias por programa + modalidad)
            query6 = adapt_query(f"""
                SELECT programa_estudiante || ' - ' || modalidad as programa_completo, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY programa_estudiante, modalidad
                ORDER BY total DESC
                LIMIT 15
            """)
            df_programas = pd.read_sql_query(query6, conn, params=params)
            
            # 7. Análisis cruzado Programa x Evento (incluyendo modalidad)
            query7 = adapt_query(f"""
                SELECT 
                    nombre_evento,
                    programa_estudiante || ' - ' || modalidad as programa_completo,
                    COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY nombre_evento, programa_estudiante, modalidad
                ORDER BY nombre_evento, total DESC
            """)
            df_cruzado = pd.read_sql_query(query7, conn, params=params)
            
            # 8. Tendencia mensual
            query8 = adapt_query(f"""
                SELECT 
                    SUBSTR(fecha_evento, 1, 7) as mes,
                    COUNT(*) as total
                FROM asistencias
                WHERE fecha_evento IS NOT NULL 
                    AND fecha_evento != ''
                    {' AND ' + ' AND '.join(where_clauses) if where_clauses else ''}
                GROUP BY mes
                ORDER BY mes
            """)
            df_mensual = pd.read_sql_query(query8, conn, params=params if where_clauses else [])
            
            # 9. Top 5 programas por evento
            if USE_POSTGRES:
                query_top = f"""
                    SELECT * FROM (
                        SELECT 
                            nombre_evento,
                            programa_estudiante,
                            COUNT(*) as total,
                            ROW_NUMBER() OVER (PARTITION BY nombre_evento ORDER BY COUNT(*) DESC) as ranking
                        FROM asistencias {where_sql}
                        GROUP BY nombre_evento, programa_estudiante
                    ) AS subquery
                    WHERE ranking <= 5
                    ORDER BY nombre_evento, ranking
                """
            else:
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
            query_top = adapt_query(query_top)
            df_top_por_evento = pd.read_sql_query(query_top, conn, params=params)
            
            # 10. Obtener listas únicas para filtros
            df_eventos_lista = pd.read_sql_query("SELECT DISTINCT nombre_evento FROM asistencias ORDER BY nombre_evento", conn)
            df_programas_lista = pd.read_sql_query("SELECT DISTINCT programa_estudiante as nombre FROM asistencias ORDER BY programa_estudiante", conn)
            
            # 11. Tipo de asistentes
            query11 = adapt_query(f"""
                SELECT tipo_asistente, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY tipo_asistente
                ORDER BY total DESC
            """)
            df_tipo_asistente = pd.read_sql_query(query11, conn, params=params)
            
            # 12. Modalidad
            query12 = adapt_query(f"""
                SELECT modalidad, COUNT(*) as total
                FROM asistencias {where_sql}
                GROUP BY modalidad
                ORDER BY total DESC
            """)
            df_modalidad = pd.read_sql_query(query12, conn, params=params)

        # Verificar si hay datos
        if df_eventos.empty:
            return render_template("estadisticas_avanzadas.html", 
                                 mensaje="No hay datos disponibles para mostrar estadísticas",
                                 anos_con_datos=anos_con_datos,
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
        
        programa_labels = df_programas['programa_completo'].tolist()
        programa_valores = [int(x) for x in df_programas['total'].tolist()]
        
        # Procesar datos cruzados para matriz
        matriz_cruzada = {}
        for _, row in df_cruzado.iterrows():
            evento = row['nombre_evento']
            programa = row['programa_completo']
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
                             
                             # Años con datos para botones de filtro rápido
                             anos_con_datos=anos_con_datos,
                             
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
        import traceback
        traceback.print_exc()
        return render_template("estadisticas_avanzadas.html", 
                             error=f"Error cargando estadísticas: {str(e)}",
                             anos_con_datos=[],
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
                "calidad_contenido", "metodologia", "lenguaje_comprensible",
                "manejo_grupo", "solucion_inquietudes"
            ]
            
            evaluacion = {}
            for campo in campos_evaluacion:
                valor = request.form.get(campo, "").strip()
                if not valor or not valor.isdigit() or int(valor) < 1 or int(valor) > 5:
                    with get_db_connection() as conn:
                        cursor = get_cursor(conn)
                        query = adapt_query("SELECT * FROM asistencias WHERE id = ?")
                        cursor.execute(query, (asistencia_id,))
                        asistencia_raw = cursor.fetchone()
                    
                    asistencia = asistencia_to_tuple(asistencia_raw)
                    
                    return render_template("formulario_evaluacion.html",
                                         asistencia=asistencia,
                                         error=f"El campo '{campo.replace('_', ' ').title()}' debe ser un valor entre 1 y 5",
                                         form_data=request.form,
                                         es_acceso_publico=es_acceso_publico)
                evaluacion[campo] = int(valor)
            
            # Obtener comentarios (opcional)
            comentarios = request.form.get("comentarios", "").strip()
            
            # Insertar la evaluación
            with get_db_connection() as conn:
                cursor = get_cursor(conn)
                query = adapt_query("""
                    INSERT INTO evaluaciones_capacitaciones (
                        asistencia_id, calidad_contenido, metodologia,
                        lenguaje_comprensible, manejo_grupo, solucion_inquietudes, comentarios
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """)
                cursor.execute(query, (asistencia_id, 
                      evaluacion["calidad_contenido"],
                      evaluacion["metodologia"],
                      evaluacion["lenguaje_comprensible"],
                      evaluacion["manejo_grupo"],
                      evaluacion["solucion_inquietudes"],
                      comentarios if comentarios else None))
                conn.commit()
            
            # Redirigir a página de éxito
            if es_acceso_publico:
                return redirect(url_for('evaluacion_success', publico='1'))
            else:
                return redirect(url_for('evaluacion_success'))
                
        except Exception as e:
            if "UNIQUE" in str(e) or "unique" in str(e):
                with get_db_connection() as conn:
                    cursor = get_cursor(conn)
                    query = adapt_query("SELECT * FROM asistencias WHERE id = ?")
                    cursor.execute(query, (asistencia_id,))
                    asistencia_raw = cursor.fetchone()
                
                asistencia = asistencia_to_tuple(asistencia_raw)
                
                return render_template("formulario_evaluacion.html",
                                     asistencia=asistencia,
                                     error="Ya existe una evaluación para esta asistencia",
                                     form_data=request.form,
                                     es_acceso_publico=es_acceso_publico)
            
            with get_db_connection() as conn:
                cursor = get_cursor(conn)
                query = adapt_query("SELECT * FROM asistencias WHERE id = ?")
                cursor.execute(query, (asistencia_id,))
                asistencia_raw = cursor.fetchone()
            
            asistencia = asistencia_to_tuple(asistencia_raw)
            
            return render_template("formulario_evaluacion.html",
                                 asistencia=asistencia,
                                 error=f"Error al registrar evaluación: {str(e)}",
                                 form_data=request.form,
                                 es_acceso_publico=es_acceso_publico)
    
    # GET request
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("SELECT * FROM asistencias WHERE id = ?")
            cursor.execute(query, (asistencia_id,))
            asistencia_raw = cursor.fetchone()
            
            if not asistencia_raw:
                return redirect(url_for('formulario'))
        
        asistencia = asistencia_to_tuple(asistencia_raw)
        
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
            cursor = get_cursor(conn)
            
            # Obtener programas
            query_programas = adapt_query("""
                SELECT id, nombre, activo, 
                       fecha_creacion, fecha_modificacion
                FROM programas
                ORDER BY nombre ASC
            """)
            cursor.execute(query_programas)
            programas_raw = cursor.fetchall()
            
            # Obtener modalidades
            query_modalidades = adapt_query("""
                SELECT id, nombre, activo, 
                       fecha_creacion, fecha_modificacion
                FROM modalidades
                ORDER BY nombre ASC
            """)
            cursor.execute(query_modalidades)
            modalidades_raw = cursor.fetchall()
        
        # Convertir a tuplas para compatibilidad con template
        programas = programas_to_tuples(programas_raw)
        modalidades = programas_to_tuples(modalidades_raw)
        
        return render_template("gestion_programas.html", 
                             programas=programas,
                             modalidades=modalidades)
    except Exception as e:
        return render_template("gestion_programas.html", 
                             error=f"Error al cargar datos: {str(e)}",
                             programas=[],
                             modalidades=[])

@app.route("/api/programas/activos")
def get_programas_activos():
    """API para obtener solo programas activos (para el formulario)"""
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("""
                SELECT nombre
                FROM programas
                WHERE activo = 1
                ORDER BY nombre ASC
            """)
            cursor.execute(query)
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
            cursor = get_cursor(conn)
            # Obtener estado actual
            query_select = adapt_query("SELECT activo FROM programas WHERE id = ?")
            cursor.execute(query_select, (programa_id,))
            programa = cursor.fetchone()
            
            if not programa:
                return {"success": False, "error": "Programa no encontrado"}, 404
            
            # Cambiar estado
            nuevo_estado = 0 if programa["activo"] == 1 else 1
            query_update = adapt_query("""
                UPDATE programas 
                SET activo = ?,
                    fecha_modificacion = CURRENT_TIMESTAMP
                WHERE id = ?
            """)
            cursor.execute(query_update, (nuevo_estado, programa_id))
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
            cursor = get_cursor(conn)
            
            if USE_POSTGRES:
                # PostgreSQL: usar RETURNING id
                query = """
                    INSERT INTO programas (nombre, activo)
                    VALUES (%s, 1)
                    RETURNING id
                """
                cursor.execute(query, (nombre,))
                programa_id = cursor.fetchone()['id']
            else:
                # SQLite: usar lastrowid
                query = """
                    INSERT INTO programas (nombre, activo)
                    VALUES (?, 1)
                """
                cursor.execute(query, (nombre,))
                programa_id = cursor.lastrowid
            
            conn.commit()
        
        return {
            "success": True,
            "mensaje": "Programa agregado exitosamente",
            "programa_id": programa_id
        }
    except Exception as e:
        if "UNIQUE" in str(e) or "unique" in str(e):
            return {"success": False, "error": "Ya existe un programa con ese nombre"}, 400
        return {"success": False, "error": str(e)}, 500

@app.route("/api/programas/eliminar/<int:programa_id>", methods=["DELETE"])
def eliminar_programa(programa_id):
    """Eliminar un programa académico (solo si no tiene registros asociados)"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            
            # Verificar si tiene asistencias registradas
            query_check = adapt_query("""
                SELECT COUNT(*) as count FROM asistencias 
                WHERE programa_estudiante IN (
                    SELECT nombre FROM programas WHERE id = ?
                ) OR programa_docente IN (
                    SELECT nombre FROM programas WHERE id = ?
                )
            """)
            cursor.execute(query_check, (programa_id, programa_id))
            
            result = cursor.fetchone()
            count = result["count"]
            
            if count > 0:
                return {
                    "success": False, 
                    "error": f"No se puede eliminar. El programa tiene {count} registro(s) asociado(s). Considere deshabilitarlo en su lugar."
                }, 400
            
            # Eliminar programa
            query_delete = adapt_query("DELETE FROM programas WHERE id = ?")
            cursor.execute(query_delete, (programa_id,))
            conn.commit()
        
        return {
            "success": True,
            "mensaje": "Programa eliminado exitosamente"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}, 500
    
# ========== API: GESTIÓN DE MODALIDADES ==========

@app.route("/api/modalidades/agregar", methods=["POST"])
def api_agregar_modalidad():
    """Agregar nueva modalidad"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        data = request.get_json()
        nombre = data.get('nombre', '').strip()
        
        if not nombre:
            return {"success": False, "error": "El nombre es requerido"}, 400
        
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            
            try:
                if USE_POSTGRES:
                    cursor.execute("""
                        INSERT INTO modalidades (nombre, activo)
                        VALUES (%s, 1)
                        RETURNING id
                    """, (nombre,))
                    modalidad_id = cursor.fetchone()['id']
                else:
                    cursor.execute("""
                        INSERT INTO modalidades (nombre, activo)
                        VALUES (?, 1)
                    """, (nombre,))
                    modalidad_id = cursor.lastrowid
                
                conn.commit()
                
                return {
                    "success": True,
                    "message": "Modalidad agregada exitosamente",
                    "id": modalidad_id
                }
            
            except Exception as e:
                if "UNIQUE constraint failed" in str(e) or "duplicate key" in str(e):
                    return {"success": False, "error": "Esta modalidad ya existe"}, 400
                
                return {"success": False, "error": str(e)}, 500
    
    except Exception as e:
        return {"success": False, "error": str(e)}, 500


@app.route("/api/modalidades/toggle/<int:modalidad_id>", methods=["POST"])
def api_toggle_modalidad(modalidad_id):
    """Activar/Desactivar modalidad"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            
            # Obtener estado actual
            if USE_POSTGRES:
                cursor.execute("SELECT activo FROM modalidades WHERE id = %s", (modalidad_id,))
            else:
                cursor.execute("SELECT activo FROM modalidades WHERE id = ?", (modalidad_id,))
            
            row = cursor.fetchone()
            
            if not row:
                return {"success": False, "error": "Modalidad no encontrada"}, 404
            
            # Determinar nuevo estado
            if isinstance(row, dict):
                current_activo = row['activo']
            else:
                current_activo = row[0]
            
            nuevo_estado = 0 if current_activo == 1 else 1
            
            # Actualizar estado
            if USE_POSTGRES:
                cursor.execute("""
                    UPDATE modalidades 
                    SET activo = %s, fecha_modificacion = CURRENT_TIMESTAMP
                    WHERE id = %s
                """, (nuevo_estado, modalidad_id))
            else:
                cursor.execute("""
                    UPDATE modalidades 
                    SET activo = ?, fecha_modificacion = CURRENT_TIMESTAMP
                    WHERE id = ?
                """, (nuevo_estado, modalidad_id))
            
            conn.commit()
        
        return {
            "success": True,
            "message": "Estado actualizado",
            "nuevo_estado": nuevo_estado
        }
    
    except Exception as e:
        return {"success": False, "error": str(e)}, 500


@app.route("/api/modalidades/eliminar/<int:modalidad_id>", methods=["DELETE"])
def api_eliminar_modalidad(modalidad_id):
    """Eliminar modalidad"""
    if 'usuario' not in session:
        return {"success": False, "error": "No autorizado"}, 401
    
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            
            # Verificar si tiene registros asociados
            if USE_POSTGRES:
                cursor.execute("""
                    SELECT COUNT(*) as count FROM asistencias WHERE modalidad = (
                        SELECT nombre FROM modalidades WHERE id = %s
                    )
                """, (modalidad_id,))
            else:
                cursor.execute("""
                    SELECT COUNT(*) as count FROM asistencias WHERE modalidad = (
                        SELECT nombre FROM modalidades WHERE id = ?
                    )
                """, (modalidad_id,))
            
            row = cursor.fetchone()
            count = row['count'] if isinstance(row, dict) else row[0]
            
            if count > 0:
                return {
                    "success": False, 
                    "error": f"No se puede eliminar. Hay {count} registros usando esta modalidad"
                }, 400
            
            # Eliminar modalidad
            if USE_POSTGRES:
                cursor.execute("DELETE FROM modalidades WHERE id = %s", (modalidad_id,))
            else:
                cursor.execute("DELETE FROM modalidades WHERE id = ?", (modalidad_id,))
            
            conn.commit()
        
        return {
            "success": True,
            "message": "Modalidad eliminada exitosamente"
        }
    
    except Exception as e:
        return {"success": False, "error": str(e)}, 500

@app.route("/api/modalidades/activas")
def api_modalidades_activas():
    """API para obtener solo modalidades activas (para el formulario)"""
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("""
                SELECT nombre
                FROM modalidades
                WHERE activo = 1
                ORDER BY nombre ASC
            """)
            cursor.execute(query)
            modalidades = cursor.fetchall()
        
        return {
            "success": True,
            "modalidades": [{"value": m["nombre"], "label": m["nombre"]} for m in modalidades]
        }
    except Exception as e:
        return {"success": False, "error": str(e)}, 500

@app.route("/admin/evaluaciones_capacitadores")
def ver_evaluaciones_capacitadores():
    """
    Página para ver las evaluaciones de los capacitadores.
    Muestra estadísticas por capacitador y evaluaciones individuales.
    """
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            
            # Obtener estadísticas por capacitador
            query_stats = adapt_query("""
                SELECT 
                    a.dictado_por as capacitador,
                    COUNT(DISTINCT e.id) as total_evaluaciones,
                    ROUND(AVG(e.calidad_contenido), 2) as prom_calidad,
                    ROUND(AVG(e.metodologia), 2) as prom_metodologia,
                    ROUND(AVG(e.lenguaje_comprensible), 2) as prom_lenguaje,
                    ROUND(AVG(e.manejo_grupo), 2) as prom_manejo,
                    ROUND(AVG(e.solucion_inquietudes), 2) as prom_solucion,
                    ROUND(AVG((e.calidad_contenido + e.metodologia + e.lenguaje_comprensible + 
                               e.manejo_grupo + e.solucion_inquietudes) / 5.0), 2) as promedio_general
                FROM evaluaciones_capacitaciones e
                INNER JOIN asistencias a ON e.asistencia_id = a.id
                GROUP BY a.dictado_por
                ORDER BY promedio_general DESC, total_evaluaciones DESC
            """)
            cursor.execute(query_stats)
            stats_capacitadores = cursor.fetchall()
            
            # Obtener todas las evaluaciones con detalles
            query_evaluaciones = adapt_query("""
                SELECT 
                    e.id,
                    e.asistencia_id,
                    e.calidad_contenido,
                    e.metodologia,
                    e.lenguaje_comprensible,
                    e.manejo_grupo,
                    e.solucion_inquietudes,
                    e.comentarios,
                    e.fecha_registro,
                    a.nombre_evento,
                    a.dictado_por as capacitador,
                    a.fecha_evento,
                    a.nombre_completo as evaluador,
                    a.tipo_asistente,
                    a.programa_estudiante,
                    ROUND((e.calidad_contenido + e.metodologia + e.lenguaje_comprensible + 
                           e.manejo_grupo + e.solucion_inquietudes) / 5.0, 2) as promedio_individual
                FROM evaluaciones_capacitaciones e
                INNER JOIN asistencias a ON e.asistencia_id = a.id
                ORDER BY e.fecha_registro DESC
            """)
            cursor.execute(query_evaluaciones)
            evaluaciones = cursor.fetchall()
            
            # Calcular promedio general del sistema
            query_promedio_global = adapt_query("""
                SELECT 
                    COUNT(*) as total,
                    ROUND(AVG((calidad_contenido + metodologia + lenguaje_comprensible + 
                               manejo_grupo + solucion_inquietudes) / 5.0), 2) as promedio
                FROM evaluaciones_capacitaciones
            """)
            cursor.execute(query_promedio_global)
            promedio_global = cursor.fetchone()
        
        return render_template("evaluaciones_capacitadores.html",
                             stats_capacitadores=stats_capacitadores,
                             evaluaciones=evaluaciones,
                             promedio_global=promedio_global)
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return render_template("panel.html",
                             error=f"Error al cargar evaluaciones: {str(e)}",
                             total_registros=0)

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
                                 total_registros=0)
        else:
            return render_template("panel.html",
                                 error=resultado['mensaje'],
                                 total_registros=0)
    except Exception as e:
        return render_template("panel.html",
                             error=f"Error al limpiar datos: {str(e)}",
                             total_registros=0)

@app.route('/importar-bd-temp', methods=['GET', 'POST'])
def importar_bd_temp():
    if "usuario" not in session:
        return redirect(url_for("login"))
    
    if request.method == 'POST':
        archivo = request.files.get('backup')
        if not archivo:
            return "No se subió ningún archivo", 400
        
        import subprocess, tempfile
        db_url = os.environ.get('DATABASE_URL', '')
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        
        with tempfile.NamedTemporaryFile(suffix='.sql', delete=False, mode='wb') as f:
            archivo.save(f)
            tmp_path = f.name
        
        result = subprocess.run(
            ['psql', db_url, '-f', tmp_path],
            capture_output=True, text=True
        )
        os.unlink(tmp_path)
        
        if result.returncode != 0:
            return f"<pre>Error:\n{result.stderr}</pre>", 500
        
        return "<h2>✅ Backup restaurado exitosamente!</h2>"
    
    return '''
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="backup" accept=".sql">
            <button type="submit">Restaurar</button>
        </form>
    '''


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


# ============================================
# RUTA DE DIAGNÓSTICO (OPCIONAL)
# ============================================

@app.route("/admin/init-db")
def init_db_route():
    """Ruta para forzar inicialización de base de datos - SOLO SI ES NECESARIO"""
    # Verificar si ya hay usuarios (seguridad)
    try:
        with get_db_connection() as conn:
            cursor = get_cursor(conn)
            query = adapt_query("SELECT COUNT(*) as count FROM usuarios")
            cursor.execute(query)
            result = cursor.fetchone()
            if result['count'] > 0:
                return "⚠️ Base de datos ya inicializada. Acceso denegado.", 403
    except:
        # La tabla no existe, continuar
        pass
    
    try:
        mensaje = init_db()
        return f"✅ Base de datos inicializada correctamente.<br><br>{mensaje if mensaje else ''}<br><br><a href='/'>Ir al inicio</a>"
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        return f"❌ Error al inicializar base de datos:<br><pre>{str(e)}</pre><br><pre>{error_detail}</pre>", 500


# ============================================
# INICIALIZACIÓN AUTOMÁTICA
# ============================================

# CRÍTICO: Inicializar DB SIEMPRE (incluso con gunicorn)
# Esto se ejecuta al importar el módulo, antes de que gunicorn inicie
try:
    mensaje_limpieza_duplicados = init_db()
    if mensaje_limpieza_duplicados:
        mensaje_limpieza_global = mensaje_limpieza_duplicados
except Exception as e:
    import traceback
    traceback.print_exc()


if __name__ == "__main__":
    # Solo se ejecuta con python app.py, no con gunicorn
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=not USE_POSTGRES)