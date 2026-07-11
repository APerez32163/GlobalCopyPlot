import re
import os
import zipfile
import tempfile
import shutil
import magic
import random, string
import struct
from pypdf import PdfReader
from flask import Flask, render_template, request, redirect, url_for, flash, session
from sqlalchemy import text, func, extract
from sqlalchemy.orm.attributes import flag_modified
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta, date
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from email_service import enviar_correo_restablecimiento, enviar_correo_confirmacion, enviar_correo_listo
from extensions import db 
from models import Usuario, Pedido, DetallePedido, ArchivoPedido, Catalogo, Configuracion, ServicioImpresion, ServicioImpresionTamano
from functools import wraps 
from pptx import Presentation
from flask_wtf.csrf import CSRFProtect
from dotenv import load_dotenv
load_dotenv()

def detectar_paginas(filepath, filename):

    ext = filename.rsplit('.', 1)[-1].lower()
    
    try:
        if ext == 'pdf':
            reader = PdfReader(filepath)
            paginas = len(reader.pages)
            return paginas, None
        
        elif ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif'):
            return 1, None
        
        elif ext == 'pptx':
            prs = Presentation(filepath)
            paginas = len(prs.slides)
            return paginas, None

        elif ext == 'zip':
            tmpdir = tempfile.mkdtemp()
            try:
                with zipfile.ZipFile(filepath, 'r') as zf:
                    # Verificar que no haya subcarpetas
                    for member in zf.namelist():
                        if member.endswith('/'):  # es una carpeta
                            return None, "El ZIP no debe contener carpetas internas."
                    zf.extractall(tmpdir)

                # Listar solo archivos en la raíz del tmpdir (sin recursión)
                archivos_internos = [f for f in os.listdir(tmpdir) if os.path.isfile(os.path.join(tmpdir, f))]
                
                imagenes_encontradas = 0
                for f in archivos_internos:
                    ext_interna = f.rsplit('.', 1)[-1].lower() if '.' in f else ''
                    if ext_interna in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif'):
                        imagenes_encontradas += 1
                    else:
                        return None, f"El ZIP contiene un archivo no soportado: {f}. Solo se permiten imágenes."

                if imagenes_encontradas == 0:
                    return None, "No se encontraron imágenes en el ZIP."
                
                return imagenes_encontradas, None
            except zipfile.BadZipFile:
                return None, "El archivo ZIP está dañado."
            except Exception as e:
                print(f"Error al procesar ZIP: {e}")
                return None, f"No se pudo leer el ZIP. ({str(e)[:80]})"
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)
        
        else:
            return None, f"Formato no soportado: .{ext}. Solo se permiten PDF, Canvas, Imágenes y ZIP de imágenes."
    
    except Exception as e:
        print(f"Error al detectar páginas de {filename}: {e}")
        return None, f"No se pudo leer el archivo. Error: {str(e)[:100]}"

def validar_mime(filepath, categorias_permitidas):
    """
    Verifica que el archivo tenga un MIME válido según la lista de categorías.
    Retorna (True, None) si es válido, o (False, mensaje) si no lo es.
    """
    mime = magic.from_file(filepath, mime=True)
    
    mime_permitidos = {
        'pdf': ['application/pdf'],
        'imagen': ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff'],
        'pptx': ['application/vnd.openxmlformats-officedocument.presentationml.presentation'],
        'zip': ['application/zip']
    }
    
    permitidos = []
    for cat in categorias_permitidas:
        if cat in mime_permitidos:
            permitidos.extend(mime_permitidos[cat])
    
    if mime in permitidos:
        return True, None
    else:
        return False, f"Tipo de archivo no permitido. MIME detectado: {mime}"

def verificar_integridad(filepath, ext):
    """
    Comprueba que el archivo no está corrupto usando su librería nativa.
    Retorna (True, None) si es íntegro, o (False, mensaje) si está corrupto.
    """
    try:
        # PDF
        if ext == 'pdf':
            # Validación rápida sin cargar todo el archivo
            ok, msg = validar_pdf(filepath)
            if not ok:
                return False, msg
            # Si pasa, intentar abrir con PyPDF2 para verificar páginas
            reader = PdfReader(filepath)
            _ = len(reader.pages)
        
        # Imágenes
        elif ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif'):
            from PIL import Image
            with Image.open(filepath) as img:
                img.load()   # fuerza la carga de datos, más tolerante que verify()
        
        # PowerPoint
        elif ext == 'pptx':
            from pptx import Presentation
            _ = Presentation(filepath)  # abre la presentación
        
        # ZIP
        elif ext == 'zip':
            import zipfile
            with zipfile.ZipFile(filepath, 'r') as zf:
                _ = zf.namelist()       # lista el contenido; si es corrupto lanza excepción
        
        # Si no hay excepción, el archivo es íntegro
        return True, None

    except Exception as e:
        # Capturamos cualquier error que indique corrupción
        return False, f"El archivo está corrupto o no se puede leer. ({str(e)[:80]})"

def validar_pdf(filepath):
    """
    Validación segura de PDF sin abrir completamente el archivo con PdfReader.
    Retorna (True, None) o (False, mensaje).
    """
    try:
        # Leer solo los primeros 1024 bytes para comprobar el header
        with open(filepath, 'rb') as f:
            header = f.read(1024)
        
        # 1. Verificar firma PDF (%PDF-)
        if not header.startswith(b'%PDF-'):
            return False, "El archivo no es un PDF válido (firma incorrecta)."
        
        # 2. Verificar que no empiece con basura binaria (bytes no imprimibles)
        # Un PDF real empieza con %PDF- seguido de la versión (ej: 1.4)
        # Si el header contiene muchos bytes nulos o basura, es sospechoso
        non_printable = sum(1 for b in header[:20] if b < 32 and b not in (10, 13, 9))
        if non_printable > 5:  # más de 5 bytes no imprimibles en el header es sospechoso
            return False, "El archivo PDF parece corrupto o contiene datos binarios anómalos."
        
        # 3. Leer últimos 1024 bytes para verificar el trailer
        f.seek(0, 2)  # final del archivo
        size = f.tell()
        if size > 10 * 1024 * 1024:  # 10 MB máximo para PDF
            return False, "El archivo PDF es demasiado grande (máx. 10 MB)."
        
        f.seek(max(0, size - 1024))
        trailer = f.read(1024)
        if b'%%EOF' not in trailer:
            return False, "El archivo PDF está incompleto (falta marca de fin)."
        
        return True, None
    except Exception as e:
        return False, f"No se pudo validar el PDF: {str(e)[:80]}"

app = Flask(__name__)

UPLOAD_FOLDER = 'static/uploads'
UPLOAD_FOLDER_IMPRESION = os.path.join(UPLOAD_FOLDER, 'impresion')
UPLOAD_FOLDER_COMPROBANTES = os.path.join(UPLOAD_FOLDER, 'comprobantes')

os.makedirs(UPLOAD_FOLDER_COMPROBANTES, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(UPLOAD_FOLDER_IMPRESION, exist_ok=True)
os.makedirs(os.path.join(UPLOAD_FOLDER_IMPRESION, 'temp'), exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL')
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'clave-temporal-solo-dev')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_pre_ping': True,
    'pool_recycle': 300,
    'pool_timeout': 30,
}

db.init_app(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

serializer = URLSafeTimedSerializer(app.config['SECRET_KEY'])

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.ES_ADMIN:
            flash('Acceso denegado. Solo administradores.', 'danger')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def operador_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or (not current_user.ES_OPERADOR and not current_user.ES_ADMIN):
            flash('Acceso denegado. Solo operadores.', 'danger')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

# ------------------------------------------------------------
# Página de inicio
# ------------------------------------------------------------

@app.route('/')
def index():
    # Obtener configuración del negocio (valores de la tabla configuracion)
    config = {}
    try:
        configuraciones = Configuracion.query.all()
        config = {c.CLAVE: c.VALOR for c in configuraciones}
    except Exception as e:
        print("Error al cargar configuración:", e)

    # Obtener imágenes del catálogo (si usas la tabla catalogo)
    try:
        with db.engine.connect() as conn:
            result = conn.execute(text("SELECT IMAGEN FROM catalogo ORDER BY ORDEN"))
            imagenes = [row[0] for row in result]
    except Exception as e:
        imagenes = []
        print("Error al consultar catálogo:", e)

    return render_template('index.html', imagenes=imagenes, config=config)

# ------------------------------------------------------------
# Login
# ------------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.args.get('logout'):
        flash('Has cerrado sesión.', 'info')
    if request.args.get('registered'):
        flash('Cuenta creada correctamente. Ahora inicia sesión.', 'success')
    if request.args.get('reset'):
        flash('Contraseña restablecida correctamente. Inicia sesión.', 'success')
    if request.args.get('confirmed'):
        flash('Correo confirmado. Ya puedes iniciar sesión.', 'success')
    if request.method == 'POST':
        cedula = request.form.get('usuario', '').strip()
        cedula = re.sub(r'^[VEve]-', '', cedula).strip()
        contrasena = request.form.get('contrasena', '')

        error_cedula = validar_cedula(cedula)
        if error_cedula:
            flash(error_cedula, 'danger')
            return render_template('login.html')

        user = Usuario.query.filter_by(ID_USUARIO=cedula).first()

        # --- NUEVO: Verificar si el usuario está bloqueado temporalmente ---
        if user and user.BLOQUEADO_HASTA and user.BLOQUEADO_HASTA > datetime.now():
            minutos_restantes = int((user.BLOQUEADO_HASTA - datetime.now()).total_seconds() // 60) + 1
            flash(f'Demasiados intentos fallidos. Espera {minutos_restantes} minuto(s) o restablece tu contraseña.', 'danger')
            return render_template('login.html')

        if user and check_password_hash(user.CONTRASEÑA, contrasena):
            if not user.CONFIRMADO:
                flash('Debes confirmar tu correo electrónico antes de iniciar sesión.', 'warning')
            else:
                # --- NUEVO: Resetear intentos fallidos y desbloquear ---
                user.INTENTOS_FALLIDOS = 0
                user.BLOQUEADO_HASTA = None
                db.session.commit()

                login_user(user)
                if not user.ES_ADMIN and not user.ES_OPERADOR:
                    flash('¡Ahora puedes realizar peticiones en línea desde el centro de impresión!', 'info')
                flash(f'¡Bienvenido {user.NOMBRE}!', 'success')
                return redirect(url_for('index'))
        else:
            # --- NUEVO: Incrementar contador de intentos fallidos ---
            if user:
                user.INTENTOS_FALLIDOS = (user.INTENTOS_FALLIDOS or 0) + 1
                if user.INTENTOS_FALLIDOS >= 5:
                    user.BLOQUEADO_HASTA = datetime.now() + timedelta(minutes=1)
                    user.INTENTOS_FALLIDOS = 0  # reiniciar contador para cuando se desbloquee
                    db.session.commit()
                    flash('Has superado el límite de intentos. Espera 1 minuto o restablece tu contraseña.', 'danger')
                    return redirect(url_for('index', bloqueado=1))
                else:
                    db.session.commit()
                    flash('Cédula o contraseña incorrectos.', 'danger')
            else:
                flash('Cédula o contraseña incorrectos.', 'danger')

    return render_template('login.html')
# ------------------------------------------------------------
# Registro
# ------------------------------------------------------------
@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        nombre = request.form.get('nombre', '').strip()
        apellido = request.form.get('apellido', '').strip()
        email = request.form.get('email', '').strip().lower()
        cedula = request.form.get('usuario_id', '').strip()
        cedula = re.sub(r'^[VEve]-?', '', cedula).strip() 
        cedula = re.sub(r'^[VEve]-', '', cedula).strip()
        telefono = request.form.get('telefono', '').strip()
        contrasena = request.form.get('contrasena', '')
        pregunta1 = request.form.get('pregunta1', '').strip()
        respuesta1 = request.form.get('respuesta1', '').strip()
        pregunta2 = request.form.get('pregunta2', '').strip()
        respuesta2 = request.form.get('respuesta2', '').strip()

        field_errors = {}

        # Validaciones existentes
        err_nombre = validar_nombre_apellido(nombre, "Nombre")
        if err_nombre:
            field_errors['nombre'] = err_nombre

        err_apellido = validar_nombre_apellido(apellido, "Apellido")
        if err_apellido:
            field_errors['apellido'] = err_apellido

        err_email = validar_email(email)
        if err_email:
            field_errors['email'] = err_email

        err_cedula = validar_cedula(cedula)
        if err_cedula:
            field_errors['usuario_id'] = err_cedula

        err_telefono = validar_telefono(telefono)
        if err_telefono:
            field_errors['telefono'] = err_telefono

        err_password = validar_contraseña(contrasena)
        if err_password:
            field_errors['contrasena'] = err_password

        # Validar preguntas de seguridad
        if not pregunta1:
            pregunta1 = '¿Nombre de tu primera mascota?'
        else:
            pregunta1 = pregunta1.strip()
            if len(pregunta1) < 4:
                field_errors['pregunta1'] = 'La pregunta debe tener al menos 4 caracteres.'
            elif len(pregunta1) > 20:
                field_errors['pregunta1'] = 'La pregunta no puede tener más de 20 caracteres.'
            elif not re.match(r'^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s?]+$', pregunta1):
                field_errors['pregunta1'] = 'Solo se permiten letras y espacios.'

        if not respuesta1:
            field_errors['respuesta1'] = 'Debes escribir una respuesta.'
        else:
            respuesta1 = respuesta1.strip()
            if len(respuesta1) > 20:
                field_errors['respuesta1'] = 'La respuesta no puede tener más de 20 caracteres.'

        if not pregunta2:
            pregunta2 = '¿Ciudad donde naciste?'
        else:
            pregunta2 = pregunta2.strip()
            if len(pregunta2) < 4:
                field_errors['pregunta2'] = 'La pregunta debe tener al menos 4 caracteres.'
            elif len(pregunta2) > 20:
                field_errors['pregunta2'] = 'La pregunta no puede tener más de 20 caracteres.'
            elif not re.match(r'^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s?]+$', pregunta2):
                field_errors['pregunta2'] = 'Solo se permiten letras y espacios.'

        if not respuesta2:
            field_errors['respuesta2'] = 'Debes escribir una respuesta.'
        else:
            respuesta2 = respuesta2.strip()
            if len(respuesta2) > 20:
                field_errors['respuesta2'] = 'La respuesta no puede tener más de 20 caracteres.'

        # Validar que las preguntas no sean iguales entre sí
        if pregunta1 and pregunta2 and pregunta1.lower() == pregunta2.lower():
            field_errors['pregunta2'] = 'Las preguntas no pueden ser iguales.'

        # Validar que las respuestas no sean iguales entre sí
        if respuesta1 and respuesta2 and respuesta1.lower() == respuesta2.lower():
            field_errors['respuesta2'] = 'Las respuestas no pueden ser iguales.'

        if field_errors:
            return render_template('registro.html', form_data=request.form, field_errors=field_errors)

        # Verificar duplicados (igual que antes)
        if Usuario.query.filter_by(ID_USUARIO=cedula).first():
            flash('Esta cédula ya está registrada.', 'danger')
            return render_template('registro.html', form_data=request.form, field_errors={})
        if Usuario.query.filter_by(EMAIL=email).first():
            flash('El correo ya está en uso.', 'danger')
            return render_template('registro.html', form_data=request.form, field_errors={})
        if Usuario.query.filter_by(TELEFONO=telefono).first():
            flash('El número de teléfono ya está registrado.', 'danger')
            return render_template('registro.html', form_data=request.form, field_errors={})

        # Crear usuario con nuevas columnas
        nuevo = Usuario(
            ID_USUARIO=cedula,
            NOMBRE=nombre,
            APELLIDO=apellido,
            EMAIL=email,
            TELEFONO=telefono,
            CONTRASEÑA=generate_password_hash(contrasena),
            CONFIRMADO=False,
            PREGUNTA1=pregunta1,
            RESPUESTA1=respuesta1.strip().upper(), 
            PREGUNTA2=pregunta2,
            RESPUESTA2=respuesta2.strip().upper()
        )
        db.session.add(nuevo)
        db.session.commit()

        token = serializer.dumps(nuevo.ID, salt='confirmar-email')
        exito, _ = enviar_correo_confirmacion(email, token)
        if exito:
            flash('Cuenta creada. Te hemos enviado un enlace de confirmación a tu correo.', 'success')
        else:
            flash('Cuenta creada, pero no pudimos enviar el correo de confirmación. Contacta soporte.', 'warning')
        return redirect(url_for('login'))

    return render_template('registro.html', form_data=None, field_errors={})

# ------------------------------------------------------------
# confirmar el correo
# ------------------------------------------------------------

@app.route('/confirmar/<token>')
def confirmar_email(token):
    try:
        user_id = serializer.loads(token, salt='confirmar-email', max_age=86400)  # 24h
    except SignatureExpired:
        flash('El enlace de confirmación ha expirado. Regístrate de nuevo.', 'danger')
        return redirect(url_for('registro'))
    except BadSignature:
        flash('Enlace inválido.', 'danger')
        return redirect(url_for('registro'))

    user = Usuario.query.get(user_id)
    if not user:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('registro'))

    if user.CONFIRMADO:
        flash('La cuenta ya estaba confirmada. Inicia sesión.', 'info')
    else:
        user.CONFIRMADO = True
        db.session.commit()
    return redirect(url_for('login', confirmed=1))

# ------------------------------------------------------------
# recuperar por preguntas
# ------------------------------------------------------------

@app.route('/recuperar-preguntas', methods=['GET', 'POST'])
def recuperar_preguntas():
    cedula = session.get('recuperar_cedula')
    if not cedula:
        flash('Sesión expirada. Intenta de nuevo.', 'danger')
        return redirect(url_for('recuperar'))

    user = Usuario.query.filter_by(ID_USUARIO=cedula).first()
    if not user:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('recuperar'))

    # Límite de intentos: 3 fallos y se bloquea
    intentos = session.get('intentos_pregunta', 0)

    if request.method == 'POST':
        respuesta = request.form.get('respuesta', '').strip().upper()
        pregunta_idx = session.get('pregunta_aleatoria')

        if pregunta_idx is None:
            flash('Error de sesión. Intenta de nuevo.', 'danger')
            return redirect(url_for('recuperar'))

        respuesta_real = user.RESPUESTA1 if pregunta_idx == 0 else user.RESPUESTA2
        if respuesta_real:
            respuesta_real = respuesta_real.strip().upper()

        if respuesta == respuesta_real:
            # Correcta: reiniciar intentos y redirigir
            session['reset_cedula'] = cedula
            session.pop('pregunta_aleatoria', None)
            session.pop('intentos_pregunta', None)
            session.pop('recuperar_cedula', None)
            return redirect(url_for('cambiar_contrasena'))
        else:
            # Fallo: alternar a la otra pregunta
            session['intentos_pregunta'] = intentos + 1
            if session['intentos_pregunta'] >= 3:
                flash('Demasiados intentos fallidos. Solicita un nuevo enlace.', 'danger')
                session.pop('pregunta_aleatoria', None)
                session.pop('intentos_pregunta', None)
                session.pop('recuperar_cedula', None)
                return redirect(url_for('recuperar'))

            # Alternar pregunta
            nuevo_idx = 1 if session['pregunta_aleatoria'] == 0 else 0
            session['pregunta_aleatoria'] = nuevo_idx
            pregunta = user.PREGUNTA1 if nuevo_idx == 0 else user.PREGUNTA2
            flash('Respuesta incorrecta. Intenta con la otra pregunta.', 'warning')
            return render_template('preguntas_seguridad.html', pregunta=pregunta)

    # GET: elegir pregunta al azar (solo la primera vez)
    if 'pregunta_aleatoria' not in session:
        import random
        session['pregunta_aleatoria'] = random.randint(0, 1)
        session['intentos_pregunta'] = 0

    idx = session['pregunta_aleatoria']
    pregunta = user.PREGUNTA1 if idx == 0 else user.PREGUNTA2

    return render_template('preguntas_seguridad.html', pregunta=pregunta)

# ------------------------------------------------------------
# Recuperar acceso
# ------------------------------------------------------------

@app.route('/recuperar', methods=['GET', 'POST'])
def recuperar():
    if request.method == 'POST':
        # --- SI YA HAY CÉDULA EN SESIÓN (usuario eligió "Enviar enlace al correo") ---
        if session.get('recuperar_cedula'):
            cedula = session['recuperar_cedula']
            user = Usuario.query.filter_by(ID_USUARIO=cedula).first()
            if user:
                try:
                    user.EMAIL.encode('ascii')
                except UnicodeEncodeError:
                    flash('El correo asociado contiene caracteres no permitidos.', 'danger')
                    return redirect(url_for('recuperar'))
                token = serializer.dumps(user.ID, salt='recuperar-contrasena')
                exito, error_msg = enviar_correo_restablecimiento(user.EMAIL, token)
                if exito:
                    flash('Te hemos enviado un enlace de recuperación a tu correo.', 'success')
                else:
                    flash(f'No se pudo enviar el correo: {error_msg}', 'danger')
                session.pop('recuperar_cedula', None)
                return redirect(url_for('login'))
            else:
                flash('Usuario no encontrado.', 'danger')
                return redirect(url_for('recuperar'))

        # --- PASO 1: Ingreso de cédula ---
        cedula = request.form.get('cedula', '').strip()
        cedula = re.sub(r'^[VEve]-', '', cedula).strip()
        error_cedula = validar_cedula(cedula)
        if error_cedula:
            flash(error_cedula, 'danger')
            return redirect(url_for('recuperar'))

        user = Usuario.query.filter_by(ID_USUARIO=cedula).first()
        if not user:
            flash('No existe una cuenta con esa cédula.', 'danger')
            return redirect(url_for('recuperar'))

        try:
            user.EMAIL.encode('ascii')
        except UnicodeEncodeError:
            flash('El correo asociado contiene caracteres no permitidos.', 'danger')
            return redirect(url_for('recuperar'))

        # Guardar cédula en sesión y mostrar opciones (SIEMPRE)
        session['recuperar_cedula'] = cedula
        tiene_preguntas = bool(user.PREGUNTA1 and user.PREGUNTA2)
        return render_template('recuperar.html', paso=2, cedula=cedula, tiene_preguntas=tiene_preguntas)

    # GET
    return render_template('recuperar.html', paso=1)

# ------------------------------------------------------------
# reset-password/<token
# ------------------------------------------------------------

@app.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    try:
        user_id = serializer.loads(token, salt='recuperar-contrasena', max_age=1800)
    except SignatureExpired:
        flash('El enlace ha expirado.', 'danger')
        return redirect(url_for('recuperar'))
    except BadSignature:
        flash('Enlace inválido.', 'danger')
        return redirect(url_for('recuperar'))

    user = Usuario.query.get(user_id)
    if not user:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('recuperar'))

    if request.method == 'POST':
        nueva = request.form.get('nueva', '')
        repetir = request.form.get('repetir', '')

        if nueva != repetir:
            flash('Las contraseñas no coinciden.', 'warning')
            return render_template('cambiar_contraseña.html', token=token)

        # Validar la nueva contraseña con los criterios definidos
        error_pass = validar_contraseña(nueva)
        if error_pass:
            flash(error_pass, 'warning')
            return render_template('cambiar_contraseña.html', token=token)

        # Actualizar la contraseña
        user.CONTRASEÑA = generate_password_hash(nueva)
        db.session.commit()
        flash('Contraseña actualizada correctamente. Inicia sesión.', 'success')

        # Limpiar sesión
        session.pop('reset_cedula', None)
        session.pop('reset_code', None)
        session.pop('reset_code_time', None)
        session.pop('recuperar_cedula', None)
        session.pop('recuperar_email_real', None)
        return redirect(url_for('login'))

    return render_template('cambiar_contraseña.html', token=token)

# ------------------------------------------------------------
# Cambiar contraseña
# ------------------------------------------------------------

@app.route('/cambiar-contrasena', methods=['GET', 'POST'])
def cambiar_contrasena():
    # Solo necesitamos la cédula en sesión
    if 'reset_cedula' not in session:
        flash('Acceso no autorizado.', 'danger')
        return redirect(url_for('recuperar'))

    cedula = session['reset_cedula']
    user = Usuario.query.filter_by(ID_USUARIO=cedula).first()
    if not user:
        flash('Usuario no encontrado.', 'danger')
        return redirect(url_for('recuperar'))

    if request.method == 'POST':
        nueva = request.form.get('nueva', '')
        repetir = request.form.get('repetir', '')

        if nueva != repetir:
            flash('Las contraseñas no coinciden.', 'warning')
        else:
            error_pass = validar_contraseña(nueva)
            if error_pass:
                flash(error_pass, 'warning')
            else:
                user.CONTRASEÑA = generate_password_hash(nueva)
                db.session.commit()

                # Limpiar TODAS las variables de sesión relacionadas con recuperación
                session.pop('reset_cedula', None)
                session.pop('reset_code', None)
                session.pop('reset_code_time', None)
                session.pop('recuperar_cedula', None)
                session.pop('recuperar_email_real', None)
                session.pop('pregunta_aleatoria', None)

                return redirect(url_for('login', reset=1))

    return render_template('cambiar_contraseña.html')

# ------------------------------------------------------------
# PANEL DE ADMINISTRACIÓN
# ------------------------------------------------------------

@app.route('/admin')
@login_required
@admin_required
def panel_admin():
    # Leer el período guardado en sesión desde Reportes/Ingresos (por defecto mensual)
    periodo = session.get('periodo_ingresos', 'mensual')
    hoy = datetime.now()
    
    # Calcular rango de fechas según el período
    if periodo == 'diario':
        desde = hoy.replace(hour=0, minute=0, second=0, microsecond=0)
        hasta = hoy.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif periodo == 'semanal':
        desde = hoy - timedelta(days=hoy.weekday())
        desde = desde.replace(hour=0, minute=0, second=0, microsecond=0)
        hasta = desde + timedelta(days=6, hours=23, minutes=59, seconds=59)
    elif periodo == 'mensual':
        desde = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        if hoy.month == 12:
            hasta = hoy.replace(year=hoy.year+1, month=1, day=1) - timedelta(seconds=1)
        else:
            hasta = hoy.replace(month=hoy.month+1, day=1) - timedelta(seconds=1)
    elif periodo == 'anual':
        desde = hoy.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
        hasta = hoy.replace(month=12, day=31, hour=23, minute=59, second=59, microsecond=999999)
    else:
        # Por defecto mensual
        desde = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        if hoy.month == 12:
            hasta = hoy.replace(year=hoy.year+1, month=1, day=1) - timedelta(seconds=1)
        else:
            hasta = hoy.replace(month=hoy.month+1, day=1) - timedelta(seconds=1)
    
    # Calcular ingresos SOLO de pagos confirmados y posteriores
    ingresos = db.session.query(func.sum(Pedido.TOTAL)).filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo', 'Entregado']),
        Pedido.FECHA >= desde,
        Pedido.FECHA <= hasta
    ).scalar() or 0
    
    pendientes = Pedido.query.filter_by(ESTADO='Pendiente de pago').count()
    nuevos_pendientes = Pedido.query.filter_by(ESTADO='Pendiente de pago', VISTO_ADMIN=False).count()
    nuevos_admin = Pedido.query.filter_by(VISTO_ADMIN=False).count()
    en_proceso = Pedido.query.filter(
        Pedido.ESTADO.in_(['Esperando validación', 'Pago confirmado'])
    ).count()
    listos = Pedido.query.filter_by(ESTADO='Listo').count()

    nuevos_proceso = Pedido.query.filter(
        Pedido.ESTADO.in_(['Esperando validación', 'Pago confirmado']),
        Pedido.VISTO_ADMIN == False
    ).count()
    
    return render_template('admin/dashboard.html',
                         pendientes=pendientes,
                         proceso=en_proceso,
                         listos=listos,
                         ingresos=ingresos,
                         periodo_actual=periodo,
                         nuevos_pendientes=nuevos_pendientes,
                         nuevos_proceso=nuevos_proceso)

# ------------------------------------------------------------
# PANEL ADMIN – SOLICITUDES
# ------------------------------------------------------------

@app.route('/admin/solicitudes')
@login_required
@admin_required
def admin_solicitudes():
    estado = request.args.get('estado')
    if estado:
        estados = [e.strip() for e in estado.split(',') if e.strip()]
        if len(estados) > 1:
            pedidos = Pedido.query.filter(Pedido.ESTADO.in_(estados)).order_by(Pedido.FECHA.desc()).all()
        else:
            pedidos = Pedido.query.filter_by(ESTADO=estados[0]).order_by(Pedido.FECHA.desc()).all()
    else:
        pedidos = Pedido.query.order_by(Pedido.FECHA.desc()).all()

    pedidos_con_usuarios = []
    for pedido in pedidos:
        usuario = Usuario.query.get(pedido.ID_USUARIO)
        pedidos_con_usuarios.append({'pedido': pedido, 'usuario': usuario})

    return render_template('admin/solicitudes.html',
                           pedidos=pedidos_con_usuarios,
                           estado_filtro=estado)

@app.route('/admin/solicitud/<int:pedido_id>')
@login_required
@admin_required
def admin_solicitud_detalle(pedido_id):
    pedido = db.session.get(Pedido, pedido_id) or abort(404)
    usuario = db.session.get(Usuario, pedido.ID_USUARIO)
    detalles = DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).all()
    archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).all()

    # Obtener la lista de archivos con sus configuraciones
    if pedido.DETALLE_ARCHIVOS:
        # Caso múltiple
        archivos_info = pedido.DETALLE_ARCHIVOS
    else:
        # Caso único: creamos una lista de un solo elemento con los datos del pedido
        archivos_info = [{
            'nombre': archivos[0].NOMBRE_ARCHIVO if archivos else 'Sin archivo',
            'paginas': pedido.PAGINAS,
            'servicio_id': pedido.SERVICIO_ID,
            'tamano': pedido.TAMANO,
            'paginas_color': pedido.PAGINAS_COLOR,    
            'comentarios': pedido.COMENTARIOS
        }]

    # Enriquecer archivos_info con los nombres de servicio y precio
    for item in archivos_info:
        servicio_id = item.get('servicio_id')
        if servicio_id:
            servicio = ServicioImpresion.query.get(servicio_id)
            item['servicio_nombre'] = servicio.TITULO if servicio else 'Desconocido'
            # Obtener precio del tamaño
            tamano_nombre = item.get('tamano')
            if tamano_nombre and servicio:
                tamano_obj = ServicioImpresionTamano.query.filter_by(
                    SERVICIO_ID=servicio_id, NOMBRE=tamano_nombre
                ).first()
                item['precio'] = float(tamano_obj.PRECIO_BN) if tamano_obj else None
        else:
            item['servicio_nombre'] = None
            item['precio'] = None

    # Marcar como visto por el admin
    if not pedido.VISTO_ADMIN:
        pedido.VISTO_ADMIN = True
        db.session.commit()

    historial = [
        {"fecha": pedido.FECHA, "estado": "Pedido creado"},
        {"fecha": datetime.now(), "estado": pedido.ESTADO}
    ]

    return render_template('admin/solicitud_detalle.html', 
                           pedido=pedido,
                           usuario=usuario,
                           detalles=detalles,
                           archivos=archivos,
                           historial=historial,
                           archivos_info=archivos_info) 

@app.route('/admin/solicitud/<int:pedido_id>/confirmar-pago', methods=['POST'])
@login_required
@admin_required
def admin_confirmar_pago(pedido_id):
    pedido = db.session.get(Pedido, pedido_id)
    if pedido.ESTADO != 'Esperando validación':
        flash('El pedido no está en espera de validación.', 'warning')
        return redirect(url_for('admin_solicitud_detalle', pedido_id=pedido.ID))

    pedido.ESTADO = 'Pago confirmado'
    # Generar código de ticket si no existe
    if not pedido.CODIGO_TICKET:
        codigo = f"TICK-{pedido.ID}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"
        pedido.CODIGO_TICKET = codigo
    db.session.commit()
    flash('Pago confirmado. El cliente ya puede acceder a su ticket.', 'success')
    return redirect(url_for('admin_solicitud_detalle', pedido_id=pedido.ID))

@app.route('/admin/solicitud/<int:pedido_id>/actualizar', methods=['POST'])
@login_required
@admin_required
def admin_actualizar_estado(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    nuevo_estado = request.form.get('estado')
    estados_validos = ['Pendiente de pago', 'Esperando validación', 'Pago confirmado', 'Listo', 'Entregado', 'Cancelado']
    
    if nuevo_estado not in estados_validos:
        flash('Estado no válido.', 'danger')
        return redirect(url_for('admin_solicitud_detalle', pedido_id=pedido.ID))

    pedido.ESTADO = nuevo_estado
    db.session.commit()

    # Si el estado es "Listo", enviar correo al cliente
    if nuevo_estado == 'Listo':
        usuario = Usuario.query.get(pedido.ID_USUARIO)
        if usuario:
            ticket_url = url_for('ticket_impresion', pedido_id=pedido.ID, _external=True)
            # Generar código de ticket si no existe
            if not pedido.CODIGO_TICKET:
                import random, string
                codigo = f"TICK-{pedido.ID}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"
                pedido.CODIGO_TICKET = codigo
                db.session.commit()
            exito, error = enviar_correo_listo(usuario.EMAIL, usuario.NOMBRE, ticket_url)
            if exito:
                flash('Estado actualizado y correo enviado al cliente.', 'success')
            else:
                flash(f'Estado actualizado pero no se pudo enviar el correo: {error}', 'warning')
        else:
            flash('Estado actualizado.', 'success')
    else:
        flash('Estado actualizado.', 'success')

    return redirect(url_for('admin_solicitud_detalle', pedido_id=pedido.ID))

@app.route('/admin/solicitud/<int:pedido_id>/eliminar', methods=['POST'])
@login_required
@admin_required
def admin_eliminar_solicitud(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    
    # Eliminar primero los detalles asociados
    DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).delete()
    
    # Eliminar archivos asociados
    ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).delete()
    
    # Eliminar el pedido
    db.session.delete(pedido)
    db.session.commit()
    
    flash('Solicitud eliminada correctamente.', 'success')
    return redirect(url_for('admin_solicitudes'))

# ------------------------------------------------------------
# PANEL ADMIN – CATALOGO
# ------------------------------------------------------------

@app.route('/admin/catalogos')
@login_required
@admin_required
def admin_catalogos():
    tab = request.args.get('tab', 'servicios')
    imagenes = Catalogo.query.order_by(Catalogo.ORDEN).all()          
    servicios = ServicioImpresion.query.order_by(ServicioImpresion.TITULO).all()

    return render_template('admin/catalogos.html', 
                           imagenes=imagenes,
                           servicios=servicios,
                           tab_activa=tab)

@app.route('/admin/catalogos/agregar', methods=['POST'])
@login_required
@admin_required
def admin_agregar_imagen():
    imagen = request.files.get('imagen')
    if not imagen or imagen.filename == '':
        flash('Debes seleccionar una imagen.', 'danger')
        return redirect(url_for('admin_catalogos'))

    filename = secure_filename(imagen.filename)
    filepath = os.path.join('static/img', filename)
    imagen.save(filepath)

    # Verificar si ya existe en la base de datos
    if Catalogo.query.filter_by(IMAGEN=filename).first():
        flash('Ya existe una imagen con ese nombre.', 'warning')
        return redirect(url_for('admin_catalogos'))

    nuevo = Catalogo(IMAGEN=filename)
    db.session.add(nuevo)
    db.session.commit()
    flash('Imagen agregada correctamente.', 'success')
    return redirect(url_for('admin_catalogos'))

@app.route('/admin/catalogos/eliminar/<int:imagen_id>', methods=['POST'])
@login_required
@admin_required
def admin_eliminar_imagen(imagen_id):
    imagen = Catalogo.query.get_or_404(imagen_id)
    # Eliminar archivo físico
    ruta = os.path.join('static/img', imagen.IMAGEN)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(imagen)
    db.session.commit()
    flash('Imagen eliminada correctamente.', 'success')
    return redirect(url_for('admin_catalogos'))

@app.route('/admin/catalogos/actualizar/<int:imagen_id>', methods=['POST'])
@login_required
@admin_required
def admin_actualizar_imagen(imagen_id):
    imagen = Catalogo.query.get_or_404(imagen_id)
    nueva_imagen = request.files.get('imagen')
    if nueva_imagen and nueva_imagen.filename:
        # Eliminar la anterior
        ruta_anterior = os.path.join('static/img', imagen.IMAGEN)
        if os.path.exists(ruta_anterior):
            os.remove(ruta_anterior)
        # Guardar la nueva
        filename = secure_filename(nueva_imagen.filename)
        nueva_imagen.save(os.path.join('static/img', filename))
        imagen.IMAGEN = filename
        db.session.commit()
        flash('Imagen actualizada correctamente.', 'success')
    else:
        flash('No se seleccionó una nueva imagen.', 'info')
    return redirect(url_for('admin_catalogos'))

@app.route('/admin/catalogos/orden-actual')
@login_required
@admin_required
def admin_orden_actual():
    imagenes = Catalogo.query.order_by(Catalogo.ORDEN).all()
    orden = [{'id': img.ID, 'nombre': img.IMAGEN} for img in imagenes]
    return {'orden': orden}

@app.route('/admin/catalogos/reordenar', methods=['POST'])
@login_required
@admin_required
def admin_reordenar_catalogo():
    data = request.get_json()
    orden_ids = data.get('orden', [])
    for idx, img_id in enumerate(orden_ids):
        Catalogo.query.filter_by(ID=img_id).update({'ORDEN': idx})
    db.session.commit()
    return {'success': True}

# SERVICIOS DE IMPRESIÓN
@app.route('/admin/servicios-impresion')
@login_required
@admin_required
def admin_servicios_impresion():
    servicios = ServicioImpresion.query.order_by(ServicioImpresion.TITULO).all()
    return render_template('admin/servicios_impresion.html', servicios=servicios)

@app.route('/admin/servicio-impresion/agregar', methods=['POST'])
@login_required
@admin_required
def admin_agregar_servicio_impresion():
    titulo = request.form.get('titulo', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    activo = 'activo' in request.form
    es_mixto = 'es_mixto' in request.form

    if not titulo:
        flash('El título es obligatorio.', 'danger')
        return redirect(url_for('admin_catalogos', tab='serviciosImpresion'))

    nuevo = ServicioImpresion(TITULO=titulo, 
                              DESCRIPCION=descripcion, 
                              ACTIVO=activo,
                              ES_MIXTO=es_mixto)
    
    db.session.add(nuevo)
    db.session.flush() 

    # Procesar tamaños
    nombres = request.form.getlist('tam_nombre[]')
    precios_bn = request.form.getlist('tam_precio_bn[]')
    precios_color = request.form.getlist('tam_precio_color[]')
    
    # Iterar con índice para poder acceder al precio color correcto
    for i, nom in enumerate(nombres):
        pre_bn = precios_bn[i] if i < len(precios_bn) else ''
        if nom.strip() and pre_bn.strip():
            try:
                precio_bn = float(pre_bn)
            except ValueError:
                flash('Precio B/N inválido.', 'danger')
                db.session.rollback()
                return redirect(url_for('admin_catalogos', tab='serviciosImpresion'))
            
            # Precio color: tomar el correspondiente, o si no existe, usar el B/N
            pre_color = precios_color[i] if i < len(precios_color) else pre_bn
            try:
                precio_color = float(pre_color)
            except ValueError:
                precio_color = precio_bn
            
            t = ServicioImpresionTamano(
                SERVICIO_ID=nuevo.ID,
                NOMBRE=nom.strip(),
                PRECIO_BN=precio_bn,
                PRECIO_COLOR=precio_color
            )
            db.session.add(t)

    db.session.commit()
    flash('Servicio de impresión agregado correctamente.', 'success')
    return redirect(url_for('admin_catalogos', tab='serviciosImpresion'))

@app.route('/admin/servicio-impresion/<int:servicio_id>/editar', methods=['POST'])
@login_required
@admin_required
def admin_editar_servicio_impresion(servicio_id):
    servicio = ServicioImpresion.query.get_or_404(servicio_id)
    servicio.TITULO = request.form.get('titulo', '').strip()
    servicio.DESCRIPCION = request.form.get('descripcion', '').strip()
    servicio.ACTIVO = 'activo' in request.form
    servicio.ES_MIXTO = 'es_mixto' in request.form

    # Reemplazar tamaños existentes por los nuevos enviados
    ServicioImpresionTamano.query.filter_by(SERVICIO_ID=servicio_id).delete()
    nombres = request.form.getlist('tam_nombre[]')
    precios_bn = request.form.getlist('tam_precio_bn[]')
    precios_color = request.form.getlist('tam_precio_color[]')
    
    for i, nom in enumerate(nombres):
        pre_bn = precios_bn[i] if i < len(precios_bn) else ''
        if nom.strip() and pre_bn.strip():
            pre_color = precios_color[i] if i < len(precios_color) else pre_bn
            try:
                precio_color = float(pre_color)
            except ValueError:
                precio_color = float(pre_bn)
            t = ServicioImpresionTamano(
                SERVICIO_ID=servicio_id,
                NOMBRE=nom.strip(),
                PRECIO_BN=float(pre_bn),
                PRECIO_COLOR=precio_color
            )
            db.session.add(t)
            
    db.session.commit()

    flash('Servicio actualizado.', 'success')
    return redirect(url_for('admin_catalogos', tab='serviciosImpresion'))

@app.route('/admin/servicio-impresion/<int:servicio_id>/eliminar', methods=['POST'])
@login_required
@admin_required
def admin_eliminar_servicio_impresion(servicio_id):
    servicio = ServicioImpresion.query.get_or_404(servicio_id)
    
    # Desvincular pedidos que usan este servicio (poner SERVICIO_ID a NULL)
    Pedido.query.filter_by(SERVICIO_ID=servicio_id).update({'SERVICIO_ID': None})
    
    db.session.delete(servicio)
    db.session.commit()
    flash('Servicio de impresión eliminado.', 'success')
    return redirect(url_for('admin_catalogos', tab='serviciosImpresion'))

# ------------------------------------------------------------
# PANEL ADMIN – USUARIOS
# ------------------------------------------------------------

@app.route('/admin/usuarios')
@login_required
@admin_required
def admin_usuarios():
    busqueda = request.args.get('buscar', '').strip()
    if busqueda:
        # Buscar por cédula, nombre, apellido o email
        usuarios = Usuario.query.filter(
            (Usuario.ID_USUARIO.contains(busqueda)) |
            (Usuario.NOMBRE.contains(busqueda)) |
            (Usuario.APELLIDO.contains(busqueda)) |
            (Usuario.EMAIL.contains(busqueda))
        ).order_by(Usuario.APELLIDO, Usuario.NOMBRE).all()
    else:
        usuarios = Usuario.query.order_by(Usuario.APELLIDO, Usuario.NOMBRE).all()
    return render_template('admin/usuarios.html', usuarios=usuarios, busqueda=busqueda)

@app.route('/admin/usuario/<int:user_id>')
@login_required
@admin_required
def admin_usuario_detalle(user_id):
    usuario = Usuario.query.get_or_404(user_id)
    pedidos = Pedido.query.filter_by(ID_USUARIO=usuario.ID).order_by(Pedido.FECHA.desc()).all()
    return render_template('admin/usuario_detalle.html', usuario=usuario, pedidos=pedidos)

@app.route('/admin/usuario/<int:user_id>/eliminar', methods=['POST'])
@login_required
@admin_required
def admin_eliminar_usuario(user_id):
    usuario = Usuario.query.get_or_404(user_id)
    
    if usuario.ID == current_user.ID:
        flash('No puedes eliminar tu propio usuario.', 'danger')
        return redirect(url_for('admin_usuarios'))
    
    # Eliminar pedidos asociados (y sus detalles, archivos)
    pedidos = Pedido.query.filter_by(ID_USUARIO=usuario.ID).all()
    for pedido in pedidos:
        DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).delete()
        ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).delete()
        db.session.delete(pedido)
    
    # Ya no se eliminan mensajes de chat porque la tabla fue borrada
    
    db.session.delete(usuario)
    db.session.commit()
    
    flash('Usuario eliminado correctamente.', 'success')
    return redirect(url_for('admin_usuarios'))

@app.route('/admin/perfil', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_perfil():
    if request.method == 'POST':
        telefono = request.form.get('telefono', '').strip()
        email = request.form.get('email', '').strip().lower()

        # Validaciones
        errores = []
        if not telefono or not telefono.isdigit() or len(telefono) != 11:
            errores.append('El teléfono debe tener exactamente 11 dígitos numéricos.')
        if not email or '@' not in email or '.' not in email.split('@')[-1]:
            errores.append('El formato del correo no es válido.')

        if errores:
            for error in errores:
                flash(error, 'danger')
        else:
            current_user.TELEFONO = telefono
            current_user.EMAIL = email
            db.session.commit()
            flash('Perfil actualizado correctamente.', 'success')
        return redirect(url_for('admin_perfil'))

    # GET: mostrar formulario con los datos actuales
    return render_template('admin/perfil.html')

from sqlalchemy import func, extract

# ------------------------------------------------------------
# ADMIN - OPERADORES
# ------------------------------------------------------------
@app.route('/admin/operadores')
@login_required
@admin_required
def admin_operadores():
    # Obtener todos los usuarios que ya son operadores
    operadores = Usuario.query.filter_by(ES_OPERADOR=1).all()
    return render_template('admin/operadores.html', operadores=operadores)

@app.route('/admin/operadores/buscar', methods=['POST'])
@login_required
@admin_required
def admin_operadores_buscar():
    cedula = request.form.get('cedula', '').strip()
    cedula = re.sub(r'^[VEve]-', '', cedula).strip()
    if not cedula:
        return {'error': 'Debe ingresar una cédula'}, 400
    
    usuario = Usuario.query.filter_by(ID_USUARIO=cedula).first()
    if not usuario:
        return {'error': 'No se encontró un usuario con esa cédula'}, 404

    # Contar cuántos operadores hay actualmente
    cantidad_operadores = Usuario.query.filter_by(ES_OPERADOR=1).count()
    
    return {
        'id': usuario.ID,
        'nombre': usuario.NOMBRE,
        'apellido': usuario.APELLIDO,
        'cedula': usuario.ID_USUARIO,
        'email': usuario.EMAIL,
        'telefono': usuario.TELEFONO,
        'es_operador': usuario.ES_OPERADOR,
        'cantidad_operadores': cantidad_operadores
    }

@app.route('/admin/operadores/asignar', methods=['POST'])
@login_required
@admin_required
def admin_operadores_asignar():
    user_id = request.form.get('user_id')
    if not user_id:
        flash('ID de usuario no proporcionado', 'danger')
        return redirect(url_for('admin_operadores'))
    
    usuario = Usuario.query.get(int(user_id))
    if not usuario:
        flash('Usuario no encontrado', 'danger')
        return redirect(url_for('admin_operadores'))
    
    # Si ya es operador, lo desasignamos
    if usuario.ES_OPERADOR:
        usuario.ES_OPERADOR = False
        db.session.commit()
        flash(f'{usuario.NOMBRE} {usuario.APELLIDO} ya no es operador.', 'success')
    else:
        # Verificar límite de 6 operadores
        if Usuario.query.filter_by(ES_OPERADOR=1).count() >= 6:
            flash('No se pueden agregar más de 6 operadores.', 'danger')
        else:
            usuario.ES_OPERADOR = True
            db.session.commit()
            flash(f'{usuario.NOMBRE} {usuario.APELLIDO} ahora es operador.', 'success')
    
    return redirect(url_for('admin_operadores'))

# ------------------------------------------------------------
# PANEL ADMIN – REPORTES
# ------------------------------------------------------------

@app.route('/admin/reportes')
@login_required
@admin_required
def admin_reportes():
    # --- Parámetros de período y pestaña (ya existentes) ---
    periodo = request.args.get('periodo', session.get('periodo_ingresos', 'mensual'))
    session['periodo_ingresos'] = periodo
    tab_activa = request.args.get('tab', 'ventas')
    # Solicitudes por fecha (con validación)
        # Solicitudes por fecha (con validación)
    desde_fecha = request.args.get('desde', '')
    hasta_fecha = request.args.get('hasta', '')

    query = Pedido.query

    if desde_fecha:
        query = query.filter(Pedido.FECHA >= datetime.strptime(desde_fecha, '%Y-%m-%d'))
    if hasta_fecha:
        query = query.filter(Pedido.FECHA <= datetime.strptime(hasta_fecha, '%Y-%m-%d') + timedelta(days=1))

    solicitudes_fecha = query.order_by(Pedido.FECHA.desc()).all()
    buscar = request.args.get('buscar', '').strip()

    hoy = datetime.now()

    # Calcular rango de fechas según período (para ingresos)
    if periodo == 'diario':
        desde = hoy.replace(hour=0, minute=0, second=0, microsecond=0)
        hasta = hoy.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif periodo == 'semanal':
        desde = hoy - timedelta(days=hoy.weekday())
        desde = desde.replace(hour=0, minute=0, second=0, microsecond=0)
        hasta = desde + timedelta(days=6, hours=23, minutes=59, seconds=59)
    elif periodo == 'mensual':
        desde = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        if hoy.month == 12:
            hasta = hoy.replace(year=hoy.year+1, month=1, day=1) - timedelta(seconds=1)
        else:
            hasta = hoy.replace(month=hoy.month+1, day=1) - timedelta(seconds=1)
    elif periodo == 'anual':
        desde = hoy.replace(month=1, day=1, hour=0, minute=0, second=0, microsecond=0)
        hasta = hoy.replace(month=12, day=31, hour=23, minute=59, second=59, microsecond=999999)
    else:
        desde = hoy.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        if hoy.month == 12:
            hasta = hoy.replace(year=hoy.year+1, month=1, day=1) - timedelta(seconds=1)
        else:
            hasta = hoy.replace(month=hoy.month+1, day=1) - timedelta(seconds=1)

    # --- Servicios más solicitados (ranking) ---
    servicios_ranking = db.session.query(
        ServicioImpresion.ID,
        ServicioImpresion.TITULO,
        ServicioImpresion.DESCRIPCION,
        func.count(Pedido.ID).label('total_pedidos')
    ).join(Pedido, Pedido.SERVICIO_ID == ServicioImpresion.ID)\
     .filter(Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo', 'Entregado']))\
     .group_by(ServicioImpresion.ID)\
     .order_by(func.count(Pedido.ID).desc())\
     .all()

    ranking = []
    for idx, (id_srv, titulo, descripcion, total) in enumerate(servicios_ranking, start=1):
        ranking.append({
            'posicion': idx,
            'titulo': titulo,
            'descripcion': descripcion,
            'total': total
        })

    # Ingresos del período
    ingresos_periodo = db.session.query(func.sum(Pedido.TOTAL)).filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo', 'Entregado']),
        Pedido.FECHA >= desde,
        Pedido.FECHA <= hasta
    ).scalar() or 0

    # --- Ventas con búsqueda y datos de usuario ---
    # Consulta base
    ventas_query = Pedido.query

    # Si hay búsqueda, filtrar por datos del usuario
    if buscar:
        # Buscar IDs de usuarios que coincidan con la búsqueda
        usuarios_ids = Usuario.query.filter(
            (Usuario.ID_USUARIO.contains(buscar)) |
            (Usuario.NOMBRE.contains(buscar)) |
            (Usuario.APELLIDO.contains(buscar)) |
            (Usuario.EMAIL.contains(buscar))
        ).with_entities(Usuario.ID).all()
        ids = [u[0] for u in usuarios_ids]
        ventas_query = ventas_query.filter(Pedido.ID_USUARIO.in_(ids) if ids else Pedido.ID_USUARIO == None)

    ventas = ventas_query.order_by(Pedido.FECHA.desc()).all()

    # Añadir usuario a cada pedido
    ventas_con_usuarios = []
    for pedido in ventas:
        usuario = Usuario.query.get(pedido.ID_USUARIO)
        ventas_con_usuarios.append({'pedido': pedido, 'usuario': usuario})

    # Solicitudes por fecha (con filtro de fechas)
    query = Pedido.query
    if desde_fecha:
        query = query.filter(Pedido.FECHA >= datetime.strptime(desde_fecha, '%Y-%m-%d'))
    if hasta_fecha:
        query = query.filter(Pedido.FECHA <= datetime.strptime(hasta_fecha, '%Y-%m-%d') + timedelta(days=1))
    solicitudes_fecha = query.order_by(Pedido.FECHA.desc()).all()

    # Ingresos mensuales (histórico)
    ingresos_mensuales = db.session.query(
        extract('year', Pedido.FECHA).label('anio'),
        extract('month', Pedido.FECHA).label('mes'),
        func.sum(Pedido.TOTAL).label('total')
    ).filter(Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo', 'Entregado']))\
     .group_by('anio', 'mes')\
     .order_by('anio', 'mes').all()

    filtro_activo = bool(desde_fecha or hasta_fecha)

    return render_template('admin/reportes.html',
                           ventas=ventas_con_usuarios,
                           solicitudes_fecha=solicitudes_fecha,
                           ingresos_mensuales=ingresos_mensuales,
                           ingresos_periodo=ingresos_periodo,
                           periodo_actual=periodo,
                           desde=desde_fecha,
                           hasta=hasta_fecha,
                           tab_activa=tab_activa,
                           buscar=buscar,
                           filtro_activo=filtro_activo,
                           hoy=datetime.now(),
                           ranking=ranking)

@app.route('/admin/api/pedido/<int:pedido_id>')
@login_required
@admin_required
def admin_api_pedido(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    usuario = Usuario.query.get(pedido.ID_USUARIO)
    archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido_id).all()
    detalles = DetallePedido.query.filter_by(PEDIDO_ID=pedido_id).all()

    # Construir archivos_detalle (siempre, para uno o varios archivos)
    archivos_detalle = []
    if pedido.DETALLE_ARCHIVOS:
        for item in pedido.DETALLE_ARCHIVOS:
            archivo_real = next((a for a in archivos if a.NOMBRE_ARCHIVO == item['nombre'] and 'comprobantes' not in a.RUTA), None)
            ruta_descarga = None
            if archivo_real:
                ruta_descarga = url_for('static', filename=archivo_real.RUTA.replace('static/', '', 1).replace('\\', '/'))
            servicio_nombre = None
            precio_bn = None
            precio_color = None
            if item.get('servicio_id'):
                srv = ServicioImpresion.query.get(item['servicio_id'])
                if srv:
                    servicio_nombre = srv.TITULO
                    if item.get('tamano'):
                        tam = ServicioImpresionTamano.query.filter_by(
                            SERVICIO_ID=item['servicio_id'], NOMBRE=item['tamano']
                        ).first()
                        if tam:
                            precio_bn = float(tam.PRECIO_BN)    # ← corregido
                            precio_color = float(tam.PRECIO_COLOR) if srv.ES_MIXTO else None
            archivos_detalle.append({
                'nombre': item['nombre'],
                'paginas': item['paginas'],
                'servicio_nombre': servicio_nombre,
                'tamano': item.get('tamano'),
                'precio': precio_bn,                    # usado en modal
                'precio_color': precio_color,           # nuevo
                'ruta_descarga': ruta_descarga,
                'paginas_color': item.get('paginas_color'),      # nuevo
                'comentarios': item.get('comentarios')          # nuevo
            })
    else:
        for a in archivos:
            if 'comprobantes' not in a.RUTA:
                servicio_nombre = None
                precio_bn = None
                precio_color = None
                if pedido.SERVICIO_ID:
                    srv = ServicioImpresion.query.get(pedido.SERVICIO_ID)
                    if srv:
                        servicio_nombre = srv.TITULO
                        if pedido.TAMANO:
                            tam = ServicioImpresionTamano.query.filter_by(
                                SERVICIO_ID=pedido.SERVICIO_ID, NOMBRE=pedido.TAMANO
                            ).first()
                            if tam:
                                precio_bn = float(tam.PRECIO_BN)
                                precio_color = float(tam.PRECIO_COLOR) if srv.ES_MIXTO else None
                archivos_detalle.append({
                    'nombre': a.NOMBRE_ARCHIVO,
                    'paginas': pedido.PAGINAS,
                    'servicio_nombre': servicio_nombre,
                    'tamano': pedido.TAMANO,
                    'precio': precio_bn,
                    'precio_color': precio_color,
                    'ruta_descarga': url_for('static', filename=a.RUTA.replace('static/', '', 1).replace('\\', '/')),
                    'paginas_color': pedido.PAGINAS_COLOR if servicio_nombre and srv.ES_MIXTO else None,
                    'comentarios': pedido.COMENTARIOS if servicio_nombre and srv.ES_MIXTO else None
                })

    return {
        'pedido': {
            'id': pedido.ID,
            'fecha': pedido.FECHA.strftime('%d/%m/%Y %H:%M'),
            'total': float(pedido.TOTAL),
            'estado': pedido.ESTADO,
            'codigo_ticket': pedido.CODIGO_TICKET,
            'fecha_retiro': pedido.FECHA_RETIRO.strftime('%d/%m/%Y') if pedido.FECHA_RETIRO else None,
            'hora_retiro': pedido.HORA_RETIRO.strftime('%I:%M %p') if pedido.HORA_RETIRO else None,
            'comentarios': pedido.COMENTARIOS,
            'tamano': pedido.TAMANO,
            'referencia_pago': pedido.REFERENCIA_PAGO
        },
        'usuario': {
            'nombre': usuario.NOMBRE,
            'apellido': usuario.APELLIDO,
            'cedula': usuario.ID_USUARIO,
            'email': usuario.EMAIL,
            'telefono': usuario.TELEFONO
        },
        'archivos_detalle': archivos_detalle,
        'detalles': [{
            'cantidad': d.CANTIDAD,
            'precio_unitario': float(d.PRECIO_UNITARIO) if d.PRECIO_UNITARIO else None,
            'subtotal': float(d.SUBTOTAL)
        } for d in detalles]
    }

# ------------------------------------------------------------
# PANEL ADMIN – CRONOGRAMA
# ------------------------------------------------------------

@app.route('/admin/cronograma')
@login_required
@admin_required
def admin_cronograma():
    pedidos = Pedido.query.filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo'])
    ).order_by(Pedido.FECHA_RETIRO.asc(), Pedido.HORA_RETIRO.asc()).all()
    
    # Agrupar por fecha
    from collections import defaultdict
    grupos = defaultdict(list)
    for pedido in pedidos:
        if pedido.FECHA_RETIRO:
            usuario = Usuario.query.get(pedido.ID_USUARIO)
            grupos[pedido.FECHA_RETIRO.isoformat()].append({
                'pedido': pedido,
                'usuario': usuario
            })
    
    # Ordenar las fechas
    fechas_ordenadas = sorted(grupos.keys())
    
    # Preparar datos para la plantilla
    cronograma = []
    hoy = date.today()
    for fecha_str in fechas_ordenadas:
        fecha = date.fromisoformat(fecha_str)
        if fecha == hoy:
            etiqueta = "Hoy"
        elif fecha == hoy + timedelta(days=1):
            etiqueta = "Mañana"
        else:
            dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
            etiqueta = dias_semana[fecha.weekday()]
        
        cronograma.append({
            'fecha': fecha,
            'etiqueta': etiqueta,
            'fecha_formateada': fecha.strftime('%d/%m/%Y'),
            'pedidos': grupos[fecha_str]
        })
    
    return render_template('admin/cronograma.html', cronograma=cronograma, hoy=hoy)

# ------------------------------------------------------------
# PANEL ADMIN – CONFIGURACION
# ------------------------------------------------------------

@app.route('/admin/configuracion', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_configuracion():
    if request.method == 'POST':
        # Actualizar cada clave con el valor enviado desde el formulario
        for clave, valor in request.form.items():
            if clave.startswith('config_'):
                clave_real = clave[7:]  # quitar prefijo 'config_'
                config = Configuracion.query.filter_by(CLAVE=clave_real).first()
                if config:
                    config.VALOR = valor
        db.session.commit()
        flash('Configuración actualizada.', 'success')
        return redirect(url_for('admin_configuracion'))

    # Cargar toda la configuración actual
    configuraciones = Configuracion.query.all()
    config_dict = {c.CLAVE: c.VALOR for c in configuraciones}
    return render_template('admin/configuracion.html', config=config_dict)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

def validar_nombre_apellido(texto, campo="Nombre"):
    if not texto:
        return f"El {campo} es obligatorio."
    if len(texto) < 2 or len(texto) > 20:
        return f"El {campo} debe tener entre 2 y 20 caracteres."
    if not re.match(r'^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s]+$', texto):
        return f"El {campo} solo puede contener letras y espacios."
    return None

def validar_cedula(cedula):
    if not cedula:
        return "La cédula es obligatoria."
    cedula = cedula.strip().upper()
    
    # Aceptar solo dígitos de 7 a 12 caracteres
    if re.match(r'^\d{7,12}$', cedula):
        return None
    
    return "Formato de cédula no válido. Use de 7 a 12 dígitos."

def validar_telefono(telefono):
    if not telefono:
        return "El teléfono es obligatorio."
    # Eliminar cualquier carácter no numérico
    telefono = re.sub(r'\D', '', telefono)
    
    # Verificar longitud exacta (11 dígitos para Venezuela)
    if len(telefono) != 11:
        return "El teléfono debe tener exactamente 11 dígitos."
    
    # Lista blanca de prefijos válidos (móviles + un fijo para Caracas)
    prefijos_validos = ['0412', '0422', '0414', '0424', '0416', '0426', '0212']
    
    prefijo = telefono[:4]
    if prefijo not in prefijos_validos:
        return "Prefijo no válido. Solo se permiten números venezolanos (0412, 0414, 0416, 0422, 0424, 0426, 0212)."
    
    return None  # Válido

def validar_email(email):
    if not email:
        return "El correo es obligatorio."
    if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', email):
        return "El formato del correo no es válido."
    return None

def validar_contraseña(password):
    errores = []
    if len(password) < 8:
        errores.append("al menos 8 caracteres")
    if not re.search(r"[A-Z]", password):
        errores.append("una mayúscula")
    if not re.search(r"[a-z]", password):
        errores.append("una minúscula")
    if not re.search(r"\d", password):
        errores.append("un número")
    if not re.search(r"[!@#$%^&*(),.?\":{}|<>_\-+=~]", password):
        errores.append("un carácter especial (ej: ! @ # $ %)")

    if errores:
        return f"La contraseña debe contener: {', '.join(errores)}."
    return None

@app.route('/admin/configuracion/vaciar-historial', methods=['POST'])
@login_required
@admin_required
def admin_vaciar_historial():
    try:
        # Eliminar archivos físicos
        carpetas = ['static/uploads/impresion', 'static/uploads/comprobantes']
        for carpeta in carpetas:
            if os.path.exists(carpeta):
                for archivo in os.listdir(carpeta):
                    ruta_archivo = os.path.join(carpeta, archivo)
                    if os.path.isfile(ruta_archivo):
                        os.remove(ruta_archivo)

        # SQL directo para eliminar registros
        db.session.execute(text("DELETE FROM detalle"))
        db.session.execute(text("DELETE FROM archivo_pedido"))
        db.session.execute(text("DELETE FROM pedido"))
        db.session.commit()

        flash('Historial vaciado correctamente.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error al vaciar el historial: {str(e)}', 'danger')
    
    return redirect(url_for('admin_configuracion'))

# ------------------------------------------------------------
# PANEL DE OPERADOR
# ------------------------------------------------------------

@app.route('/operador')
@login_required
@operador_required
def panel_operador():
    # Pedidos pendientes de imprimir
    pedidos = Pedido.query.filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso'])
    ).order_by(Pedido.FECHA.asc()).all()

    # Historial reciente (pedidos listos o entregados)
    historial = Pedido.query.filter(
        Pedido.ESTADO.in_(['Listo', 'Entregado'])
    ).order_by(Pedido.FECHA.desc()).limit(20).all()

    # Contador de pedidos no vistos por el operador
    nuevos_operador = Pedido.query.filter_by(VISTO_OPERADOR=False).filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso'])
    ).count()

    # Preparar los datos enriquecidos para las tarjetas
    pedidos_data = []
    for pedido in pedidos:
        usuario = Usuario.query.get(pedido.ID_USUARIO)
        archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).all()
        # Solo archivos de impresión
        archivos_impresion = [a for a in archivos if 'comprobantes' not in a.RUTA.lower()]
        detalles = DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).all()

        # Obtener servicio de impresión
        servicio = None
        if pedido.SERVICIO_ID:
            servicio = ServicioImpresion.query.get(pedido.SERVICIO_ID)

        pedidos_data.append({
            'pedido': pedido,
            'usuario': usuario,
            'archivos': archivos_impresion,
            'detalles': detalles,
            'servicio': servicio,          # ← nuevo
            'referencia_pago': pedido.REFERENCIA_PAGO   # ← nuevo
        })

    # --- Cronograma (misma lógica que en admin) ---
    from collections import defaultdict
    pedidos_cronograma = Pedido.query.filter(
        Pedido.ESTADO.in_(['Pago confirmado', 'En proceso', 'Listo'])
    ).order_by(Pedido.FECHA_RETIRO.asc(), Pedido.HORA_RETIRO.asc()).all()

    grupos = defaultdict(list)
    for pedido in pedidos_cronograma:
        if pedido.FECHA_RETIRO:
            usuario = Usuario.query.get(pedido.ID_USUARIO)
            grupos[pedido.FECHA_RETIRO.isoformat()].append({
                'pedido': pedido,
                'usuario': usuario
            })

    fechas_ordenadas = sorted(grupos.keys())
    cronograma = []
    hoy = date.today()
    for fecha_str in fechas_ordenadas:
        fecha = date.fromisoformat(fecha_str)
        if fecha == hoy:
            etiqueta = "Hoy"
        elif fecha == hoy + timedelta(days=1):
            etiqueta = "Mañana"
        else:
            dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
            etiqueta = dias_semana[fecha.weekday()]
        cronograma.append({
            'fecha': fecha,
            'etiqueta': etiqueta,
            'fecha_formateada': fecha.strftime('%d/%m/%Y'),
            'pedidos': grupos[fecha_str]
        })

    # Recuperar mensaje de sesión (si existe)
    mensaje = session.pop('mensaje_operador', None)

    return render_template('operador/dashboard.html',
                           pedidos=pedidos_data,
                           historial=historial,
                           nuevos_operador=nuevos_operador,
                           mensaje=mensaje,
                           cronograma=cronograma)

@app.route('/operador/solicitud/<int:pedido_id>')
@login_required
@operador_required
def operador_detalle(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    usuario = Usuario.query.get(pedido.ID_USUARIO)
    archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).all()
    detalles = DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).all()

    # Solo archivos de impresión
    archivos_impresion = [a for a in archivos if 'comprobantes' not in a.RUTA.lower()]

    # Servicio de impresión
    servicio = None
    if pedido.SERVICIO_ID:
        servicio = ServicioImpresion.query.get(pedido.SERVICIO_ID)

    # Referencia de pago
    referencia_pago = pedido.REFERENCIA_PAGO

    # Marcar como visto
    if not pedido.VISTO_OPERADOR:
        pedido.VISTO_OPERADOR = True
        db.session.commit()

    if pedido.DETALLE_ARCHIVOS:
        archivos_info = pedido.DETALLE_ARCHIVOS
    else:
        archivos_info = [{
            'nombre': archivos[0].NOMBRE_ARCHIVO if archivos else 'Sin archivo',
            'paginas': pedido.PAGINAS,
            'servicio_id': pedido.SERVICIO_ID,
            'tamano': pedido.TAMANO,
            'paginas_color': pedido.PAGINAS_COLOR,
            'comentarios': pedido.COMENTARIOS
        }]

    # Obtener nombre real de cada servicio de impresión
    todos_servicios = {srv.ID: srv.TITULO for srv in ServicioImpresion.query.all()}

    return render_template('operador/detalle.html',
                           pedido=pedido,
                           usuario=usuario,
                           archivos=archivos_impresion,
                           detalles=detalles,
                           servicio=servicio,
                           referencia_pago=referencia_pago,
                           todos_servicios=todos_servicios,
                           archivos_info=archivos_info)

@app.route('/operador/solicitud/<int:pedido_id>/marcar-listo', methods=['POST'])
@login_required
@operador_required
def operador_marcar_listo(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ESTADO in ['Pago confirmado', 'En proceso']:
        # Generar código de ticket si no existe
        if not pedido.CODIGO_TICKET:
            import random, string
            pedido.CODIGO_TICKET = f"TICK-{pedido.ID}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"

        pedido.ESTADO = 'Listo'
        db.session.commit()

        # Enviar correo al cliente
        usuario = Usuario.query.get(pedido.ID_USUARIO)
        if usuario and usuario.EMAIL:
            enviar_correo_listo(usuario.EMAIL, usuario.NOMBRE, pedido.CODIGO_TICKET, pedido.ID)

        session['mensaje_operador'] = 'Pedido marcado como Listo. Se ha notificado al cliente.'
        return redirect(url_for('panel_operador'))

    return redirect(url_for('panel_operador'))

@app.route('/operador/solicitud/<int:pedido_id>/marcar-entregado', methods=['POST'])
@login_required
@operador_required
def operador_marcar_entregado(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ESTADO in ['Listo', 'Pago confirmado', 'En proceso']:
        pedido.ESTADO = 'Entregado'
        db.session.commit()
        return {'success': True}
    return {'error': 'No se puede marcar como entregado en este estado'}, 400

# ------------------------------------------------------------
# IMPRIMIR
# ------------------------------------------------------------

@app.route('/imprimir', methods=['GET'])
@login_required
def imprimir():
    pedido_id = request.args.get('pedido_id', type=int)
    configurado = request.args.get('configurado') == '1'
    pedido = None
    if pedido_id:
        pedido = Pedido.query.get(pedido_id)
        if pedido and pedido.ID_USUARIO != current_user.ID:
            pedido = None  # no autorizado
    return render_template('tienda/imprimir.html', pedido=pedido, configurado=configurado, active_page='imprimir')

@app.route('/detectar-paginas', methods=['POST'])
@login_required
def detectar_paginas_ajax():
    if 'archivo' not in request.files:
        return {'error': 'No se recibió archivo'}, 400
    archivo = request.files['archivo']
    if archivo.filename == '':
        return {'error': 'Archivo vacío'}, 400

    filename = secure_filename(archivo.filename)
    ruta_destino = os.path.join(UPLOAD_FOLDER_IMPRESION, filename)
    archivo.save(ruta_destino)

    # Obtener extensión una sola vez
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''

    try:
        # ---------- 1. Validación MIME ----------
        if ext in ('pdf', 'png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif', 'pptx'):
            valido, error = validar_mime(ruta_destino, ['pdf', 'imagen', 'pptx'])
        elif ext == 'zip':
            valido, error = validar_mime(ruta_destino, ['zip'])
        else:
            valido, error = False, f"Formato no soportado: .{ext}"
        if not valido:
            os.remove(ruta_destino)
            return {'error': error}, 400

        # ---------- 2. Verificación de integridad (excepto PDF) ----------
        if ext != 'pdf':
            integro, error_integridad = verificar_integridad(ruta_destino, ext)
            if not integro:
                os.remove(ruta_destino)
                return {'error': error_integridad}, 400

        # ---------- 3. Validación específica para PDF ----------
        if ext == 'pdf':
            size = os.path.getsize(ruta_destino)
            if size > 10 * 1024 * 1024:  # 10 MB
                os.remove(ruta_destino)
                return {'error': 'El PDF no puede superar los 10 MB.'}, 400

            # Verificar cabecera mágica %PDF
            with open(ruta_destino, 'rb') as f:
                header = f.read(4)
                if not header.startswith(b'%PDF'):
                    os.remove(ruta_destino)
                    return {'error': 'El archivo PDF no es válido (cabecera incorrecta).'}, 400

        # ---------- 4. Detección de páginas ----------
        paginas, mensaje = detectar_paginas(ruta_destino, filename)
        if paginas is None:
            os.remove(ruta_destino)
            return {'error': mensaje or 'Formato no soportado'}, 400

        # ---------- 5. Crear pedido borrador ----------
        nuevo = Pedido(
            ID_USUARIO=current_user.ID,
            FECHA=datetime.now(),
            ESTADO='borrador',
            TOTAL=0.0,
            PAGINAS=paginas
        )
        db.session.add(nuevo)
        db.session.flush()

        archivo_bd = ArchivoPedido(
            PEDIDO_ID=nuevo.ID,
            NOMBRE_ARCHIVO=filename,
            RUTA=ruta_destino
        )
        db.session.add(archivo_bd)
        db.session.commit()

        # ---------- 6. Limpiar borradores viejos ----------
        limite = datetime.now() - timedelta(minutes=5)
        pedidos_viejos = Pedido.query.filter(
            Pedido.ID_USUARIO == current_user.ID,
            Pedido.ESTADO == 'borrador',
            Pedido.FECHA < limite
        ).all()

        for p in pedidos_viejos:
            for archivo in ArchivoPedido.query.filter_by(PEDIDO_ID=p.ID).all():
                if os.path.exists(archivo.RUTA):
                    os.remove(archivo.RUTA)
                db.session.delete(archivo)
            DetallePedido.query.filter_by(PEDIDO_ID=p.ID).delete()
            db.session.delete(p)

        db.session.commit()

        return {
            'success': True,
            'pedido_id': nuevo.ID,
            'paginas': paginas,
            'mensaje': mensaje or ''
        }

    except Exception as e:
        # Si ocurre cualquier error inesperado, limpiar archivo y avisar
        if os.path.exists(ruta_destino):
            os.remove(ruta_destino)
        print(f"Error en detectar_paginas_ajax: {e}")
        return {'error': f'Error al procesar el archivo. ({str(e)[:80]})'}, 500

@app.route('/detectar-paginas-multiples', methods=['POST'])
@login_required
def detectar_paginas_multiples():
    archivos = request.files.getlist('archivos[]')
    if not archivos or len(archivos) == 0:
        return {'error': 'No se recibieron archivos'}, 400

    docs = 0
    imgs = 0
    for f in archivos:
        ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ''
        if ext in ('pdf', 'pptx', 'zip'):
            docs += 1
        elif ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif'):
            imgs += 1
        else:
            return {'error': f'Formato no soportado: {f.filename}'}, 400

    if docs > 10:
        return {'error': 'Máximo 10 documentos permitidos'}, 400
    if imgs > 20:
        return {'error': 'Máximo 20 imágenes permitidas'}, 400

    nuevo = Pedido(
        ID_USUARIO=current_user.ID,
        FECHA=datetime.now(),
        ESTADO='borrador',
        TOTAL=0.0,
        PAGINAS=0
    )
    db.session.add(nuevo)
    db.session.flush()

    total_paginas = 0
    detalle_archivos = []

    try:
        for f in archivos:
            filename = secure_filename(f.filename)
            ruta_destino = os.path.join(UPLOAD_FOLDER_IMPRESION, filename)
            f.save(ruta_destino)

            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''

            # 1. Validación MIME
            if ext in ('pdf', 'png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'tif', 'pptx'):
                valido, error = validar_mime(ruta_destino, ['pdf', 'imagen', 'pptx'])
            elif ext == 'zip':
                valido, error = validar_mime(ruta_destino, ['zip'])
            else:
                valido, error = False, f"Formato no soportado: .{ext}"
            if not valido:
                os.remove(ruta_destino)
                raise Exception(error)

            # 2. Integridad (excepto PDF)
            if ext != 'pdf':
                integro, error_integridad = verificar_integridad(ruta_destino, ext)
                if not integro:
                    os.remove(ruta_destino)
                    raise Exception(error_integridad)

            # 3. Cabecera y tamaño para PDF
            if ext == 'pdf':
                size = os.path.getsize(ruta_destino)
                if size > 10 * 1024 * 1024:
                    os.remove(ruta_destino)
                    raise Exception('El PDF no puede superar los 10 MB.')
                with open(ruta_destino, 'rb') as fh:
                    header = fh.read(4)
                    if not header.startswith(b'%PDF'):
                        os.remove(ruta_destino)
                        raise Exception('El archivo PDF no es válido (cabecera incorrecta).')

            # 4. Detección de páginas
            paginas, mensaje = detectar_paginas(ruta_destino, filename)
            if paginas is None:
                os.remove(ruta_destino)
                raise Exception(f'Error en {filename}: {mensaje}')

            total_paginas += paginas
            detalle_archivos.append({
                'nombre': filename,
                'paginas': paginas,
                'servicio_id': None,
                'tamano': None
            })

            archivo_bd = ArchivoPedido(
                PEDIDO_ID=nuevo.ID,
                NOMBRE_ARCHIVO=filename,
                RUTA=ruta_destino
            )
            db.session.add(archivo_bd)

        nuevo.PAGINAS = total_paginas
        nuevo.DETALLE_ARCHIVOS = detalle_archivos
        db.session.commit()

        # ---------- Limpiar borradores viejos (CORREGIDO) ----------
        limite = datetime.now() - timedelta(minutes=5)
        pedidos_viejos = Pedido.query.filter(
            Pedido.ID_USUARIO == current_user.ID,
            Pedido.ESTADO == 'borrador',
            Pedido.FECHA < limite
        ).all()

        for p in pedidos_viejos:
            # 1. Eliminar archivos físicos y registros de ArchivoPedido
            for a in ArchivoPedido.query.filter_by(PEDIDO_ID=p.ID).all():
                if os.path.exists(a.RUTA):
                    os.remove(a.RUTA)
                db.session.delete(a)
            # 2. Eliminar detalles del pedido
            DetallePedido.query.filter_by(PEDIDO_ID=p.ID).delete()
            # 3. Ahora sí, eliminar el pedido
            db.session.delete(p)

        db.session.commit()

        return {
            'success': True,
            'pedido_id': nuevo.ID,
            'paginas_totales': total_paginas,
            'detalle_archivos': detalle_archivos
        }

    except Exception as e:
        # Cualquier error → eliminar archivo temporal y hacer rollback
        if 'ruta_destino' in locals() and os.path.exists(ruta_destino):
            os.remove(ruta_destino)
        db.session.rollback()          # revierte la creación del pedido y sus archivos
        print(f"Error en detectar_paginas_multiples: {e}")
        return {'error': f'Error al procesar los archivos. ({str(e)[:80]})'}, 500

@app.route('/cancelar-pedido/<int:pedido_id>', methods=['POST'])
@login_required
def cancelar_pedido(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return {'error': 'No autorizado'}, 403
    if pedido.ESTADO not in ['borrador', 'Pendiente de pago']:
        return {'error': 'No se puede cancelar en este estado'}, 400

    # Eliminar archivos asociados (físicos y registros)
    archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido_id).all()
    for archivo in archivos:
        # Borrar archivo físico si existe
        if os.path.exists(archivo.RUTA):
            os.remove(archivo.RUTA)
        db.session.delete(archivo)

    # Eliminar detalles
    DetallePedido.query.filter_by(PEDIDO_ID=pedido_id).delete()
    # Eliminar pedido
    db.session.delete(pedido)
    db.session.commit()
    return {'success': True}

@app.route('/configurar/<int:pedido_id>', methods=['GET', 'POST'])
@login_required
def configurar_impresion(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return redirect(url_for('imprimir'))

    if pedido.ESTADO not in ['borrador', 'Pendiente de pago']:
        return redirect(url_for('mis_solicitudes'))

    if pedido.DETALLE_ARCHIVOS:
        return redirect(url_for('configurar_impresion_multiple', pedido_id=pedido_id))

    servicios = ServicioImpresion.query.filter_by(ACTIVO=True).order_by(ServicioImpresion.TITULO).all()

    archivo = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido_id).first()
    archivo_nombre = archivo.NOMBRE_ARCHIVO if archivo else 'Sin archivo'

    # Si se solicita reiniciar la configuración, limpiar selección previa
    if request.args.get('reset') == '1':
        pedido.SERVICIO_ID = None
        pedido.TAMANO = None
        db.session.commit()

    if request.method == 'POST':
        servicio_id = request.form.get('servicio_id', type=int)
        tamano_nombre = request.form.get('tamano_nombre', '').strip()
        comentarios = request.form.get('comentarios', '').strip()

        if not servicio_id or not tamano_nombre:
            
            return redirect(url_for('configurar_impresion', pedido_id=pedido.ID))

        pedido.SERVICIO_ID = servicio_id
        pedido.TAMANO = tamano_nombre

        servicio = ServicioImpresion.query.get(servicio_id)

        if servicio and servicio.ES_MIXTO:
            pedido.PAGINAS_COLOR = request.form.get('paginas_color', '').strip() or None
        else:
            pedido.PAGINAS_COLOR = None

        pedido.COMENTARIOS = comentarios  # 'comentarios' ya viene del formulario

        db.session.commit()
        return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

    return render_template('tienda/configurar_impresion.html',
                           pedido=pedido,
                           servicios=servicios,
                           archivo_nombre=archivo_nombre) 

@app.route('/configurar-multiple/<int:pedido_id>', methods=['GET', 'POST'])
@login_required
def configurar_impresion_multiple(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return redirect(url_for('imprimir'))

    if pedido.ESTADO not in ['borrador', 'Pendiente de pago']:
        return redirect(url_for('mis_solicitudes'))

    if not pedido.DETALLE_ARCHIVOS:
        return redirect(url_for('configurar_impresion', pedido_id=pedido.ID))

    if request.args.get('reset') == '1':
        for item in pedido.DETALLE_ARCHIVOS:
            item['servicio_id'] = None
            item['tamano'] = None
        db.session.commit()

    servicios = ServicioImpresion.query.filter_by(ACTIVO=True).order_by(ServicioImpresion.TITULO).all()

    if request.method == 'POST':
        detalle = pedido.DETALLE_ARCHIVOS
        for i, item in enumerate(detalle):
            serv_id = request.form.get(f'servicio_{i}')
            tam = request.form.get(f'tamano_{i}')
            pag_color = request.form.get(f'paginas_color_{i}')
            com = request.form.get(f'comentarios_{i}')
            if serv_id:
                item['servicio_id'] = int(serv_id)
            if tam:
                item['tamano'] = tam
            if pag_color:
                item['paginas_color'] = pag_color
            if com:
                item['comentarios'] = com

        pedido.DETALLE_ARCHIVOS = detalle
        pedido.COMENTARIOS = request.form.get('comentarios', '').strip()
        flag_modified(pedido, 'DETALLE_ARCHIVOS')
        db.session.commit()
        return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

    # GET
    return render_template('tienda/configurar_impresion_multiple.html',
                           pedido=pedido,
                           servicios=servicios) 

@app.route('/api/pedido/<int:pedido_id>')
@login_required
def api_pedido(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return {'error': 'No autorizado'}, 403

    archivos_reg = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido_id).all()

    # Construir lista de archivos con sus páginas
    if pedido.DETALLE_ARCHIVOS:
        archivos_info = pedido.DETALLE_ARCHIVOS
    else:
        archivos_info = [{
            'nombre': archivos_reg[0].NOMBRE_ARCHIVO if archivos_reg else 'Sin archivo',
            'paginas': pedido.PAGINAS
        }]

    archivo_nombre = archivos_info[0]['nombre'] if archivos_info else 'Sin archivo'

    servicio_nombre = None
    if pedido.SERVICIO_ID:
        srv = ServicioImpresion.query.get(pedido.SERVICIO_ID)
        if srv:
            servicio_nombre = srv.TITULO

    return {
        'paginas': pedido.PAGINAS,
        'servicio': servicio_nombre,
        'tamano': pedido.TAMANO,
        'fecha_retiro': pedido.FECHA_RETIRO.strftime('%Y-%m-%d') if pedido.FECHA_RETIRO else '',
        'hora_retiro': pedido.HORA_RETIRO.strftime('%H:%M') if pedido.HORA_RETIRO else '',
        'total': float(pedido.TOTAL),
        'archivo_nombre': archivo_nombre,
        'archivos': archivos_info,           # ← lista de {nombre, paginas}
        'detalle_archivos': pedido.DETALLE_ARCHIVOS  # ← array completo o None
    }

@app.route('/programar-retiro/<int:pedido_id>', methods=['GET', 'POST'])
@login_required
def programar_retiro(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return redirect(url_for('imprimir'))

    if pedido.ESTADO not in ['borrador', 'Pendiente de pago']:
        return redirect(url_for('mis_solicitudes'))

    if request.method == 'POST':
        fecha_retiro = request.form.get('fecha_retiro')
        hora_retiro = request.form.get('hora_retiro')

        if not fecha_retiro or not hora_retiro:
            flash('Selecciona fecha y hora.', 'warning')
            return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

        try:
            fecha_elegida = datetime.strptime(fecha_retiro, '%Y-%m-%d').date()
            hora_elegida = datetime.strptime(hora_retiro, '%H:%M').time()
        except ValueError:
            flash('Formato de fecha u hora no válido.', 'danger')
            return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

        if fecha_elegida.weekday() == 6:
            flash('No se pueden programar retiros los domingos.', 'warning')
            return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

        if fecha_elegida.weekday() == 5:
            horarios_validos = [(9,0),(9,30),(10,0),(10,30),(11,0),(11,30),(12,0)]
            permitido = any(h == hora_elegida.hour and m == hora_elegida.minute for h, m in horarios_validos)
            if not permitido:
                flash('Los sábados solo se puede retirar de 9:00 AM a 12:00 PM.', 'warning')
                return redirect(url_for('programar_retiro', pedido_id=pedido.ID))
        else:
            manana = [(7,0),(7,30),(8,0),(8,30),(9,0),(9,30),(10,0),(10,30),(11,0),(11,30),(12,0)]
            tarde  = [(13,0),(13,30),(14,0),(14,30),(15,0),(15,30),(16,0),(16,30),(17,0),(17,30)]
            permitido = any((h == hora_elegida.hour and m == hora_elegida.minute) for h, m in manana + tarde)
            if not permitido:
                flash('Horario no válido. L-V: 7:00 AM - 12:00 PM y 1:00 PM - 5:30 PM.', 'warning')
                return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

        fecha_minima = datetime.now().date() + timedelta(days=1)
        fecha_limite = datetime.now().date() + timedelta(days=7)
        if fecha_elegida < fecha_minima or fecha_elegida > fecha_limite:
            flash('La fecha de retiro debe ser entre mañana y los próximos 7 días.', 'warning')
            return redirect(url_for('programar_retiro', pedido_id=pedido.ID))

        paginas = pedido.PAGINAS or 1
        total = 0
        tamano_obj = None

        if pedido.DETALLE_ARCHIVOS:
            for item in pedido.DETALLE_ARCHIVOS:
                paginas_archivo = item['paginas']
                serv_id = item.get('servicio_id')
                tamano_nombre = item.get('tamano')
                if serv_id and tamano_nombre:
                    servicio = ServicioImpresion.query.get(serv_id)
                    tamano_obj = ServicioImpresionTamano.query.filter_by(
                        SERVICIO_ID=serv_id, NOMBRE=tamano_nombre
                    ).first()
                    if tamano_obj:
                        if servicio and servicio.ES_MIXTO and item.get('paginas_color'):
                            try:
                                paginas_color = int(item['paginas_color'])
                            except ValueError:
                                paginas_color = 0
                            if paginas_color > paginas_archivo:
                                paginas_color = paginas_archivo
                            paginas_bn = paginas_archivo - paginas_color
                            total += round(float(tamano_obj.PRECIO_BN) * paginas_bn + float(tamano_obj.PRECIO_COLOR) * paginas_color, 2)
                        else:
                            total += round(float(tamano_obj.PRECIO_BN) * paginas_archivo, 2)
        else:
            paginas = pedido.PAGINAS or 1
            if pedido.SERVICIO_ID and pedido.TAMANO:
                servicio = ServicioImpresion.query.get(pedido.SERVICIO_ID)
                tamano_obj = ServicioImpresionTamano.query.filter_by(
                    SERVICIO_ID=pedido.SERVICIO_ID, NOMBRE=pedido.TAMANO
                ).first()
                if tamano_obj:
                    if servicio and servicio.ES_MIXTO and pedido.PAGINAS_COLOR:
                        try:
                            paginas_color = int(pedido.PAGINAS_COLOR)
                        except ValueError:
                            paginas_color = 0
                        if paginas_color > paginas:
                            paginas_color = paginas
                        paginas_bn = paginas - paginas_color
                        total = round(float(tamano_obj.PRECIO_BN) * paginas_bn + float(tamano_obj.PRECIO_COLOR) * paginas_color, 2)
                    else:
                        total = round(float(tamano_obj.PRECIO_BN) * paginas, 2)

        pedido.FECHA_RETIRO = fecha_elegida
        pedido.HORA_RETIRO = hora_elegida
        pedido.TOTAL = total
        pedido.ESTADO = 'Pendiente de pago'

        detalle_existente = DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).first()
        precio_unitario = float(tamano_obj.PRECIO_BN) if tamano_obj else 0
        if detalle_existente:
            detalle_existente.CANTIDAD = paginas
            detalle_existente.PRECIO_UNITARIO = precio_unitario
            detalle_existente.SUBTOTAL = total
        else:
            nuevo_detalle = DetallePedido(
                PEDIDO_ID=pedido.ID,
                CANTIDAD=paginas,
                PRECIO_UNITARIO=precio_unitario,
                SUBTOTAL=total)
            db.session.add(nuevo_detalle)

        db.session.commit()
        return redirect(url_for('pagar_impresion', pedido_id=pedido.ID))

    # GET: obtener el nombre del archivo y la lista de archivos
    archivos_reg = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido_id).all()
    archivos_info = []
    if pedido.DETALLE_ARCHIVOS:
        archivos_info = pedido.DETALLE_ARCHIVOS
    else:
        archivos_info = [{'nombre': a.NOMBRE_ARCHIVO, 'paginas': pedido.PAGINAS} for a in archivos_reg]

    archivo_nombre = archivos_info[0]['nombre'] if archivos_info else 'Sin archivo'

    return render_template('tienda/programar_retiro.html',
                           pedido=pedido,
                           archivo_nombre=archivo_nombre,
                           archivos_info=archivos_info)

@app.route('/pagar-impresion/<int:pedido_id>', methods=['GET', 'POST'])
@login_required
def pagar_impresion(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if request.method == 'POST':
        referencia = request.form.get('referencia')
        comprobante = request.files.get('comprobante')
        if comprobante:
            filename = secure_filename(comprobante.filename)
            filepath = os.path.join(UPLOAD_FOLDER_COMPROBANTES, filename)  # antes UPLOAD_FOLDER_IMPRESION
            comprobante.save(filepath)
            archivo = ArchivoPedido(
                PEDIDO_ID=pedido.ID,
                NOMBRE_ARCHIVO=filename,
                RUTA=filepath
            )
            db.session.add(archivo)
        pedido.REFERENCIA_PAGO = referencia
        pedido.ESTADO = 'Esperando validación'
        db.session.commit()
        return redirect(url_for('espera_validacion', pedido_id=pedido.ID))

    # Obtener datos de PagoMóvil desde la tabla configuracion
    config_pago = {}
    configs = Configuracion.query.filter(Configuracion.CLAVE.startswith('pago_movil')).all()
    for c in configs:
        config_pago[c.CLAVE] = c.VALOR

    return render_template('tienda/pagar_impresion.html', pedido=pedido, config_pago=config_pago)

@app.route('/cancelar-impresion/<int:pedido_id>', methods=['POST'])
@login_required
def cancelar_impresion(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return {'error': 'No autorizado'}, 403

    if pedido.ESTADO in ['Cancelado', 'Pago confirmado', 'Listo', 'Entregado']:
        return {'error': 'El pedido ya no puede ser cancelado.'}, 400

    pedido.ESTADO = 'Cancelado'
    db.session.commit()

    return {'ok': True}

@app.route('/espera-validacion/<int:pedido_id>')
@login_required
def espera_validacion(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return redirect(url_for('imprimir'))
    return render_template('tienda/espera_validacion.html', pedido=pedido)

@app.route('/api/estado-pedido/<int:pedido_id>')
@login_required
def api_estado_pedido(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return {'error': 'No autorizado'}, 403
    return {'estado': pedido.ESTADO}

@app.route('/ticket/<int:pedido_id>')
@login_required
def ticket_impresion(pedido_id):
    pedido = Pedido.query.get_or_404(pedido_id)
    if pedido.ID_USUARIO != current_user.ID:
        return redirect(url_for('imprimir'))

    if pedido.ESTADO not in ['Pago confirmado', 'Listo', 'Entregado']:
        return redirect(url_for('espera_validacion', pedido_id=pedido.ID))

    usuario = Usuario.query.get(pedido.ID_USUARIO)
    if not pedido.CODIGO_TICKET:
        # Generar por si acaso (redundante)
        import random, string
        codigo = f"TICK-{pedido.ID}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"
        pedido.CODIGO_TICKET = codigo
        db.session.commit()
    return render_template('tienda/ticket_impresion.html', pedido=pedido, usuario=usuario)

# ------------------------------------------------------------
# SOLICITUDES
# ------------------------------------------------------------

@app.route('/mis-solicitudes')
@login_required
def mis_solicitudes():
    pedidos = Pedido.query.filter_by(ID_USUARIO=current_user.ID).order_by(Pedido.FECHA.desc()).all()
    
    pedidos_data = []
    for pedido in pedidos:
        archivos = ArchivoPedido.query.filter_by(PEDIDO_ID=pedido.ID).all()
        detalles = DetallePedido.query.filter_by(PEDIDO_ID=pedido.ID).all()
        
        # Obtener nombre del servicio de impresión (funciona para único y múltiple)
        servicio_nombre = None
        tamano = pedido.TAMANO
        
        if pedido.DETALLE_ARCHIVOS:
            # Pedido múltiple: obtener los nombres de servicio de cada archivo
            servicios_nombres = []
            for item in pedido.DETALLE_ARCHIVOS:
                if item.get('servicio_id'):
                    srv = ServicioImpresion.query.get(item['servicio_id'])
                    if srv:
                        servicios_nombres.append(f"{srv.TITULO} · {item.get('tamano', 'Sin tamaño')}")
            servicio_nombre = ', '.join(servicios_nombres) if servicios_nombres else None
        elif pedido.SERVICIO_ID:
            # Pedido único
            srv = ServicioImpresion.query.get(pedido.SERVICIO_ID)
            if srv:
                servicio_nombre = srv.TITULO
        
        # Información adicional para servicios mixtos
        paginas_color = pedido.PAGINAS_COLOR if pedido.PAGINAS_COLOR else None
        comentarios = pedido.COMENTARIOS if pedido.COMENTARIOS else None
        
        pedidos_data.append({
            'pedido': pedido,
            'archivos': archivos,
            'detalles': detalles,
            'servicio_nombre': servicio_nombre,
            'tamano': tamano,
            'paginas_color': paginas_color,
            'comentarios': comentarios
        })
    
    return render_template('tienda/mis_solicitudes.html', pedidos=pedidos_data, active_page='solicitudes')

# ------------------------------------------------------------
# CERRAR SESION
# ------------------------------------------------------------

@app.route('/logout')
@login_required
def logout():
    logout_user()
    # Limpiar todos los mensajes flash anteriores para que no aparezcan en el login
    session.clear()
    return redirect(url_for('login', logout=1))

# ------------------------------------------------------------
# TERMINOS
# ------------------------------------------------------------

@app.route('/politicas')
def politicas():
    config = Configuracion.query.filter_by(CLAVE='politicas').first()
    texto = config.VALOR if config else 'Políticas y términos no configurados aún.'
    return render_template('politicas.html', texto=texto)

# ------------------------------------------------------------
# MI PERFIL
# ------------------------------------------------------------

@app.route('/mi-perfil', methods=['GET', 'POST'])
@login_required
def mi_perfil():
    user = current_user

    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        telefono = request.form.get('telefono', '').strip()
        pregunta1 = request.form.get('pregunta1', '').strip()
        respuesta1 = request.form.get('respuesta1', '').strip()
        pregunta2 = request.form.get('pregunta2', '').strip()
        respuesta2 = request.form.get('respuesta2', '').strip()

        field_errors = {}

        # Validar email
        if not email or '@' not in email or '.' not in email.split('@')[-1]:
            field_errors['email'] = 'Ingresa un correo válido.'
        else:
            # Verificar que no esté en uso por otro usuario
            existente = Usuario.query.filter(Usuario.EMAIL == email, Usuario.ID != user.ID).first()
            if existente:
                field_errors['email'] = 'Este correo ya está en uso por otro usuario.'

        # Validar teléfono
        if not telefono or not telefono.isdigit() or len(telefono) != 11:
            field_errors['telefono'] = 'El teléfono debe tener exactamente 11 dígitos.'
        else:
            existente = Usuario.query.filter(Usuario.TELEFONO == telefono, Usuario.ID != user.ID).first()
            if existente:
                field_errors['telefono'] = 'Este número de teléfono ya está registrado.'
                
        # Validar preguntas de seguridad
        if pregunta1:
            pregunta1 = pregunta1.strip()
            if len(pregunta1) < 4:
                field_errors['pregunta1'] = 'La pregunta debe tener al menos 4 caracteres.'
            elif len(pregunta1) > 20:
                field_errors['pregunta1'] = 'La pregunta no puede tener más de 20 caracteres.'
            elif not re.match(r'^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s?]+$', pregunta1):
                field_errors['pregunta1'] = 'Solo se permiten letras y espacios.'
        if respuesta1:
            respuesta1 = respuesta1.strip()
            if len(respuesta1) > 20:
                field_errors['respuesta1'] = 'La respuesta no puede tener más de 20 caracteres.'

        if pregunta2:
            pregunta2 = pregunta2.strip()
            if len(pregunta2) < 4:
                field_errors['pregunta2'] = 'La pregunta debe tener al menos 4 caracteres.'
            elif len(pregunta2) > 20:
                field_errors['pregunta2'] = 'La pregunta no puede tener más de 20 caracteres.'
            elif not re.match(r'^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s?]+$', pregunta2):
                field_errors['pregunta2'] = 'Solo se permiten letras y espacios.'
        if respuesta2:
            respuesta2 = respuesta2.strip()
            if len(respuesta2) > 20:
                field_errors['respuesta2'] = 'La respuesta no puede tener más de 20 caracteres.'

        # Validar que las preguntas no sean iguales entre sí
        if pregunta1 and pregunta2 and pregunta1.lower() == pregunta2.lower():
            field_errors['pregunta2'] = 'Las preguntas no pueden ser iguales.'

        # Validar que las respuestas no sean iguales entre sí
        if respuesta1 and respuesta2 and respuesta1.lower() == respuesta2.lower():
            field_errors['respuesta2'] = 'Las respuestas no pueden ser iguales.'

        # Validar que las preguntas no sean iguales
        p1 = request.form.get('pregunta1', '').strip()
        p2 = request.form.get('pregunta2', '').strip()
        if p1 and p2 and p1.lower() == p2.lower():
            field_errors['pregunta2'] = 'Las preguntas de seguridad no pueden ser iguales.'

        # Validar que las respuestas no sean iguales
        r1 = request.form.get('respuesta1', '').strip()
        r2 = request.form.get('respuesta2', '').strip()
        if r1 and r2 and r1.lower() == r2.lower():
            field_errors['respuesta2'] = 'Las respuestas de seguridad no pueden ser iguales.'

        if field_errors:
            return render_template('mi_perfil.html', form_data=request.form, field_errors=field_errors)

        # Actualizar datos del usuario
        user.EMAIL = email
        user.TELEFONO = telefono
        if pregunta1:
            user.PREGUNTA1 = pregunta1
        if respuesta1:
            user.RESPUESTA1 = respuesta1.strip().upper()
        if pregunta2:
            user.PREGUNTA2 = pregunta2
        if respuesta2:
            user.RESPUESTA2 = respuesta2.strip().upper()

        db.session.commit()
        flash('Perfil actualizado correctamente.', 'success')
        return redirect(url_for('mi_perfil'))

    # GET: mostrar formulario con datos actuales
    return render_template('mi_perfil.html', form_data=None, field_errors={})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)