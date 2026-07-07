from flask_login import UserMixin
from extensions import db
from datetime import datetime

class Usuario(UserMixin, db.Model):
    __tablename__ = 'usuario'
    ID = db.Column(db.Integer, primary_key=True)
    ID_USUARIO = db.Column(db.String(15), unique=True, nullable=False)
    NOMBRE = db.Column(db.String(50), nullable=False)
    APELLIDO = db.Column(db.String(50), nullable=False)
    EMAIL = db.Column(db.String(120), unique=True, nullable=False)
    CONTRASEÑA = db.Column('CONTRASEÑA', db.String(255), nullable=False)
    TELEFONO = db.Column(db.String(11), nullable=False)
    CONFIRMADO = db.Column(db.Boolean, default=False)
    ES_ADMIN = db.Column(db.Boolean, default=False)
    ES_OPERADOR = db.Column(db.Boolean, default=False)
    PREGUNTA1 = db.Column(db.String(200), nullable=True)
    RESPUESTA1 = db.Column(db.String(200), nullable=True)
    PREGUNTA2 = db.Column(db.String(200), nullable=True)
    RESPUESTA2 = db.Column(db.String(200), nullable=True)
    INTENTOS_FALLIDOS = db.Column(db.Integer, default=0)
    BLOQUEADO_HASTA = db.Column(db.DateTime, nullable=True)

    def get_id(self):
        return str(self.ID)

class Catalogo(db.Model):
    __tablename__ = 'catalogo'
    ID = db.Column(db.Integer, primary_key=True)
    IMAGEN = db.Column(db.String(50), unique=True, nullable=False)
    ORDEN = db.Column(db.Integer, default=0)

class Pedido(db.Model):
    __tablename__ = 'pedido'
    ID = db.Column(db.Integer, primary_key=True)
    ID_USUARIO = db.Column(db.Integer, db.ForeignKey('usuario.ID'), nullable=False)
    FECHA = db.Column(db.DateTime, default=datetime.utcnow)
    ESTADO = db.Column(db.String(20), default='Pendiente')
    TOTAL = db.Column(db.Numeric(10,2), default=0.00)
    COMENTARIOS = db.Column(db.Text, nullable=True)       
    TAMANO = db.Column(db.String(20), nullable=True)         
    FECHA_RETIRO = db.Column(db.Date, nullable=True)
    HORA_RETIRO = db.Column(db.Time, nullable=True)
    CODIGO_TICKET = db.Column(db.String(20), unique=True, nullable=True)
    VISTO_ADMIN = db.Column(db.Boolean, default=False)
    VISTO_OPERADOR = db.Column(db.Boolean, default=False)
    PAGINAS = db.Column(db.Integer, nullable=True)   # páginas detectadas
    SERVICIO_ID = db.Column(db.Integer, db.ForeignKey('servicio_impresion.ID'),nullable=True)
    DETALLE_ARCHIVOS = db.Column(db.JSON, nullable=True)
    REFERENCIA_PAGO = db.Column(db.String(100), nullable=True)
    PAGINAS_COLOR = db.Column(db.String(255), nullable=True)
    
class DetallePedido(db.Model):
    __tablename__ = 'detalle'
    ID = db.Column(db.Integer, primary_key=True)
    PEDIDO_ID = db.Column(db.Integer, db.ForeignKey('pedido.ID'), nullable=False)
    CANTIDAD = db.Column(db.Integer, nullable=False)
    PRECIO_UNITARIO = db.Column(db.Numeric(10,2), nullable=True)
    SUBTOTAL = db.Column(db.Numeric(10,2), nullable=False)

class ArchivoPedido(db.Model):
    __tablename__ = 'archivo_pedido'
    ID = db.Column(db.Integer, primary_key=True)
    PEDIDO_ID = db.Column(db.Integer, db.ForeignKey('pedido.ID'), nullable=False)
    NOMBRE_ARCHIVO = db.Column(db.String(255), nullable=False)
    RUTA = db.Column(db.String(255), nullable=False)

class Configuracion(db.Model):
    __tablename__ = 'configuracion'
    ID = db.Column(db.Integer, primary_key=True)
    CLAVE = db.Column(db.String(100), unique=True, nullable=False)
    VALOR = db.Column(db.Text)
    DESCRIPCION = db.Column(db.String(255))

class ServicioImpresion(db.Model):
    __tablename__ = 'servicio_impresion'
    ID = db.Column(db.Integer, primary_key=True)
    TITULO = db.Column(db.String(100), nullable=False)
    DESCRIPCION = db.Column(db.Text, nullable=True)
    ACTIVO = db.Column(db.Boolean, default=True)
    ES_MIXTO = db.Column(db.Boolean, default=False)
    tamanos = db.relationship('ServicioImpresionTamano', backref='servicio', cascade='all, delete-orphan', lazy='select')

class ServicioImpresionTamano(db.Model):
    __tablename__ = 'servicio_impresion_tamano'
    ID = db.Column(db.Integer, primary_key=True)
    SERVICIO_ID = db.Column(db.Integer, db.ForeignKey('servicio_impresion.ID'),
                            nullable=False)
    NOMBRE = db.Column(db.String(50), nullable=False)
    PRECIO_BN = db.Column(db.Numeric(10, 2), nullable=False)
    PRECIO_COLOR = db.Column(db.Numeric(10, 2), nullable=False)