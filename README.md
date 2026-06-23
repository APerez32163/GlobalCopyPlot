DISEÑO DE SISTEMA WEB PARA LA GESTIÓN DE SERVICIOS Y SOLICITUDES EN LÍNEA DEL CENTRO DE COPIADO GLOBALCOPYPLOT:

Este sistema aun se encuentra en modelo "prototipo" probado y aprobado (por nuestro cliente), aun no se encuentra en un host
ya que se esta acordando con nuestro cliente precios y fechas tentativas para la implementación del mismo en un host de pago

¿Qué hace el sistema?

GlobalCopyPlot es un sistema Web-to-Print que permite a clientes subir archivos, configurar parámetros de impresión (tamaño, tipo de papel, color, cantidad de páginas a color, etc.) y realizar pedidos que luego son gestionados por administradores y operadores.
Automatiza la cotización (cálculo de precios), el seguimiento de pedidos (estados como Pendiente, En proceso, Listo), la generación de tickets y la comunicación por correo electrónico.
Base de datos (estructura principal)

    usuario: Clientes, administradores y operadores (cédula, nombre, email, contraseña, rol).

    pedido: Cada solicitud de impresión (cliente, fecha, estado, total, tamaño, páginas, código de ticket, archivos adjuntos).

    detalle: Línea de facturación del pedido (cantidad de páginas, precio unitario, subtotal).

    archivo_pedido: Archivos subidos por el cliente (documento a imprimir, comprobante de pago).

    servicio_impresion y servicio_impresion_tamano: Catálogo de servicios (ej. “Impresión B/N Bond”) con sus tamaños (Carta, Oficio…) y precios (normal y color).

    configuracion: Ajustes del negocio (nombre, dirección, datos de PagoMóvil, políticas, etc.).

    catalogo: Imágenes de la galería principal del sitio.

¿Cómo se usa?

    Cliente: Se registra, sube uno o varios archivos (PDF, imágenes, ZIP), elige el servicio y tamaño, define fecha/hora de retiro y realiza el pago mediante transferencia (subiendo comprobante).

    Administrador: Gestiona servicios (crea tarjetas de impresión con precios), confirma pagos, actualiza estados (Listo/Entregado) y visualiza reportes (ventas, ingresos, servicios más solicitados).

    Operador: Consulta los pedidos pendientes, imprime los documentos y los marca como “Listo” para que el cliente reciba su ticket.

    Notificaciones: Al confirmar el pago o al estar listo el pedido, se envían correos automáticos al cliente.

¿Cómo se instala?

    Requisitos: Python 3.10+ instalado en tu PC, servidor MariaDB/MySQL corriendo (ej. XAMPP), y un editor de código (VS Code, etc.).

    Clona o copia la carpeta del proyecto (GlobalCopyplot).

    Crea un entorno virtual con python -m venv venv y actívalo.

    Instala las dependencias con pip install -r requirements.txt.

    Configura la base de datos:

        Crea una base de datos llamada globalcopyplot en MySQL.

        Importa el archivo SQL (dump) que contiene la estructura y datos iniciales (si no lo tienes, ejecuta las sentencias SQL que generaste al limpiar la base de datos).

    Ajusta el archivo app.py con las credenciales de tu servidor MySQL (usuario, contraseña) y una SECRET_KEY segura.

    Ejecuta python app.py y accede a http://localhost:5000.

El panel de administración se accede con un usuario que tenga el campo ES_ADMIN = 1.
El panel de operador, con ES_OPERADOR = 1.

Admin = Usuario: 11111111 (8 veces 1)
        Contraseña : Aa111111.
Operador = Usuario: 00000000 (8 veces 0)
           Contraseña: Aa111111. (igual que admin)
Cliente = Usuario: 22222222 (8 veces 2)
          Contraseña: Aa111111. (igual queadmin)

El panel de administración se accede con un usuario que tenga el campo ES_ADMIN = 1.
El panel de operador, con ES_OPERADOR = 1.
