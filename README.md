# GlobalCopyPlot

Sistema web para pedidos de impresión en línea. Permite subir documentos, configurar impresiones, realizar pagos y gestionar todo desde un panel de administración y operador.

## 🌐 Acceso

https://globalcopyplot.onrender.com/

La aplicación está desplegada en **Render** (plan gratuito) y la base de datos en **Clever Cloud**.  
*El host gratuito puede hibernar; la primera carga puede tardar unos segundos.*

## ⚙️ Funcionalidades principales

- Subida de archivos (PDF, PPTX, imágenes, ZIP de imágenes)
- Detección automática de páginas
- Selección de tamaño y tipo de impresión (B/N o color)
- Pagos (pago movil) con comprobante
- Panel de administración para gestionar pedidos, usuarios, servicios y reportes, etc...
- Roles: cliente, operador, administrador
- Recuperación de contraseña

## 📧 Envío de correos

**Actualmente deshabilitado** para evitar errores en el entorno gratuito. Las funciones de `email_service.py` están en modo simulación (solo imprimen en consola). Se puede reactivar configurando la variable `SMTP_PASSWORD` en el entorno.

## 🛠️ Roles

**Admin** :
- V-11111111
- contraseña: Aa111111.

**Operador**:
- V-22222222
- contraseña: Aa111111.

**Cliente**:
- V-33333333
- contraseña: Aa111111.

- **Backend**: Python + Flask
- **Base de datos**: MySQL/MariaDB
- **Frontend**: HTML, CSS, Bootstrap 5, JavaScript

---

*Proyecto desarrollado con fines educativos.*
*Andres Perez, Kenyerson Crespo, Yovanny Monrroy, Franyelber Salas*
