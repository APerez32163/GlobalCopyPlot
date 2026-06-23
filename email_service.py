# email_service.py
import smtplib
import os
from email.message import EmailMessage
from flask import url_for

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USER = 'SupportGlobalCopyPlot@gmail.com'
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD') 

def enviar_correo_restablecimiento(email_destino, token):
    enlace = url_for('reset_password', token=token, _external=True)
    asunto = 'Restablece tu contraseña - GlobalCopyPlot'
    cuerpo = (
        "Has solicitado restablecer tu contraseña.\n\n"
        f"Para continuar, haz clic en el siguiente enlace (válido por 30 minutos):\n{enlace}\n\n"
        "Si no fuiste tú, ignora este mensaje."
    )

    try:
        msg = EmailMessage()
        msg['Subject'] = asunto
        msg['From'] = SMTP_USER
        msg['To'] = email_destino
        msg.set_content(cuerpo, charset='utf-8')

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def enviar_correo_confirmacion(email_destino, token):
    enlace = url_for('confirmar_email', token=token, _external=True)
    asunto = 'Confirma tu correo - GlobalCopyPlot'
    cuerpo = (
        "Gracias por registrarte en GlobalCopyPlot.\n\n"
        f"Para activar tu cuenta, haz clic en el siguiente enlace (válido por 24 horas):\n{enlace}\n\n"
        "Si no creaste esta cuenta, ignora este mensaje."
    )
    try:
        msg = EmailMessage()
        msg['Subject'] = asunto
        msg['From'] = SMTP_USER
        msg['To'] = email_destino
        msg.set_content(cuerpo, charset='utf-8')
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def enviar_correo_listo(email_destino, nombre, codigo_ticket, pedido_id):
    asunto = 'Tu pedido está listo - GlobalCopyPlot'
    cuerpo = f"""Hola {nombre},

Tu pedido #{pedido_id} ya está listo para retirar.

Código de ticket: {codigo_ticket}

Preséntalo junto con tu cédula al momento del retiro.

Gracias por confiar en GlobalCopyPlot."""
    try:
        msg = EmailMessage()
        msg['Subject'] = asunto
        msg['From'] = SMTP_USER
        msg['To'] = email_destino
        msg.set_content(cuerpo, charset='utf-8')

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)