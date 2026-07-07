# email_service.py (correo desactivado para pruebas) se utiliza este ya que es para simulaciones, "email_service_orig.py" es el que realmente hace la funcion, no se implementa porque no funciona en hosting gratuito", al pgar un host se implementará

def enviar_correo_restablecimiento(email, token):
    print(f"[SIMULADO] Se enviaría enlace de restablecimiento a {email} con token {token}")
    return True, None

def enviar_correo_confirmacion(email, token):
    print(f"[SIMULADO] Se enviaría enlace de confirmación a {email} con token {token}")
    return True, None

def enviar_correo_listo(*args, **kwargs):
    print("[SIMULADO] Se enviaría correo de pedido listo.")
    return True, None