"""Diagnóstico de la conexión al servidor de correo, paso a paso.

Uso:   python diagnostico_correo.py                (prueba smtp.gmail.com)
       python diagnostico_correo.py smtp.otro.com

No necesita la tienda ni la base de datos: solo comprueba si esta computadora puede hablar con el
servidor y en qué paso se corta (nombre, conexión, cifrado o saludo del servidor)."""
import os
import smtplib
import socket
import ssl
import sys

if '--tienda' in sys.argv:
    # Prueba con la configuración que la tienda tiene guardada y su mismo código de envío,
    # mostrando el detalle técnico completo del error.
    import traceback
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import app as tienda
    from email.message import EmailMessage
    cfg = tienda.leer_config_smtp()
    if not cfg:
        sys.exit("El correo no está activado o falta el servidor en Datos del Negocio.")
    print(f"Configuración guardada: servidor={cfg['host']} puerto={cfg['puerto']} usuario={cfg['usuario']!r} "
          f"contraseña={'cargada (' + str(len(cfg['password'])) + ' caracteres)' if cfg['password'] else 'VACÍA'} "
          f"destino={cfg['destino_admin']!r}")
    destino = cfg['destino_admin'] or cfg['usuario']
    msg = EmailMessage()
    msg['Subject'] = 'Prueba desde la tienda (diagnóstico)'
    msg['To'] = destino
    msg.set_content('Si leés esto, la tienda puede enviar correos.')
    ok, motivo = tienda.enviar_email(msg, cfg, 'prueba', None, esperar=True)
    print('RESULTADO:', 'ENVIADO a ' + destino if ok else 'FALLA: ' + str(motivo))
    if not ok:
        try:
            tienda.contexto_tls()
            clase = smtplib.SMTP_SSL if cfg['puerto'] == 465 else smtplib.SMTP
            with clase(cfg['host'], cfg['puerto'], timeout=30) as s:
                if cfg['puerto'] != 465:
                    s.starttls(context=tienda.contexto_tls())
                s.login(cfg['usuario'], cfg['password'].replace(' ', ''))
                s.send_message(msg)
        except Exception:
            print('\nDetalle técnico:')
            traceback.print_exc()
    sys.exit(0)

host = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith('--') else 'smtp.gmail.com'
print(f"Probando {host}\n")

try:
    ips = sorted({i[4][0] for i in socket.getaddrinfo(host, 465, proto=socket.IPPROTO_TCP)})
    print(f"1. Nombre del servidor: OK ({', '.join(ips)})")
except OSError as e:
    print(f"1. Nombre del servidor: FALLA ({e}). Revisá la conexión a internet o el DNS.")
    sys.exit(1)

try:
    import certifi
    contexto = ssl.create_default_context(cafile=certifi.where())
    print("   Certificados: se usa certifi")
except ImportError:
    contexto = ssl.create_default_context()
    print("   Certificados: los del sistema (certifi no está instalado: pip install certifi)")

for puerto, modo in ((465, 'SSL directo'), (587, 'STARTTLS')):
    print(f"\nPuerto {puerto} ({modo})")
    try:
        with socket.create_connection((host, puerto), timeout=15):
            print("   2. Conexión TCP: OK")
    except OSError as e:
        print(f"   2. Conexión TCP: FALLA ({e}). Algo bloquea este puerto (red, firewall o antivirus).")
        continue
    try:
        if puerto == 465:
            servidor = smtplib.SMTP_SSL(host, puerto, timeout=15, context=contexto)
        else:
            servidor = smtplib.SMTP(host, puerto, timeout=15)
        with servidor:
            print(f"   3. Saludo del servidor: OK ({servidor.ehlo()[1].decode(errors='replace').splitlines()[0]})")
            if puerto == 587:
                servidor.starttls(context=contexto)
                print("   4. Cifrado STARTTLS: OK")
            print("   => Este puerto funciona. Usalo en el panel.")
    except ssl.SSLError as e:
        print(f"   3. Cifrado: FALLA ({e}). Problema de certificados: pip install --upgrade certifi")
    except smtplib.SMTPServerDisconnected as e:
        print(f"   3. Saludo del servidor: FALLA, la conexión se cortó ({e}). "
              "Suele ser un antivirus, firewall, proxy o proveedor de internet que interfiere con el correo.")
    except Exception as e:
        print(f"   3. Saludo del servidor: FALLA ({type(e).__name__}: {e})")
print("\n--- Prueba completa (opcional): inicio de sesión y envío real ---")
usuario = input("Tu Gmail (Enter para omitir): ").strip()
if usuario:
    import getpass
    clave = getpass.getpass("Contraseña de aplicación (no se ve al escribir): ").replace(" ", "")
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = usuario, usuario, "Prueba de diagnostico_correo.py"
    msg.set_content("Si leés esto, el envío funciona.")
    for puerto in (465, 587):
        paso = "conectar"
        try:
            if puerto == 465:
                servidor = smtplib.SMTP_SSL(host, puerto, timeout=30, context=contexto)
            else:
                servidor = smtplib.SMTP(host, puerto, timeout=30)
            with servidor:
                if puerto == 587:
                    paso = "STARTTLS"
                    servidor.starttls(context=contexto)
                paso = "iniciar sesión"
                servidor.login(usuario, clave)
                paso = "enviar"
                servidor.send_message(msg)
            print(f"Puerto {puerto}: OK, mensaje enviado a {usuario}. Revisá tu bandeja.")
            break
        except Exception as e:
            print(f"Puerto {puerto}: FALLA al {paso} -> {type(e).__name__}: {e}")

print("\nSi ningún puerto funciona, probá desde otra red (datos del celular) o desactivá un momento el "
      "análisis de correo/SSL del antivirus.")
