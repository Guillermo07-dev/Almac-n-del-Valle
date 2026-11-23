"""Punto de entrada para producción.

    gunicorn --workers 1 --threads 4 --bind 127.0.0.1:8000 wsgi:app

Se usa un solo proceso (con varios hilos) porque la base es SQLite y el hilo de
backups automáticos debe existir una única vez.
"""
import threading

import app as tienda

tienda.init_db()
threading.Thread(target=tienda._backup_scheduler, daemon=True).start()

app = tienda.app  # gunicorn busca la variable "app" en este módulo
