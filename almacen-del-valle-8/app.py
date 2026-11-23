import difflib
import hashlib
import io
import os
import random
import re
import secrets
import shutil
import smtplib
import socket
import ssl
import sqlite3
import string
import threading
import time
import unicodedata
import uuid
import base64
import json
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from functools import wraps
from html import escape

import openpyxl
from flask import (Flask, render_template, request, redirect,
                   url_for, session, jsonify, send_file, flash, abort, has_request_context)
from flask_wtf.csrf import CSRFError, CSRFProtect, generate_csrf
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode.code128 import Code128
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer)
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

# En producción conviene definir SECRET_KEY como variable de entorno.
# Los valores por defecto sirven solo para desarrollo local.
app.secret_key = os.environ.get('SECRET_KEY', 'almacendelvalle_secret_2026')

# ---------------------------------------------------------------------------
# Seguridad: CSRF, cookies de sesión y HTTPS
# ---------------------------------------------------------------------------

# Con USE_HTTPS=1 (producción, detrás de Caddy o nginx con certificado SSL) el sitio
# fuerza HTTPS y las cookies de sesión solo viajan por conexiones cifradas.
USAR_HTTPS = os.environ.get('USE_HTTPS', '0') == '1'

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,       # el JavaScript de la página no puede leer la cookie
    SESSION_COOKIE_SAMESITE='Lax',      # no se envía en peticiones originadas en otros sitios
    SESSION_COOKIE_SECURE=USAR_HTTPS,
    PREFERRED_URL_SCHEME='https' if USAR_HTTPS else 'http',
    WTF_CSRF_TIME_LIMIT=None,           # el token vale mientras dure la sesión
)

if USAR_HTTPS:
    # El proxy (Caddy/nginx) termina el SSL y avisa por cabeceras si la conexión original era HTTPS
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

# Todo formulario y toda llamada fetch() que modifique datos debe llevar un token válido
csrf = CSRFProtect(app)

DB_PATH        = 'data/almacendelvalle.db'
ADMIN_USERNAME = 'admin'
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'almacen123')

BACKUP_FOLDER = 'data/backups'
os.makedirs(BACKUP_FOLDER, exist_ok=True)

# Secciones del panel a las que se les puede dar o quitar acceso por usuario
PERMISOS_DISPONIBLES = {
    'dashboard': 'Dashboard',
    'productos': 'Productos',
    'categorias': 'Categorías',
    'pedidos': 'Pedidos',
    'clientes': 'Clientes',
    'stock': 'Alertas de stock',
    'reportes': 'Reportes',
    'ofertas': 'Ofertas de la semana',
    'gastos': 'Gastos / Compras',
    'etiquetas': 'Etiquetas de precio',
    'backups': 'Backups',
    'negocio': 'Datos del Negocio',
    'cupones': 'Cupones de Descuento',
    'chatbot': 'Chatbot',
    'historial': 'Historial de precios',
    'facturacion': 'Facturación',
}

# IVA aplicado a las facturas (Argentina, condición Consumidor Final)
IVA_PORC = 0.21

# Debajo de esta cantidad de unidades, un producto se marca como "stock bajo"
LOW_STOCK_THRESHOLD = 10

MESES_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
            'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# Estados de un pedido. "etiqueta" se usa en el panel y en las listas; "titulo" y "detalle"
# son lo que lee el cliente en la línea de tiempo de seguimiento.
ESTADOS_PEDIDO = {
    'pendiente_pago': {'etiqueta': 'Pendiente de pago', 'titulo': 'Pedido recibido',
                       'detalle': 'Registramos tu pedido y esperamos que se acredite el pago.'},
    'pagado':         {'etiqueta': 'Pagado', 'titulo': 'Pago confirmado',
                       'detalle': 'Recibimos tu pago y emitimos la factura.'},
    'preparando':     {'etiqueta': 'Preparando', 'titulo': 'Estamos preparando tu envío',
                       'detalle': 'Estamos armando tu pedido para despacharlo.'},
    'enviado':        {'etiqueta': 'Enviado', 'titulo': 'Hemos enviado tu pedido',
                       'detalle': 'Tu pedido salió del local y va hacia tu domicilio.'},
    'listo_retiro':   {'etiqueta': 'Listo para retirar', 'titulo': 'Listo para retirar',
                       'detalle': 'Tu pedido está listo: pasá a retirarlo por el local.'},
    'entregado':      {'etiqueta': 'Entregado', 'titulo': 'Tu pedido ha sido entregado',
                       'detalle': '¡Gracias por tu compra!'},
    'cancelado':      {'etiqueta': 'Cancelado', 'titulo': 'Pedido cancelado',
                       'detalle': 'Este pedido fue cancelado.'},
}

# Recorrido de un pedido según cómo se entrega
PASOS_DOMICILIO = ['pendiente_pago', 'pagado', 'preparando', 'enviado', 'entregado']
PASOS_RETIRO    = ['pendiente_pago', 'pagado', 'preparando', 'listo_retiro', 'entregado']

# Desde "pagado" en adelante el pedido ya está pago, así que debe tener su factura
ESTADOS_PAGADOS = {'pagado', 'preparando', 'enviado', 'listo_retiro', 'entregado'}

# Nombre legible de cada medio de pago, para mostrar en los correos
NOMBRES_METODO_PAGO = {
    'mercado_pago': 'Mercado Pago',
    'tarjeta_credito': 'Tarjeta de crédito',
    'tarjeta_debito': 'Tarjeta de débito',
    'transferencia': 'Transferencia bancaria',
    'efectivo': 'Efectivo (Rapipago / Pago Fácil)',
}

os.makedirs('data', exist_ok=True)
os.makedirs('static/uploads', exist_ok=True)


# ---------------------------------------------------------------------------
# Decoradores de sesión
# ---------------------------------------------------------------------------

def admin_required(fallback):
    """Exige sesión de admin; si no hay sesión, devuelve fallback() en su lugar."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not session.get('admin_logged_in'):
                return fallback()
            return fn(*a, **kw)
        return wrapper
    return deco


def permisos_de_usuario(fila):
    """Lista de secciones permitidas de un usuario. Sin permisos cargados no tiene acceso a nada."""
    return [p for p in (fila['permisos'] or '').split(',') if p in PERMISOS_DISPONIBLES]


def tiene_permiso(seccion):
    """El dueño siempre tiene acceso total; el resto depende de sus permisos guardados."""
    if session.get('es_dueno'):
        return True
    return seccion in (session.get('permisos') or [])


app.jinja_env.globals['puede'] = tiene_permiso


def permiso_required(seccion):
    """Exige sesión de admin Y permiso sobre esa sección puntual del panel."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not session.get('admin_logged_in'):
                return redirect(url_for('admin_login'))
            if not tiene_permiso(seccion):
                if request.is_json:
                    return jsonify({'error': 'No tenés permiso para esta sección'}), 403
                flash('No tenés permiso para acceder a esa sección del panel.', 'error')
                return redirect(url_for('admin'))
            return fn(*a, **kw)
        return wrapper
    return deco


def dueno_required(fallback):
    """Exige sesión de admin y que sea el dueño del negocio; si no hay sesión, devuelve fallback()."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not session.get('admin_logged_in'):
                return fallback()
            if not session.get('es_dueno'):
                return redirect(url_for('admin'))
            return fn(*a, **kw)
        return wrapper
    return deco


def cliente_required(fallback):
    """Exige sesión de cliente logueado; si no hay sesión, devuelve fallback()."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not session.get('cliente_id'):
                return fallback()
            return fn(*a, **kw)
        return wrapper
    return deco


# ---------------------------------------------------------------------------
# Base de datos
# ---------------------------------------------------------------------------

def get_db():
    """Devuelve una conexión a SQLite con Row como tipo de fila."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """Crea las tablas si no existen."""
    with get_db() as db:
        db.executescript("""
            CREATE TABLE IF NOT EXISTS categories (
                id          TEXT PRIMARY KEY,
                name        TEXT NOT NULL,
                icon        TEXT DEFAULT '📦',
                color       TEXT DEFAULT '#DCEDE6'
            );

            CREATE TABLE IF NOT EXISTS products (
                id            TEXT PRIMARY KEY,
                code          TEXT DEFAULT '',
                name          TEXT NOT NULL,
                category_id   TEXT NOT NULL REFERENCES categories(id) ON DELETE CASCADE,
                brand         TEXT DEFAULT '',
                unidad        TEXT DEFAULT 'unidad',
                price         REAL NOT NULL DEFAULT 0,
                old_price     REAL,
                stock         INTEGER DEFAULT 0,
                description   TEXT DEFAULT '',
                info_adicional TEXT DEFAULT '',
                icon          TEXT DEFAULT '📦',
                image         TEXT DEFAULT '',
                destacado     INTEGER DEFAULT 0,
                active        INTEGER DEFAULT 1,
                created_at    TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS price_history (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id  TEXT NOT NULL REFERENCES products(id) ON DELETE CASCADE,
                price_old   REAL NOT NULL,
                price_new   REAL NOT NULL,
                changed_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Ofertas de la semana: se muestran como banner gigante en la tienda
            CREATE TABLE IF NOT EXISTS ofertas_semana (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo      TEXT NOT NULL,
                texto       TEXT NOT NULL,
                producto_id TEXT REFERENCES products(id) ON DELETE SET NULL,
                categoria_id TEXT REFERENCES categories(id) ON DELETE SET NULL,
                color       TEXT DEFAULT 'verde',
                activa      INTEGER DEFAULT 1,
                orden       INTEGER DEFAULT 0,
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Gastos / compras: mercadería, sueldos, servicios, impuestos, etc.
            CREATE TABLE IF NOT EXISTS gastos (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                descripcion      TEXT NOT NULL,
                monto            REAL NOT NULL,
                categoria        TEXT DEFAULT 'General',
                proveedor        TEXT DEFAULT '',
                cuit             TEXT DEFAULT '',
                razon_social     TEXT DEFAULT '',
                domicilio_fiscal TEXT DEFAULT '',
                condicion_iva    TEXT DEFAULT '',
                nro_factura      TEXT DEFAULT '',
                tipo_factura     TEXT DEFAULT '',
                nombre           TEXT DEFAULT '',
                apellido         TEXT DEFAULT '',
                tipo_impuesto    TEXT DEFAULT '',
                nota             TEXT DEFAULT '',
                year             INTEGER NOT NULL,
                month            INTEGER NOT NULL,
                day              INTEGER NOT NULL,
                created_at       TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Clientes de la tienda (login público, distinto del admin)
            CREATE TABLE IF NOT EXISTS clientes (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre          TEXT NOT NULL,
                apellido        TEXT NOT NULL,
                email           TEXT UNIQUE NOT NULL,
                password_hash   TEXT NOT NULL,
                telefono        TEXT DEFAULT '',
                calle           TEXT DEFAULT '',
                numero          TEXT DEFAULT '',
                barrio          TEXT DEFAULT '',
                dni_cuit        TEXT DEFAULT '',
                created_at      TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Pedidos (carrito confirmado / orden de compra)
            CREATE TABLE IF NOT EXISTS pedidos (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id      INTEGER NOT NULL REFERENCES clientes(id),
                subtotal        REAL NOT NULL DEFAULT 0,
                envio           REAL NOT NULL DEFAULT 0,
                total           REAL NOT NULL DEFAULT 0,
                metodo_entrega  TEXT DEFAULT 'domicilio',
                direccion_envio TEXT DEFAULT '',
                estado          TEXT DEFAULT 'pendiente_pago',
                created_at      TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Historial de estados de cada pedido, para mostrar la línea de tiempo con sus fechas
            CREATE TABLE IF NOT EXISTS pedido_estados (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id   INTEGER NOT NULL REFERENCES pedidos(id) ON DELETE CASCADE,
                estado      TEXT NOT NULL,
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Chatbot: respuestas que carga el administrador ("entrenamiento")
            CREATE TABLE IF NOT EXISTS chatbot_respuestas (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                pregunta        TEXT NOT NULL,
                palabras_clave  TEXT DEFAULT '',
                respuesta       TEXT NOT NULL,
                activa          INTEGER DEFAULT 1,
                usos            INTEGER DEFAULT 0,
                created_at      TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Chatbot: cada consulta recibida; las no resueltas alimentan el entrenamiento
            CREATE TABLE IF NOT EXISTS chatbot_mensajes (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id  INTEGER,
                mensaje     TEXT NOT NULL,
                resuelto    INTEGER DEFAULT 1,
                descartado  INTEGER DEFAULT 0,
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Chatbot (Alice): frases que aprendió sola al ver cómo reformulan los clientes.
            -- destino: 'resp:<id>' (respuesta cargada) o 'int:<nombre>' (intención integrada)
            CREATE TABLE IF NOT EXISTS chatbot_aprendido (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                frase           TEXT NOT NULL,
                destino         TEXT NOT NULL,
                confirmaciones  INTEGER DEFAULT 1,
                sesiones        TEXT DEFAULT '',
                estado          TEXT DEFAULT 'pendiente',
                usos            INTEGER DEFAULT 0,
                created_at      TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE (frase, destino)
            );

            -- Intentos fallidos de ingreso (y pedidos de recuperación) por cuenta o por IP
            CREATE TABLE IF NOT EXISTS intentos_login (
                clave            TEXT PRIMARY KEY,
                intentos         INTEGER DEFAULT 0,
                bloqueado_hasta  REAL DEFAULT 0,
                ultimo           REAL DEFAULT 0
            );

            -- Enlaces de recuperación de contraseña; se guarda el hash del token, nunca el token
            CREATE TABLE IF NOT EXISTS recuperacion_clave (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                ambito      TEXT NOT NULL,
                usuario_id  INTEGER NOT NULL,
                token_hash  TEXT NOT NULL UNIQUE,
                vence       REAL NOT NULL,
                usado       INTEGER DEFAULT 0,
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Registro de cada correo que la tienda intentó mandar y cómo salió
            CREATE TABLE IF NOT EXISTS correos_enviados (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo          TEXT DEFAULT '',
                pedido_id     INTEGER,
                destinatario  TEXT,
                asunto        TEXT,
                estado        TEXT,
                detalle       TEXT DEFAULT '',
                created_at    TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS pedido_items (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id       INTEGER NOT NULL REFERENCES pedidos(id) ON DELETE CASCADE,
                product_id      TEXT NOT NULL,
                product_name    TEXT NOT NULL,
                product_price   REAL NOT NULL,
                qty             INTEGER NOT NULL,
                subtotal        REAL NOT NULL
            );

            -- Pagos: registra el intento/resultado de cada pasarela
            CREATE TABLE IF NOT EXISTS pagos (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id       INTEGER NOT NULL REFERENCES pedidos(id) ON DELETE CASCADE,
                metodo          TEXT NOT NULL,
                estado          TEXT DEFAULT 'aprobado',
                monto           REAL NOT NULL,
                referencia      TEXT DEFAULT '',
                detalle         TEXT DEFAULT '',
                created_at      TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Facturas emitidas por pedido
            CREATE TABLE IF NOT EXISTS facturas (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id       INTEGER NOT NULL UNIQUE REFERENCES pedidos(id) ON DELETE CASCADE,
                numero          TEXT NOT NULL UNIQUE,
                tipo            TEXT DEFAULT 'B',
                cliente_nombre  TEXT NOT NULL,
                cliente_email   TEXT NOT NULL,
                cliente_dni     TEXT DEFAULT '',
                subtotal        REAL NOT NULL,
                iva             REAL NOT NULL,
                total           REAL NOT NULL,
                created_at      TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS usuarios (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre      TEXT NOT NULL,
                apellido    TEXT DEFAULT '',
                usuario     TEXT NOT NULL UNIQUE,
                password    TEXT NOT NULL,
                email       TEXT DEFAULT '',
                permisos    TEXT DEFAULT '',
                es_dueno    INTEGER DEFAULT 0,
                activo      INTEGER DEFAULT 1,
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS backup_log (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                filename   TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Datos del negocio: horarios, direcciones, contactos y redes sociales
            CREATE TABLE IF NOT EXISTS horarios_atencion (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                dias        TEXT NOT NULL,
                hora_inicio TEXT NOT NULL,
                hora_fin    TEXT NOT NULL,
                orden       INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS direcciones (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre      TEXT DEFAULT 'Casa Central',
                direccion   TEXT NOT NULL,
                ciudad      TEXT DEFAULT '',
                provincia   TEXT DEFAULT '',
                orden       INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS contactos (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo        TEXT NOT NULL DEFAULT 'Teléfono',
                etiqueta    TEXT DEFAULT '',
                valor       TEXT NOT NULL,
                orden       INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS redes_sociales (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                plataforma  TEXT NOT NULL,
                url         TEXT NOT NULL,
                orden       INTEGER DEFAULT 0
            );

            -- Configuración general (envío, etc.) como clave/valor
            CREATE TABLE IF NOT EXISTS configuracion (
                clave  TEXT PRIMARY KEY,
                valor  TEXT NOT NULL
            );

            -- Favoritos de clientes
            CREATE TABLE IF NOT EXISTS favoritos (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id  INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
                product_id  TEXT NOT NULL REFERENCES products(id) ON DELETE CASCADE,
                created_at  TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE(cliente_id, product_id)
            );

            -- Reseñas y calificaciones de productos
            CREATE TABLE IF NOT EXISTS resenas (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id     TEXT NOT NULL REFERENCES products(id) ON DELETE CASCADE,
                cliente_id     INTEGER NOT NULL REFERENCES clientes(id),
                cliente_nombre TEXT NOT NULL,
                estrellas      INTEGER NOT NULL,
                comentario     TEXT DEFAULT '',
                created_at     TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Cupones de descuento
            CREATE TABLE IF NOT EXISTS cupones (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo           TEXT NOT NULL UNIQUE,
                tipo             TEXT NOT NULL DEFAULT 'porcentaje',
                valor            REAL NOT NULL,
                monto_minimo     REAL DEFAULT 0,
                usos_maximos     INTEGER DEFAULT 0,
                usos_actuales    INTEGER DEFAULT 0,
                fecha_expiracion TEXT DEFAULT '',
                activo           INTEGER DEFAULT 1,
                created_at       TEXT DEFAULT (datetime('now','localtime'))
            );

            -- Carrito guardado en el servidor por cliente (persiste entre sesiones/dispositivos)
            CREATE TABLE IF NOT EXISTS carrito_guardado (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id  INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
                product_id  TEXT NOT NULL,
                qty         INTEGER NOT NULL,
                updated_at  TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE(cliente_id, product_id)
            );

            -- Registro de auditoría de acciones del panel
            CREATE TABLE IF NOT EXISTS auditoria (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario     TEXT NOT NULL,
                accion      TEXT NOT NULL,
                detalle     TEXT DEFAULT '',
                created_at  TEXT DEFAULT (datetime('now','localtime'))
            );
        """)

        # El DNI identifica al cliente en el login, por eso no puede repetirse.
        # El índice es parcial: las cuentas sin DNI cargado (valor vacío) no entran en la regla.
        db.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_dni
            ON clientes(dni_cuit) WHERE dni_cuit <> ''
        """)

        # Pedidos creados antes del historial: se les registra el estado en el que están
        db.execute("""
            INSERT INTO pedido_estados (pedido_id, estado, created_at)
            SELECT id, estado, created_at FROM pedidos
            WHERE id NOT IN (SELECT pedido_id FROM pedido_estados)
        """)

        cols_pedidos = [r['name'] for r in db.execute("PRAGMA table_info(pedidos)").fetchall()]
        if 'cupon_codigo' not in cols_pedidos:
            db.execute("ALTER TABLE pedidos ADD COLUMN cupon_codigo TEXT DEFAULT ''")
        if 'descuento' not in cols_pedidos:
            db.execute("ALTER TABLE pedidos ADD COLUMN descuento REAL DEFAULT 0")

        # Confirmación del correo: las cuentas que ya existían quedan confirmadas
        cols_clientes = [r['name'] for r in db.execute("PRAGMA table_info(clientes)").fetchall()]
        if 'email_confirmado' not in cols_clientes:
            db.execute("ALTER TABLE clientes ADD COLUMN email_confirmado INTEGER DEFAULT 1")

        # Se agregaron los permisos "Historial de precios" y "Facturación", que antes venían con
        # Productos y con Datos del Negocio: a quienes ya tenían aquellos se les conservan, una sola vez.
        if not db.execute("SELECT 1 FROM configuracion WHERE clave='migr_permisos_historial_facturacion'").fetchone():
            for u in db.execute("SELECT id, permisos FROM usuarios WHERE es_dueno=0").fetchall():
                lista = [p for p in (u['permisos'] or '').split(',') if p]
                if 'productos' in lista and 'historial' not in lista:
                    lista.append('historial')
                if 'negocio' in lista and 'facturacion' not in lista:
                    lista.append('facturacion')
                db.execute("UPDATE usuarios SET permisos=? WHERE id=?", (','.join(lista), u['id']))
            db.execute("INSERT INTO configuracion (clave, valor) VALUES ('migr_permisos_historial_facturacion','1')")

        # Pedido de Factura A: CUIT y razón social del cliente que la solicita
        for columna in ('factura_a_cuit', 'factura_a_razon'):
            if columna not in cols_pedidos:
                db.execute(f"ALTER TABLE pedidos ADD COLUMN {columna} TEXT DEFAULT ''")

        # Datos de autorización (tipo ARCA) de cada factura
        cols_facturas = [r['name'] for r in db.execute("PRAGMA table_info(facturas)").fetchall()]
        for columna, definicion in [('cae', "TEXT DEFAULT ''"), ('cae_vencimiento', "TEXT DEFAULT ''"),
                                    ('modo_arca', "TEXT DEFAULT ''"), ('punto_venta', 'INTEGER DEFAULT 0'),
                                    ('nro_comprobante', 'INTEGER DEFAULT 0'),
                                    ('cliente_condicion_iva', "TEXT DEFAULT 'Consumidor Final'"),
                                    ('cliente_domicilio', "TEXT DEFAULT ''"), ('condicion_venta', "TEXT DEFAULT ''")]:
            if columna not in cols_facturas:
                db.execute(f"ALTER TABLE facturas ADD COLUMN {columna} {definicion}")

        # Alice registra de dónde salió cada respuesta y si al cliente le sirvió
        cols_mensajes = [r['name'] for r in db.execute("PRAGMA table_info(chatbot_mensajes)").fetchall()]
        if 'origen' not in cols_mensajes:
            db.execute("ALTER TABLE chatbot_mensajes ADD COLUMN origen TEXT DEFAULT ''")
        if 'util' not in cols_mensajes:
            db.execute("ALTER TABLE chatbot_mensajes ADD COLUMN util INTEGER")

        # Asegura que existan las claves de configuración nuevas en bases ya creadas
        db.executemany(
            "INSERT OR IGNORE INTO configuracion (clave, valor) VALUES (?,?)",
            [('envio_gratis_desde', '50000'), ('costo_envio', '2500'), ('compra_minima', '0'),
             ('smtp_activo', '0'), ('smtp_host', ''), ('smtp_puerto', '587'),
             ('smtp_usuario', ''), ('smtp_password', ''), ('smtp_email_destino', '')])

        # Dos respuestas de ejemplo, solo la primera vez, para que el panel del chatbot no arranque vacío
        if not db.execute("SELECT 1 FROM configuracion WHERE clave='chatbot_semilla'").fetchone():
            db.executemany(
                "INSERT INTO chatbot_respuestas (pregunta, palabras_clave, respuesta) VALUES (?,?,?)", [
                    ('¿Puedo retirar mi pedido en el local?', 'retirar pedido, retiro en el local, pasar a buscar',
                     'Sí. Al finalizar la compra elegí "Retiro en sucursal": no tiene costo y en el seguimiento '
                     'de tu pedido vas a ver cuándo está listo para retirar.'),
                    ('¿Cómo recibo mi factura?', 'recibir factura, factura por mail, descargar factura',
                     'Cuando se confirma el pago te enviamos la factura por correo, con el PDF adjunto. '
                     'También la podés descargar desde "Mis pedidos".'),
                ])
            db.execute("INSERT INTO configuracion (clave, valor) VALUES ('chatbot_semilla', '1')")

    seed_data()


def seed_data():
    """Carga datos de ejemplo la primera vez que se crea la base."""
    with get_db() as db:
        existing = db.execute("SELECT COUNT(*) FROM categories").fetchone()[0]
        if existing > 0:
            return

        categorias = [
            ('frutasyverduras', 'Frutas y Verduras', '🍎', '#E9F3D9'),
            ('frescos',         'Frescos',            '🥬', '#DCEDE6'),
            ('carniceria',      'Carnicería',         '🥩', '#F5DADA'),
            ('lacteos',         'Lácteos',            '🥛', '#E4EEF7'),
            ('panificados',     'Panificados',        '🍞', '#F4E3C6'),
            ('almacen',         'Almacén',            '🥫', '#F1E3EC'),
            ('bebidas',         'Bebidas',            '🥤', '#E1F0FB'),
            ('limpieza',        'Limpieza',           '🧽', '#E6F3EC'),
        ]
        db.executemany(
            "INSERT INTO categories (id, name, icon, color) VALUES (?,?,?,?)",
            categorias)

        productos = [
            ('Tomate perita x kg', 'frutasyverduras', 'Frutas y Verduras', 'kg', 1890, 2400, 120, '🍅', 1),
            ('Banana ecuatoriana x kg', 'frutasyverduras', 'Frutas y Verduras', 'kg', 1490, None, 80, '🍌', 0),
            ('Leche entera La Vaquita 1L', 'lacteos', 'La Vaquita', 'litro', 1350, None, 210, '🥛', 0),
            ('Yerba Mate Campo Verde 1kg', 'almacen', 'Campo Verde', 'kg', 4780, 5600, 65, '🧉', 1),
            ('Pan francés x 4 unidades', 'panificados', 'Panadería propia', 'unidad', 1200, None, 40, '🥖', 0),
            ('Milanesa de nalga x kg', 'carniceria', 'Carnicería', 'kg', 8990, None, 35, '🥩', 0),
            ('Agua saborizada 1.5L', 'bebidas', 'Cielo', 'litro', 1650, None, 300, '🥤', 0),
            ('Detergente 750ml', 'limpieza', 'Magistral', 'unidad', 2340, 2890, 150, '🧴', 0),
            ('Queso cremoso x kg', 'frescos', 'Frescos', 'kg', 7450, None, 55, '🧀', 0),
            ('Fideos guiseros 500g', 'almacen', 'Matarazzo', 'unidad', 980, None, 200, '🍝', 0),
            ('Vino Malbec 750ml', 'bebidas', 'Colón', 'unidad', 5200, 6100, 90, '🍷', 1),
            ('Yogur bebible 1L', 'lacteos', 'Serenísima', 'litro', 2100, None, 130, '🥤', 0),
        ]
        product_ids = {}
        for name, cat, brand, unidad, price, old_price, stock, icon, destacado in productos:
            pid = uuid.uuid4().hex
            product_ids[name] = pid
            db.execute("""
                INSERT INTO products (id, code, name, category_id, brand, unidad, price, old_price,
                                       stock, icon, destacado, active)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,1)
            """, (pid, pid[:8].upper(), name, cat, brand, unidad, price, old_price, stock, icon, destacado))

        db.execute("""
            INSERT OR IGNORE INTO usuarios (nombre, apellido, usuario, password, email, permisos, es_dueno, activo)
            VALUES (?,?,?,?,?,?,1,1)
        """, ('Administrador', 'Principal', ADMIN_USERNAME, generate_password_hash(ADMIN_PASSWORD),
              'admin@almacendelvalle.com.ar', ','.join(PERMISOS_DISPONIBLES.keys())))

        # Ofertas de la semana de ejemplo (banner gigante en la tienda)
        ofertas = [
            ('Semana del ahorro', 'Hasta 30% off en frescos y almacén, toda la semana.', None, None, 'verde', 1),
            ('2x1 en yerba Campo Verde', 'Llevando 2 paquetes de un 1kg, pagás uno solo.', product_ids['Yerba Mate Campo Verde 1kg'], None, 'mango', 1),
            ('Envío gratis en 25 de Mayo', 'En compras desde $50.000, sin cargo a tu domicilio.', None, None, 'ciruela', 1),
        ]
        for i, (titulo, texto, prod_id, cat_id, color, activa) in enumerate(ofertas):
            db.execute("""
                INSERT INTO ofertas_semana (titulo, texto, producto_id, categoria_id, color, activa, orden)
                VALUES (?,?,?,?,?,?,?)
            """, (titulo, texto, prod_id, cat_id, color, activa, i))

        # Datos del negocio de ejemplo
        db.execute("INSERT INTO horarios_atencion (dias, hora_inicio, hora_fin, orden) VALUES (?,?,?,0)",
                    ('Lunes a Sábado', '07:30', '20:30'))
        db.execute("""
            INSERT INTO direcciones (nombre, direccion, ciudad, provincia, orden)
            VALUES (?,?,?,?,0)
        """, ('Casa Central', 'Colectora Apóstoles Norte 951, Ruta Nac. 14', '25 de Mayo', 'Misiones'))
        db.executemany(
            "INSERT INTO contactos (tipo, etiqueta, valor, orden) VALUES (?,?,?,?)",
            [('Teléfono', 'Ventas', '+54 9 3755 650860', 0),
             ('Email', '', 'hola@almacendelvalle.com.ar', 1)])
        db.executemany(
            "INSERT INTO redes_sociales (plataforma, url, orden) VALUES (?,?,?)",
            [('Facebook', 'https://facebook.com', 0),
             ('Instagram', 'https://instagram.com', 1),
             ('WhatsApp', 'https://wa.me/5493755650860', 2)])

        # Configuración de envío por defecto
        db.executemany(
            "INSERT OR IGNORE INTO configuracion (clave, valor) VALUES (?,?)",
            [('envio_gratis_desde', '50000'), ('costo_envio', '2500'), ('compra_minima', '0'),
             ('smtp_activo', '0'), ('smtp_host', ''), ('smtp_puerto', '587'),
             ('smtp_usuario', ''), ('smtp_password', ''), ('smtp_email_destino', '')])

        # Cupón de ejemplo
        db.execute("""
            INSERT INTO cupones (codigo, tipo, valor, monto_minimo, usos_maximos, activo)
            VALUES ('BIENVENIDO10', 'porcentaje', 10, 5000, 0, 1)
        """)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def money(v):
    """Convierte un valor de formulario a número; devuelve 0.0 si está vacío o no es válido."""
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def ar_money(value):
    """Formatea un número con separador de miles al estilo argentino: 50.000"""
    try:
        value = float(value or 0)
    except (TypeError, ValueError):
        value = 0
    return f"{value:,.0f}".replace(',', '.')


def ar_money2(value):
    """Formatea un importe con centavos al estilo argentino: 1.234,50"""
    try:
        value = float(value or 0)
    except (TypeError, ValueError):
        value = 0.0
    return f"{value:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')


app.jinja_env.filters['arm'] = ar_money
app.jinja_env.filters['arm2'] = ar_money2


def info_list_to_db(lst):
    """Une las líneas de información adicional en un solo texto, una por renglón, para guardarlo."""
    return '\n'.join(l.strip() for l in lst if l.strip())


def info_db_to_list(text):
    """Separa el texto guardado en la base en una lista de líneas no vacías."""
    return [l for l in (text or '').split('\n') if l.strip()]


def get_config(clave, default=''):
    """Lee un valor de la tabla de configuración; devuelve default si la clave no existe."""
    with get_db() as db:
        row = db.execute("SELECT valor FROM configuracion WHERE clave=?", (clave,)).fetchone()
    return row['valor'] if row else default


def get_config_float(clave, default=0.0):
    """Igual que get_config, pero convertido a número."""
    try:
        return float(get_config(clave, default))
    except (TypeError, ValueError):
        return default


def set_config(clave, valor):
    """Guarda (o actualiza) un valor en la tabla de configuración."""
    with get_db() as db:
        db.execute("INSERT INTO configuracion (clave, valor) VALUES (?,?) "
                   "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor", (clave, valor))


# ---------------------------------------------------------------------------
# Correo electrónico (SMTP)
# Un mismo servidor SMTP, configurado en "Datos del Negocio", se usa para
# avisarle al admin de un pedido nuevo y para enviarle la factura al cliente.
# Un fallo de correo nunca debe romper una compra: todo error se registra en el
# log y el flujo sigue.
# ---------------------------------------------------------------------------

def leer_config_smtp():
    """Devuelve la configuración SMTP como diccionario, o None si el envío de
    correos está desactivado o falta el servidor."""
    if get_config('smtp_activo') != '1':
        return None
    host = get_config('smtp_host', '').strip()
    if not host:
        return None
    try:
        puerto = int(get_config('smtp_puerto', '587') or 587)
    except ValueError:
        puerto = 587
    return {
        'host': host,
        'puerto': puerto,
        'usuario': get_config('smtp_usuario', '').strip(),
        'password': get_config('smtp_password', ''),
        'destino_admin': get_config('smtp_email_destino', '').strip(),
    }


def contexto_tls():
    """Contexto TLS para el correo. Usa los certificados de `certifi` si está instalado, porque
    en algunas instalaciones de Python (sobre todo en Mac) los del sistema no están disponibles."""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


def explicar_error_smtp(error, cfg):
    """Traduce el error del servidor de correo a algo que se pueda resolver."""
    if isinstance(error, smtplib.SMTPAuthenticationError):
        return ("El servidor rechazó el usuario o la contraseña. Con Gmail hay que usar una "
                "\"contraseña de aplicación\" (la contraseña normal no funciona) y tener activada la verificación en dos pasos.")
    if isinstance(error, smtplib.SMTPRecipientsRefused):
        return "El servidor rechazó la dirección de destino. Revisá que el correo esté bien escrito."
    if isinstance(error, smtplib.SMTPSenderRefused):
        return "El servidor no aceptó el remitente. El usuario SMTP debe ser una casilla de correo válida del servidor."
    if isinstance(error, ssl.SSLError):
        return (f"La conexión segura con {cfg['host']} falló ({error}). En Mac suele pasar porque Python no tiene "
                f"los certificados instalados: ejecutá `pip install --upgrade certifi` en el entorno de la tienda "
                f"y reiniciá la aplicación. Si estás en una red de empresa o con antivirus que revisa HTTPS, probá desde otra red.")
    if isinstance(error, socket.gaierror):
        return (f"No se encontró el servidor {cfg['host']}: revisá que esté bien escrito y que haya internet ({error}).")
    if isinstance(error, smtplib.SMTPServerDisconnected):
        return (f"El servidor {cfg['host']}:{cfg['puerto']} cortó la conexión apenas se abrió. Con Gmail eso casi siempre "
                f"es algo en el medio que interfiere: un antivirus que revisa el correo o las conexiones seguras, un "
                f"firewall, un proxy o el proveedor de internet. Probá desde otra red (datos del celular) o desactivá "
                f"un momento el análisis de correo del antivirus. Para ver en qué paso se corta, ejecutá "
                f"`python diagnostico_correo.py` en la carpeta de la tienda.")
    if isinstance(error, (smtplib.SMTPConnectError, ConnectionError, TimeoutError, socket.timeout, OSError)):
        return (f"No se pudo conectar con {cfg['host']}:{cfg['puerto']} ({type(error).__name__}: {error}). "
                f"Probá con el puerto 465 (SSL) en lugar de 587, o desde otra red: algunos proveedores de "
                f"internet, firewalls y antivirus bloquean el puerto 587.")
    return f"{type(error).__name__}: {error}"


def _enviar_por_smtp(cfg, msg):
    """Entrega el mensaje al servidor SMTP. Devuelve None si salió bien o el texto del error,
    indicando en qué paso se cortó (conexión, inicio de sesión o envío)."""
    usa_ssl = cfg['puerto'] == 465  # 465 = SSL directo; 587 y otros = STARTTLS
    paso = 'al conectar'
    try:
        contexto = contexto_tls()
        if usa_ssl:
            servidor = smtplib.SMTP_SSL(cfg['host'], cfg['puerto'], timeout=30, context=contexto)
        else:
            servidor = smtplib.SMTP(cfg['host'], cfg['puerto'], timeout=30)
        with servidor:
            if not usa_ssl:
                paso = 'al activar el cifrado (STARTTLS)'
                servidor.starttls(context=contexto)
            if cfg['usuario'] and cfg['password']:
                paso = 'al iniciar sesión con el usuario y la contraseña'
                clave = cfg['password']
                if cfg['host'].endswith('gmail.com'):
                    clave = clave.replace(' ', '').replace('\u00a0', '').strip()   # Google las muestra en grupos de 4
                servidor.login(cfg['usuario'].strip(), clave)
            paso = 'al enviar el mensaje'
            servidor.send_message(msg)
        return None
    except Exception as e:
        explicacion = f"[{paso}] " + explicar_error_smtp(e, cfg)
        app.logger.warning("No se pudo enviar el correo '%s' a %s: %s (%s)",
                           msg['Subject'], msg['To'], explicacion, e)
        return explicacion


def entregar_correo(msg, cfg, tipo='', pedido_id=None, esperar=False):
    """Entrega el correo y deja anotado en `correos_enviados` si salió o por qué falló.
    Con esperar=False corre en segundo plano (el cliente no espera al servidor de correo);
    con esperar=True devuelve el error (o None si salió bien), para mostrárselo al administrador."""
    def tarea():
        error = _enviar_por_smtp(cfg, msg)
        try:
            with get_db() as db:
                db.execute("INSERT INTO correos_enviados (tipo, pedido_id, destinatario, asunto, estado, detalle) "
                           "VALUES (?,?,?,?,?,?)",
                           (tipo, pedido_id, msg['To'], msg['Subject'], 'error' if error else 'enviado', error or ''))
        except Exception:
            app.logger.exception("No se pudo registrar el resultado del correo")
        return error
    if esperar:
        return tarea()
    threading.Thread(target=tarea, name='envio-email', daemon=True).start()
    return None


def enviar_email(msg, cfg, tipo='', pedido_id=None, esperar=False):
    """Completa remitente, fecha e identificador del mensaje y lo entrega. Devuelve False si no hay
    un remitente válido; con esperar=True devuelve en cambio (ok, error)."""
    remitente = cfg['usuario'] if '@' in cfg['usuario'] else cfg['destino_admin']
    if not remitente:
        app.logger.warning("Correo no enviado: falta un remitente (usuario SMTP o email del admin).")
        return (False, "Falta el usuario SMTP (tiene que ser un correo) o el email del administrador.") if esperar else False
    for cabecera in ('From', 'Date', 'Message-ID'):
        del msg[cabecera]
    msg['From'] = formataddr(('Almacén del Valle', remitente))
    msg['Date'] = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid(domain=remitente.split('@')[-1])
    error = entregar_correo(msg, cfg, tipo, pedido_id, esperar)
    return (error is None, error) if esperar else True


def enviar_notificacion_email(asunto, cuerpo):
    """Avisa por correo al administrador (por ejemplo, de un pedido nuevo).
    No hace nada si no hay SMTP configurado o falta el email de destino."""
    cfg = leer_config_smtp()
    if not cfg or not cfg['destino_admin']:
        return
    msg = EmailMessage()
    msg['Subject'] = asunto
    msg['To'] = cfg['destino_admin']
    msg.set_content(cuerpo)
    enviar_email(msg, cfg)


def registrar_auditoria(accion, detalle='', db=None):
    """Anota quién hizo qué en el registro de auditoría. Con db reutiliza esa conexión."""
    usuario = session.get('admin_nombre') or session.get('cliente_nombre') or 'Sistema'
    if db is not None:
        db.execute("INSERT INTO auditoria (usuario, accion, detalle) VALUES (?,?,?)",
                   (usuario, accion, detalle))
    else:
        with get_db() as conn:
            conn.execute("INSERT INTO auditoria (usuario, accion, detalle) VALUES (?,?,?)",
                         (usuario, accion, detalle))


app.jinja_env.filters['infolist'] = info_db_to_list


# (la numeración de las facturas está en la sección FACTURACIÓN ELECTRÓNICA, más abajo)


def generar_referencia_pago(prefijo):
    """Código de referencia simulado para un pago: el prefijo más 10 dígitos al azar."""
    sufijo = ''.join(random.choices(string.digits, k=10))
    return f"{prefijo}-{sufijo}"


def get_cart():
    """El carrito vive en la sesión del servidor: [{product_id, qty}]"""
    return session.setdefault('cart', [])


def guardar_carrito_bd(cliente_id):
    """Guarda el carrito actual en la base para que persista entre sesiones/dispositivos."""
    cart = get_cart()
    with get_db() as db:
        db.execute("DELETE FROM carrito_guardado WHERE cliente_id=?", (cliente_id,))
        for entry in cart:
            db.execute("INSERT INTO carrito_guardado (cliente_id, product_id, qty) VALUES (?,?,?)",
                       (cliente_id, entry['product_id'], entry['qty']))


def cargar_carrito_bd(cliente_id):
    """Al iniciar sesión, fusiona el carrito guardado del cliente con el de la sesión actual."""
    with get_db() as db:
        rows = db.execute("SELECT product_id, qty FROM carrito_guardado WHERE cliente_id=?",
                           (cliente_id,)).fetchall()
    if not rows:
        return
    cart = get_cart()
    existing_ids = {e['product_id'] for e in cart}
    for r in rows:
        if r['product_id'] not in existing_ids:
            cart.append({'product_id': r['product_id'], 'qty': r['qty']})
    session.modified = True


def sincronizar_carrito_si_logueado():
    """Si hay un cliente con sesión iniciada, guarda su carrito en la base."""
    if session.get('cliente_id'):
        guardar_carrito_bd(session['cliente_id'])


def leer_cantidad(datos, defecto=1):
    """Lee la cantidad del JSON recibido; si no es un número entero devuelve el valor por defecto."""
    try:
        return int(datos.get('qty', defecto))
    except (TypeError, ValueError):
        return defecto


def respuesta_carrito():
    """JSON con el estado actual del carrito, que usa el panel lateral de la tienda."""
    items, subtotal = cart_details()
    return jsonify({'ok': True, 'items': items, 'subtotal': subtotal,
                    'count': sum(i['qty'] for i in items)})


def colores_categorias(db):
    """Diccionario {id de categoría: color}, para pintar las tarjetas de producto."""
    return {c['id']: c['color'] for c in db.execute("SELECT id, color FROM categories").fetchall()}


_cache_version = {'hasta': 0, 'valor': 0}


def version_de_estaticos():
    """Número que cambia cuando cambia cualquier archivo de static/: se agrega a los enlaces de
    CSS y JS para que el navegador nunca use una copia vieja guardada de una versión anterior.
    Se recalcula como mucho cada 5 segundos."""
    if time.time() < _cache_version['hasta']:
        return _cache_version['valor']
    ultimo = 0
    for carpeta, _, archivos in os.walk(os.path.join(app.root_path, 'static')):
        for nombre in archivos:
            if nombre.endswith(('.js', '.css')):
                try:
                    ultimo = max(ultimo, int(os.path.getmtime(os.path.join(carpeta, nombre))))
                except OSError:
                    pass
    _cache_version.update(hasta=time.time() + 5, valor=ultimo)
    return ultimo


app.jinja_env.globals['version_estaticos'] = version_de_estaticos


def numero_seguro(valor, defecto=0):
    """Convierte a número un valor de la base que puede venir vacío o como texto."""
    try:
        return float(valor) if '.' in str(valor) else int(valor)
    except (TypeError, ValueError):
        return defecto


def cart_details():
    """Resuelve el carrito de la sesión contra la base de datos actual."""
    cart = get_cart()
    if not cart:
        return [], 0.0
    with get_db() as db:
        items = []
        subtotal = 0.0
        for entry in cart:
            row = db.execute("SELECT * FROM products WHERE id=?", (entry['product_id'],)).fetchone()
            if not row:
                continue
            stock = numero_seguro(row['stock'])   # un stock vacío o cargado como texto no debe romper el carrito
            qty = min(entry['qty'], stock) if stock > 0 else 0
            if qty <= 0:
                continue
            line_total = row['price'] * qty
            subtotal += line_total
            items.append({
                'product_id': row['id'],
                'name': row['name'],
                'icon': row['icon'],
                'image': row['image'],
                'price': row['price'],
                'qty': qty,
                'stock': stock,
                'line_total': line_total,
            })
    return items, subtotal


# ---------------------------------------------------------------------------
# Context processor: variables disponibles en TODAS las plantillas de la tienda
# (el mega-menú y el footer usan 'categories' y el contador del carrito)
# ---------------------------------------------------------------------------

@app.context_processor
def inject_shop_globals():
    """Categorías, ofertas, datos del negocio y contador del carrito para todas las plantillas."""
    with get_db() as db:
        categories = [dict(r) for r in db.execute("SELECT * FROM categories ORDER BY name").fetchall()]
        ofertas_activas = [dict(r) for r in db.execute("""
            SELECT o.*, p.name AS producto_nombre, p.icon AS producto_icon, p.price AS producto_precio,
                   c.name AS categoria_nombre
            FROM ofertas_semana o
            LEFT JOIN products p ON p.id = o.producto_id
            LEFT JOIN categories c ON c.id = o.categoria_id
            WHERE o.activa = 1 ORDER BY o.orden
        """).fetchall()]
        horarios = [dict(r) for r in db.execute("SELECT * FROM horarios_atencion ORDER BY orden").fetchall()]
        direcciones = [dict(r) for r in db.execute("SELECT * FROM direcciones ORDER BY orden").fetchall()]
        contactos = [dict(r) for r in db.execute("SELECT * FROM contactos ORDER BY orden").fetchall()]
        redes = [dict(r) for r in db.execute("SELECT * FROM redes_sociales ORDER BY orden").fetchall()]
    items, _ = cart_details()
    return {
        'categories': categories,
        'cart_count': sum(i['qty'] for i in items),
        'ofertas_activas': ofertas_activas,
        'horarios_negocio': horarios,
        'direcciones_negocio': direcciones,
        'contactos_negocio': contactos,
        'redes_negocio': redes,
    }


# ---------------------------------------------------------------------------
# TIENDA — páginas públicas
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    """Portada de la tienda con todos los productos activos."""
    with get_db() as db:
        categories = [dict(r) for r in db.execute("SELECT * FROM categories ORDER BY name").fetchall()]
        products = [dict(r) for r in db.execute(
            "SELECT * FROM products WHERE active=1 ORDER BY created_at DESC").fetchall()]
    cat_colors = {c['id']: c['color'] for c in categories}
    return render_template('index.html', categories=categories, products=products, cat_colors=cat_colors)


@app.route('/producto/<product_id>')
def producto_detalle(product_id):
    """Ficha de un producto con sus reseñas, si es favorito y productos relacionados."""
    with get_db() as db:
        p = db.execute("""
            SELECT p.*, c.name AS category_name, c.color AS category_color
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            WHERE p.id=?
        """, (product_id,)).fetchone()
        if not p:
            return render_template('404.html'), 404
        p = dict(p)

        resenas = [dict(r) for r in db.execute(
            "SELECT * FROM resenas WHERE product_id=? ORDER BY created_at DESC", (product_id,)).fetchall()]

        es_favorito = False
        if session.get('cliente_id'):
            fav = db.execute("SELECT id FROM favoritos WHERE cliente_id=? AND product_id=?",
                              (session['cliente_id'], product_id)).fetchone()
            es_favorito = fav is not None

        # "También compraron": productos que aparecieron en los mismos pedidos
        tambien_compraron = [dict(r) for r in db.execute("""
            SELECT p2.*, COUNT(*) AS veces
            FROM pedido_items pi1
            JOIN pedido_items pi2 ON pi1.pedido_id = pi2.pedido_id AND pi2.product_id != pi1.product_id
            JOIN products p2 ON p2.id = pi2.product_id
            WHERE pi1.product_id=? AND p2.active=1
            GROUP BY p2.id ORDER BY veces DESC LIMIT 4
        """, (product_id,)).fetchall()]

        if len(tambien_compraron) < 4:
            faltan = 4 - len(tambien_compraron)
            ya_ids = [t['id'] for t in tambien_compraron] + [product_id]
            ph = ','.join('?' * len(ya_ids))
            extra = db.execute(f"""
                SELECT * FROM products WHERE category_id=? AND active=1 AND id NOT IN ({ph})
                ORDER BY created_at DESC LIMIT ?
            """, [p['category_id'], *ya_ids, faltan]).fetchall()
            tambien_compraron += [dict(r) for r in extra]

    promedio = round(sum(r['estrellas'] for r in resenas) / len(resenas), 1) if resenas else 0
    with get_db() as db:
        cat_colors = colores_categorias(db)
    return render_template('producto.html', p=p, resenas=resenas, promedio=promedio,
                            es_favorito=es_favorito, tambien_compraron=tambien_compraron,
                            cat_colors=cat_colors)


@app.route('/producto/<product_id>/resena', methods=['POST'])
@cliente_required(lambda: redirect(url_for('login')))
def producto_resena(product_id):
    """Guarda la reseña (de 1 a 5 estrellas) de un cliente con sesión iniciada."""
    estrellas = int(request.form.get('estrellas', 5))
    comentario = request.form.get('comentario', '').strip()
    with get_db() as db:
        db.execute("""
            INSERT INTO resenas (product_id, cliente_id, cliente_nombre, estrellas, comentario)
            VALUES (?,?,?,?,?)
        """, (product_id, session['cliente_id'], session.get('cliente_nombre', 'Cliente'),
              max(1, min(5, estrellas)), comentario))
    return redirect(url_for('producto_detalle', product_id=product_id))


@app.route('/api/favoritos/toggle', methods=['POST'])
@cliente_required(lambda: (jsonify({'error': 'Iniciá sesión primero'}), 401))
def api_favoritos_toggle():
    """Agrega o quita un producto de los favoritos del cliente."""
    d = request.get_json() or {}
    product_id = d.get('product_id')
    with get_db() as db:
        existe = db.execute("SELECT id FROM favoritos WHERE cliente_id=? AND product_id=?",
                             (session['cliente_id'], product_id)).fetchone()
        if existe:
            db.execute("DELETE FROM favoritos WHERE id=?", (existe['id'],))
            es_favorito = False
        else:
            db.execute("INSERT INTO favoritos (cliente_id, product_id) VALUES (?,?)",
                       (session['cliente_id'], product_id))
            es_favorito = True
    return jsonify({'ok': True, 'es_favorito': es_favorito})


@app.route('/favoritos')
@cliente_required(lambda: redirect(url_for('login', next=url_for('favoritos'))))
def favoritos():
    """Lista de productos favoritos del cliente."""
    with get_db() as db:
        productos = [dict(r) for r in db.execute("""
            SELECT p.* FROM favoritos f JOIN products p ON p.id = f.product_id
            WHERE f.cliente_id=? ORDER BY f.created_at DESC
        """, (session['cliente_id'],)).fetchall()]
        cat_colors = colores_categorias(db)
    return render_template('favoritos.html', productos=productos, cat_colors=cat_colors)


@app.route('/api/productos/buscar')
def api_productos_buscar():
    """Sugerencias del buscador: hasta 6 productos cuyo nombre o marca contiene el texto."""
    q = request.args.get('q', '').strip().lower()
    if len(q) < 2:
        return jsonify([])
    with get_db() as db:
        rows = db.execute("""
            SELECT id, name, brand, icon, image, price FROM products
            WHERE active=1 AND (LOWER(name) LIKE ? OR LOWER(brand) LIKE ?)
            ORDER BY name LIMIT 6
        """, (f'%{q}%', f'%{q}%')).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/carrito')
def api_carrito():
    """Contenido actual del carrito."""
    return respuesta_carrito()


@app.route('/api/carrito/agregar', methods=['POST'])
def api_carrito_agregar():
    """Suma un producto al carrito."""
    d = request.get_json() or {}
    product_id = d.get('product_id')
    qty = max(1, leer_cantidad(d))
    with get_db() as db:
        row = db.execute("SELECT id, name, stock FROM products WHERE id=? AND active=1", (product_id,)).fetchone()
        if not row:
            return jsonify({'error': 'Producto no encontrado'}), 404
        if numero_seguro(row['stock']) <= 0:
            return jsonify({'error': f"{row['name']} está sin stock por el momento."}), 409
    cart = get_cart()
    for entry in cart:
        if entry['product_id'] == product_id:
            entry['qty'] += qty
            break
    else:
        cart.append({'product_id': product_id, 'qty': qty})
    session.modified = True
    sincronizar_carrito_si_logueado()
    return respuesta_carrito()


@app.route('/api/carrito/actualizar', methods=['POST'])
def api_carrito_actualizar():
    """Cambia la cantidad de un producto del carrito; con cantidad 0 lo quita."""
    d = request.get_json() or {}
    product_id = d.get('product_id')
    qty = leer_cantidad(d)
    cart = get_cart()
    if qty <= 0:
        cart[:] = [e for e in cart if e['product_id'] != product_id]
    else:
        for entry in cart:
            if entry['product_id'] == product_id:
                entry['qty'] = qty
                break
    session.modified = True
    sincronizar_carrito_si_logueado()
    return respuesta_carrito()


@app.route('/api/carrito/vaciar', methods=['POST'])
def api_carrito_vaciar():
    """Vacía el carrito de la sesión."""
    session['cart'] = []
    session.modified = True
    return jsonify({'ok': True})


# ---------------------------------------------------------------------------
# CLIENTES — registro / login / logout
# ---------------------------------------------------------------------------

def normalizar_dni(texto):
    """Devuelve el DNI solo con dígitos (se quitan puntos, guiones y espacios),
    o una cadena vacía si no tiene 7 u 8 dígitos como un DNI argentino."""
    digitos = re.sub(r'[.\-\s]', '', texto or '')
    return digitos if digitos.isdigit() and 7 <= len(digitos) <= 8 else ''


def buscar_cliente(db, identificador):
    """Busca un cliente por correo (si lo escrito tiene arroba) o por DNI."""
    identificador = (identificador or '').strip()
    if '@' in identificador:
        return db.execute("SELECT * FROM clientes WHERE email=?",
                          (identificador.lower(),)).fetchone()
    dni = normalizar_dni(identificador)
    if not dni:
        return None
    return db.execute("SELECT * FROM clientes WHERE dni_cuit=?", (dni,)).fetchone()


def guardar_dni_cliente(db, cliente_id, dni_texto):
    """Guarda el DNI de un cliente que todavía no lo tenía cargado.
    Devuelve un mensaje de error, o None si se guardó correctamente."""
    dni = normalizar_dni(dni_texto)
    if not dni:
        return 'El DNI debe tener 7 u 8 números.'
    en_uso = db.execute("SELECT id FROM clientes WHERE dni_cuit=? AND id<>?",
                        (dni, cliente_id)).fetchone()
    if en_uso:
        return 'Ese DNI ya está registrado en otra cuenta.'
    db.execute("UPDATE clientes SET dni_cuit=? WHERE id=?", (dni, cliente_id))
    return None


def iniciar_sesion_cliente(cliente_id, nombre):
    """Deja al cliente logueado y recupera su carrito guardado."""
    session['cliente_id'] = cliente_id
    session['cliente_nombre'] = nombre
    cargar_carrito_bd(cliente_id)


def destino_seguro(url):
    """Acepta solo rutas internas del sitio ("/checkout"). Evita que el parámetro
    ?next= sirva para redirigir al cliente a una página externa."""
    if url and url.startswith('/') and not url.startswith('//'):
        return url
    return url_for('index')


@app.route('/registro', methods=['GET', 'POST'])
def registro():
    """Alta de una cuenta de cliente con nombre, correo, DNI y contraseña."""
    if request.method == 'POST':
        nombre   = request.form.get('nombre', '').strip()
        apellido = request.form.get('apellido', '').strip()
        email    = request.form.get('email', '').strip().lower()
        dni      = normalizar_dni(request.form.get('dni', ''))
        password = request.form.get('password', '')
        telefono = request.form.get('telefono', '').strip()

        if not (nombre and apellido and email and password):
            flash('Completá todos los campos obligatorios.', 'error')
            return render_template('registro.html')
        if not dni:
            flash('Ingresá un DNI válido, de 7 u 8 números.', 'error')
            return render_template('registro.html')
        if len(password) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'error')
            return render_template('registro.html')

        with get_db() as db:
            if db.execute("SELECT id FROM clientes WHERE email=?", (email,)).fetchone():
                flash('Ya existe una cuenta con ese correo.', 'error')
                return render_template('registro.html')
            if db.execute("SELECT id FROM clientes WHERE dni_cuit=?", (dni,)).fetchone():
                flash('Ya existe una cuenta con ese DNI.', 'error')
                return render_template('registro.html')

            cursor = db.execute("""
                INSERT INTO clientes (nombre, apellido, email, dni_cuit, password_hash, telefono)
                VALUES (?,?,?,?,?,?)
            """, (nombre, apellido, email, dni, generate_password_hash(password), telefono))
            cliente_id = cursor.lastrowid

        # Con el correo configurado la cuenta queda pendiente hasta que la persona abra el enlace
        # que le llega; sin correo configurado no hay forma de confirmar, así que queda activa.
        if leer_config_smtp():
            with get_db() as db:
                db.execute("UPDATE clientes SET email_confirmado=0 WHERE id=?", (cliente_id,))
            if enviar_confirmacion(cliente_id, email, nombre):
                flash(f'¡Casi listo! Te enviamos un correo a {email}. Abrí el enlace que trae para activar tu cuenta '
                      f'(revisá también el spam).', 'success')
                return redirect(url_for('login'))
            with get_db() as db:   # no se pudo preparar el correo: no se deja la cuenta trabada
                db.execute("UPDATE clientes SET email_confirmado=1 WHERE id=?", (cliente_id,))
        iniciar_sesion_cliente(cliente_id, nombre)
        return redirect(url_for('index'))

    return render_template('registro.html')


# ---------------------------------------------------------------------------
# ACCESO — bloqueo por intentos fallidos y recuperación de contraseña
# Tres intentos fallidos bloquean esa cuenta durante tres minutos. Los intentos se cuentan
# por cuenta (correo, DNI o usuario) y, con un tope más alto, por dirección IP, para frenar
# también a quien prueba muchas cuentas distintas. Se cuentan igual las cuentas que no
# existen, así la respuesta no revela qué correos o DNI están registrados.
# ---------------------------------------------------------------------------

MAX_INTENTOS_LOGIN = 3
SEGUNDOS_BLOQUEO = 3 * 60
MAX_FALLOS_POR_IP = 15
SEGUNDOS_OLVIDO_FALLOS = 15 * 60   # pasado este tiempo sin fallar, el contador vuelve a cero
SEGUNDOS_VALIDEZ_ENLACE = 60 * 60  # el enlace de recuperación dura una hora


def clave_de_intentos(ambito, identificador):
    """Clave con la que se cuentan los intentos: el DNI siempre se normaliza, para que
    '12.345.678' y '12345678' sean la misma cuenta."""
    texto = (identificador or '').strip().lower()
    if '@' not in texto:
        texto = normalizar_dni(texto) or texto
    return f"{ambito}:{texto[:100]}"


def clave_de_ip(ambito):
    return f"{ambito}-ip:{request.remote_addr or 'desconocida'}"


def segundos_de_bloqueo(db, *claves):
    """Segundos que faltan para que termine el bloqueo más largo entre esas claves (0 si no hay)."""
    ahora = time.time()
    restante = 0.0
    for clave in claves:
        fila = db.execute("SELECT bloqueado_hasta FROM intentos_login WHERE clave=?", (clave,)).fetchone()
        if fila:
            restante = max(restante, fila['bloqueado_hasta'] - ahora)
    return int(restante) + 1 if restante > 0 else 0


def sumar_intento(db, clave, maximo=MAX_INTENTOS_LOGIN):
    """Cuenta un intento fallido. Al llegar al máximo bloquea la clave y devuelve 0;
    si no, devuelve cuántos intentos quedan."""
    ahora = time.time()
    fila = db.execute("SELECT * FROM intentos_login WHERE clave=?", (clave,)).fetchone()
    intentos = fila['intentos'] if fila and ahora - fila['ultimo'] < SEGUNDOS_OLVIDO_FALLOS else 0
    intentos += 1
    bloqueado_hasta = 0.0
    if intentos >= maximo:
        bloqueado_hasta = ahora + SEGUNDOS_BLOQUEO
        intentos = 0  # al terminar el bloqueo vuelve a tener todos sus intentos
    db.execute("""INSERT INTO intentos_login (clave, intentos, bloqueado_hasta, ultimo) VALUES (?,?,?,?)
                  ON CONFLICT(clave) DO UPDATE SET intentos=excluded.intentos,
                  bloqueado_hasta=excluded.bloqueado_hasta, ultimo=excluded.ultimo""",
               (clave, intentos, bloqueado_hasta, ahora))
    # Limpieza de registros viejos para que la tabla no crezca sin límite
    db.execute("DELETE FROM intentos_login WHERE ultimo < ? AND bloqueado_hasta < ?",
               (ahora - 86400, ahora))
    return 0 if bloqueado_hasta else maximo - intentos


def limpiar_intentos(db, *claves):
    """Borra el contador de las claves (después de un ingreso correcto o de cambiar la contraseña)."""
    for clave in claves:
        db.execute("DELETE FROM intentos_login WHERE clave=?", (clave,))


def texto_de_espera(segundos):
    minutos, resto = divmod(int(segundos), 60)
    return f"{minutos} min {resto:02d} s" if minutos else f"{resto} s"


def mensaje_de_bloqueo(segundos):
    return (f"Superaste los {MAX_INTENTOS_LOGIN} intentos permitidos. Por seguridad, esperá "
            f"{texto_de_espera(segundos)} antes de volver a intentar.")


def procesar_intento_de_ingreso(db, ambito, identificador, cuenta_id, es_valido, texto_fallo):
    """Lógica común del ingreso de clientes y del panel. Devuelve (ingreso_ok, mensaje, bloqueado).
    Los intentos se cuentan por cuenta (cuenta_id), así entrar con el correo o con el DNI
    no da intentos de más; si la cuenta no existe se cuentan por lo que se escribió.
    Con la cuenta bloqueada ni siquiera se mira la contraseña."""
    clave_cuenta = f"{ambito}:id:{cuenta_id}" if cuenta_id else clave_de_intentos(ambito, identificador)
    clave_ip = clave_de_ip(ambito)
    espera = segundos_de_bloqueo(db, clave_cuenta, clave_ip)
    if espera:
        return False, mensaje_de_bloqueo(espera), True
    if es_valido():
        limpiar_intentos(db, clave_cuenta)
        return True, '', False
    restantes_cuenta = sumar_intento(db, clave_cuenta)
    restantes_ip = sumar_intento(db, clave_ip, MAX_FALLOS_POR_IP)
    if restantes_cuenta == 0 or restantes_ip == 0:
        return False, mensaje_de_bloqueo(segundos_de_bloqueo(db, clave_cuenta, clave_ip)), True
    palabra = 'intento' if restantes_cuenta == 1 else 'intentos'
    verbo = 'queda' if restantes_cuenta == 1 else 'quedan'
    return False, f"{texto_fallo} Te {verbo} {restantes_cuenta} {palabra}.", False


def url_publica():
    """Dirección base del sitio para armar enlaces en los correos. En producción conviene
    definir SITE_URL (por ejemplo https://almacendelvalle.com.ar) en vez de depender del
    encabezado Host que envía el navegador."""
    return (os.environ.get('SITE_URL', '').strip() or request.url_root).rstrip('/')


def hash_de_token(token):
    return hashlib.sha256(token.encode('utf-8')).hexdigest()


def crear_token_de_recuperacion(db, ambito, usuario_id, validez=None):
    """Genera un enlace de un solo uso. Anula los anteriores de la misma cuenta y guarda
    solo el hash del token: quien lea la base no puede usarlo."""
    token = secrets.token_urlsafe(32)
    db.execute("UPDATE recuperacion_clave SET usado=1 WHERE ambito=? AND usuario_id=? AND usado=0",
               (ambito, usuario_id))
    db.execute("INSERT INTO recuperacion_clave (ambito, usuario_id, token_hash, vence) VALUES (?,?,?,?)",
               (ambito, usuario_id, hash_de_token(token), time.time() + (validez or SEGUNDOS_VALIDEZ_ENLACE)))
    return token


def buscar_token_de_recuperacion(db, ambito, token):
    """Devuelve el registro del token si existe, no se usó y no venció."""
    return db.execute("""SELECT * FROM recuperacion_clave
                         WHERE token_hash=? AND ambito=? AND usado=0 AND vence>?""",
                      (hash_de_token(token or ''), ambito, time.time())).fetchone()


def enviar_correo_simple(destinatario, asunto, cuerpo):
    """Manda un correo de texto con el SMTP del negocio. Devuelve False si no se pudo preparar."""
    cfg = leer_config_smtp()
    if not cfg:
        app.logger.warning("Correo '%s' no enviado: el correo (SMTP) no está configurado en Datos del Negocio.", asunto)
        return False
    msg = EmailMessage()
    msg['Subject'] = asunto
    msg['To'] = destinatario
    msg.set_content(cuerpo)
    return enviar_email(msg, cfg)


SEGUNDOS_VALIDEZ_CONFIRMACION = 24 * 60 * 60


def enviar_confirmacion(cliente_id, email, nombre):
    """Manda el enlace para activar la cuenta. Sirve también para comprobar que el correo
    ingresado es de quien se registró. Devuelve False si no se pudo preparar el envío."""
    with get_db() as db:
        token = crear_token_de_recuperacion(db, 'confirmar', cliente_id, SEGUNDOS_VALIDEZ_CONFIRMACION)
    enlace = f"{url_publica()}{url_for('confirmar_correo', token=token)}"
    return enviar_correo_simple(
        email, 'Activá tu cuenta - Almacén del Valle',
        f"Hola {nombre},\n\n"
        f"Gracias por registrarte en Almacén del Valle. Para activar tu cuenta entrá a este enlace "
        f"(vence en 24 horas):\n\n{enlace}\n\n"
        f"Después podés ingresar con este correo o con tu DNI. Si no creaste esta cuenta, ignorá este mensaje.\n\n"
        f"Almacén del Valle")


def enviar_enlace_de_recuperacion(destinatario, nombre, enlace):
    minutos = SEGUNDOS_VALIDEZ_ENLACE // 60
    enviar_correo_simple(
        destinatario, 'Recuperá tu contraseña - Almacén del Valle',
        f"Hola {nombre},\n\n"
        f"Recibimos un pedido para cambiar la contraseña de tu cuenta. Para elegir una nueva "
        f"entrá a este enlace (sirve una sola vez y vence en {minutos} minutos):\n\n{enlace}\n\n"
        f"Si no fuiste vos, ignorá este correo: tu contraseña actual sigue funcionando.\n\n"
        f"Almacén del Valle")


def avisar_cambio_de_clave(destinatario, nombre):
    enviar_correo_simple(
        destinatario, 'Cambiaste tu contraseña - Almacén del Valle',
        f"Hola {nombre},\n\nTe avisamos que la contraseña de tu cuenta se cambió hace un momento. "
        f"Si no fuiste vos, pedí otro cambio de contraseña y escribinos cuanto antes.\n\nAlmacén del Valle")


MENSAJE_ENLACE_ENVIADO = ("Si el correo corresponde a una cuenta, te enviamos un enlace para cambiar la contraseña. "
                          "Revisá también la carpeta de spam. Si pediste varios enlaces seguidos, esperá unos minutos.")


def pedir_recuperacion(ambito, buscar_cuenta):
    """Pasos comunes del pedido de recuperación (clientes y panel). buscar_cuenta(db) devuelve
    la fila de la cuenta (con id, nombre y email) o None. Cada pedido cuenta como un intento:
    tras tres seguidos se pausa tres minutos, para que nadie use el formulario para llenar de
    correos la casilla de otra persona. La respuesta es siempre la misma."""
    enlace_para_enviar = None
    with get_db() as db:
        claves = (clave_de_ip(f'recuperar-{ambito}'),)
        if not segundos_de_bloqueo(db, *claves):
            sumar_intento(db, claves[0], 10)
            cuenta = buscar_cuenta(db)
            if cuenta and cuenta['email']:
                clave_cuenta = clave_de_intentos(f'recuperar-{ambito}', cuenta['email'])
                if not segundos_de_bloqueo(db, clave_cuenta):
                    sumar_intento(db, clave_cuenta)
                    token = crear_token_de_recuperacion(db, ambito, cuenta['id'])
                    destino = 'restablecer_clave' if ambito == 'cliente' else 'admin_restablecer_clave'
                    enlace_para_enviar = (cuenta['email'], cuenta['nombre'],
                                          url_publica() + url_for(destino, token=token))
    if enlace_para_enviar:
        enviar_enlace_de_recuperacion(*enlace_para_enviar)
    flash(MENSAJE_ENLACE_ENVIADO, 'success')


@app.route('/recuperar-clave', methods=['GET', 'POST'])
def recuperar_clave():
    """Pide el correo de la cuenta y le envía un enlace para elegir una contraseña nueva."""
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        pedir_recuperacion('cliente', lambda db: db.execute(
            "SELECT id, nombre, email FROM clientes WHERE email=?", (email,)).fetchone() if '@' in email else None)
        return redirect(url_for('recuperar_clave'))
    return render_template('recuperar_clave.html')


@app.route('/restablecer-clave/<token>', methods=['GET', 'POST'])
def restablecer_clave(token):
    """Pantalla a la que llega el cliente desde el correo para elegir su contraseña nueva."""
    with get_db() as db:
        registro_token = buscar_token_de_recuperacion(db, 'cliente', token)
    if not registro_token:
        return render_template('restablecer_clave.html', valido=False), 410

    if request.method == 'POST':
        nueva, repetida = request.form.get('password', ''), request.form.get('password2', '')
        if len(nueva) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'error')
        elif nueva != repetida:
            flash('Las dos contraseñas no coinciden.', 'error')
        else:
            with get_db() as db:
                cliente = db.execute("SELECT id, nombre, email, dni_cuit FROM clientes WHERE id=?",
                                     (registro_token['usuario_id'],)).fetchone()
                # Abrir el enlace enviado al correo también prueba que el correo es suyo
                db.execute("UPDATE clientes SET password_hash=?, email_confirmado=1 WHERE id=?",
                           (generate_password_hash(nueva), cliente['id']))
                db.execute("UPDATE recuperacion_clave SET usado=1 WHERE id=?", (registro_token['id'],))
                # Quien cambia su clave por correo queda habilitado para ingresar aunque estuviera bloqueado
                limpiar_intentos(db, f"cliente:id:{cliente['id']}")
            avisar_cambio_de_clave(cliente['email'], cliente['nombre'])
            flash('Listo, cambiaste tu contraseña. Ya podés ingresar.', 'success')
            return redirect(url_for('login'))
    return render_template('restablecer_clave.html', valido=True, token=token)


@app.route('/login', methods=['GET', 'POST'])
def login():
    """Ingreso de clientes con su correo o su DNI y la contraseña."""
    if request.method == 'POST':
        # El cliente puede ingresar con su correo o con su DNI en el mismo campo
        identificador = request.form.get('identificador', '')
        password      = request.form.get('password', '')
        with get_db() as db:
            cliente = buscar_cliente(db, identificador)
            ingreso_ok, mensaje, bloqueado = procesar_intento_de_ingreso(
                db, 'cliente', identificador, cliente['id'] if cliente else None,
                lambda: check_password_hash(cliente['password_hash'], password) if cliente else False,
                'Correo, DNI o contraseña incorrectos.')

        if ingreso_ok and not cliente['email_confirmado']:
            flash('Todavía no activaste tu cuenta. Abrí el enlace del correo que te enviamos.', 'error')
            return render_template('login.html', sin_confirmar=identificador)

        if ingreso_ok:
            iniciar_sesion_cliente(cliente['id'], cliente['nombre'])
            return redirect(destino_seguro(request.args.get('next')))

        # El mensaje es el mismo si falla el usuario o la contraseña, para no
        # revelar qué correos o DNI tienen cuenta
        flash(mensaje, 'error')
        return render_template('login.html'), 429 if bloqueado else 200

    return render_template('login.html')


@app.route('/confirmar-correo/<token>')
def confirmar_correo(token):
    """Destino del enlace del correo: activa la cuenta y deja al cliente adentro."""
    with get_db() as db:
        registro_token = buscar_token_de_recuperacion(db, 'confirmar', token)
        if not registro_token:
            flash('El enlace no es válido o ya venció. Ingresá con tu cuenta para pedir uno nuevo.', 'error')
            return redirect(url_for('login'))
        db.execute("UPDATE clientes SET email_confirmado=1 WHERE id=?", (registro_token['usuario_id'],))
        db.execute("UPDATE recuperacion_clave SET usado=1 WHERE id=?", (registro_token['id'],))
        cliente = db.execute("SELECT id, nombre FROM clientes WHERE id=?", (registro_token['usuario_id'],)).fetchone()
    iniciar_sesion_cliente(cliente['id'], cliente['nombre'])
    flash('¡Cuenta activada! Ya podés comprar.', 'success')
    return redirect(url_for('index'))


@app.route('/reenviar-confirmacion', methods=['POST'])
def reenviar_confirmacion():
    """Vuelve a mandar el enlace de activación. La respuesta es siempre la misma y se espera
    un minuto entre envíos, para que no sirva para llenar de correos la casilla de otra persona."""
    identificador = request.form.get('identificador', '')
    with get_db() as db:
        cliente = buscar_cliente(db, identificador)
        if cliente and not cliente['email_confirmado']:
            ultimo = db.execute("SELECT vence FROM recuperacion_clave WHERE ambito='confirmar' AND usuario_id=? "
                                "ORDER BY id DESC LIMIT 1", (cliente['id'],)).fetchone()
            if not ultimo or time.time() - (ultimo['vence'] - SEGUNDOS_VALIDEZ_CONFIRMACION) > 60:
                enviar_confirmacion(cliente['id'], cliente['email'], cliente['nombre'])
    flash('Si la cuenta está pendiente de activar, te enviamos el enlace de nuevo. Revisá también el spam.', 'success')
    return redirect(url_for('login'))


@app.route('/logout')
def logout():
    """Cierra la sesión del cliente."""
    session.pop('cliente_id', None)
    session.pop('cliente_nombre', None)
    return redirect(url_for('index'))


# ---------------------------------------------------------------------------
# CHECKOUT — requiere login de cliente
# ---------------------------------------------------------------------------

@app.route('/checkout')
@cliente_required(lambda: redirect(url_for('login', next=url_for('checkout'))))
def checkout():
    """Pantalla para finalizar la compra (requiere cliente con sesión iniciada)."""
    items, subtotal = cart_details()
    if not items:
        return redirect(url_for('index'))
    with get_db() as db:
        cliente = db.execute("SELECT * FROM clientes WHERE id=?", (session['cliente_id'],)).fetchone()
    envio_gratis_desde = get_config_float('envio_gratis_desde', 50000)
    costo_envio = get_config_float('costo_envio', 2500)
    envio = 0 if subtotal >= envio_gratis_desde else costo_envio
    total = subtotal + envio
    return render_template('checkout.html', items=items, subtotal=subtotal,
                            envio=envio, total=total, cliente=cliente,
                            envio_gratis_desde=envio_gratis_desde,
                            permite_factura_a=config_facturacion()['condicion_iva'] == 'responsable_inscripto')


def validar_cupon(codigo, subtotal):
    """Devuelve (cupon_row, descuento, error) según el estado del cupón."""
    if not codigo:
        return None, 0, None
    with get_db() as db:
        cupon = db.execute("SELECT * FROM cupones WHERE codigo=? AND activo=1",
                            (codigo.strip().upper(),)).fetchone()
    if not cupon:
        return None, 0, 'El cupón no existe o ya no está activo.'
    if cupon['fecha_expiracion'] and cupon['fecha_expiracion'] < datetime.now().strftime('%Y-%m-%d'):
        return None, 0, 'Este cupón ya venció.'
    if cupon['usos_maximos'] and cupon['usos_actuales'] >= cupon['usos_maximos']:
        return None, 0, 'Este cupón alcanzó el límite de usos.'
    if subtotal < cupon['monto_minimo']:
        return None, 0, f"Este cupón requiere una compra mínima de ${ar_money(cupon['monto_minimo'])}."
    descuento = subtotal * (cupon['valor'] / 100) if cupon['tipo'] == 'porcentaje' else cupon['valor']
    descuento = min(descuento, subtotal)
    return cupon, descuento, None


@app.route('/checkout/confirmar', methods=['POST'])
@cliente_required(lambda: redirect(url_for('login')))
def checkout_confirmar():
    """Cobra, registra el pedido, descuenta stock, emite la factura y envía los correos."""
    items, subtotal = cart_details()
    if not items:
        return redirect(url_for('index'))

    metodo_entrega = request.form.get('metodo_entrega', 'domicilio')
    direccion      = request.form.get('direccion', '').strip()
    metodo_pago    = request.form.get('metodo_pago', 'mercado_pago')
    cupon_codigo   = request.form.get('cupon_codigo', '').strip().upper()
    dni_ingresado  = request.form.get('dni', '').strip()

    # Factura A: solo si quien vende es Responsable Inscripto y el cliente da un CUIT válido
    factura_a_cuit = factura_a_razon = ''
    if request.form.get('factura_a') and config_facturacion()['condicion_iva'] == 'responsable_inscripto':
        factura_a_cuit = normalizar_cuit(request.form.get('factura_a_cuit', ''))
        factura_a_razon = request.form.get('factura_a_razon', '').strip()[:100]
        if not factura_a_cuit:
            flash('El CUIT para la Factura A no es válido. Revisá los 11 dígitos.', 'error')
            return redirect(url_for('checkout'))
        if not factura_a_razon:
            flash('Para la Factura A necesitamos la razón social.', 'error')
            return redirect(url_for('checkout'))

    # Los clientes sin DNI cargado pueden completarlo acá; queda en la factura
    # y les permite ingresar con él la próxima vez.
    if dni_ingresado:
        with get_db() as db:
            tiene_dni = db.execute("SELECT dni_cuit FROM clientes WHERE id=?",
                                   (session['cliente_id'],)).fetchone()['dni_cuit']
            error_dni = None if tiene_dni else guardar_dni_cliente(
                db, session['cliente_id'], dni_ingresado)
        if error_dni:
            flash(error_dni, 'error')
            return redirect(url_for('checkout'))

    cupon, descuento, error_cupon = validar_cupon(cupon_codigo, subtotal)
    if cupon_codigo and error_cupon:
        flash(error_cupon, 'error')
        return redirect(url_for('checkout'))

    envio, total = calcular_envio_y_total(subtotal, descuento, metodo_entrega)

    datos = {'metodo_entrega': metodo_entrega, 'direccion': direccion, 'metodo_pago': metodo_pago,
             'cupon_codigo': cupon_codigo, 'factura_a_cuit': factura_a_cuit, 'factura_a_razon': factura_a_razon}
    if metodo_pago not in NOMBRES_METODO_PAGO:
        flash('Elegí un medio de pago.', 'error')
        return redirect(url_for('checkout'))

    # Mercado Pago y tarjetas: primero se pasa por la pantalla de pago y el pedido se registra
    # recién cuando el cobro sale bien. Transferencia y efectivo registran el pedido como pendiente
    # y le muestran al cliente cómo pagar.
    if metodo_pago in METODOS_CON_PANTALLA_DE_COBRO:
        session['pago_pendiente'] = datos
        return redirect(url_for('checkout_pago'))

    pedido_id = registrar_pedido(items, subtotal, descuento, cupon, envio, total, datos, pago_diferido(metodo_pago))
    return redirect(url_for('pedido_confirmado', pedido_id=pedido_id))


def calcular_envio_y_total(subtotal, descuento, metodo_entrega):
    """Costo de envío (gratis desde cierto monto o si retira) y total a cobrar."""
    envio_gratis_desde = get_config_float('envio_gratis_desde', 50000)
    costo_envio = get_config_float('costo_envio', 2500)
    envio = 0 if subtotal >= envio_gratis_desde or metodo_entrega == 'retiro' else costo_envio
    return envio, max(subtotal - descuento, 0) + envio


def registrar_pedido(items, subtotal, descuento, cupon, envio, total, datos, pago_info):
    """Guarda el pedido con su pago, descuenta stock, emite la factura si ya está pago y manda los correos."""
    metodo_pago = datos['metodo_pago']
    with get_db() as db:
        db.execute("""
            INSERT INTO pedidos (cliente_id, subtotal, envio, total, metodo_entrega, direccion_envio,
                                  estado, cupon_codigo, descuento, factura_a_cuit, factura_a_razon)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, (session['cliente_id'], subtotal, envio, total, datos['metodo_entrega'], datos['direccion'],
              pago_info['estado_pedido'], cupon['codigo'] if cupon else '', descuento,
              datos['factura_a_cuit'], datos['factura_a_razon']))
        pedido_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        registrar_estado(db, pedido_id, 'pendiente_pago')
        if pago_info['estado_pedido'] != 'pendiente_pago':
            registrar_estado(db, pedido_id, pago_info['estado_pedido'])

        for it in items:
            db.execute("""
                INSERT INTO pedido_items (pedido_id, product_id, product_name, product_price, qty, subtotal)
                VALUES (?,?,?,?,?,?)
            """, (pedido_id, it['product_id'], it['name'], it['price'], it['qty'], it['line_total']))
            db.execute("UPDATE products SET stock = stock - ? WHERE id=?", (it['qty'], it['product_id']))

        referencia = pago_info['referencia']
        if metodo_pago == 'efectivo':
            referencia = codigo_de_pago_efectivo(pedido_id)   # lleva el número de pedido, por eso se arma acá
        db.execute("""
            INSERT INTO pagos (pedido_id, metodo, estado, monto, referencia, detalle)
            VALUES (?,?,?,?,?,?)
        """, (pedido_id, metodo_pago, pago_info['estado'], total, referencia, pago_info['detalle']))

        if cupon:
            db.execute("UPDATE cupones SET usos_actuales = usos_actuales + 1 WHERE id=?", (cupon['id'],))

        # Se emite la factura automáticamente si el pago quedó aprobado
        factura_emitida = False
        if pago_info['estado'] == 'aprobado':
            factura_emitida = emitir_factura(db, pedido_id)

        db.execute("DELETE FROM carrito_guardado WHERE cliente_id=?", (session['cliente_id'],))

    # Los correos se envían recién cuando la compra ya quedó guardada en la base
    if factura_emitida:
        enviar_factura_por_email(pedido_id)

    enviar_notificacion_email(
        f"Nuevo pedido #{pedido_id} - Almacén del Valle",
        f"Se registró un nuevo pedido.\n\n"
        f"Pedido: #{pedido_id}\n"
        f"Cliente: {session.get('cliente_nombre', '')}\n"
        f"Total: ${ar_money2(total)}\n"
        f"Método de pago: {NOMBRES_METODO_PAGO.get(metodo_pago, metodo_pago)}\n"
        f"Estado: {pago_info['estado_pedido']}\n"
    )

    session['cart'] = []
    session.pop('pago_pendiente', None)
    session.modified = True
    return pedido_id


@app.route('/checkout/pago', methods=['GET', 'POST'])
@cliente_required(lambda: redirect(url_for('login')))
def checkout_pago():
    """Pantalla de cobro de Mercado Pago y de tarjetas. Si el pago sale bien, registra el pedido."""
    datos = session.get('pago_pendiente')
    items, subtotal = cart_details()
    if not datos or not items:
        return redirect(url_for('checkout'))
    cupon, descuento, error_cupon = validar_cupon(datos['cupon_codigo'], subtotal)
    if datos['cupon_codigo'] and error_cupon:
        flash(error_cupon, 'error')
        return redirect(url_for('checkout'))
    envio, total = calcular_envio_y_total(subtotal, descuento, datos['metodo_entrega'])
    metodo = datos['metodo_pago']

    error = None
    if request.method == 'POST':
        pago_info = procesar_pago_simulado(metodo, total, request.form)
        if pago_info['aprobado']:
            pedido_id = registrar_pedido(items, subtotal, descuento, cupon, envio, total, datos, pago_info)
            return redirect(url_for('pedido_confirmado', pedido_id=pedido_id))
        error = pago_info['detalle']

    cuotas = [(n, total / n) for n in CUOTAS_DISPONIBLES] if metodo == 'tarjeta_credito' else []
    return render_template('checkout_pago.html', metodo=metodo, nombre_metodo=NOMBRES_METODO_PAGO[metodo],
                           items=items, subtotal=subtotal, descuento=descuento, envio=envio, total=total,
                           cuotas=cuotas, error=error, form=request.form)


@app.route('/api/cupon/validar', methods=['POST'])
@cliente_required(lambda: (jsonify({'error': 'Iniciá sesión primero'}), 401))
def api_validar_cupon():
    """Valida un cupón contra el carrito actual y devuelve el descuento que corresponde."""
    d = request.get_json() or {}
    _, subtotal = cart_details()
    cupon, descuento, error = validar_cupon(d.get('codigo', ''), subtotal)
    if error:
        return jsonify({'ok': False, 'error': error})
    return jsonify({'ok': True, 'descuento': descuento, 'codigo': cupon['codigo']})


# --- Medios de pago -------------------------------------------------------
# No hay credenciales de ninguna pasarela real: el cobro se simula, pero cada medio muestra lo
# mismo que vería el cliente de verdad (formulario de tarjeta completo, cupón de pago con código
# de barras, datos de la cuenta del negocio). Los datos de la tarjeta nunca se guardan: de ella
# solo quedan la marca y los últimos 4 dígitos.

METODOS_CON_PANTALLA_DE_COBRO = ('mercado_pago', 'tarjeta_credito', 'tarjeta_debito')
CUOTAS_DISPONIBLES = (1, 3, 6)
HORAS_VALIDEZ_CUPON = 72


def luhn_valido(numero):
    """Control matemático que cumplen todos los números de tarjeta reales."""
    total = 0
    for i, c in enumerate(reversed(numero)):
        n = int(c)
        if i % 2:
            n = n * 2 - 9 if n * 2 > 9 else n * 2
        total += n
    return total % 10 == 0


def marca_de_tarjeta(numero):
    if numero.startswith('4'):
        return 'Visa'
    if re.match(r'^(5[1-5]|2(2[2-9][1-9]|2[3-9]|[3-6]|7[01]|720))', numero):
        return 'Mastercard'
    if re.match(r'^3[47]', numero):
        return 'American Express'
    if re.match(r'^(50|5[6-9]|6)', numero):
        return 'Maestro'
    return ''


def validar_tarjeta(form):
    """Revisa los campos de la tarjeta como lo haría una pasarela. Devuelve (error, marca, ultimos4)."""
    numero = re.sub(r'\D', '', form.get('tarjeta_numero', ''))
    if not 13 <= len(numero) <= 19 or not luhn_valido(numero):
        return 'El número de tarjeta no es válido. Revisalo.', '', ''
    marca = marca_de_tarjeta(numero)
    if len(re.sub(r'\s', '', form.get('tarjeta_nombre', ''))) < 3:
        return 'Ingresá el nombre tal como figura en la tarjeta.', '', ''
    m = re.fullmatch(r'(\d{2})\s*/\s*(\d{2})', form.get('tarjeta_venc', '').strip())
    if not m or not 1 <= int(m.group(1)) <= 12:
        return 'La fecha de vencimiento no es válida. Usá el formato MM/AA.', '', ''
    hoy = datetime.now()
    if (2000 + int(m.group(2)), int(m.group(1))) < (hoy.year, hoy.month):
        return 'La tarjeta está vencida.', '', ''
    largo_cvv = 4 if marca == 'American Express' else 3
    if not re.fullmatch(r'\d{%d}' % largo_cvv, form.get('tarjeta_cvv', '').strip()):
        return f'El código de seguridad debe tener {largo_cvv} dígitos.', '', ''
    if not re.fullmatch(r'\d{7,8}', re.sub(r'\D', '', form.get('tarjeta_dni', ''))):
        return 'Ingresá el DNI del titular de la tarjeta.', '', ''
    return None, marca, numero[-4:]


def procesar_pago_simulado(metodo, total, form):
    """Simula la respuesta de la pasarela para Mercado Pago y tarjetas. Para probar un rechazo,
    usar una tarjeta cuyo número termine en 0002."""
    if metodo == 'mercado_pago':
        return {'aprobado': True, 'estado': 'aprobado', 'estado_pedido': 'pagado',
                'referencia': generar_referencia_pago('MP'), 'detalle': 'Pago aprobado con Mercado Pago (simulado).'}

    if metodo in ('tarjeta_credito', 'tarjeta_debito'):
        error, marca, ultimos = validar_tarjeta(form)
        if error:
            return {'aprobado': False, 'estado': 'rechazado', 'estado_pedido': 'pendiente_pago',
                    'referencia': '', 'detalle': error}
        if ultimos == '0002':
            return {'aprobado': False, 'estado': 'rechazado', 'estado_pedido': 'pendiente_pago',
                    'referencia': '', 'detalle': 'El banco rechazó el pago por fondos insuficientes. Probá con otra tarjeta.'}
        tipo = 'crédito' if metodo == 'tarjeta_credito' else 'débito'
        cuotas = 1
        if metodo == 'tarjeta_credito' and form.get('cuotas', '1').isdigit() and int(form['cuotas']) in CUOTAS_DISPONIBLES:
            cuotas = int(form['cuotas'])
        detalle = f"Pago aprobado con {marca or 'tarjeta'} de {tipo} terminada en {ultimos}"
        detalle += f", en {cuotas} cuotas de ${ar_money2(total / cuotas)}" if cuotas > 1 else ""
        return {'aprobado': True, 'estado': 'aprobado', 'estado_pedido': 'pagado',
                'referencia': generar_referencia_pago('CARD'), 'detalle': detalle + ' (simulado).'}

    return {'aprobado': False, 'estado': 'rechazado', 'estado_pedido': 'pendiente_pago',
            'referencia': '', 'detalle': 'Método de pago no reconocido.'}


def pago_diferido(metodo):
    """Transferencia y efectivo: el pedido queda pendiente hasta que el local confirme que cobró."""
    if metodo == 'transferencia':
        return {'aprobado': True, 'estado': 'pendiente', 'estado_pedido': 'pendiente_pago',
                'referencia': generar_referencia_pago('TRANSF'),
                'detalle': 'Esperando acreditación de la transferencia bancaria.'}
    return {'aprobado': True, 'estado': 'pendiente', 'estado_pedido': 'pendiente_pago',
            'referencia': '', 'detalle': 'Pendiente de pago en Rapipago o Pago Fácil con el cupón generado.'}


def codigo_de_pago_efectivo(pedido_id):
    """Código numérico de 20 dígitos del cupón de pago: el número de pedido más 12 dígitos al azar."""
    return f"{pedido_id:08d}" + ''.join(random.choices(string.digits, k=12))


DATOS_TRANSFERENCIA_DE_EJEMPLO = {'banco': 'Banco de ejemplo', 'titular': 'Almacén del Valle',
                                  'cuit': '', 'cbu': '0000003100012345678901', 'alias': 'ALMACEN.VALLE'}


def datos_transferencia():
    """Cuenta del negocio para recibir transferencias, cargada en el panel. Si todavía no se
    cargó, devuelve datos de ejemplo marcados como tales para que nadie transfiera a ciegas."""
    datos = {k: get_config('transf_' + k, '') for k in DATOS_TRANSFERENCIA_DE_EJEMPLO}
    datos['cargados'] = bool(datos['cbu'] or datos['alias'])
    if not datos['cargados']:
        datos = dict(DATOS_TRANSFERENCIA_DE_EJEMPLO, cargados=False)
    return datos


def svg_codigo_de_barras(valor, alto=70, modulo=1.6):
    """Código de barras Code 128 del cupón de pago, como SVG para mostrarlo en la página.
    Se usa la codificación de reportlab: mayúsculas son barras y minúsculas espacios, y la
    letra indica el grosor (A/a = 1 módulo, B/b = 2, ...)."""
    codigo = Code128(valor, quiet=False)
    codigo.validate()
    codigo.encode()
    codigo.decompose()
    x, barras = 0.0, []
    for letra in codigo.decomposed:
        ancho = (ord(letra.lower()) - ord('a') + 1) * modulo
        if letra.isupper():
            barras.append(f'<rect x="{x:.2f}" y="0" width="{ancho:.2f}" height="{alto}" fill="#000"/>')
        x += ancho
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {x:.2f} {alto}" width="{x:.0f}" height="{alto}" '
            f'preserveAspectRatio="none">{"".join(barras)}</svg>')


def vencimiento_cupon(pago):
    """Fecha hasta la que se puede pagar el cupón."""
    try:
        creado = datetime.strptime(pago['created_at'][:19], '%Y-%m-%d %H:%M:%S')
    except (TypeError, ValueError):
        creado = datetime.now()
    return creado + timedelta(hours=HORAS_VALIDEZ_CUPON)


def etiqueta_estado(estado):
    """Nombre legible de un estado de pedido ("pendiente_pago" -> "Pendiente de pago")."""
    datos = ESTADOS_PEDIDO.get(estado)
    return datos['etiqueta'] if datos else str(estado).replace('_', ' ').capitalize()


def fecha_ar(valor):
    """Fecha de la base a formato argentino: '2026-10-31 10:21:17' -> '31/10/2026 10:21'."""
    try:
        return datetime.strptime(valor[:16], '%Y-%m-%d %H:%M').strftime('%d/%m/%Y %H:%M')
    except (TypeError, ValueError):
        return valor or ''


def pasos_de(pedido):
    """Recorrido del pedido: con envío a domicilio o con retiro en sucursal."""
    return PASOS_RETIRO if pedido['metodo_entrega'] == 'retiro' else PASOS_DOMICILIO


def siguiente_paso(pedido):
    """Próximo estado del recorrido del pedido y el texto del botón del panel, o None si ya terminó.
    Un pedido que espera el pago ofrece "Confirmar pago" (transferencia o efectivo)."""
    pasos = pasos_de(pedido)
    if pedido['estado'] not in pasos or pedido['estado'] == pasos[-1]:
        return None
    destino = pasos[pasos.index(pedido['estado']) + 1]
    retiro = pedido['metodo_entrega'] == 'retiro'
    texto = {
        'pendiente_pago': 'Confirmar pago',
        'pagado': 'Preparar pedido' if retiro else 'Preparar envío',
        'preparando': 'Listo para retirar' if retiro else 'Marcar como enviado',
        'enviado': 'Marcar como entregado',
        'listo_retiro': 'Marcar como entregado',
    }[pedido['estado']]
    return destino, texto


def estados_disponibles(pedido):
    """Estados que el panel ofrece para un pedido: su recorrido más "cancelado"."""
    estados = pasos_de(pedido) + ['cancelado']
    if pedido['estado'] not in estados:
        estados.insert(0, pedido['estado'])
    return [(estado, etiqueta_estado(estado)) for estado in estados]


app.jinja_env.filters['estado_pedido'] = etiqueta_estado
app.jinja_env.filters['fecha_ar'] = fecha_ar
app.jinja_env.globals['estados_disponibles'] = estados_disponibles
app.jinja_env.globals['siguiente_paso'] = siguiente_paso
app.jinja_env.globals['datos_transferencia'] = datos_transferencia
app.jinja_env.globals['correo_configurado'] = lambda: leer_config_smtp() is not None


def registrar_estado(db, pedido_id, estado):
    """Anota en el historial que el pedido pasó a un estado, con la fecha y hora actuales."""
    db.execute("INSERT INTO pedido_estados (pedido_id, estado) VALUES (?,?)", (pedido_id, estado))


def linea_de_tiempo(db, pedido):
    """Arma el seguimiento de un pedido para mostrarlo al cliente.
    Cada paso indica si ya se cumplió ('done'), si es el actual ('current') o si falta ('pending'),
    junto con la fecha en que se alcanzó."""
    pasos = pasos_de(pedido)
    fechas = {r['estado']: r['fecha'] for r in db.execute(
        "SELECT estado, MIN(created_at) AS fecha FROM pedido_estados WHERE pedido_id=? GROUP BY estado",
        (pedido['id'],)).fetchall()}

    cancelado = pedido['estado'] == 'cancelado'
    if cancelado:
        # Se muestra hasta el último paso que llegó a cumplirse
        alcanzados = [i for i, paso in enumerate(pasos) if paso in fechas]
        indice_actual = max(alcanzados) if alcanzados else 0
    else:
        indice_actual = pasos.index(pedido['estado']) if pedido['estado'] in pasos else 0

    resultado = []
    for i, estado in enumerate(pasos):
        if i < indice_actual or cancelado and i == indice_actual:
            clase = 'done'
        elif i == indice_actual:
            clase = 'done' if estado == 'entregado' else 'current'  # el último paso ya es el final
        else:
            clase = 'pending'
        resultado.append({'estado': estado, 'titulo': ESTADOS_PEDIDO[estado]['titulo'],
                          'detalle': ESTADOS_PEDIDO[estado]['detalle'], 'clase': clase,
                          'fecha': fecha_ar(fechas.get(estado)) if clase != 'pending' else ''})
    return {'pasos': resultado, 'cancelado': cancelado,
            'cancelado_fecha': fecha_ar(fechas.get('cancelado'))}


def calcular_iva(total):
    """Separa un total con IVA incluido en (neto sin IVA, IVA).
    Los precios de la tienda ya incluyen IVA, como corresponde a un consumidor final."""
    neto = round(total / (1 + IVA_PORC), 2)
    return neto, round(total - neto, 2)


# ---------------------------------------------------------------------------
# FACTURACIÓN ELECTRÓNICA (estilo ARCA)
# Cada factura lleva letra (A, B o C), punto de venta, número correlativo, CAE y vencimiento del
# CAE. Por ahora solo funciona el modo DEMOSTRACIÓN: genera un CAE de mentira y marca el
# comprobante como "sin validez fiscal". La conexión real con ARCA (WSAA + WSFEv1) se enchufa
# en solicitar_cae(), que es lo único que habría que reemplazar; el resto del sistema ya trabaja
# con los mismos datos que devolvería ARCA.
# ---------------------------------------------------------------------------

MODOS_ARCA = {'demo': 'Demostración (sin validez fiscal)'}
CODIGOS_COMPROBANTE = {'A': 1, 'B': 6, 'C': 11}  # códigos de tipo de comprobante de ARCA
CONDICIONES_IVA = {'responsable_inscripto': 'IVA Responsable Inscripto', 'monotributista': 'Responsable Monotributo'}
DIAS_VALIDEZ_CAE_DEMO = 10


class ErrorArca(Exception):
    """No se pudo obtener la autorización del comprobante."""


def config_facturacion():
    """Datos del emisor y modo de facturación, tal como están en Datos del Negocio."""
    try:
        punto_venta = int(get_config('arca_punto_venta', '1'))
    except ValueError:
        punto_venta = 1
    return {
        'modo': get_config('arca_modo', 'demo') if get_config('arca_modo', 'demo') in MODOS_ARCA else 'demo',
        'razon_social': get_config('arca_razon_social', '') or 'Almacén del Valle',
        'cuit': get_config('arca_cuit', ''),
        'condicion_iva': get_config('arca_condicion_iva', 'responsable_inscripto'),
        'punto_venta': punto_venta,
        'domicilio': get_config('arca_domicilio', '') or 'Colectora Apóstoles Norte 951, 25 de Mayo, Misiones',
        'ingresos_brutos': get_config('arca_iibb', ''),
        'inicio_actividades': get_config('arca_inicio_actividades', ''),
    }


def normalizar_cuit(texto):
    """Devuelve el CUIT/CUIL como 11 dígitos si es válido (dígito verificador incluido), o ''."""
    digitos = re.sub(r'\D', '', texto or '')
    if len(digitos) != 11:
        return ''
    pesos = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    resto = sum(int(d) * p for d, p in zip(digitos[:10], pesos)) % 11
    verificador = {0: 0}.get(resto, 11 - resto)
    return digitos if verificador == int(digitos[10]) else ''


def formatear_cuit(digitos):
    return f"{digitos[:2]}-{digitos[2:10]}-{digitos[10:]}" if len(digitos) == 11 else digitos


def letra_de_comprobante(condicion_iva, pide_factura_a=False):
    """Letra de la factura según quién emite y a quién:
    Monotributo -> C; Responsable Inscripto -> A para otro Responsable Inscripto (con CUIT)
    y B para consumidores finales, monotributistas y exentos."""
    if condicion_iva == 'monotributista':
        return 'C'
    return 'A' if pide_factura_a else 'B'


def solicitar_cae(db, letra, punto_venta, modo):
    """Pide la autorización del comprobante. Devuelve (número, CAE, vencimiento del CAE).
    En modo demostración se simula; ARCA entregaría estos mismos tres datos."""
    if modo != 'demo':
        raise ErrorArca("La conexión real con ARCA todavía no está configurada.")
    numero = db.execute("SELECT COALESCE(MAX(nro_comprobante), 0) FROM facturas WHERE punto_venta=? AND tipo=?",
                        (punto_venta, letra)).fetchone()[0] + 1
    while db.execute("SELECT 1 FROM facturas WHERE numero=?", (formatear_numero(letra, punto_venta, numero),)).fetchone():
        numero += 1  # facturas anteriores a esta versión usaban otro esquema de numeración
    cae = ''.join(secrets.choice('0123456789') for _ in range(14))
    vencimiento = (datetime.now() + timedelta(days=DIAS_VALIDEZ_CAE_DEMO)).strftime('%Y-%m-%d')
    return numero, cae, vencimiento


def formatear_numero(letra, punto_venta, numero):
    return f"{letra}-{punto_venta:04d}-{numero:08d}"


UNIDADES_FACTURA = {'unidad': 'unidades', 'kg': 'kilogramos', 'g': 'gramos', 'l': 'litros', 'litro': 'litros',
                    'ml': 'mililitros', 'pack': 'packs', 'docena': 'docenas'}


def lineas_de_factura(pedido, items, db=None):
    """Renglones de la factura: productos, más el envío y el descuento si los hay, así los
    renglones suman exactamente el total. Los precios incluyen IVA; el neto de cada renglón
    (precio / 1,21) es lo que se muestra en la Factura A."""
    if db is None:
        with get_db() as conexion:
            return lineas_de_factura(pedido, items, conexion)
    lineas = []
    for it in items:
        producto = db.execute("SELECT code, unidad FROM products WHERE id=?", (it['product_id'],)).fetchone()
        unidad = ((producto['unidad'] if producto else '') or 'unidad').strip().lower()
        lineas.append({'codigo': (producto['code'] if producto else '') or '', 'descripcion': it['product_name'],
                       'cantidad': it['qty'], 'unidad': UNIDADES_FACTURA.get(unidad, unidad),
                       'precio_unitario': it['product_price'], 'subtotal': it['subtotal']})
    if pedido['descuento']:
        lineas.append({'codigo': '', 'descripcion': f"Descuento {pedido['cupon_codigo'] or ''}".strip(),
                       'cantidad': 1, 'unidad': 'unidades', 'precio_unitario': -pedido['descuento'],
                       'subtotal': -pedido['descuento']})
    if pedido['envio']:
        lineas.append({'codigo': '', 'descripcion': 'Envío', 'cantidad': 1, 'unidad': 'unidades',
                       'precio_unitario': pedido['envio'], 'subtotal': pedido['envio']})
    for linea in lineas:
        linea['neto'] = round(linea['subtotal'] / (1 + IVA_PORC), 2)
    return lineas


def emitir_factura(db, pedido_id):
    """Genera la factura de un pedido ya pagado y le pide su CAE.
    Devuelve True si emitió una factura nueva y False si el pedido ya tenía una (o si no
    se pudo autorizar: en ese caso queda registrado en el log y la compra no se interrumpe)."""
    existe = db.execute("SELECT id FROM facturas WHERE pedido_id=?", (pedido_id,)).fetchone()
    if existe:
        return False

    pedido = db.execute("SELECT * FROM pedidos WHERE id=?", (pedido_id,)).fetchone()
    cliente = db.execute("SELECT * FROM clientes WHERE id=?", (pedido['cliente_id'],)).fetchone()
    items = db.execute("SELECT * FROM pedido_items WHERE pedido_id=?", (pedido_id,)).fetchall()
    cfg = config_facturacion()
    es_a = bool(pedido['factura_a_cuit']) and cfg['condicion_iva'] == 'responsable_inscripto'
    letra = letra_de_comprobante(cfg['condicion_iva'], es_a)
    try:
        numero, cae, vencimiento = solicitar_cae(db, letra, cfg['punto_venta'], cfg['modo'])
    except ErrorArca:
        app.logger.exception("No se pudo autorizar la factura del pedido #%s", pedido_id)
        return False

    # En la tabla, "subtotal" de la factura es el neto gravado (sin IVA).
    # La Factura C no discrimina IVA: el neto es el total.
    total = pedido['total']
    if letra == 'C':
        subtotal_sin_iva, iva = total, 0.0
    elif letra == 'A':
        # En la A el neto es la suma de los netos de cada renglón, para que el detalle cierre al centavo
        subtotal_sin_iva = round(sum(l['neto'] for l in lineas_de_factura(pedido, items, db)), 2)
        iva = round(total - subtotal_sin_iva, 2)
    else:
        subtotal_sin_iva, iva = calcular_iva(total)

    if es_a:
        nombre, documento, condicion_cliente = (pedido['factura_a_razon'], formatear_cuit(pedido['factura_a_cuit']),
                                                'IVA Responsable Inscripto')
    else:
        nombre, documento, condicion_cliente = (f"{cliente['nombre']} {cliente['apellido']}",
                                                cliente['dni_cuit'] or '', 'Consumidor Final')
    domicilio = pedido['direccion_envio'] or ' '.join(x for x in (cliente['calle'], cliente['numero'], cliente['barrio']) if x)
    pago = db.execute("SELECT metodo FROM pagos WHERE pedido_id=? ORDER BY id DESC LIMIT 1", (pedido_id,)).fetchone()
    condicion_venta = NOMBRES_METODO_PAGO.get(pago['metodo'], pago['metodo']) if pago else 'Contado'

    db.execute("""
        INSERT INTO facturas (pedido_id, numero, tipo, cliente_nombre, cliente_email, cliente_dni,
                               subtotal, iva, total, cae, cae_vencimiento, modo_arca, punto_venta, nro_comprobante,
                               cliente_condicion_iva, cliente_domicilio, condicion_venta)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (pedido_id, formatear_numero(letra, cfg['punto_venta'], numero), letra, nombre, cliente['email'],
          documento, subtotal_sin_iva, iva, total, cae, vencimiento, cfg['modo'], cfg['punto_venta'], numero,
          condicion_cliente, domicilio, condicion_venta))
    return True


def url_de_verificacion(factura):
    """Dirección que lleva el QR de la factura. En demostración apunta a la página de verificación
    de esta misma tienda; con ARCA real sería el formato oficial (https://www.afip.gob.ar/fe/qr/?p=...)."""
    if factura['modo_arca'] == 'demo' or not factura['modo_arca']:
        base = url_publica() if has_request_context() else os.environ.get('SITE_URL', '').rstrip('/')
        return f"{base}{url_for('verificar_factura', numero=factura['numero'])}" if has_request_context() \
            else f"{base}/verificar-factura/{factura['numero']}"
    cfg = config_facturacion()
    datos = {'ver': 1, 'fecha': factura['created_at'][:10], 'cuit': int(re.sub(r'\D', '', cfg['cuit']) or 0),
             'ptoVta': factura['punto_venta'], 'tipoCmp': CODIGOS_COMPROBANTE.get(factura['tipo'], 6),
             'nroCmp': factura['nro_comprobante'], 'importe': factura['total'], 'moneda': 'PES', 'ctz': 1,
             'tipoDocRec': 80 if len(re.sub(r'\D', '', factura['cliente_dni'])) == 11 else (96 if factura['cliente_dni'] else 99),
             'nroDocRec': int(re.sub(r'\D', '', factura['cliente_dni']) or 0),
             'tipoCodAut': 'E', 'codAut': int(factura['cae'] or 0)}
    return "https://www.afip.gob.ar/fe/qr/?p=" + base64.urlsafe_b64encode(json.dumps(datos).encode()).decode()


@app.route('/verificar-factura/<numero>')
def verificar_factura(numero):
    """Página pública a la que lleva el QR: confirma que la factura existe y muestra sus datos
    de autorización. No muestra datos personales del cliente."""
    with get_db() as db:
        factura = db.execute("SELECT * FROM facturas WHERE numero=?", (numero,)).fetchone()
    if not factura:
        abort(404)
    return render_template('verificar_factura.html', factura=factura, emisor=config_facturacion())


@app.route('/pedido/<int:pedido_id>/confirmado')
@cliente_required(lambda: redirect(url_for('login')))
def pedido_confirmado(pedido_id):
    """Pantalla de pedido confirmado, con su estado, su pago y su factura."""
    with get_db() as db:
        pedido = db.execute("SELECT * FROM pedidos WHERE id=? AND cliente_id=?",
                             (pedido_id, session['cliente_id'])).fetchone()
        if not pedido:
            return redirect(url_for('index'))
        items = db.execute("SELECT * FROM pedido_items WHERE pedido_id=?", (pedido_id,)).fetchall()
        pago = db.execute("SELECT * FROM pagos WHERE pedido_id=? ORDER BY id DESC LIMIT 1", (pedido_id,)).fetchone()
        factura = db.execute("SELECT * FROM facturas WHERE pedido_id=?", (pedido_id,)).fetchone()
        timeline = linea_de_tiempo(db, pedido)
    return render_template('pedido_confirmado.html', pedido=pedido, items=items, pago=pago,
                            factura=factura, timeline=timeline, transferencia=datos_transferencia(),
                            codigo_barras=svg_codigo_de_barras(pago['referencia'])
                                if pago['metodo'] == 'efectivo' and pago['estado'] == 'pendiente' else '',
                            vence_cupon=vencimiento_cupon(pago))


def cargar_cupon_de_pago(pedido_id):
    """Pedido, pago y vencimiento del cupón de un pedido del cliente que paga en efectivo."""
    with get_db() as db:
        pedido = db.execute("SELECT * FROM pedidos WHERE id=? AND cliente_id=?",
                            (pedido_id, session['cliente_id'])).fetchone()
        pago = db.execute("SELECT * FROM pagos WHERE pedido_id=? ORDER BY id DESC LIMIT 1", (pedido_id,)).fetchone()
    if not pedido or not pago or pago['metodo'] != 'efectivo':
        abort(404)
    return pedido, pago, vencimiento_cupon(pago)


@app.route('/pedido/<int:pedido_id>/cupon-de-pago')
@cliente_required(lambda: redirect(url_for('login')))
def cupon_de_pago(pedido_id):
    """Cupón para pagar en efectivo en Rapipago o Pago Fácil, listo para imprimir."""
    pedido, pago, vence = cargar_cupon_de_pago(pedido_id)
    return render_template('cupon_pago.html', pedido=pedido, pago=pago, vence=vence,
                           codigo_barras=svg_codigo_de_barras(pago['referencia']),
                           pagado=pago['estado'] == 'aprobado')


@app.route('/pedido/<int:pedido_id>/cupon-de-pago.pdf')
@cliente_required(lambda: redirect(url_for('login')))
def cupon_de_pago_pdf(pedido_id):
    """El mismo cupón en PDF."""
    pedido, pago, vence = cargar_cupon_de_pago(pedido_id)
    nombre_negocio = get_config('arca_razon_social', '') or 'Almacén del Valle'
    buffer = io.BytesIO()
    c = pdfcanvas.Canvas(buffer, pagesize=A4)
    ancho, alto = A4
    x0, y = 3 * cm, alto - 3 * cm
    c.setFont('Helvetica-Bold', 18); c.drawString(x0, y, 'Cupón de pago')
    c.setFont('Helvetica', 10); c.drawString(x0, y - 16, nombre_negocio)
    c.setFont('Helvetica', 11)
    filas = [('Pedido', f"#{pedido['id']}"), ('Importe a pagar', f"$ {ar_money2(pago['monto'])}"),
             ('Vence', vence.strftime('%d/%m/%Y %H:%M')), ('Código de pago', pago['referencia'])]
    y -= 56
    for etiqueta, valor in filas:
        c.setFont('Helvetica', 10); c.setFillColor(colors.grey); c.drawString(x0, y, etiqueta)
        c.setFont('Helvetica-Bold', 13); c.setFillColor(colors.black); c.drawString(x0 + 4.5 * cm, y, valor)
        y -= 24
    codigo = Code128(pago['referencia'], barHeight=70, barWidth=1.3, quiet=False)
    codigo.drawOn(c, x0, y - 90)
    c.setFont('Helvetica', 9)
    c.drawString(x0, y - 106, pago['referencia'])
    y -= 140
    c.setFont('Helvetica', 10)
    for linea in ('Cómo pagar:', '1. Presentá este cupón (impreso o en el celular) en cualquier sucursal de Rapipago o Pago Fácil.',
                  '2. Pagá el importe exacto en efectivo.',
                  '3. Cuando el local confirme el cobro, tu pedido pasa a preparación y recibís tu factura por correo.',
                  f'El cupón vence el {vence.strftime("%d/%m/%Y")}. Pasada esa fecha, el pedido se puede cancelar.',
                  'Cupón de demostración: sin validez de cobro real.'):
        c.drawString(x0, y, linea); y -= 16
    c.showPage(); c.save()
    buffer.seek(0)
    return send_file(buffer, mimetype='application/pdf', as_attachment=True,
                     download_name=f"cupon-pedido-{pedido['id']}.pdf")


@app.route('/api/pedido/<int:pedido_id>/estado')
@cliente_required(lambda: (jsonify({'error': 'Iniciá sesión primero'}), 401))
def api_pedido_estado(pedido_id):
    """Estado actual de un pedido del cliente; la pantalla de seguimiento lo consulta para actualizarse sola."""
    with get_db() as db:
        pedido = db.execute("SELECT estado FROM pedidos WHERE id=? AND cliente_id=?",
                            (pedido_id, session['cliente_id'])).fetchone()
    if not pedido:
        return jsonify({'error': 'Pedido no encontrado'}), 404
    return jsonify({'estado': pedido['estado']})


@app.route('/mis-pedidos')
@cliente_required(lambda: redirect(url_for('login')))
def mis_pedidos():
    """Historial de pedidos del cliente."""
    with get_db() as db:
        pedidos = db.execute("""
            SELECT p.*, f.numero AS factura_numero
            FROM pedidos p LEFT JOIN facturas f ON f.pedido_id = p.id
            WHERE p.cliente_id=? ORDER BY p.created_at DESC
        """, (session['cliente_id'],)).fetchall()
    return render_template('mis_pedidos.html', pedidos=pedidos)


# ---------------------------------------------------------------------------
# FACTURA — datos, PDF, versión imprimible y envío por correo
# ---------------------------------------------------------------------------

def cargar_factura(db, pedido_id):
    """Devuelve (factura, pedido, items) de un pedido, o None si el pedido no existe
    o todavía no tiene factura emitida."""
    pedido = db.execute("SELECT * FROM pedidos WHERE id=?", (pedido_id,)).fetchone()
    factura = db.execute("SELECT * FROM facturas WHERE pedido_id=?", (pedido_id,)).fetchone()
    if not pedido or not factura:
        return None
    items = db.execute("SELECT * FROM pedido_items WHERE pedido_id=?", (pedido_id,)).fetchall()
    return factura, pedido, items


def resumen_factura(factura, pedido):
    """Desglose de importes que se muestra en el correo, el PDF y la versión imprimible.
    Los precios ya incluyen IVA, así que el neto y el IVA se calculan sobre el total:
    productos - descuento + envío = total = neto sin IVA + IVA."""
    return {
        'productos': pedido['subtotal'],
        'descuento': pedido['descuento'] or 0,
        'cupon': pedido['cupon_codigo'] or '',
        'envio': pedido['envio'],
        'neto': factura['subtotal'],
        'iva': factura['iva'],
        'iva_porc': int(IVA_PORC * 100),
        'total': factura['total'],
        'discrimina_iva': factura['tipo'] != 'C',
    }


def _get_factura_o_403(pedido_id):
    """Devuelve (factura, pedido, items) o una respuesta de error si no corresponde."""
    with get_db() as db:
        pedido = db.execute("SELECT * FROM pedidos WHERE id=?", (pedido_id,)).fetchone()
        if not pedido:
            return ("Pedido no encontrado", 404), None, None

        es_dueño = session.get('cliente_id') == pedido['cliente_id']
        es_admin = session.get('admin_logged_in')
        if not (es_dueño or es_admin):
            return ("No autorizado", 403), None, None

        datos = cargar_factura(db, pedido_id)
        if not datos:
            return ("Este pedido todavía no tiene una factura emitida", 404), None, None

    return datos


@app.route('/factura/<int:pedido_id>/pdf')
def factura_pdf(pedido_id):
    """Muestra la factura en PDF dentro del navegador."""
    factura, pedido, items = _get_factura_o_403(pedido_id)
    if isinstance(factura, tuple):
        return factura  # fue un error (mensaje, código)
    return enviar_pdf_factura(factura, pedido, items, attachment=False)


@app.route('/factura/<int:pedido_id>/descargar')
def factura_descargar(pedido_id):
    """Descarga la factura en PDF."""
    factura, pedido, items = _get_factura_o_403(pedido_id)
    if isinstance(factura, tuple):
        return factura
    return enviar_pdf_factura(factura, pedido, items, attachment=True)


@app.route('/factura/<int:pedido_id>/imprimir')
def factura_imprimir(pedido_id):
    """Para imprimir se usa el mismo PDF que se descarga, así el comprobante es siempre idéntico."""
    return redirect(url_for('factura_pdf', pedido_id=pedido_id))


class LienzoNumerado(pdfcanvas.Canvas):
    """Lienzo que conoce el total de páginas, para escribir "Pág. 1/1" como los comprobantes de ARCA."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._estados = []

    def showPage(self):
        self._estados.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._estados)
        for estado in self._estados:
            self.__dict__.update(estado)
            self.setFont('Helvetica', 8)
            self.drawRightString(A4[0] - 1.2 * cm, 3.5 * cm, f"Pág. {self._pageNumber}/{total}")
            super().showPage()
        super().save()


def fecha_corta(texto):
    """'2026-11-16 18:51:03' o '2026-11-16' -> '16/11/2026'. Vacío si no hay fecha."""
    try:
        return datetime.strptime((texto or '')[:10], '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return '-'


def construir_pdf_factura(factura, pedido, items):
    """Dibuja la factura con la disposición de los comprobantes electrónicos de ARCA y devuelve el PDF
    como bytes: ORIGINAL, encabezado en tres bloques (emisor, letra con su código, datos del
    comprobante), receptor, detalle, totales y, al pie, QR y CAE."""
    emisor = config_facturacion()
    letra = factura['tipo']
    es_demo = (factura['modo_arca'] or 'demo') == 'demo' and bool(factura['cae'])
    lineas = lineas_de_factura(pedido, items)
    fecha = fecha_corta(factura['created_at'])
    verde = colors.HexColor('#16493D')

    base = getSampleStyleSheet()['Normal']
    chico = ParagraphStyle('ChicoAV', parent=base, fontSize=8, leading=10.5)
    centrado = ParagraphStyle('CentradoAV', parent=chico, alignment=1)
    celda = ParagraphStyle('CeldaFAV', parent=base, fontSize=7.5, leading=9)

    def dato(rotulo, valor):
        return Paragraph(f"<b>{rotulo}</b> {escape(str(valor)) if valor not in (None, '') else '-'}", chico)

    elementos = []
    caja_original = Table([[Paragraph("<b>ORIGINAL</b>", ParagraphStyle('OrigAV', parent=centrado, fontSize=11, leading=14))]],
                          colWidths=[18.6 * cm])
    caja_original.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.8, colors.black)]))
    elementos.append(caja_original)

    emisor_bloque = [
        Paragraph(f"<b>{escape(emisor['razon_social'])}</b>", ParagraphStyle('RazonAV', parent=chico, fontSize=13, leading=16, textColor=verde)),
        Spacer(1, 4),
        dato('Razón social:', emisor['razon_social']),
        dato('Domicilio comercial:', emisor['domicilio']),
        dato('Condición frente al IVA:', CONDICIONES_IVA.get(emisor['condicion_iva'], '')),
    ]
    letra_bloque = [
        Paragraph(f"<b>{letra}</b>", ParagraphStyle('LetraAV', parent=centrado, fontSize=30, leading=34)),
        Paragraph(f"COD. {CODIGOS_COMPROBANTE.get(letra, 6):02d}", ParagraphStyle('CodAV', parent=centrado, fontSize=7)),
    ]
    inicio = fecha_corta(emisor['inicio_actividades']) if emisor['inicio_actividades'] else '-'
    comprobante_bloque = [
        Paragraph("<b>FACTURA</b>", ParagraphStyle('TipoAV', parent=chico, fontSize=15, leading=18)),
        Spacer(1, 2),
        dato('Punto de Venta:', f"{factura['punto_venta']:05d}   Comp. Nro: {factura['nro_comprobante']:08d}"
             if factura['nro_comprobante'] else factura['numero']),
        dato('Fecha de Emisión:', fecha),
        dato('CUIT:', emisor['cuit']),
        dato('Ingresos Brutos:', emisor['ingresos_brutos']),
        dato('Fecha de Inicio de Actividades:', inicio),
    ]
    encabezado = Table([[emisor_bloque, letra_bloque, comprobante_bloque]], colWidths=[8.3 * cm, 2.0 * cm, 8.3 * cm])
    encabezado.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.8, colors.black),
                                    ('LINEAFTER', (0, 0), (1, 0), 0.8, colors.black),
                                    ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                                    ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]))
    elementos.append(encabezado)

    periodo = Table([[dato('Período Facturado Desde:', fecha), dato('Hasta:', fecha),
                      dato('Fecha de Vto. para el pago:', fecha)]], colWidths=[7.2 * cm, 4.2 * cm, 7.2 * cm])
    periodo.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.8, colors.black)]))
    elementos.append(periodo)

    receptor = Table([
        [dato('CUIT/DNI:', factura['cliente_dni']), dato('Apellido y Nombre / Razón Social:', factura['cliente_nombre'])],
        [dato('Condición frente al IVA:', factura['cliente_condicion_iva'] or 'Consumidor Final'),
         dato('Domicilio:', factura['cliente_domicilio'])],
        [dato('Condición de venta:', factura['condicion_venta'] or 'Contado'), ''],
    ], colWidths=[7.2 * cm, 11.4 * cm])
    receptor.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.8, colors.black)]))
    elementos.append(receptor)
    elementos.append(Spacer(1, 8))

    # Detalle: la Factura A muestra los importes sin IVA, su alícuota y el subtotal con IVA;
    # la B y la C muestran el precio final
    if letra == 'A':
        columnas = ['Código', 'Producto / Servicio', 'Cantidad', 'U. medida', 'Precio Unit.', '% Bonif',
                    'Imp. Bonif.', 'Subtotal', 'Alícuota IVA', 'Subtotal c/IVA']
        anchos = [1.7, 3.6, 1.4, 1.5, 1.8, 1.2, 1.5, 1.9, 1.8, 2.2]
        filas = [[l['codigo'], Paragraph(escape(l['descripcion']), celda), ar_money2(l['cantidad']), l['unidad'],
                  ar_money2(round(l['precio_unitario'] / (1 + IVA_PORC), 2)), '0,00', '0,00',
                  ar_money2(l['neto']), f"{IVA_PORC * 100:.2f}%".replace('.', ','), ar_money2(l['subtotal'])]
                 for l in lineas]
    else:
        columnas = ['Código', 'Producto / Servicio', 'Cantidad', 'U. medida', 'Precio Unit.', '% Bonif',
                    'Imp. Bonif.', 'Subtotal']
        anchos = [1.7, 5.9, 1.8, 1.9, 2.3, 1.3, 1.5, 2.2]
        filas = [[l['codigo'], Paragraph(escape(l['descripcion']), celda), ar_money2(l['cantidad']), l['unidad'],
                  ar_money2(l['precio_unitario']), '0,00', '0,00', ar_money2(l['subtotal'])] for l in lineas]
    tabla = Table([columnas] + filas, colWidths=[a * cm for a in anchos], repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E7ECE8')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOX', (0, 0), (-1, -1), 0.8, colors.black), ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 10))

    # Totales; en la Factura B va además la leyenda de transparencia fiscal al consumidor
    derecha_estilo = ParagraphStyle('DerAV', parent=chico, alignment=TA_RIGHT)
    if letra == 'A':
        total_filas = [['Importe Neto Gravado: $', ar_money2(factura['subtotal'])]]
        total_filas += [[f"IVA {alicuota}%: $", ar_money2(factura['iva'] if alicuota == '21' else 0)]
                        for alicuota in ('27', '21', '10,5', '5', '2,5', '0')]
        total_filas += [['Importe Otros Tributos: $', '0,00'], ['Importe Total: $', ar_money2(factura['total'])]]
    else:
        total_filas = [['Subtotal: $', ar_money2(factura['total'])], ['Importe Otros Tributos: $', '0,00'],
                       ['Importe Total: $', ar_money2(factura['total'])]]
    tabla_totales = Table([[Paragraph(f"<b>{r}</b>" if r.startswith('Importe Total') else r, derecha_estilo),
                            Paragraph(f"<b>{v}</b>" if r.startswith('Importe Total') else v, derecha_estilo)]
                           for r, v in total_filas], colWidths=[4.6 * cm, 2.6 * cm])
    tabla_totales.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 0.8, colors.black),
                                       ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2)]))
    transparencia = ''
    if letra == 'B':
        transparencia = [Paragraph("<b>Régimen de Transparencia Fiscal al Consumidor (Ley 27.743)</b>", chico),
                         dato('IVA Contenido: $', ar_money2(factura['iva'])),
                         dato('Otros Impuestos Nacionales Indirectos: $', '0,00')]
    totales = Table([[transparencia, tabla_totales]], colWidths=[11.4 * cm, 7.2 * cm])
    totales.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0)]))
    elementos.append(totales)

    def dibujar_pie(lienzo, _doc):
        """Parte fija al pie de cada página: QR, estado del comprobante y CAE."""
        ancho, alto = A4
        lienzo.saveState()
        if es_demo:
            lienzo.setFillColor(colors.HexColor('#B3261E'))
            lienzo.setFont('Helvetica-Bold', 9)
            lienzo.drawCentredString(ancho / 2, alto - 0.8 * cm, "COMPROBANTE DE DEMOSTRACIÓN — SIN VALIDEZ FISCAL")
            lienzo.setFillColor(colors.black)
        lienzo.setLineWidth(0.8)
        lienzo.line(1.2 * cm, 4.0 * cm, ancho - 1.2 * cm, 4.0 * cm)
        if factura['cae']:
            qr = QrCodeWidget(url_de_verificacion(factura))
            x0, y0, x1, y1 = qr.getBounds()
            lado = 2.8 * cm
            dibujo = Drawing(lado, lado, transform=[lado / (x1 - x0), 0, 0, lado / (y1 - y0), 0, 0])
            dibujo.add(qr)
            renderPDF.draw(dibujo, lienzo, 1.2 * cm, 0.9 * cm)
            lienzo.setFont('Helvetica-Bold', 10)
            lienzo.drawString(4.9 * cm, 3.0 * cm, "Comprobante de demostración" if es_demo else "Comprobante Autorizado")
            if not es_demo:
                lienzo.setFont('Helvetica-Oblique', 7)
                lienzo.drawString(4.9 * cm, 2.5 * cm, "Esta Administración Federal no se responsabiliza por los datos")
                lienzo.drawString(4.9 * cm, 2.1 * cm, "ingresados en el detalle de la operación")
            lienzo.setFont('Helvetica-Bold', 9)
            lienzo.drawRightString(ancho - 1.2 * cm, 2.9 * cm, f"CAE N°: {factura['cae']}")
            lienzo.drawRightString(ancho - 1.2 * cm, 2.3 * cm, f"Fecha de Vto. de CAE: {fecha_corta(factura['cae_vencimiento'])}")
        lienzo.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5 * cm, bottomMargin=4.4 * cm,
                             leftMargin=1.2 * cm, rightMargin=1.2 * cm)
    doc.build(elementos, onFirstPage=dibujar_pie, onLaterPages=dibujar_pie, canvasmaker=LienzoNumerado)
    return buffer.getvalue()


def enviar_pdf_factura(factura, pedido, items, attachment=False):
    """Responde la petición con el PDF de la factura."""
    pdf = construir_pdf_factura(factura, pedido, items)
    return send_file(io.BytesIO(pdf), mimetype='application/pdf', as_attachment=attachment,
                      download_name=f"factura_{factura['numero']}.pdf")


def construir_mail_factura(pedido_id):
    """Arma el correo con la factura de un pedido: datos del cliente, detalle de la
    compra, subtotal, IVA y total, más el PDF adjunto. Devuelve None si no hay factura."""
    with get_db() as db:
        datos = cargar_factura(db, pedido_id)
        if not datos:
            return None
        factura, pedido, items = datos
        cliente = db.execute("SELECT * FROM clientes WHERE id=?", (pedido['cliente_id'],)).fetchone()
        pago = db.execute("SELECT * FROM pagos WHERE pedido_id=? ORDER BY id DESC LIMIT 1",
                          (pedido_id,)).fetchone()

    contexto = {
        'factura': factura, 'pedido': pedido, 'items': items, 'cliente': cliente,
        'resumen': resumen_factura(factura, pedido),
        'medio_pago': NOMBRES_METODO_PAGO.get(pago['metodo'], pago['metodo']) if pago else '',
        'referencia_pago': pago['referencia'] if pago else '',
    }

    msg = EmailMessage()
    msg['Subject'] = f"Tu factura {factura['numero']} - Almacén del Valle"
    msg['To'] = factura['cliente_email']
    msg.set_content(render_template('email_factura.txt', **contexto))
    msg.add_alternative(render_template('email_factura.html', **contexto), subtype='html')
    msg.add_attachment(construir_pdf_factura(factura, pedido, items),
                       maintype='application', subtype='pdf',
                       filename=f"factura_{factura['numero']}.pdf")
    return msg


def enviar_factura_por_email(pedido_id, esperar=False):
    """Envía al cliente la factura de su pedido apenas se confirma el pago.
    Nunca interrumpe la compra. Con esperar=True devuelve (ok, motivo) para mostrarlo en el panel."""
    cfg = leer_config_smtp()
    if not cfg:
        motivo = "El correo no está configurado: activalo en Datos del Negocio > Correo electrónico."
        app.logger.warning("Factura del pedido #%s no enviada: %s", pedido_id, motivo)
        return (False, motivo) if esperar else None
    try:
        msg = construir_mail_factura(pedido_id)
        if not msg:
            return (False, "El pedido todavía no tiene factura.") if esperar else None
        resultado = enviar_email(msg, cfg, 'factura', pedido_id, esperar)
        return resultado if esperar else None
    except Exception as e:
        app.logger.exception("No se pudo preparar el correo de la factura del pedido #%s", pedido_id)
        return (False, f"No se pudo armar el correo: {e}") if esperar else None


@app.route('/admin/pedidos/<int:pedido_id>/reenviar-factura', methods=['POST'])
@permiso_required('pedidos')
def reenviar_factura(pedido_id):
    """Vuelve a mandar la factura al correo del cliente y dice si salió o por qué no."""
    ok, motivo = enviar_factura_por_email(pedido_id, esperar=True)
    if ok:
        flash(f'Factura del pedido #{pedido_id} enviada al cliente.', 'success')
        registrar_auditoria('Factura reenviada por correo', f'Pedido #{pedido_id}')
    else:
        flash(f'No se pudo enviar la factura del pedido #{pedido_id}: {motivo}', 'error')
    return redirect(url_for('admin', view='pedidos'))


# ---------------------------------------------------------------------------
# REPORTES — helpers de datos (ventas, stock, valorización, facturación)
# ---------------------------------------------------------------------------

def build_ventas_data(year, month=None):
    """Devuelve {categoria: [filas]} con lo vendido en pedidos pagados/entregados."""
    query = """
        SELECT c.name AS category, pi.product_name AS name, p.brand AS brand,
               pi.product_price AS price, SUM(pi.qty) AS total_qty,
               SUM(pi.subtotal) AS total
        FROM pedido_items pi
        JOIN pedidos ped ON ped.id = pi.pedido_id
        LEFT JOIN products p ON p.id = pi.product_id
        LEFT JOIN categories c ON c.id = p.category_id
        WHERE ped.estado != 'cancelado'
          AND strftime('%Y', ped.created_at) = ?
          {mes_filtro}
        GROUP BY pi.product_name
        ORDER BY category, total_qty DESC
    """
    params = [str(year)]
    if month:
        query = query.format(mes_filtro="AND strftime('%m', ped.created_at) = ?")
        params.append(f"{month:02d}")
    else:
        query = query.format(mes_filtro="")

    with get_db() as db:
        rows = db.execute(query, params).fetchall()

    cat_data = {}
    for r in rows:
        cat = r['category'] or 'Sin categoría'
        cat_data.setdefault(cat, []).append({
            'name': r['name'], 'brand': r['brand'] or '-',
            'price': r['price'], 'qty': r['total_qty'], 'total': r['total'],
        })
    return cat_data


def build_stock_bajo_data():
    """Productos con stock igual o menor al umbral de stock bajo."""
    with get_db() as db:
        rows = db.execute("""
            SELECT p.name, p.brand, c.name AS category, p.stock, p.price
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            WHERE p.stock <= ? ORDER BY p.stock ASC, p.name
        """, (LOW_STOCK_THRESHOLD,)).fetchall()
    return [dict(r) for r in rows]


def build_stock_general_data():
    """Todos los productos con su stock, ordenados por categoría."""
    with get_db() as db:
        rows = db.execute("""
            SELECT p.name, p.brand, c.name AS category, p.stock, p.price, p.active
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            ORDER BY category, p.name
        """).fetchall()
    return [dict(r) for r in rows]


def build_valor_stock_data():
    """Stock general con el valor de cada producto (stock por precio)."""
    rows = build_stock_general_data()
    for r in rows:
        r['valor'] = r['stock'] * r['price']
    return rows


def build_facturacion_data(year, month=None):
    """Facturas emitidas en un año, o en un mes puntual si se indica."""
    query = """
        SELECT f.numero, f.created_at, f.cliente_nombre, f.subtotal, f.iva, f.total
        FROM facturas f
        WHERE strftime('%Y', f.created_at) = ? {mes_filtro}
        ORDER BY f.created_at DESC
    """
    params = [str(year)]
    if month:
        query = query.format(mes_filtro="AND strftime('%m', f.created_at) = ?")
        params.append(f"{month:02d}")
    else:
        query = query.format(mes_filtro="")
    with get_db() as db:
        rows = db.execute(query, params).fetchall()
    return [dict(r) for r in rows]


# ---------------------------------------------------------------------------
# REPORTES — generación de PDF (reportlab)
# ---------------------------------------------------------------------------

VERDE_PDF = colors.HexColor('#16493D')
MANGO_PDF = colors.HexColor('#DB8E14')
ROJO_PDF  = colors.HexColor('#C1554A')
LGRAY_PDF = colors.HexColor('#F2F2F2')


class _NumberedCanvas(pdfcanvas.Canvas):
    """Canvas que guarda páginas y añade 'Página N de M' al pie."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        """Guarda el estado de la página en lugar de cerrarla, para poder numerarlas al final."""
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        """Dibuja "Página N de M" en cada página y recién entonces cierra el PDF."""
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(num_pages)
            super().showPage()
        super().save()

    def _draw_page_number(self, page_count):
        """Dibuja el pie de página con la numeración."""
        self.saveState()
        self.setFont('Helvetica', 8)
        self.setFillColor(colors.HexColor('#888888'))
        text = f"Página {self._pageNumber} de {page_count}   •   Almacén del Valle"
        self.drawCentredString(A4[0] / 2, 0.7 * cm, text)
        self.setStrokeColor(colors.HexColor('#DDDDDD'))
        self.setLineWidth(0.3)
        self.line(1.5*cm, 1.0*cm, A4[0]-1.5*cm, 1.0*cm)
        self.restoreState()


def _report_styles():
    """Estilos de texto compartidos por los reportes en PDF."""
    S = ParagraphStyle
    return {
        'title': S('t', fontSize=18, textColor=VERDE_PDF, fontName='Helvetica-Bold', spaceAfter=4),
        'sub':   S('s', fontSize=10, textColor=colors.grey, spaceAfter=14),
        'warn':  S('w', fontSize=10, textColor=ROJO_PDF, spaceAfter=14),
        'cat':   S('c', fontSize=12, textColor=VERDE_PDF, fontName='Helvetica-Bold', spaceBefore=14, spaceAfter=5),
    }


def _table_style(header_color=VERDE_PDF, left_align_col=2):
    """Estilo de tabla compartido por los reportes en PDF; left_align_col es la columna alineada a la izquierda."""
    return TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (left_align_col, 1), (left_align_col, -1), 'LEFT'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, LGRAY_PDF]),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#DDDDDD')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ])


def pdf_ventas(year, month=None):
    """Reporte PDF de ventas por producto, de un mes o de un año completo."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    tipo = "Mensual" if month else "Anual"
    cat_data = build_ventas_data(year, month)

    story = [
        Paragraph("ALMACÉN DEL VALLE", st['title']),
        Paragraph(f"Reporte {tipo} de Ventas — {period}", st['sub']),
    ]
    if not cat_data:
        story.append(Paragraph("No hay ventas registradas para el período seleccionado.",
                                getSampleStyleSheet()['Normal']))
    else:
        for cat, items in cat_data.items():
            story.append(Paragraph(f"▸ {cat}", st['cat']))
            rows = [['#', 'Producto', 'Marca', 'Precio unit.', 'Unidades', 'Total $']]
            for i, it in enumerate(items, 1):
                rows.append([str(i), it['name'], it['brand'], f"${ar_money(it['price'])}",
                             str(it['qty']), f"${ar_money(it['total'])}"])
            t = Table(rows, colWidths=[0.8*cm, 6.5*cm, 3*cm, 2.5*cm, 2.2*cm, 2.5*cm])
            t.setStyle(_table_style(left_align_col=1))
            story.append(t)
    story += [Spacer(1, 0.5*cm), Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def pdf_stock_bajo():
    """Reporte PDF de los productos con stock bajo."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    data = build_stock_bajo_data()
    story = [
        Paragraph("ALMACÉN DEL VALLE", st['title']),
        Paragraph(f"Alerta de Stock Bajo — productos con {LOW_STOCK_THRESHOLD} unidades o menos", st['warn']),
    ]
    if not data:
        story.append(Paragraph("No hay productos con stock bajo.", getSampleStyleSheet()['Normal']))
    else:
        rows = [['#', 'Producto', 'Marca', 'Categoría', 'Stock', 'Precio $']]
        for i, it in enumerate(data, 1):
            rows.append([str(i), it['name'], it['brand'] or '-', it['category'] or '-',
                         str(it['stock']), f"${ar_money(it['price'])}"])
        t = Table(rows, colWidths=[0.8*cm, 6*cm, 3*cm, 3*cm, 1.8*cm, 2.5*cm])
        t.setStyle(_table_style(header_color=ROJO_PDF, left_align_col=1))
        story.append(t)
    story += [Spacer(1, 0.5*cm), Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def pdf_stock_general():
    """Reporte PDF con el stock de todos los productos."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    data = build_stock_general_data()
    story = [Paragraph("ALMACÉN DEL VALLE", st['title']), Paragraph("Stock General de Productos", st['sub'])]
    rows = [['#', 'Producto', 'Marca', 'Categoría', 'Stock', 'Precio $', 'Estado']]
    for i, it in enumerate(data, 1):
        rows.append([str(i), it['name'], it['brand'] or '-', it['category'] or '-',
                     str(it['stock']), f"${ar_money(it['price'])}", 'Activo' if it['active'] else 'Inactivo'])
    t = Table(rows, colWidths=[0.8*cm, 5.5*cm, 2.8*cm, 2.8*cm, 1.6*cm, 2.2*cm, 2*cm])
    t.setStyle(_table_style(left_align_col=1))
    story.append(t)
    story += [Spacer(1, 0.5*cm), Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def pdf_valor_stock():
    """Reporte PDF con el valor del inventario (stock por precio)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    data = build_valor_stock_data()
    total_general = sum(r['valor'] for r in data)
    story = [Paragraph("ALMACÉN DEL VALLE", st['title']), Paragraph("Valorización de Stock (inventario a precio de venta)", st['sub'])]
    rows = [['#', 'Producto', 'Categoría', 'Stock', 'Precio $', 'Valor total $']]
    for i, it in enumerate(data, 1):
        rows.append([str(i), it['name'], it['category'] or '-', str(it['stock']),
                     f"${ar_money(it['price'])}", f"${ar_money(it['valor'])}"])
    t = Table(rows, colWidths=[0.8*cm, 6.5*cm, 3.2*cm, 1.8*cm, 2.5*cm, 2.7*cm])
    t.setStyle(_table_style(left_align_col=1))
    story.append(t)
    story += [Spacer(1, 0.4*cm),
              Paragraph(f"<b>Valor total del inventario: ${ar_money(total_general)}</b>", getSampleStyleSheet()['Normal']),
              Spacer(1, 0.3*cm),
              Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def pdf_facturacion(year, month=None):
    """Reporte PDF de las facturas emitidas, con neto, IVA y total."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    tipo = "Mensual" if month else "Anual"
    data = build_facturacion_data(year, month)
    story = [Paragraph("ALMACÉN DEL VALLE", st['title']), Paragraph(f"Reporte {tipo} de Facturación — {period}", st['sub'])]
    if not data:
        story.append(Paragraph("No se emitieron facturas en el período seleccionado.", getSampleStyleSheet()['Normal']))
    else:
        rows = [['N° Factura', 'Fecha', 'Cliente', 'Subtotal $', 'IVA $', 'Total $']]
        total_gral = 0
        for f in data:
            rows.append([f['numero'], f['created_at'][:10], f['cliente_nombre'],
                         f"${ar_money(f['subtotal'])}", f"${ar_money(f['iva'])}", f"${ar_money(f['total'])}"])
            total_gral += f['total']
        t = Table(rows, colWidths=[3.2*cm, 2.5*cm, 5*cm, 2.3*cm, 2*cm, 2.5*cm])
        t.setStyle(_table_style(left_align_col=2))
        story.append(t)
        story += [Spacer(1, 0.4*cm), Paragraph(f"<b>Total facturado: ${ar_money(total_gral)}</b>", getSampleStyleSheet()['Normal'])]
    story += [Spacer(1, 0.3*cm), Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# REPORTES — generación de Excel (openpyxl)
# ---------------------------------------------------------------------------

def _xl_fill(hex_color):
    """Relleno de un color sólido (hexadecimal) para celdas de Excel."""
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')


def _xl_header(ws, row, cols, color='16493D'):
    """Escribe una fila de encabezados con el estilo del sistema."""
    fill = _xl_fill(color)
    font = Font(bold=True, color='FFFFFF', size=10)
    brd = Border(bottom=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'))
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.border = brd
        cell.alignment = Alignment(horizontal='center', vertical='center')


def excel_ventas(year, month=None):
    """Reporte de ventas por producto en Excel (.xlsx)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    tipo = "Mensual" if month else "Anual"
    cat_data = build_ventas_data(year, month)
    alt = _xl_fill("F7F7F7")

    if not cat_data:
        ws = wb.create_sheet("Sin datos")
        ws['A1'] = "No hay ventas registradas para el período seleccionado."
    else:
        for cat, items in cat_data.items():
            ws = wb.create_sheet(cat[:31])
            ws.merge_cells('A1:F1')
            ws['A1'] = f"ALMACÉN DEL VALLE — {tipo} — {cat} — {period}"
            ws['A1'].font = Font(bold=True, size=12, color='16493D')
            ws['A1'].alignment = Alignment(horizontal='center')
            hdrs = ['#', 'Producto', 'Marca', 'Precio unit.', 'Unidades vendidas', 'Total $']
            for c, h in enumerate(hdrs, 1):
                ws.cell(row=3, column=c, value=h)
            _xl_header(ws, 3, len(hdrs))
            for i, it in enumerate(items, 1):
                r = i + 3
                vals = [i, it['name'], it['brand'], it['price'], it['qty'], it['total']]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.alignment = Alignment(horizontal='left' if c == 2 else 'center')
                    if i % 2 == 0:
                        cell.fill = alt
            for col, w in zip('ABCDEF', [5, 34, 16, 14, 18, 14]):
                ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _excel_simple(filename_title, headers, rows_data, col_widths, extra_footer=None):
    """Arma un libro .xlsx de una hoja (título, encabezados, filas y pie opcional) y lo devuelve en memoria."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = filename_title[:31]
    ncols = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws['A1'] = f"ALMACÉN DEL VALLE — {filename_title}"
    ws['A1'].font = Font(bold=True, size=12, color='16493D')
    ws['A1'].alignment = Alignment(horizontal='center')
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    _xl_header(ws, 3, ncols)
    alt = _xl_fill("F7F7F7")
    for i, row in enumerate(rows_data, 1):
        r = i + 3
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if i % 2 == 0:
                cell.fill = alt
    if extra_footer:
        r = len(rows_data) + 5
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        cell = ws.cell(row=r, column=1, value=extra_footer)
        cell.font = Font(bold=True, size=10, color='16493D')
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def excel_stock_bajo():
    """Reporte Excel (.xlsx) de los productos con stock bajo."""
    data = build_stock_bajo_data()
    rows = [[i, it['name'], it['brand'] or '-', it['category'] or '-', it['stock'], it['price']]
            for i, it in enumerate(data, 1)]
    return _excel_simple("Stock Bajo", ['#', 'Producto', 'Marca', 'Categoría', 'Stock', 'Precio $'],
                          rows, [5, 34, 16, 16, 10, 14])


def excel_stock_general():
    """Reporte Excel (.xlsx) con el stock de todos los productos."""
    data = build_stock_general_data()
    rows = [[i, it['name'], it['brand'] or '-', it['category'] or '-', it['stock'],
             it['price'], 'Activo' if it['active'] else 'Inactivo'] for i, it in enumerate(data, 1)]
    return _excel_simple("Stock General", ['#', 'Producto', 'Marca', 'Categoría', 'Stock', 'Precio $', 'Estado'],
                          rows, [5, 32, 16, 16, 10, 14, 12])


def excel_valor_stock():
    """Reporte Excel (.xlsx) con el valor del inventario."""
    data = build_valor_stock_data()
    total = sum(r['valor'] for r in data)
    rows = [[i, it['name'], it['category'] or '-', it['stock'], it['price'], it['valor']]
            for i, it in enumerate(data, 1)]
    return _excel_simple("Valorización de Stock", ['#', 'Producto', 'Categoría', 'Stock', 'Precio $', 'Valor total $'],
                          rows, [5, 34, 16, 10, 14, 16], extra_footer=f"Valor total del inventario: ${ar_money(total)}")


def excel_facturacion(year, month=None):
    """Reporte Excel (.xlsx) de las facturas emitidas."""
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    data = build_facturacion_data(year, month)
    rows = [[f['numero'], f['created_at'][:10], f['cliente_nombre'], f['subtotal'], f['iva'], f['total']]
            for f in data]
    total = sum(f['total'] for f in data)
    return _excel_simple(f"Facturación {period}", ['N° Factura', 'Fecha', 'Cliente', 'Subtotal $', 'IVA $', 'Total $'],
                          rows, [16, 14, 30, 14, 12, 14], extra_footer=f"Total facturado: ${ar_money(total)}")


# ---------------------------------------------------------------------------
# REPORTES — rutas de descarga
# ---------------------------------------------------------------------------

def _send_pdf(buf, nombre):
    """Responde con un PDF para descargar."""
    return send_file(buf, as_attachment=True, download_name=f"{nombre}.pdf", mimetype='application/pdf')


def _send_excel(buf, nombre):
    """Responde con un archivo Excel (.xlsx) para descargar."""
    return send_file(buf, as_attachment=True, download_name=f"{nombre}.xlsx",
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/reporte/ventas/mensual/<fmt>')
@permiso_required('reportes')
def reporte_ventas_mensual(fmt):
    """Descarga el reporte de ventas del mes, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    month = int(request.args.get('month', datetime.now().month))
    nombre = f"AlmacenDelValle_Ventas_Mensual_{MESES_ES[month-1]}_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_ventas(year, month), nombre)
    return _send_excel(excel_ventas(year, month), nombre)


@app.route('/admin/reporte/ventas/anual/<fmt>')
@permiso_required('reportes')
def reporte_ventas_anual(fmt):
    """Descarga el reporte de ventas del año, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    nombre = f"AlmacenDelValle_Ventas_Anual_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_ventas(year), nombre)
    return _send_excel(excel_ventas(year), nombre)


@app.route('/admin/reporte/stock-bajo/<fmt>')
@permiso_required('reportes')
def reporte_stock_bajo(fmt):
    """Descarga el reporte de stock bajo, en PDF o Excel (.xlsx) según fmt."""
    nombre = f"AlmacenDelValle_StockBajo_{datetime.now().strftime('%d_%m_%Y')}"
    if fmt == 'pdf':
        return _send_pdf(pdf_stock_bajo(), nombre)
    return _send_excel(excel_stock_bajo(), nombre)


@app.route('/admin/reporte/stock-general/<fmt>')
@permiso_required('reportes')
def reporte_stock_general(fmt):
    """Descarga el reporte de stock general, en PDF o Excel (.xlsx) según fmt."""
    nombre = f"AlmacenDelValle_StockGeneral_{datetime.now().strftime('%d_%m_%Y')}"
    if fmt == 'pdf':
        return _send_pdf(pdf_stock_general(), nombre)
    return _send_excel(excel_stock_general(), nombre)


@app.route('/admin/reporte/valor-stock/<fmt>')
@permiso_required('reportes')
def reporte_valor_stock(fmt):
    """Descarga el reporte de valor del inventario, en PDF o Excel (.xlsx) según fmt."""
    nombre = f"AlmacenDelValle_ValorStock_{datetime.now().strftime('%d_%m_%Y')}"
    if fmt == 'pdf':
        return _send_pdf(pdf_valor_stock(), nombre)
    return _send_excel(excel_valor_stock(), nombre)


@app.route('/admin/reporte/facturacion/mensual/<fmt>')
@permiso_required('reportes')
def reporte_facturacion_mensual(fmt):
    """Descarga el reporte de facturación del mes, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    month = int(request.args.get('month', datetime.now().month))
    nombre = f"AlmacenDelValle_Facturacion_Mensual_{MESES_ES[month-1]}_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_facturacion(year, month), nombre)
    return _send_excel(excel_facturacion(year, month), nombre)


@app.route('/admin/reporte/facturacion/anual/<fmt>')
@permiso_required('reportes')
def reporte_facturacion_anual(fmt):
    """Descarga el reporte de facturación del año, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    nombre = f"AlmacenDelValle_Facturacion_Anual_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_facturacion(year), nombre)
    return _send_excel(excel_facturacion(year), nombre)


# ---------------------------------------------------------------------------
# ADMIN — login
# ---------------------------------------------------------------------------

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    """Ingreso al panel de administración con usuario y contraseña."""
    if request.method == 'POST':
        usuario  = request.form.get('usuario', '').strip()
        password = request.form.get('password', '')
        with get_db() as db:
            row = db.execute("SELECT * FROM usuarios WHERE usuario=? AND activo=1", (usuario,)).fetchone()
            ingreso_ok, mensaje, bloqueado = procesar_intento_de_ingreso(
                db, 'admin', usuario, row['id'] if row else None,
                lambda: check_password_hash(row['password'], password) if row else False,
                'Usuario o contraseña incorrectos.')
        if ingreso_ok:
            session['admin_logged_in'] = True
            session['admin_nombre'] = f"{row['nombre']} {row['apellido']}".strip()
            session['admin_id'] = row['id']
            session['es_dueno'] = bool(row['es_dueno'])
            session['permisos'] = permisos_de_usuario(row)
            registrar_auditoria('Inicio de sesión en el panel')
            return redirect(url_for('admin'))
        flash(mensaje, 'error')
        return render_template('admin_login.html', mostrar_clave_inicial=not USAR_HTTPS), 429 if bloqueado else 200
    return render_template('admin_login.html', mostrar_clave_inicial=not USAR_HTTPS)


@app.route('/admin/recuperar-clave', methods=['GET', 'POST'])
def admin_recuperar_clave():
    """Recuperación de contraseña del panel: el enlace va al email del usuario; si el dueño
    no tiene uno cargado, al email de avisos de Datos del Negocio."""
    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip()

        def buscar(db):
            fila = db.execute("SELECT id, nombre, email, es_dueno FROM usuarios WHERE usuario=? AND activo=1",
                              (usuario,)).fetchone()
            if not fila:
                return None
            email = (fila['email'] or '').strip()
            if not email and fila['es_dueno']:
                email = get_config('smtp_email_destino', '').strip()
            return {'id': fila['id'], 'nombre': fila['nombre'], 'email': email}

        pedir_recuperacion('admin', buscar)
        return redirect(url_for('admin_recuperar_clave'))
    return render_template('admin_clave.html', modo='pedir')


@app.route('/admin/restablecer-clave/<token>', methods=['GET', 'POST'])
def admin_restablecer_clave(token):
    """Pantalla a la que llega el usuario del panel desde el correo para elegir su contraseña nueva."""
    with get_db() as db:
        registro_token = buscar_token_de_recuperacion(db, 'admin', token)
    if not registro_token:
        return render_template('admin_clave.html', modo='vencido'), 410

    if request.method == 'POST':
        nueva, repetida = request.form.get('password', ''), request.form.get('password2', '')
        if len(nueva) < 6:
            flash('La contraseña debe tener al menos 6 caracteres.', 'error')
        elif nueva != repetida:
            flash('Las dos contraseñas no coinciden.', 'error')
        else:
            with get_db() as db:
                usuario = db.execute("SELECT id, nombre, usuario, email, es_dueno FROM usuarios WHERE id=?",
                                     (registro_token['usuario_id'],)).fetchone()
                db.execute("UPDATE usuarios SET password=? WHERE id=?",
                           (generate_password_hash(nueva), usuario['id']))
                db.execute("UPDATE recuperacion_clave SET usado=1 WHERE id=?", (registro_token['id'],))
                limpiar_intentos(db, f"admin:id:{usuario['id']}")
                registrar_auditoria('Contraseña del panel cambiada por correo', usuario['usuario'], db=db)
            destino = (usuario['email'] or '').strip() or get_config('smtp_email_destino', '').strip()
            if destino:
                avisar_cambio_de_clave(destino, usuario['nombre'])
            flash('Listo, cambiaste tu contraseña. Ya podés ingresar.', 'success')
            return redirect(url_for('admin_login'))
    return render_template('admin_clave.html', modo='nueva', token=token)


@app.route('/admin/logout')
def admin_logout():
    """Cierra la sesión del panel de administración."""
    session.pop('admin_logged_in', None)
    session.pop('admin_nombre', None)
    session.pop('admin_id', None)
    session.pop('es_dueno', None)
    session.pop('permisos', None)
    return redirect(url_for('admin_login'))


# ---------------------------------------------------------------------------
# ADMIN — dashboard
# ---------------------------------------------------------------------------

@app.route('/admin')
@admin_required(lambda: redirect(url_for('admin_login')))
def admin():
    """Panel de administración: reúne los datos que muestran todas sus secciones."""
    with get_db() as db:
        categories = [dict(r) for r in db.execute("SELECT * FROM categories ORDER BY name").fetchall()]
        products = [dict(r) for r in db.execute("""
            SELECT p.*, c.name AS category_name, c.color AS category_color
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            ORDER BY p.created_at DESC
        """).fetchall()]
        pedidos = [dict(r) for r in db.execute("""
            SELECT p.*, cl.nombre, cl.apellido, cl.email AS cliente_email, f.numero AS factura_numero,
                   (SELECT ce.estado FROM correos_enviados ce WHERE ce.pedido_id = p.id AND ce.tipo = 'factura'
                    ORDER BY ce.id DESC LIMIT 1) AS factura_correo
            FROM pedidos p
            JOIN clientes cl ON cl.id = p.cliente_id
            LEFT JOIN facturas f ON f.pedido_id = p.id
            ORDER BY p.created_at DESC
        """).fetchall()]
        clientes = [dict(r) for r in db.execute("SELECT * FROM clientes ORDER BY created_at DESC").fetchall()]

        ofertas = [dict(r) for r in db.execute("""
            SELECT o.*, p.name AS producto_nombre, c.name AS categoria_nombre
            FROM ofertas_semana o
            LEFT JOIN products p ON p.id = o.producto_id
            LEFT JOIN categories c ON c.id = o.categoria_id
            ORDER BY o.orden
        """).fetchall()]

        usuarios_admin = [dict(r) for r in db.execute("SELECT * FROM usuarios ORDER BY es_dueno DESC, nombre").fetchall()]

        gastos_mes_total = db.execute("""
            SELECT COALESCE(SUM(monto),0) FROM gastos
            WHERE year=strftime('%Y','now') AND month=strftime('%m','now')
        """).fetchone()[0]
        gastos_recientes = build_gastos_data(datetime.now().year, datetime.now().month)
        for g in gastos_recientes:
            g['quien'] = _quien_gasto(g)

        horarios = [dict(r) for r in db.execute("SELECT * FROM horarios_atencion ORDER BY orden").fetchall()]
        direcciones = [dict(r) for r in db.execute("SELECT * FROM direcciones ORDER BY orden").fetchall()]
        contactos = [dict(r) for r in db.execute("SELECT * FROM contactos ORDER BY orden").fetchall()]
        redes = [dict(r) for r in db.execute("SELECT * FROM redes_sociales ORDER BY orden").fetchall()]

        ventas_mes = db.execute("""
            SELECT COALESCE(SUM(total),0) FROM pedidos
            WHERE estado != 'cancelado' AND strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now')
        """).fetchone()[0]
        pedidos_mes = db.execute("""
            SELECT COUNT(*) FROM pedidos WHERE strftime('%Y-%m', created_at) = strftime('%Y-%m', 'now')
        """).fetchone()[0]
        productos_activos = db.execute("SELECT COUNT(*) FROM products WHERE active=1").fetchone()[0]
        total_clientes = db.execute("SELECT COUNT(*) FROM clientes").fetchone()[0]

        ventas_por_mes = db.execute("""
            SELECT strftime('%Y-%m', created_at) AS mes, SUM(total) AS total
            FROM pedidos WHERE estado != 'cancelado'
            GROUP BY mes ORDER BY mes DESC LIMIT 12
        """).fetchall()
        ventas_por_categoria = db.execute("""
            SELECT c.name, SUM(pi.subtotal) AS total
            FROM pedido_items pi
            JOIN products p ON p.id = pi.product_id
            JOIN categories c ON c.id = p.category_id
            GROUP BY c.name ORDER BY total DESC
        """).fetchall()

        stock_bajo = [dict(r) for r in db.execute("""
            SELECT p.*, c.name AS category_name, c.color AS category_color
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            WHERE p.stock > 0 AND p.stock <= ?
            ORDER BY p.stock ASC
        """, (LOW_STOCK_THRESHOLD,)).fetchall()]

        sin_stock = [dict(r) for r in db.execute("""
            SELECT p.*, c.name AS category_name, c.color AS category_color
            FROM products p LEFT JOIN categories c ON c.id = p.category_id
            WHERE p.stock <= 0
            ORDER BY p.name
        """).fetchall()]

        cupones = [dict(r) for r in db.execute("SELECT * FROM cupones ORDER BY created_at DESC").fetchall()]

        historial_precios = [dict(r) for r in db.execute("""
            SELECT ph.*, p.name AS product_name
            FROM price_history ph JOIN products p ON p.id = ph.product_id
            ORDER BY ph.changed_at DESC LIMIT 100
        """).fetchall()]

        auditoria_log = [dict(r) for r in db.execute(
            "SELECT * FROM auditoria ORDER BY created_at DESC LIMIT 100").fetchall()]

        pedidos_pendientes = db.execute(
            "SELECT COUNT(*) FROM pedidos WHERE estado='pendiente_pago'").fetchone()[0]

        chatbot_respuestas = [dict(r) for r in db.execute(
            "SELECT * FROM chatbot_respuestas ORDER BY created_at DESC").fetchall()]
        # Consultas que el bot no supo responder, agrupadas por texto, las más repetidas primero
        chatbot_pendientes = [dict(r) for r in db.execute("""
            SELECT MIN(id) AS id, mensaje, COUNT(*) AS veces, MAX(created_at) AS ultima
            FROM chatbot_mensajes WHERE resuelto=0 AND descartado=0
            GROUP BY LOWER(mensaje) ORDER BY veces DESC, ultima DESC LIMIT 50
        """).fetchall()]
        fila = db.execute("SELECT COUNT(*) AS total, COALESCE(SUM(resuelto), 0) AS resueltas FROM chatbot_mensajes").fetchone()
        valoradas = db.execute("SELECT COUNT(*) AS n, COALESCE(SUM(util), 0) AS buenas FROM chatbot_mensajes WHERE util IS NOT NULL").fetchone()
        chatbot_stats = {'total': fila['total'],
                         'porcentaje': round(100 * fila['resueltas'] / fila['total']) if fila['total'] else 0,
                         'valoradas': valoradas['n'],
                         'utiles': round(100 * valoradas['buenas'] / valoradas['n']) if valoradas['n'] else 0}
        # Frases que Alice aprendió sola, con el significado legible de cada una
        etiquetas_resp = {r['id']: r['pregunta'] for r in chatbot_respuestas}
        chatbot_aprendido = []
        for f in db.execute("""SELECT * FROM chatbot_aprendido
                               ORDER BY CASE estado WHEN 'pendiente' THEN 0 WHEN 'activa' THEN 1 ELSE 2 END,
                               created_at DESC LIMIT 100""").fetchall():
            f = dict(f)
            tipo, _, valor = f['destino'].partition(':')
            f['significado'] = (etiquetas_resp.get(int(valor), 'Respuesta eliminada') if tipo == 'resp' and valor.isdigit()
                                else f"Tema integrado: {valor}")
            chatbot_aprendido.append(f)

        config_envio = {
            'envio_gratis_desde': get_config('envio_gratis_desde', '50000'),
            'costo_envio': get_config('costo_envio', '2500'),
        }

        config_notif = {
            'smtp_activo': get_config('smtp_activo', '0'),
            'smtp_host': get_config('smtp_host', ''),
            'smtp_puerto': get_config('smtp_puerto', '587'),
            'smtp_usuario': get_config('smtp_usuario', ''),
            'smtp_email_destino': get_config('smtp_email_destino', ''),
        }
        config_arca = config_facturacion()

    kpis = {
        'ventas_mes': ventas_mes, 'pedidos_mes': pedidos_mes,
        'productos_activos': productos_activos, 'total_clientes': total_clientes,
    }
    charts = {
        'meses': [r['mes'] for r in reversed(ventas_por_mes)],
        'ventas': [r['total'] for r in reversed(ventas_por_mes)],
        'cat_labels': [r['name'] for r in ventas_por_categoria],
        'cat_valores': [r['total'] for r in ventas_por_categoria],
    }

    return render_template('admin.html', categories=categories, products=products,
                            pedidos=pedidos, clientes=clientes, kpis=kpis, charts=charts,
                            stock_bajo=stock_bajo, sin_stock=sin_stock,
                            low_stock_threshold=LOW_STOCK_THRESHOLD,
                            anio_actual=datetime.now().year, mes_actual=datetime.now().month,
                            meses_es=MESES_ES, ofertas=ofertas, usuarios_admin=usuarios_admin,
                            permisos_disponibles=PERMISOS_DISPONIBLES,
                            gastos_mes_total=gastos_mes_total, gastos_recientes=gastos_recientes,
                            condiciones_iva=['Responsable Inscripto', 'Monotributista', 'Exento', 'Consumidor Final'],
                            horarios=horarios, direcciones=direcciones, contactos=contactos, redes=redes,
                            cupones=cupones, historial_precios=historial_precios, auditoria_log=auditoria_log,
                            pedidos_pendientes=pedidos_pendientes, config_envio=config_envio,
                            config_notif=config_notif, config_arca=config_arca, modos_arca=MODOS_ARCA,
                            condiciones_facturacion=CONDICIONES_IVA, chatbot_respuestas=chatbot_respuestas, chatbot_aprendido=chatbot_aprendido,
                            chatbot_pendientes=chatbot_pendientes, chatbot_stats=chatbot_stats)


# --- Productos ---

@app.route('/admin/productos/agregar', methods=['POST'])
@permiso_required('productos')
def admin_producto_agregar():
    """Crea un producto nuevo."""
    pid = uuid.uuid4().hex
    features_raw = request.form.get('info_adicional', '').strip()
    info_list = info_list_to_db(features_raw.split('\n'))
    image_url = request.form.get('image_url', '').strip()

    with get_db() as db:
        cat = db.execute("SELECT icon FROM categories WHERE id=?", (request.form.get('category_id'),)).fetchone()
        db.execute("""
            INSERT INTO products (id, code, name, category_id, brand, unidad, price, old_price,
                                   stock, description, info_adicional, icon, image, destacado, active)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)
        """, (pid, request.form.get('code', '').strip() or pid[:8].upper(),
              request.form.get('name', '').strip(), request.form.get('category_id'),
              request.form.get('brand', '').strip(), request.form.get('unidad', 'unidad'),
              money(request.form.get('price')),
              money(request.form.get('old_price')) or None,
              int(request.form.get('stock', 0) or 0),
              request.form.get('description', '').strip(),
              info_list,
              cat['icon'] if cat else '📦',
              image_url,
              1 if request.form.get('destacado') else 0))
    registrar_auditoria('Producto creado', request.form.get('name', '').strip())
    return redirect(url_for('admin'))


@app.route('/admin/productos/editar/<product_id>', methods=['POST'])
@permiso_required('productos')
def admin_producto_editar(product_id):
    """Edita un producto y registra el cambio de precio en el historial."""
    with get_db() as db:
        row = db.execute("SELECT price FROM products WHERE id=?", (product_id,)).fetchone()
        if not row:
            return redirect(url_for('admin'))
        nuevo_precio = money(request.form.get('price'))
        if row['price'] != nuevo_precio:
            db.execute("INSERT INTO price_history (product_id, price_old, price_new) VALUES (?,?,?)",
                       (product_id, row['price'], nuevo_precio))

        info_list = info_list_to_db(request.form.get('info_adicional', '').split('\n'))

        db.execute("""
            UPDATE products SET code=?, name=?, category_id=?, brand=?, unidad=?, price=?, old_price=?,
                                 stock=?, description=?, info_adicional=?, image=?, destacado=?, active=?
            WHERE id=?
        """, (request.form.get('code', '').strip(), request.form.get('name', '').strip(),
              request.form.get('category_id'), request.form.get('brand', '').strip(),
              request.form.get('unidad', 'unidad'), nuevo_precio,
              money(request.form.get('old_price')) or None,
              int(request.form.get('stock', 0) or 0),
              request.form.get('description', '').strip(),
              info_list, request.form.get('image_url', '').strip(),
              1 if request.form.get('destacado') else 0,
              1 if request.form.get('active') else 0,
              product_id))
    registrar_auditoria('Producto editado', request.form.get('name', '').strip())
    return redirect(url_for('admin'))


@app.route('/admin/productos/eliminar/<product_id>', methods=['POST'])
@permiso_required('productos')
def admin_producto_eliminar(product_id):
    """Elimina un producto."""
    with get_db() as db:
        row = db.execute("SELECT name FROM products WHERE id=?", (product_id,)).fetchone()
        db.execute("DELETE FROM products WHERE id=?", (product_id,))
    registrar_auditoria('Producto eliminado', row['name'] if row else product_id)
    return redirect(url_for('admin'))


@app.route('/admin/productos/<product_id>/destacar', methods=['POST'])
@permiso_required('productos')
def admin_producto_destacar(product_id):
    """Marca o desmarca un producto como destacado."""
    with get_db() as db:
        row = db.execute("SELECT destacado, name FROM products WHERE id=?", (product_id,)).fetchone()
        if not row:
            return jsonify({'error': 'no encontrado'}), 404
        nuevo = 0 if row['destacado'] else 1
        db.execute("UPDATE products SET destacado=? WHERE id=?", (nuevo, product_id))
        registrar_auditoria('Producto destacado' if nuevo else 'Producto ya no destacado', row['name'], db=db)
    return jsonify({'ok': True, 'destacado': nuevo})


# --- Categorías ---

@app.route('/admin/categorias/agregar', methods=['POST'])
@permiso_required('categorias')
def admin_categoria_agregar():
    """Crea una categoría, con un color de la paleta."""
    nombre = request.form.get('name', '').strip()
    cid = nombre.lower().replace(' ', '')
    paleta = ['#DCEDE6', '#F4E3C6', '#E4EEF7', '#F5DADA', '#E1F0FB', '#F1E3EC', '#E6F3EC', '#E9F3D9']
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM categories").fetchone()[0]
        db.execute("INSERT OR IGNORE INTO categories (id, name, icon, color) VALUES (?,?,?,?)",
                   (cid, nombre, request.form.get('icon', '📦'), paleta[n % len(paleta)]))
    registrar_auditoria('Categoría creada', nombre)
    return redirect(url_for('admin'))


@app.route('/admin/categorias/eliminar/<cat_id>', methods=['POST'])
@permiso_required('categorias')
def admin_categoria_eliminar(cat_id):
    """Elimina una categoría."""
    with get_db() as db:
        db.execute("DELETE FROM categories WHERE id=?", (cat_id,))
    registrar_auditoria('Categoría eliminada', cat_id)
    return redirect(url_for('admin'))


# --- Ofertas de la semana (banner gigante en la tienda) ---

@app.route('/admin/ofertas/agregar', methods=['POST'])
@permiso_required('ofertas')
def admin_oferta_agregar():
    """Publica una oferta de la semana (banner de la portada)."""
    titulo = request.form.get('titulo', '').strip()
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM ofertas_semana").fetchone()[0]
        db.execute("""
            INSERT INTO ofertas_semana (titulo, texto, producto_id, categoria_id, color, activa, orden)
            VALUES (?,?,?,?,?,1,?)
        """, (titulo, request.form.get('texto', '').strip(),
              request.form.get('producto_id') or None, request.form.get('categoria_id') or None,
              request.form.get('color', 'verde'), n))
        registrar_auditoria('Oferta creada', titulo, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/ofertas/<int:oferta_id>/toggle', methods=['POST'])
@permiso_required('ofertas')
def admin_oferta_toggle(oferta_id):
    """Activa o desactiva una oferta de la semana."""
    with get_db() as db:
        row = db.execute("SELECT activa, titulo FROM ofertas_semana WHERE id=?", (oferta_id,)).fetchone()
        if row:
            db.execute("UPDATE ofertas_semana SET activa=? WHERE id=?", (0 if row['activa'] else 1, oferta_id))
            registrar_auditoria('Oferta actualizada', row['titulo'], db=db)
    return redirect(url_for('admin'))


@app.route('/admin/ofertas/<int:oferta_id>/eliminar', methods=['POST'])
@permiso_required('ofertas')
def admin_oferta_eliminar(oferta_id):
    """Elimina una oferta de la semana."""
    with get_db() as db:
        row = db.execute("SELECT titulo FROM ofertas_semana WHERE id=?", (oferta_id,)).fetchone()
        db.execute("DELETE FROM ofertas_semana WHERE id=?", (oferta_id,))
        registrar_auditoria('Oferta eliminada', row['titulo'] if row else str(oferta_id), db=db)
    return redirect(url_for('admin'))


# --- Etiquetas de precio (PDF imprimible, individual o global) ---

@app.route('/admin/etiquetas', methods=['GET', 'POST'])
@permiso_required('etiquetas')
def admin_etiquetas():
    """PDF de etiquetas de precio (3 columnas por 7 filas por hoja A4), de los productos elegidos o de todos."""
    ids = request.form.getlist('product_ids') if request.method == 'POST' else request.args.getlist('ids')

    with get_db() as db:
        if ids:
            ph = ','.join('?' * len(ids))
            rows = db.execute(f"SELECT * FROM products WHERE id IN ({ph}) ORDER BY name", ids).fetchall()
        else:
            rows = db.execute("SELECT * FROM products ORDER BY category_id, name").fetchall()
    products = [dict(r) for r in rows]
    if not products:
        return redirect(url_for('admin'))

    buf = io.BytesIO()
    W, H = A4
    et_w, et_h = 6.5 * cm, 3.8 * cm
    cols, rows_n = 3, 7
    margin_x = (W - cols * et_w) / 2
    margin_y = (H - rows_n * et_h) / 2

    c = pdfcanvas.Canvas(buf, pagesize=A4)
    for i, p in enumerate(products):
        if i > 0 and i % (cols * rows_n) == 0:
            c.showPage()
        pos = i % (cols * rows_n)
        col_i, row_i = pos % cols, pos // cols
        x = margin_x + col_i * et_w
        y = H - margin_y - (row_i + 1) * et_h

        c.setStrokeColor(colors.HexColor('#CCCCCC'))
        c.setLineWidth(0.5)
        c.rect(x + 2, y + 2, et_w - 4, et_h - 4)

        c.setFont('Helvetica-Bold', 6)
        c.setFillColor(VERDE_PDF)
        c.drawString(x + 6, y + et_h - 14, 'ALMACÉN DEL VALLE')

        nombre = p['name'][:34] + ('…' if len(p['name']) > 34 else '')
        c.setFont('Helvetica-Bold', 7.5)
        c.setFillColor(colors.HexColor('#111111'))
        c.drawString(x + 6, y + et_h - 26, nombre)

        sub = ' · '.join(filter(None, [p.get('brand', ''), p.get('unidad', '')]))
        if sub:
            c.setFont('Helvetica', 6.5)
            c.setFillColor(colors.HexColor('#666666'))
            c.drawString(x + 6, y + et_h - 37, sub[:38])

        c.setStrokeColor(colors.HexColor('#EEEEEE'))
        c.line(x + 6, y + et_h - 43, x + et_w - 6, y + et_h - 43)

        c.setFont('Helvetica-Bold', 15)
        c.setFillColor(VERDE_PDF)
        c.drawString(x + 6, y + 9, f"${ar_money(p['price'])}")

        if p.get('old_price'):
            c.setFont('Helvetica', 7)
            c.setFillColor(MANGO_PDF)
            c.drawString(x + 6, y + 22, f"Antes ${ar_money(p['old_price'])}")

        if p.get('code'):
            c.setFont('Helvetica', 5.5)
            c.setFillColor(colors.HexColor('#AAAAAA'))
            c.drawRightString(x + et_w - 6, y + 3, p['code'])

    c.save()
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Etiquetas_AlmacenDelValle.pdf",
                      mimetype='application/pdf')


# --- Gastos / compras ---

def build_gastos_data(year, month=None):
    """Gastos registrados en un año, o en un mes puntual si se indica."""
    with get_db() as db:
        if month:
            rows = db.execute("SELECT * FROM gastos WHERE year=? AND month=? ORDER BY id DESC",
                               (year, month)).fetchall()
        else:
            rows = db.execute("SELECT * FROM gastos WHERE year=? ORDER BY month, id", (year,)).fetchall()
    return [dict(r) for r in rows]


def _quien_gasto(g):
    """Texto de la columna "quién": el empleado si es un sueldo, o el proveedor en los demás casos."""
    if g.get('categoria') == 'Sueldos':
        return f"{g.get('nombre', '')} {g.get('apellido', '')}".strip() or '-'
    return g.get('proveedor') or '-'


@app.route('/admin/api/gastos')
@permiso_required('gastos')
def api_gastos():
    """Gastos de un año o mes, en JSON, para la tabla del panel."""
    year = int(request.args.get('year', datetime.now().year))
    month = request.args.get('month', type=int)
    data = build_gastos_data(year, month)
    for g in data:
        g['quien'] = _quien_gasto(g)
    return jsonify(data)


@app.route('/admin/api/gastos/nuevo', methods=['POST'])
@permiso_required('gastos')
def api_gasto_nuevo():
    """Registra un gasto con la fecha de hoy."""
    d = request.get_json() or {}
    now = datetime.now()
    monto = float(d.get('monto', 0) or 0)

    def texto(clave):
        """Lee un campo del JSON como texto, sin espacios de más."""
        valor = d.get(clave)
        return valor.strip() if isinstance(valor, str) else (valor or '')

    with get_db() as db:
        db.execute("""
            INSERT INTO gastos (descripcion, monto, categoria, proveedor, cuit, razon_social,
                                 domicilio_fiscal, condicion_iva, nro_factura, tipo_factura,
                                 nombre, apellido, tipo_impuesto, nota, year, month, day)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (texto('descripcion'), monto, texto('categoria'),
              texto('proveedor'), texto('cuit'), texto('razon_social'), texto('domicilio_fiscal'),
              texto('condicion_iva'), texto('nro_factura'), texto('tipo_factura'),
              texto('nombre'), texto('apellido'), texto('tipo_impuesto'), texto('nota'),
              now.year, now.month, now.day))
        gid = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        registrar_auditoria('Gasto registrado', f"{texto('categoria')} — ${ar_money(monto)}", db=db)
    return jsonify({'ok': True, 'id': gid})


@app.route('/admin/api/gastos/<int:gid>/eliminar', methods=['POST'])
@permiso_required('gastos')
def api_gasto_eliminar(gid):
    """Elimina un gasto."""
    with get_db() as db:
        db.execute("DELETE FROM gastos WHERE id=?", (gid,))
        registrar_auditoria('Gasto eliminado', f'ID {gid}', db=db)
    return jsonify({'ok': True})


def pdf_gastos(year, month=None):
    """Reporte PDF de gastos y compras."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm,
                             topMargin=1.5*cm, bottomMargin=1.8*cm)
    st = _report_styles()
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    tipo = "Mensual" if month else "Anual"
    data = build_gastos_data(year, month)
    story = [Paragraph("ALMACÉN DEL VALLE", st['title']), Paragraph(f"Reporte {tipo} de Gastos / Compras — {period}", st['sub'])]
    if not data:
        story.append(Paragraph("No hay gastos registrados para el período seleccionado.", getSampleStyleSheet()['Normal']))
    else:
        rows = [['Fecha', 'Categoría', 'Descripción', 'Proveedor / Empleado', 'Monto $']]
        total = 0
        for g in data:
            rows.append([f"{g['day']:02d}/{g['month']:02d}/{g['year']}", g['categoria'], g['descripcion'],
                         _quien_gasto(g), f"${ar_money(g['monto'])}"])
            total += g['monto']
        t = Table(rows, colWidths=[2.2*cm, 3*cm, 5.5*cm, 4*cm, 2.5*cm])
        t.setStyle(_table_style(left_align_col=2))
        story.append(t)
        story += [Spacer(1, 0.4*cm), Paragraph(f"<b>Total de gastos: ${ar_money(total)}</b>", getSampleStyleSheet()['Normal'])]
    story += [Spacer(1, 0.3*cm), Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", st['sub'])]
    doc.build(story, canvasmaker=_NumberedCanvas)
    buf.seek(0)
    return buf


def excel_gastos(year, month=None):
    """Reporte Excel (.xlsx) de gastos y compras."""
    period = f"{MESES_ES[month-1]} {year}" if month else str(year)
    data = build_gastos_data(year, month)
    rows = [[f"{g['day']:02d}/{g['month']:02d}/{g['year']}", g['categoria'], g['descripcion'],
             _quien_gasto(g), g['monto']] for g in data]
    total = sum(g['monto'] for g in data)
    return _excel_simple(f"Gastos {period}", ['Fecha', 'Categoría', 'Descripción', 'Proveedor / Empleado', 'Monto $'],
                          rows, [12, 16, 34, 26, 14], extra_footer=f"Total de gastos: ${ar_money(total)}")


@app.route('/admin/reporte/gastos/mensual/<fmt>')
@permiso_required('gastos')
def reporte_gastos_mensual(fmt):
    """Descarga el reporte de gastos del mes, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    month = int(request.args.get('month', datetime.now().month))
    nombre = f"AlmacenDelValle_Gastos_Mensual_{MESES_ES[month-1]}_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_gastos(year, month), nombre)
    return _send_excel(excel_gastos(year, month), nombre)


@app.route('/admin/reporte/gastos/anual/<fmt>')
@permiso_required('gastos')
def reporte_gastos_anual(fmt):
    """Descarga el reporte de gastos del año, en PDF o Excel (.xlsx) según fmt."""
    year = int(request.args.get('year', datetime.now().year))
    nombre = f"AlmacenDelValle_Gastos_Anual_{year}"
    if fmt == 'pdf':
        return _send_pdf(pdf_gastos(year), nombre)
    return _send_excel(excel_gastos(year), nombre)


# --- Backups de la base de datos ---

def hacer_backup_automatico():
    """Crea una copia física de la BD en la carpeta de backups."""
    n = datetime.now()
    filename = f"auto_{n.strftime('%Y%m%d_%H%M')}.db"
    dest = os.path.join(BACKUP_FOLDER, filename)
    try:
        shutil.copy2(DB_PATH, dest)
        autos = sorted(f for f in os.listdir(BACKUP_FOLDER) if f.startswith('auto_'))
        for viejo in autos[:-7]:
            try:
                os.remove(os.path.join(BACKUP_FOLDER, viejo))
            except OSError:
                pass
        with get_db() as db:
            db.execute("INSERT INTO backup_log (filename) VALUES (?)", (filename,))
    except Exception as e:
        app.logger.warning("No se pudo generar el backup: %s", e)


def _backup_scheduler():
    """Genera un backup automático cada 24 horas (el primero, un minuto después de iniciar)."""
    time.sleep(60)
    while True:
        hacer_backup_automatico()
        time.sleep(24 * 60 * 60)


@app.route('/admin/backup')
@permiso_required('backups')
def backup_db():
    """Descarga una copia inmediata del archivo de base de datos."""
    n = datetime.now()
    filename = f"AlmacenDelValle_Backup_{n.strftime('%Y%m%d_%H%M')}.db"
    buf = io.BytesIO()
    with open(DB_PATH, 'rb') as f:
        buf.write(f.read())
    buf.seek(0)
    with get_db() as db:
        db.execute("INSERT INTO backup_log (filename) VALUES (?)", (filename,))
    return send_file(buf, as_attachment=True, download_name=filename, mimetype='application/octet-stream')


@app.route('/admin/backup/ahora', methods=['POST'])
@permiso_required('backups')
def backup_ahora():
    """Genera un backup manual inmediato en la carpeta de backups del servidor."""
    hacer_backup_automatico()
    registrar_auditoria('Backup manual generado')
    return jsonify({'ok': True})


@app.route('/admin/api/backups')
@permiso_required('backups')
def api_backups():
    """Lista los backups guardados en el servidor."""
    archivos = sorted((f for f in os.listdir(BACKUP_FOLDER) if f.endswith('.db')), reverse=True)
    result = []
    for f in archivos:
        path = os.path.join(BACKUP_FOLDER, f)
        result.append({
            'filename': f,
            'size_kb': round(os.path.getsize(path) / 1024, 1),
            'created': datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d/%m/%Y %H:%M'),
        })
    return jsonify(result)


@app.route('/admin/backup/descargar/<filename>')
@permiso_required('backups')
def backup_descargar(filename):
    safe = os.path.basename(filename)
    path = os.path.join(BACKUP_FOLDER, safe)
    if not os.path.exists(path) or not safe.endswith('.db'):
        return "Archivo no encontrado", 404
    return send_file(path, as_attachment=True, download_name=f"AlmacenDelValle_{safe}",
                      mimetype='application/octet-stream')


# --- Usuarios del panel (solo el dueño puede gestionar) ---


@app.route('/admin/usuarios/agregar', methods=['POST'])
@dueno_required(lambda: redirect(url_for('admin_login')))
def agregar_usuario():
    """Crea un usuario del panel; si no se indica contraseña, se genera una al azar."""
    permisos = ','.join(request.form.getlist('permisos'))
    with get_db() as db:
        try:
            db.execute("""
                INSERT INTO usuarios (nombre, apellido, usuario, password, email, permisos)
                VALUES (?,?,?,?,?,?)
            """, (request.form.get('nombre', '').strip(), request.form.get('apellido', '').strip(),
                  request.form.get('usuario', '').strip(),
                  generate_password_hash(request.form.get('password', '').strip() or uuid.uuid4().hex[:10]),
                  request.form.get('email', '').strip(), permisos))
            registrar_auditoria('Usuario del panel creado', request.form.get('usuario', '').strip(), db=db)
        except sqlite3.IntegrityError:
            flash('Ya existe un usuario con ese nombre de usuario.', 'error')
    return redirect(url_for('admin'))


@app.route('/admin/usuarios/editar/<int:uid>', methods=['POST'])
@dueno_required(lambda: redirect(url_for('admin_login')))
def editar_usuario(uid):
    """Edita los datos y permisos de un usuario del panel; la contraseña solo cambia si se completa."""
    permisos = ','.join(request.form.getlist('permisos'))
    nueva_pass = request.form.get('password', '').strip()
    with get_db() as db:
        if nueva_pass:
            db.execute("""
                UPDATE usuarios SET nombre=?, apellido=?, email=?, permisos=?, password=?
                WHERE id=?
            """, (request.form.get('nombre', '').strip(), request.form.get('apellido', '').strip(),
                  request.form.get('email', '').strip(), permisos,
                  generate_password_hash(nueva_pass), uid))
        else:
            db.execute("""
                UPDATE usuarios SET nombre=?, apellido=?, email=?, permisos=? WHERE id=?
            """, (request.form.get('nombre', '').strip(), request.form.get('apellido', '').strip(),
                  request.form.get('email', '').strip(), permisos, uid))
        registrar_auditoria('Usuario del panel editado', request.form.get('nombre', '').strip() or f'ID {uid}', db=db)
    return redirect(url_for('admin'))


@app.route('/admin/usuarios/eliminar/<int:uid>', methods=['POST'])
@dueno_required(lambda: redirect(url_for('admin_login')))
def eliminar_usuario(uid):
    """Elimina un usuario del panel (el dueño no se puede eliminar)."""
    with get_db() as db:
        row = db.execute("SELECT es_dueno, usuario FROM usuarios WHERE id=?", (uid,)).fetchone()
        if row and not row['es_dueno']:
            db.execute("DELETE FROM usuarios WHERE id=?", (uid,))
            registrar_auditoria('Usuario del panel eliminado', row['usuario'], db=db)
    return redirect(url_for('admin'))


@app.route('/admin/usuarios/toggle/<int:uid>', methods=['POST'])
@dueno_required(lambda: redirect(url_for('admin_login')))
def toggle_usuario(uid):
    """Activa o desactiva un usuario del panel (el dueño no se puede desactivar)."""
    with get_db() as db:
        row = db.execute("SELECT activo, es_dueno, usuario FROM usuarios WHERE id=?", (uid,)).fetchone()
        if row and not row['es_dueno']:
            nuevo_activo = 0 if row['activo'] else 1
            db.execute("UPDATE usuarios SET activo=? WHERE id=?", (nuevo_activo, uid))
            registrar_auditoria('Usuario activado' if nuevo_activo else 'Usuario desactivado', row['usuario'], db=db)
    return redirect(url_for('admin'))


# --- Datos del negocio: horarios, direcciones, contactos, redes sociales ---

@app.route('/admin/horarios/agregar', methods=['POST'])
@permiso_required('negocio')
def horario_agregar():
    """Agrega un horario de atención."""
    dias = request.form.get('dias', '').strip()
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM horarios_atencion").fetchone()[0]
        db.execute("INSERT INTO horarios_atencion (dias, hora_inicio, hora_fin, orden) VALUES (?,?,?,?)",
                   (dias, request.form.get('hora_inicio', '').strip(),
                    request.form.get('hora_fin', '').strip(), n))
        registrar_auditoria('Horario agregado', dias, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/horarios/editar/<int:hid>', methods=['POST'])
@permiso_required('negocio')
def horario_editar(hid):
    """Edita un horario de atención."""
    dias = request.form.get('dias', '').strip()
    with get_db() as db:
        db.execute("UPDATE horarios_atencion SET dias=?, hora_inicio=?, hora_fin=? WHERE id=?",
                   (dias, request.form.get('hora_inicio', '').strip(),
                    request.form.get('hora_fin', '').strip(), hid))
        registrar_auditoria('Horario editado', dias, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/horarios/eliminar/<int:hid>', methods=['POST'])
@permiso_required('negocio')
def horario_eliminar(hid):
    """Elimina un horario de atención."""
    with get_db() as db:
        row = db.execute("SELECT dias FROM horarios_atencion WHERE id=?", (hid,)).fetchone()
        db.execute("DELETE FROM horarios_atencion WHERE id=?", (hid,))
        registrar_auditoria('Horario eliminado', row['dias'] if row else str(hid), db=db)
    return redirect(url_for('admin'))


@app.route('/admin/direcciones/agregar', methods=['POST'])
@permiso_required('negocio')
def direccion_agregar():
    """Agrega una dirección o sucursal."""
    nombre = request.form.get('nombre', '').strip() or 'Sucursal'
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM direcciones").fetchone()[0]
        db.execute("""
            INSERT INTO direcciones (nombre, direccion, ciudad, provincia, orden) VALUES (?,?,?,?,?)
        """, (nombre, request.form.get('direccion', '').strip(),
              request.form.get('ciudad', '').strip(), request.form.get('provincia', '').strip(), n))
        registrar_auditoria('Dirección agregada', nombre, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/direcciones/editar/<int:did>', methods=['POST'])
@permiso_required('negocio')
def direccion_editar(did):
    """Edita una dirección o sucursal."""
    nombre = request.form.get('nombre', '').strip() or 'Sucursal'
    with get_db() as db:
        db.execute("""
            UPDATE direcciones SET nombre=?, direccion=?, ciudad=?, provincia=? WHERE id=?
        """, (nombre, request.form.get('direccion', '').strip(),
              request.form.get('ciudad', '').strip(), request.form.get('provincia', '').strip(), did))
        registrar_auditoria('Dirección editada', nombre, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/direcciones/eliminar/<int:did>', methods=['POST'])
@permiso_required('negocio')
def direccion_eliminar(did):
    """Elimina una dirección o sucursal."""
    with get_db() as db:
        row = db.execute("SELECT nombre FROM direcciones WHERE id=?", (did,)).fetchone()
        db.execute("DELETE FROM direcciones WHERE id=?", (did,))
        registrar_auditoria('Dirección eliminada', row['nombre'] if row else str(did), db=db)
    return redirect(url_for('admin'))


@app.route('/admin/contactos/agregar', methods=['POST'])
@permiso_required('negocio')
def contacto_agregar():
    """Agrega un teléfono o medio de contacto."""
    valor = request.form.get('valor', '').strip()
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM contactos").fetchone()[0]
        db.execute("INSERT INTO contactos (tipo, etiqueta, valor, orden) VALUES (?,?,?,?)",
                   (request.form.get('tipo', 'Teléfono'), request.form.get('etiqueta', '').strip(), valor, n))
        registrar_auditoria('Contacto agregado', valor, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/contactos/editar/<int:cid>', methods=['POST'])
@permiso_required('negocio')
def contacto_editar(cid):
    """Edita un teléfono o medio de contacto."""
    valor = request.form.get('valor', '').strip()
    with get_db() as db:
        db.execute("UPDATE contactos SET tipo=?, etiqueta=?, valor=? WHERE id=?",
                   (request.form.get('tipo', 'Teléfono'), request.form.get('etiqueta', '').strip(), valor, cid))
        registrar_auditoria('Contacto editado', valor, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/contactos/eliminar/<int:cid>', methods=['POST'])
@permiso_required('negocio')
def contacto_eliminar(cid):
    """Elimina un teléfono o medio de contacto."""
    with get_db() as db:
        row = db.execute("SELECT valor FROM contactos WHERE id=?", (cid,)).fetchone()
        db.execute("DELETE FROM contactos WHERE id=?", (cid,))
        registrar_auditoria('Contacto eliminado', row['valor'] if row else str(cid), db=db)
    return redirect(url_for('admin'))


@app.route('/admin/redes/agregar', methods=['POST'])
@permiso_required('negocio')
def red_agregar():
    """Agrega una red social."""
    plataforma = request.form.get('plataforma', '').strip()
    with get_db() as db:
        n = db.execute("SELECT COUNT(*) FROM redes_sociales").fetchone()[0]
        db.execute("INSERT INTO redes_sociales (plataforma, url, orden) VALUES (?,?,?)",
                   (plataforma, request.form.get('url', '').strip(), n))
        registrar_auditoria('Red social agregada', plataforma, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/redes/editar/<int:rid>', methods=['POST'])
@permiso_required('negocio')
def red_editar(rid):
    """Edita una red social."""
    plataforma = request.form.get('plataforma', '').strip()
    with get_db() as db:
        db.execute("UPDATE redes_sociales SET plataforma=?, url=? WHERE id=?",
                   (plataforma, request.form.get('url', '').strip(), rid))
        registrar_auditoria('Red social editada', plataforma, db=db)
    return redirect(url_for('admin'))


@app.route('/admin/redes/eliminar/<int:rid>', methods=['POST'])
@permiso_required('negocio')
def red_eliminar(rid):
    """Elimina una red social."""
    with get_db() as db:
        row = db.execute("SELECT plataforma FROM redes_sociales WHERE id=?", (rid,)).fetchone()
        db.execute("DELETE FROM redes_sociales WHERE id=?", (rid,))
        registrar_auditoria('Red social eliminada', row['plataforma'] if row else str(rid), db=db)
    return redirect(url_for('admin'))


# --- Cupones de descuento ---

@app.route('/admin/cupones/agregar', methods=['POST'])
@permiso_required('cupones')
def cupon_agregar():
    """Crea un cupón de descuento."""
    with get_db() as db:
        try:
            db.execute("""
                INSERT INTO cupones (codigo, tipo, valor, monto_minimo, usos_maximos, fecha_expiracion, activo)
                VALUES (?,?,?,?,?,?,1)
            """, (request.form.get('codigo', '').strip().upper(), request.form.get('tipo', 'porcentaje'),
                  money(request.form.get('valor')), money(request.form.get('monto_minimo')) or 0,
                  int(request.form.get('usos_maximos', 0) or 0), request.form.get('fecha_expiracion', '')))
            registrar_auditoria('Cupón creado', request.form.get('codigo', '').strip().upper(), db=db)
        except sqlite3.IntegrityError:
            flash('Ya existe un cupón con ese código.', 'error')
    return redirect(url_for('admin'))


@app.route('/admin/cupones/<int:cid>/toggle', methods=['POST'])
@permiso_required('cupones')
def cupon_toggle(cid):
    """Activa o desactiva un cupón."""
    with get_db() as db:
        row = db.execute("SELECT activo, codigo FROM cupones WHERE id=?", (cid,)).fetchone()
        if row:
            db.execute("UPDATE cupones SET activo=? WHERE id=?", (0 if row['activo'] else 1, cid))
            registrar_auditoria('Cupón actualizado', row['codigo'], db=db)
    return redirect(url_for('admin'))


@app.route('/admin/cupones/<int:cid>/eliminar', methods=['POST'])
@permiso_required('cupones')
def cupon_eliminar(cid):
    """Elimina un cupón."""
    with get_db() as db:
        row = db.execute("SELECT codigo FROM cupones WHERE id=?", (cid,)).fetchone()
        db.execute("DELETE FROM cupones WHERE id=?", (cid,))
        if row:
            registrar_auditoria('Cupón eliminado', row['codigo'], db=db)
    return redirect(url_for('admin'))


# --- Configuración general (envíos) ---

@app.route('/admin/configuracion/guardar', methods=['POST'])
@permiso_required('negocio')
def configuracion_guardar():
    """Guarda el costo de envío y el monto desde el cual el envío es gratis."""
    set_config('envio_gratis_desde', request.form.get('envio_gratis_desde', '50000'))
    set_config('costo_envio', request.form.get('costo_envio', '2500'))
    registrar_auditoria('Configuración de envío actualizada')
    flash('Configuración de envío actualizada.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/facturacion/guardar', methods=['POST'])
@permiso_required('facturacion')
def facturacion_guardar():
    """Guarda los datos del emisor que salen en las facturas."""
    cuit_ingresado = request.form.get('cuit', '').strip()
    cuit = normalizar_cuit(cuit_ingresado)
    condicion = request.form.get('condicion_iva', '')
    punto_venta = request.form.get('punto_venta', '1').strip()
    inicio = request.form.get('inicio_actividades', '').strip()
    if cuit_ingresado and not cuit:
        flash('El CUIT no es válido: revisá los 11 dígitos.', 'error')
    elif condicion not in CONDICIONES_IVA:
        flash('Elegí la condición frente al IVA.', 'error')
    elif not punto_venta.isdigit() or not 1 <= int(punto_venta) <= 99999:
        flash('El punto de venta debe ser un número entre 1 y 99999.', 'error')
    elif inicio and not re.fullmatch(r'\d{4}-\d{2}-\d{2}', inicio):
        flash('La fecha de inicio de actividades no es válida.', 'error')
    else:
        set_config('arca_razon_social', request.form.get('razon_social', '').strip()[:100])
        set_config('arca_cuit', formatear_cuit(cuit))
        set_config('arca_condicion_iva', condicion)
        set_config('arca_punto_venta', str(int(punto_venta)))
        set_config('arca_domicilio', request.form.get('domicilio', '').strip()[:150])
        set_config('arca_iibb', request.form.get('ingresos_brutos', '').strip()[:30])
        set_config('arca_inicio_actividades', inicio)
        registrar_auditoria('Datos de facturación actualizados')
        flash('Datos de facturación guardados. Se aplican a las facturas nuevas.', 'success')
    return redirect(url_for('admin', view='facturacion'))


@app.route('/admin/transferencia/guardar', methods=['POST'])
@permiso_required('negocio')
def transferencia_guardar():
    """Guarda la cuenta bancaria que ven los clientes que pagan por transferencia."""
    cbu = re.sub(r'\s', '', request.form.get('cbu', ''))
    alias = request.form.get('alias', '').strip().upper()
    if cbu and not re.fullmatch(r'\d{22}', cbu):
        flash('El CBU debe tener 22 dígitos.', 'error')
    elif alias and not re.fullmatch(r'[A-Z0-9.\-]{6,20}', alias):
        flash('El alias debe tener entre 6 y 20 caracteres (letras, números, puntos y guiones).', 'error')
    elif not cbu and not alias:
        flash('Cargá al menos el CBU o el alias.', 'error')
    else:
        set_config('transf_banco', request.form.get('banco', '').strip()[:60])
        set_config('transf_titular', request.form.get('titular', '').strip()[:100])
        set_config('transf_cuit', request.form.get('cuit', '').strip()[:20])
        set_config('transf_cbu', cbu)
        set_config('transf_alias', alias)
        registrar_auditoria('Datos de transferencia actualizados')
        flash('Datos para transferencias guardados.', 'success')
    return redirect(url_for('admin', view='negocio'))


@app.route('/admin/notificaciones/guardar', methods=['POST'])
@permiso_required('negocio')
def notificaciones_guardar():
    """Guarda la configuración del servidor de correo (SMTP)."""
    set_config('smtp_activo', '1' if request.form.get('smtp_activo') else '0')
    set_config('smtp_host', request.form.get('smtp_host', '').strip())
    set_config('smtp_puerto', request.form.get('smtp_puerto', '587').strip())
    set_config('smtp_usuario', request.form.get('smtp_usuario', '').strip())
    if request.form.get('smtp_password', '').strip():
        set_config('smtp_password', request.form.get('smtp_password', '').strip())
    set_config('smtp_email_destino', request.form.get('smtp_email_destino', '').strip())
    registrar_auditoria('Configuración de notificaciones por email actualizada')
    flash('Configuración de notificaciones guardada.', 'success')
    return redirect(url_for('admin', view='negocio'))


@app.route('/admin/notificaciones/probar', methods=['POST'])
@permiso_required('negocio')
def notificaciones_probar():
    """Manda un correo de prueba de verdad y devuelve si salió o el motivo del error. Va al email
    del administrador cargado abajo o, si no hay, a la propia casilla del usuario SMTP."""
    cfg = leer_config_smtp()
    if not cfg:
        return jsonify({'ok': False, 'mensaje': 'Primero marcá "Activar el envío de correos", cargá el servidor y guardá.'})
    destino = cfg['destino_admin'] or (cfg['usuario'] if '@' in cfg['usuario'] else '')
    if not destino:
        return jsonify({'ok': False, 'mensaje': 'Cargá un email de destino o un usuario SMTP que sea un correo.'})
    msg = EmailMessage()
    msg['Subject'] = "Prueba de correo - Almacén del Valle"
    msg['To'] = destino
    msg.set_content("Si recibiste este correo, la tienda ya puede enviar facturas y avisos a tus clientes.")
    ok, motivo = enviar_email(msg, cfg, 'prueba', None, esperar=True)
    if ok:
        return jsonify({'ok': True, 'mensaje': f'Correo de prueba enviado a {destino}. Revisá la bandeja (y el spam).'})
    return jsonify({'ok': False, 'mensaje': motivo})


@app.route('/admin/api/notificaciones')
@permiso_required('pedidos')
def api_notificaciones_pedidos():
    """Cantidad de pedidos pendientes de pago; el panel la consulta cada tanto para avisar."""
    with get_db() as db:
        pendientes = db.execute(
            "SELECT COUNT(*) FROM pedidos WHERE estado='pendiente_pago'").fetchone()[0]
    return jsonify({'pendientes': pendientes})


# --- Historial de precios ---

@app.route('/admin/productos/exportar')
@permiso_required('productos')
def productos_exportar():
    """Exporta todos los productos a un Excel (.xlsx) con el mismo formato que acepta la importación."""
    with get_db() as db:
        rows = db.execute("""
            SELECT p.code, p.name, p.category_id, p.brand, p.unidad, p.price, p.old_price,
                   p.stock, p.description, p.info_adicional, p.image, p.destacado, p.active
            FROM products p ORDER BY p.name
        """).fetchall()
    headers = ['code', 'name', 'category_id', 'brand', 'unidad', 'price', 'old_price',
               'stock', 'description', 'info_adicional', 'image', 'destacado', 'active']
    data = [[r[h] for h in headers] for r in rows]
    buf = _excel_simple('Productos', headers, data, [12, 30, 14, 16, 10, 10, 10, 8, 26, 26, 24, 10, 8])
    registrar_auditoria('Productos exportados a Excel')
    return send_file(buf, as_attachment=True, download_name='AlmacenDelValle_Productos.xlsx',
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/productos/importar', methods=['POST'])
@permiso_required('productos')
def productos_importar():
    """Importa productos desde un Excel (.xlsx) evitando duplicados."""
    file = request.files.get('archivo')
    # openpyxl solo lee el formato moderno (.xlsx); los .xls viejos no se pueden abrir
    if not file or not file.filename.lower().endswith('.xlsx'):
        flash('Subí un archivo Excel en formato .xlsx.', 'error')
        return redirect(url_for('admin'))

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        flash('No se pudo leer el archivo: verificá que sea un Excel (.xlsx) válido.', 'error')
        return redirect(url_for('admin'))
    ws = wb.active
    filas = list(ws.iter_rows(min_row=2, values_only=True))
    creados, actualizados = 0, 0

    with get_db() as db:
        categorias_validas = {r['id'] for r in db.execute("SELECT id FROM categories").fetchall()}
        for fila in filas:
            if not fila or not fila[1]:
                continue
            code, name, category_id, brand, unidad, price, old_price, stock, description, info_adicional, image, destacado, active = (
                list(fila) + [None] * (13 - len(fila)))[:13]
            if category_id not in categorias_validas:
                continue

            # Emparejar primero por código; si no hay código o no matchea, buscar por
            # nombre exacto (sin distinguir mayúsculas) dentro de la misma categoría.
            existente = None
            if code:
                existente = db.execute("SELECT id FROM products WHERE code=?", (code,)).fetchone()
            if not existente and name:
                existente = db.execute(
                    "SELECT id FROM products WHERE LOWER(name)=LOWER(?) AND category_id=?",
                    (name.strip(), category_id)
                ).fetchone()
            if existente:
                fila_existente = db.execute("SELECT code FROM products WHERE id=?", (existente['id'],)).fetchone()
                codigo_final = code or fila_existente['code']  # nunca borra un código ya cargado
                db.execute("""
                    UPDATE products SET code=?, name=?, category_id=?, brand=?, unidad=?, price=?, old_price=?,
                                         stock=?, description=?, info_adicional=?, image=?, destacado=?, active=?
                    WHERE id=?
                """, (codigo_final, name, category_id, brand or '', unidad or 'unidad', money(price),
                      money(old_price) or None, int(stock or 0), description or '', info_adicional or '',
                      image or '', 1 if destacado else 0, 1 if active is None or active else 0, existente['id']))
                actualizados += 1
            else:
                pid = uuid.uuid4().hex
                cat_icon = db.execute("SELECT icon FROM categories WHERE id=?", (category_id,)).fetchone()
                db.execute("""
                    INSERT INTO products (id, code, name, category_id, brand, unidad, price, old_price,
                                           stock, description, info_adicional, icon, image, destacado, active)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)
                """, (pid, code or pid[:8].upper(), name, category_id, brand or '', unidad or 'unidad',
                      money(price), money(old_price) or None, int(stock or 0), description or '',
                      info_adicional or '', cat_icon['icon'] if cat_icon else '📦', image or '',
                      1 if destacado else 0))
                creados += 1

    registrar_auditoria('Importación de productos', f'{creados} creados, {actualizados} actualizados')
    flash(f'Importación completa: {creados} productos creados, {actualizados} actualizados.', 'success')
    return redirect(url_for('admin'))


# --- Pedidos ---

def avisar_estado_al_cliente(pedido_id, estado):
    """Le avisa por correo al cliente que su pedido avanzó. Si no hay correo configurado no hace nada."""
    with get_db() as db:
        datos = db.execute("""SELECT c.nombre, c.email FROM pedidos p JOIN clientes c ON c.id = p.cliente_id
                              WHERE p.id=?""", (pedido_id,)).fetchone()
    if not datos or not datos['email']:
        return
    info = ESTADOS_PEDIDO[estado]
    try:
        enlace = f"{url_publica()}{url_for('pedido_confirmado', pedido_id=pedido_id)}"
    except RuntimeError:
        enlace = ''
    enviar_correo_simple(
        datos['email'], f"{info['titulo']} - Pedido #{pedido_id}",
        f"Hola {datos['nombre']},\n\n{info['titulo']}.\n{info['detalle']}\n\n"
        + (f"Podés seguir tu pedido acá: {enlace}\n\n" if enlace else '') + "Almacén del Valle")


@app.route('/admin/pedidos/<int:pedido_id>/estado', methods=['POST'])
@permiso_required('pedidos')
def admin_pedido_estado(pedido_id):
    """Cambia el estado de un pedido y lo anota en su historial; al pasar a "pagado" se emite la factura y se envía al cliente."""
    nuevo_estado = request.form.get('estado') or (request.get_json(silent=True) or {}).get('estado')
    if nuevo_estado not in ESTADOS_PEDIDO:
        if request.is_json:
            return jsonify({'error': 'Estado de pedido inválido'}), 400
        flash('Estado de pedido inválido.', 'error')
        return redirect(url_for('admin'))

    with get_db() as db:
        pedido = db.execute("SELECT estado FROM pedidos WHERE id=?", (pedido_id,)).fetchone()
        if not pedido:
            return redirect(url_for('admin'))
        db.execute("UPDATE pedidos SET estado=? WHERE id=?", (nuevo_estado, pedido_id))
        if pedido['estado'] != nuevo_estado:
            registrar_estado(db, pedido_id, nuevo_estado)
        # Cualquier estado posterior al pago exige que la factura exista
        factura_emitida = nuevo_estado in ESTADOS_PAGADOS and emitir_factura(db, pedido_id)
    # Al confirmar un pago (por ejemplo una transferencia), el cliente recibe su factura
    if factura_emitida:
        enviar_factura_por_email(pedido_id)
    registrar_auditoria('Estado de pedido actualizado', f'Pedido #{pedido_id} → {nuevo_estado}')
    if nuevo_estado != pedido['estado'] and nuevo_estado in ('preparando', 'enviado', 'listo_retiro', 'entregado'):
        avisar_estado_al_cliente(pedido_id, nuevo_estado)
    if request.is_json:
        return jsonify({'ok': True})
    return redirect(url_for('admin'))


# ---------------------------------------------------------------------------
# CHATBOT — Alice, la asistente de la tienda
# No usa servicios externos. Responde con (1) lo que le enseña el administrador desde el panel,
# (2) datos reales del negocio (horarios, envíos, pedidos, productos) y (3) charla básica.
# Además aprende sola de dos maneras, siempre bajo el control del administrador:
#   - Reformulaciones: si un cliente hace una pregunta que no entiende y enseguida la escribe de
#     otra manera que sí entiende, Alice anota la primera frase como sinónimo de esa respuesta.
#     Se activa cuando dos sesiones distintas confirman lo mismo, y el panel permite aprobarla,
#     rechazarla o borrarla.
#   - Valoraciones: cada respuesta lleva "me sirvió / no me sirvió"; las que no sirvieron
#     vuelven a la lista de consultas por enseñar.
# ---------------------------------------------------------------------------

PALABRAS_VACIAS = {
    'el', 'la', 'los', 'las', 'un', 'una', 'unos', 'unas', 'de', 'del', 'al', 'a', 'en', 'y', 'o', 'u',
    'que', 'como', 'cual', 'cuales', 'para', 'por', 'con', 'sin', 'se', 'mi', 'mis', 'tu', 'tus', 'su',
    'sus', 'me', 'te', 'lo', 'le', 'es', 'son', 'soy', 'esta', 'estan', 'estoy', 'puedo', 'puede',
    'pueden', 'este', 'ese', 'eso', 'esto', 'si', 'ya', 'muy', 'mas', 'favor', 'quisiera', 'queria',
}

SUGERENCIAS_CHATBOT = ['Horarios', 'Costo de envío', 'Medios de pago', 'Estado de mi pedido', 'Ofertas']


def normalizar_texto(texto):
    """Minúsculas, sin tildes ni signos: '¿Dónde están?' -> 'donde estan'."""
    texto = unicodedata.normalize('NFD', (texto or '').lower())
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^a-z0-9]+', ' ', texto).strip()


def raiz_palabra(palabra):
    """Quita el plural para que 'envíos' y 'envío' cuenten como la misma palabra."""
    if len(palabra) > 4 and palabra.endswith('es'):
        return palabra[:-2]
    if len(palabra) > 3 and palabra.endswith('s'):
        return palabra[:-1]
    return palabra


def tokens_de(texto):
    """Conjunto de palabras significativas de un texto, normalizadas."""
    return {raiz_palabra(p) for p in normalizar_texto(texto).split()
            if p not in PALABRAS_VACIAS and len(p) > 1}


def puntaje_respuesta(entrada, tokens_mensaje):
    """Qué tan bien responde una respuesta cargada por el administrador a la consulta (0 a 2).
    Una frase de "palabras clave" cuyas palabras están todas en la consulta es coincidencia segura;
    si no, se compara la consulta con la pregunta de ejemplo."""
    mejor = 0.0
    for frase in re.split(r'[,\n]', entrada['palabras_clave'] or ''):
        tokens_frase = tokens_de(frase)
        if tokens_frase and tokens_frase <= tokens_mensaje:
            mejor = max(mejor, 1 + 0.1 * len(tokens_frase))  # frases más específicas ganan
    tokens_pregunta = tokens_de(entrada['pregunta'])
    comunes = len(tokens_pregunta & tokens_mensaje)
    if comunes and tokens_pregunta:
        mejor = max(mejor, comunes / (len(tokens_pregunta) * len(tokens_mensaje)) ** 0.5)
    return mejor


def _respuesta(texto, acciones=None):
    """Formato de respuesta del chatbot: texto y botones (enlaces) opcionales."""
    return {'respuesta': texto, 'acciones': acciones or []}


def _elegir(*opciones):
    """Una de varias formas de decir lo mismo, para que Alice no suene repetida."""
    return random.choice(opciones)


def _nombre_en_saludo(ctx):
    nombre = ctx['contexto'].get('nombre')
    return f", {nombre}" if nombre else ''


def _bot_saludo(ctx):
    return _respuesta(_elegir(
        f"¡Hola{_nombre_en_saludo(ctx)}! Soy Alice, la asistente de Almacén del Valle. Puedo contarte horarios, "
        "costos de envío, medios de pago, buscar productos y decirte cómo va tu pedido. ¿Qué necesitás?",
        f"¡Buenas{_nombre_en_saludo(ctx)}! Soy Alice, la asistente de la tienda. Preguntame por productos, envíos, pagos u horarios, "
        "o por cómo va tu pedido."))


def _bot_gracias(ctx):
    return _respuesta(_elegir('¡De nada! Si necesitás algo más, escribime.',
                              '¡Un gusto ayudarte! Acá estoy para lo que necesites.',
                              '¡Gracias a vos! Que tengas un lindo día.'))


def _bot_ayuda(ctx):
    return _bot_saludo(ctx)


def _bot_identidad(ctx):
    return _respuesta('Me llamo Alice y soy la asistente virtual de Almacén del Valle. No soy una persona: soy un '
                      'programa que responde con los datos de la tienda y con lo que me enseñan, y cada día '
                      'aprendo un poco más de las preguntas de los clientes.')


def _bot_estado_animo(ctx):
    return _respuesta(_elegir('¡Muy bien, gracias por preguntar! ¿En qué te puedo ayudar?',
                              'Todo bien por acá, lista para ayudarte. ¿Qué necesitás?'))


def _bot_chiste(ctx):
    return _respuesta(_elegir('¿Qué le dice una yerba a otra? "Hoy nos toca estar bien unidas".',
                              '¿Por qué el tomate no puede dormir? Porque la ensalada lo mantiene en vilo.',
                              'Mi chiste favorito es que mi precio favorito siempre es "en oferta".'))


def _bot_aprendizaje(ctx):
    return _respuesta('Aprendo de dos maneras: el dueño de la tienda me enseña respuestas nuevas, y también '
                      'me fijo cuando alguien repite una pregunta con otras palabras para entenderla la próxima vez. '
                      'Si una respuesta no te sirvió, marcalo con el pulgar hacia abajo: me ayuda a mejorar.')


def _bot_horarios(ctx):
    filas = ctx['db'].execute("SELECT * FROM horarios_atencion ORDER BY orden").fetchall()
    if not filas:
        return _respuesta('Todavía no tengo cargados los horarios de atención. Escribinos por los medios de contacto de la tienda.')
    return _respuesta('Nuestros horarios de atención:\n' + '\n'.join(
        f"• {f['dias']}: {f['hora_inicio']} a {f['hora_fin']}" for f in filas))


def _bot_ubicacion(ctx):
    filas = ctx['db'].execute("SELECT * FROM direcciones ORDER BY orden").fetchall()
    if not filas:
        return _respuesta('Todavía no tengo cargada la dirección del local.')
    return _respuesta('Podés visitarnos en:\n' + '\n'.join(
        f"• {f['nombre']}: {f['direccion']}" + (f", {f['ciudad']}" if f['ciudad'] else '') for f in filas))


def _bot_contacto(ctx):
    filas = ctx['db'].execute("SELECT * FROM contactos ORDER BY orden").fetchall()
    if not filas:
        return _respuesta('Todavía no tengo cargados los datos de contacto.')
    return _respuesta('Podés comunicarte con nosotros por:\n' + '\n'.join(
        f"• {f['tipo']}: {f['valor']}" + (f" ({f['etiqueta']})" if f['etiqueta'] else '') for f in filas))


def _bot_envio(ctx):
    gratis_desde = get_config_float('envio_gratis_desde', 50000)
    costo = get_config_float('costo_envio', 2500)
    return _respuesta(f"El envío a domicilio cuesta ${ar_money(costo)} y es gratis en compras desde "
                      f"${ar_money(gratis_desde)}. También podés elegir retiro en sucursal, sin costo. "
                      f"Después de pagar, seguís el pedido paso a paso desde \"Mis pedidos\".")


def _bot_pago(ctx):
    medios = ', '.join(NOMBRES_METODO_PAGO.values())
    return _respuesta(f"Aceptamos: {medios}. Con tarjeta y Mercado Pago el pago se confirma al instante; "
                      f"con transferencia o efectivo el pedido queda pendiente hasta que se acredite.")


def _bot_pedido(ctx):
    if not ctx['cliente_id']:
        return _respuesta('Para ver el estado de tu pedido necesito que ingreses a tu cuenta.',
                          [{'texto': 'Ingresar', 'url': url_for('login', next=url_for('mis_pedidos'))}])
    pedido = ctx['db'].execute("SELECT id, estado FROM pedidos WHERE cliente_id=? ORDER BY id DESC LIMIT 1",
                               (ctx['cliente_id'],)).fetchone()
    if not pedido:
        return _respuesta('Todavía no hiciste ningún pedido. ¿Te ayudo a buscar algún producto?')
    datos = ESTADOS_PEDIDO.get(pedido['estado'], {})
    return _respuesta(f"Tu último pedido es el #{pedido['id']} y está: {etiqueta_estado(pedido['estado'])}. "
                      f"{datos.get('detalle', '')}",
                      [{'texto': 'Ver seguimiento', 'url': url_for('pedido_confirmado', pedido_id=pedido['id'])}])


def _bot_factura(ctx):
    return _respuesta('La factura se emite cuando se confirma el pago y te llega por correo con el PDF adjunto. '
                      'También la descargás desde "Mis pedidos".',
                      [{'texto': 'Mis pedidos', 'url': url_for('mis_pedidos')}])


def _bot_cupon(ctx):
    return _respuesta('Si tenés un cupón de descuento, ingresá el código en la pantalla de finalizar compra: '
                      'ahí se valida y se descuenta del total.')


def _bot_cuenta(ctx):
    return _respuesta('Podés ingresar con tu correo o con tu DNI. Si todavía no tenés cuenta, creala en un minuto.',
                      [{'texto': 'Ingresar', 'url': url_for('login')},
                       {'texto': 'Crear cuenta', 'url': url_for('registro')}])


def _bot_ofertas(ctx):
    db = ctx['db']
    ofertas = db.execute("SELECT titulo FROM ofertas_semana WHERE activa=1 ORDER BY orden LIMIT 3").fetchall()
    productos = db.execute("""SELECT id, name, price, old_price FROM products
                              WHERE active=1 AND old_price > price AND stock > 0 ORDER BY (old_price - price) / old_price DESC LIMIT 3""").fetchall()
    if not ofertas and not productos:
        return _respuesta('Ahora no tengo ofertas cargadas. Pasá por la tienda más tarde.')
    lineas = [f"• {o['titulo']}" for o in ofertas]
    lineas += [f"• {p['name']}: ${ar_money(p['price'])} (antes ${ar_money(p['old_price'])})" for p in productos]
    return _respuesta('Estas son las ofertas de ahora:\n' + '\n'.join(lineas),
                      [{'texto': p['name'], 'url': url_for('producto_detalle', product_id=p['id'])} for p in productos])


def buscar_productos(db, tokens, limite=3):
    """Productos activos cuyo nombre o marca contiene las palabras buscadas, los más parecidos primero."""
    if not tokens:
        return []
    encontrados = []
    for p in db.execute("SELECT id, name, brand, price, stock, unidad FROM products WHERE active=1").fetchall():
        palabras = tokens_de(f"{p['name']} {p['brand']}")
        puntos = sum(1 for t in tokens if t in palabras or any(w.startswith(t) for w in palabras if len(t) >= 4))
        if puntos:
            encontrados.append((puntos, p))
    encontrados.sort(key=lambda x: (-x[0], x[1]['name']))
    if not encontrados:
        return []
    mejor = encontrados[0][0]
    return [p for puntos, p in encontrados if puntos == mejor][:limite]


def _bot_producto(ctx):
    db = ctx['db']
    if not ctx['tokens_sin_intencion']:
        # Sin nombre de producto: si venían hablando de uno ("¿y cuánto sale?"), se sigue con ese
        ultimos = ctx['contexto'].get('ultimos_ids') or []
        if ultimos:
            marcas = ','.join('?' for _ in ultimos)
            productos = db.execute(f"SELECT id, name, price, stock FROM products WHERE active=1 AND id IN ({marcas})",
                                   ultimos).fetchall()
            if productos:
                return _respuesta('Del que veníamos hablando:\n' + '\n'.join(_linea_de_producto(p) for p in productos),
                                  [{'texto': p['name'], 'url': url_for('producto_detalle', product_id=p['id'])} for p in productos])
        return _respuesta('¿Qué producto estás buscando? Escribime el nombre, por ejemplo "yerba" o "leche".')
    productos = buscar_productos(db, ctx['tokens_sin_intencion'])
    if not productos:
        return _respuesta('No encontré ese producto. Probá con otro nombre o recorré las categorías de la tienda.')
    ctx['contexto']['ultimos_ids'] = [p['id'] for p in productos]
    return _respuesta('Encontré esto:\n' + '\n'.join(_linea_de_producto(p) for p in productos),
                      [{'texto': p['name'], 'url': url_for('producto_detalle', product_id=p['id'])} for p in productos])


def _linea_de_producto(p):
    return f"• {p['name']}: ${ar_money(p['price'])}" + (' (disponible)' if p['stock'] > 0 else ' (sin stock)')


# Cada intención: (nombre, palabras que la activan, función que arma la respuesta).
# Si dos intenciones empatan gana la primera de la lista.
INTENCIONES_CHATBOT = [
    ('saludo',    'hola buenas buenos saludos hey', _bot_saludo),
    ('gracias',   'gracias chau adios luego listo perfecto', _bot_gracias),
    ('pedido',    'pedido pedidos seguimiento rastrear estado llego demora', _bot_pedido),
    ('factura',   'factura facturas comprobante iva ticket', _bot_factura),
    ('cupon',     'cupon cupones codigo', _bot_cupon),
    ('ofertas',   'oferta ofertas promo promocion promociones descuento rebaja', _bot_ofertas),
    ('envio',     'envio envios delivery domicilio flete costo gratis reparto entregan entrega retiro retirar', _bot_envio),
    ('pago',      'pago pagar pagos tarjeta credito debito mercado transferencia efectivo rapipago cuotas medios', _bot_pago),
    ('horarios',  'horario horarios hora abren abre cierran cierra atienden atencion abierto', _bot_horarios),
    ('ubicacion', 'donde ubicacion direccion queda local sucursal mapa llegar', _bot_ubicacion),
    ('contacto',  'telefono whatsapp contacto llamar mail correo email celular numero comunicar', _bot_contacto),
    ('cuenta',    'cuenta registrar registro ingresar login sesion contrasena dni usuario', _bot_cuenta),
    ('ayuda',     'ayuda ayudar podes hacer funciones menu opciones', _bot_ayuda),
    ('identidad', 'quien sos llamas alice robot bot persona real', _bot_identidad),
    ('animo',     'estas tal andas como va', _bot_estado_animo),
    ('chiste',    'chiste chistes gracioso reir divertido', _bot_chiste),
    ('aprender',  'aprendes aprender aprendiste inteligente entrenada', _bot_aprendizaje),
    ('producto',  'tenes tienen hay vendes venden precio cuanto cuesta sale vale stock producto busco quiero', _bot_producto),
]
FUNCIONES_POR_INTENCION = {nombre: funcion for nombre, _, funcion in INTENCIONES_CHATBOT}
CONFIRMACIONES_PARA_ACTIVAR = 2   # sesiones distintas que deben repetir una reformulación para usarla
SEGUNDOS_PARA_REFORMULAR = 5 * 60


def vocabulario_chatbot(db):
    """Todas las palabras que Alice conoce: las de sus intenciones, las de las respuestas cargadas,
    las frases que aprendió y los nombres de productos. Sirve para corregir errores de tipeo."""
    palabras = set()
    for _, claves, _ in INTENCIONES_CHATBOT:
        palabras |= tokens_de(claves)
    for fila in db.execute("SELECT pregunta, palabras_clave FROM chatbot_respuestas WHERE activa=1"):
        palabras |= tokens_de(fila['pregunta']) | tokens_de(fila['palabras_clave'])
    for fila in db.execute("SELECT frase FROM chatbot_aprendido WHERE estado='activa'"):
        palabras |= tokens_de(fila['frase'])
    for fila in db.execute("SELECT name, brand FROM products WHERE active=1"):
        palabras |= tokens_de(f"{fila['name']} {fila['brand']}")
    return palabras


def corregir_errores(tokens, vocabulario):
    """Cambia cada palabra desconocida por la conocida más parecida ('horaio' -> 'horario')."""
    conocidas = sorted(vocabulario)
    corregidos = set()
    for palabra in tokens:
        if palabra in vocabulario or len(palabra) < 4:
            corregidos.add(palabra)
            continue
        # Las palabras cortas se corrigen con más cautela: 'hacen' no es un error de 'hacer'
        # y solo se consideran palabras que empiezan igual y tienen un largo parecido
        candidatas = [w for w in conocidas if w[0] == palabra[0] and abs(len(w) - len(palabra)) <= 2]
        parecidas = difflib.get_close_matches(palabra, candidatas, n=1, cutoff=0.8 if len(palabra) >= 6 else 0.88)
        corregidos.add(parecidas[0] if parecidas else palabra)
    return corregidos


def ejecutar_destino(db, destino, ctx):
    """Arma la respuesta de un destino aprendido: 'resp:<id>' (respuesta cargada) o 'int:<nombre>'."""
    tipo, _, valor = destino.partition(':')
    if tipo == 'resp':
        fila = db.execute("SELECT * FROM chatbot_respuestas WHERE id=? AND activa=1", (valor,)).fetchone()
        if fila:
            db.execute("UPDATE chatbot_respuestas SET usos = usos + 1 WHERE id=?", (fila['id'],))
            return _respuesta(fila['respuesta'])
    elif tipo == 'int' and valor in FUNCIONES_POR_INTENCION and valor != 'producto':
        return FUNCIONES_POR_INTENCION[valor](ctx)
    return None


def buscar_frase_aprendida(db, tokens):
    """La frase aprendida (activa) que más se parece a la consulta, o None."""
    mejor, mejor_puntaje = None, 0.0
    for fila in db.execute("SELECT * FROM chatbot_aprendido WHERE estado='activa'").fetchall():
        tokens_frase = tokens_de(fila['frase'])
        comunes = len(tokens_frase & tokens)
        if not comunes:
            continue
        puntaje = comunes / (len(tokens_frase) * len(tokens)) ** 0.5
        if puntaje >= 0.75 and puntaje > mejor_puntaje:
            mejor, mejor_puntaje = fila, puntaje
    return mejor


def registrar_aprendizaje(db, frase, destino, sesion):
    """Anota que 'frase' (algo que Alice no entendió) significó 'destino' para un cliente que la
    reformuló. La frase se empieza a usar cuando dos sesiones distintas lo confirman."""
    palabras = tokens_de(frase)
    frase_normal = normalizar_texto(frase)
    if not 1 <= len(palabras) <= 8 or len(frase_normal) < 4:
        return
    fila = db.execute("SELECT * FROM chatbot_aprendido WHERE frase=? AND destino=?", (frase_normal, destino)).fetchone()
    if not fila:
        db.execute("INSERT INTO chatbot_aprendido (frase, destino, sesiones) VALUES (?,?,?)",
                   (frase_normal, destino, sesion))
        return
    sesiones = [x for x in fila['sesiones'].split(',') if x]
    if fila['estado'] == 'rechazada' or sesion in sesiones:
        return
    confirmaciones = fila['confirmaciones'] + 1
    estado = 'activa' if confirmaciones >= CONFIRMACIONES_PARA_ACTIVAR else fila['estado']
    db.execute("UPDATE chatbot_aprendido SET confirmaciones=?, sesiones=?, estado=? WHERE id=?",
               (confirmaciones, ','.join((sesiones + [sesion])[-20:]), estado, fila['id']))


def detectar_nombre(mensaje):
    """'Me llamo Juan' / 'mi nombre es Juan' -> 'Juan' (solo se guarda en la sesión del cliente)."""
    coincidencia = re.search(r'\b(?:me llamo|mi nombre es)\s+([a-záéíóúñü]{2,20})\b', mensaje.lower())
    return coincidencia.group(1).capitalize() if coincidencia else None


def responder_chatbot(db, mensaje, cliente_id=None, contexto=None):
    """Busca la mejor respuesta para una consulta, en este orden:
    1) presentación del cliente, 2) respuestas cargadas por el administrador,
    3) frases que Alice aprendió, 4) intenciones con datos reales y búsqueda de productos.
    Devuelve el texto, botones, si pudo resolverla y su 'destino' (de dónde salió, para aprender).
    'contexto' es la memoria de la conversación (nombre, últimos productos) y se modifica acá."""
    contexto = contexto if contexto is not None else {}
    nombre = detectar_nombre(mensaje)
    if nombre:
        contexto['nombre'] = nombre
        return dict(_respuesta(f"¡Mucho gusto, {nombre}! Soy Alice. ¿En qué te puedo ayudar?"),
                    resuelto=True, destino='charla:nombre')

    vocabulario = vocabulario_chatbot(db)
    tokens = corregir_errores(tokens_de(mensaje), vocabulario)
    ctx = {'db': db, 'cliente_id': cliente_id, 'contexto': contexto, 'tokens_sin_intencion': set()}

    # 2) Respuestas entrenadas desde el panel
    candidatas = [(puntaje_respuesta(e, tokens), e) for e in
                  db.execute("SELECT * FROM chatbot_respuestas WHERE activa=1").fetchall()]
    candidatas = [(puntaje, e) for puntaje, e in candidatas if puntaje >= 0.5]
    if candidatas:
        _, entrada = max(candidatas, key=lambda x: x[0])
        db.execute("UPDATE chatbot_respuestas SET usos = usos + 1 WHERE id=?", (entrada['id'],))
        return dict(_respuesta(entrada['respuesta']), resuelto=True, destino=f"resp:{entrada['id']}")

    # 3) Frases que aprendió de las reformulaciones de los clientes
    aprendida = buscar_frase_aprendida(db, tokens)
    if aprendida:
        respuesta = ejecutar_destino(db, aprendida['destino'], ctx)
        if respuesta:
            db.execute("UPDATE chatbot_aprendido SET usos = usos + 1 WHERE id=?", (aprendida['id'],))
            return dict(respuesta, resuelto=True, destino=aprendida['destino'], aprendido_id=aprendida['id'])

    # 4) Intenciones con datos del negocio. La de productos es la más general ("cuánto cuesta..."),
    #    así que cede ante una intención específica, salvo que además se nombre un producto.
    palabras_de_producto = tokens_de(next(p for n, p, _ in INTENCIONES_CHATBOT if n == 'producto'))
    palabras_de_intenciones = set().union(*(tokens_de(p) for _, p, _ in INTENCIONES_CHATBOT))
    sobrantes = tokens - palabras_de_intenciones  # lo que no es palabra de ninguna intención: posible nombre de producto

    mejor_nombre, mejor_aciertos = None, 0
    for nombre_intencion, palabras, _ in INTENCIONES_CHATBOT:
        aciertos = len(tokens & tokens_de(palabras))
        if nombre_intencion != 'producto' and aciertos > mejor_aciertos:
            mejor_nombre, mejor_aciertos = nombre_intencion, aciertos
    nombra_producto = bool(sobrantes) and bool(buscar_productos(db, sobrantes))

    ctx['tokens_sin_intencion'] = sobrantes
    if mejor_nombre and not nombra_producto:
        return dict(FUNCIONES_POR_INTENCION[mejor_nombre](ctx), resuelto=True, destino=f"int:{mejor_nombre}")
    if nombra_producto or tokens & palabras_de_producto:
        return dict(_bot_producto(ctx), resuelto=True, destino='producto')

    return dict(_respuesta(_elegir(
        'Todavía no sé responder eso, pero lo anoto para aprenderlo. Mientras tanto puedo ayudarte con '
        'horarios, envíos, medios de pago, productos u ofertas.',
        'Mmm, esa no me la sé. La dejo anotada para que me enseñen. ¿Probamos con otra cosa? Sé de horarios, '
        'envíos, pagos, productos y ofertas.'),
        [{'texto': 'Ver contacto', 'url': url_for('index') + '#contacto'}]),
        resuelto=False, destino=None, sugerencias=SUGERENCIAS_CHATBOT)


def chat_demasiado_rapido():
    """Limita a 20 consultas por minuto por sesión, para que nadie llene el registro."""
    ahora = time.time()
    recientes = [t for t in session.get('chat_tiempos', []) if ahora - t < 60]
    if len(recientes) >= 20:
        return True
    session['chat_tiempos'] = recientes + [ahora]
    return False


def aprender_de_la_conversacion(db, respuesta):
    """Si el cliente reformuló una consulta que Alice no había entendido y ahora sí la entendió,
    guarda la frase anterior como sinónimo. Se llama después de cada respuesta."""
    pendiente = session.get('chat_pendiente')
    if not respuesta['resuelto']:
        return
    session.pop('chat_pendiente', None)
    destino = respuesta.get('destino') or ''
    if pendiente and time.time() - pendiente['t'] < SEGUNDOS_PARA_REFORMULAR and destino.startswith(('resp:', 'int:')) \
            and not destino.startswith('int:producto'):
        if 'chat_sid' not in session:
            session['chat_sid'] = secrets.token_hex(6)
        registrar_aprendizaje(db, pendiente['texto'], destino, session['chat_sid'])


@app.route('/api/chatbot', methods=['POST'])
def api_chatbot():
    """Recibe la consulta del cliente y devuelve la respuesta de Alice."""
    datos = request.get_json(silent=True) or {}
    mensaje = str(datos.get('mensaje', '')).strip()[:300]
    if not mensaje:
        return jsonify(dict(_respuesta('Escribime tu consulta y te ayudo.'), resuelto=True,
                            sugerencias=SUGERENCIAS_CHATBOT))
    if chat_demasiado_rapido():
        return jsonify(dict(_respuesta('Estás escribiendo muy rápido. Esperá unos segundos y probá de nuevo.'),
                            resuelto=True))
    es_prueba = bool(datos.get('prueba')) and session.get('admin_logged_in')  # las pruebas del panel no se registran y no enseñan
    contexto = dict(session.get('chat_ctx') or {})
    with get_db() as db:
        respuesta = responder_chatbot(db, mensaje, session.get('cliente_id'), contexto)
        if not es_prueba:
            origen = f"apr:{respuesta['aprendido_id']}" if respuesta.get('aprendido_id') else (respuesta.get('destino') or '')
            cursor = db.execute("INSERT INTO chatbot_mensajes (cliente_id, mensaje, resuelto, origen) VALUES (?,?,?,?)",
                                (session.get('cliente_id'), mensaje, 1 if respuesta['resuelto'] else 0, origen))
            aprender_de_la_conversacion(db, respuesta)
            if not respuesta['resuelto']:
                session['chat_pendiente'] = {'texto': mensaje[:120], 't': time.time()}
            session['chat_ctx'] = contexto
            session['chat_ids'] = (session.get('chat_ids', []) + [cursor.lastrowid])[-10:]
            respuesta['mensaje_id'] = cursor.lastrowid
    respuesta.pop('aprendido_id', None)
    respuesta.pop('destino', None)
    return jsonify(respuesta)


@app.route('/api/chatbot/valoracion', methods=['POST'])
def api_chatbot_valoracion():
    """El cliente marca si la respuesta le sirvió. Un "no" devuelve la consulta a la lista de
    pendientes del panel y, si la respuesta venía de una frase aprendida, la descarta."""
    datos = request.get_json(silent=True) or {}
    mensaje_id = datos.get('mensaje_id')
    if not isinstance(mensaje_id, int) or mensaje_id not in session.get('chat_ids', []):
        return jsonify({'error': 'Consulta desconocida.'}), 400
    sirvio = bool(datos.get('sirvio'))
    with get_db() as db:
        fila = db.execute("SELECT origen FROM chatbot_mensajes WHERE id=?", (mensaje_id,)).fetchone()
        if not fila:
            return jsonify({'error': 'Consulta desconocida.'}), 400
        db.execute("UPDATE chatbot_mensajes SET util=? WHERE id=?", (1 if sirvio else 0, mensaje_id))
        if not sirvio:
            db.execute("UPDATE chatbot_mensajes SET resuelto=0, descartado=0 WHERE id=?", (mensaje_id,))
            if (fila['origen'] or '').startswith('apr:'):
                db.execute("UPDATE chatbot_aprendido SET estado='rechazada' WHERE id=?", (fila['origen'][4:],))
    return jsonify({'ok': True})


# --- Entrenamiento desde el panel ---

@app.route('/admin/chatbot/agregar', methods=['POST'])
@permiso_required('chatbot')
def chatbot_agregar():
    """Carga una respuesta nueva; si nace de una consulta sin respuesta, esa consulta queda resuelta."""
    pregunta = request.form.get('pregunta', '').strip()
    respuesta = request.form.get('respuesta', '').strip()
    if not (pregunta and respuesta):
        flash('Completá la pregunta y la respuesta.', 'error')
        return redirect(url_for('admin', view='chatbot'))
    with get_db() as db:
        db.execute("INSERT INTO chatbot_respuestas (pregunta, palabras_clave, respuesta) VALUES (?,?,?)",
                   (pregunta, request.form.get('palabras_clave', '').strip(), respuesta))
        pendiente_id = request.form.get('pendiente_id', type=int)
        if pendiente_id:
            db.execute("""UPDATE chatbot_mensajes SET descartado=1 WHERE LOWER(mensaje) =
                          (SELECT LOWER(mensaje) FROM chatbot_mensajes WHERE id=?)""", (pendiente_id,))
        registrar_auditoria('Respuesta del chatbot creada', pregunta, db=db)
    return redirect(url_for('admin', view='chatbot'))


@app.route('/admin/chatbot/<int:rid>/editar', methods=['POST'])
@permiso_required('chatbot')
def chatbot_editar(rid):
    """Edita la pregunta, las palabras clave o la respuesta."""
    with get_db() as db:
        db.execute("UPDATE chatbot_respuestas SET pregunta=?, palabras_clave=?, respuesta=? WHERE id=?",
                   (request.form.get('pregunta', '').strip(), request.form.get('palabras_clave', '').strip(),
                    request.form.get('respuesta', '').strip(), rid))
        registrar_auditoria('Respuesta del chatbot editada', f'ID {rid}', db=db)
    return redirect(url_for('admin', view='chatbot'))


@app.route('/admin/chatbot/<int:rid>/toggle', methods=['POST'])
@permiso_required('chatbot')
def chatbot_toggle(rid):
    """Activa o desactiva una respuesta sin borrarla."""
    with get_db() as db:
        db.execute("UPDATE chatbot_respuestas SET activa = 1 - activa WHERE id=?", (rid,))
        registrar_auditoria('Respuesta del chatbot activada/desactivada', f'ID {rid}', db=db)
    return redirect(url_for('admin', view='chatbot'))


@app.route('/admin/chatbot/<int:rid>/eliminar', methods=['POST'])
@permiso_required('chatbot')
def chatbot_eliminar(rid):
    """Elimina una respuesta."""
    with get_db() as db:
        db.execute("DELETE FROM chatbot_respuestas WHERE id=?", (rid,))
        registrar_auditoria('Respuesta del chatbot eliminada', f'ID {rid}', db=db)
    return redirect(url_for('admin', view='chatbot'))


@app.route('/admin/chatbot/aprendido/<int:aid>/<accion>', methods=['POST'])
@permiso_required('chatbot')
def chatbot_aprendido_accion(aid, accion):
    """El administrador aprueba, rechaza o borra una frase que Alice aprendió sola."""
    if accion not in ('aprobar', 'rechazar', 'eliminar'):
        abort(404)
    with get_db() as db:
        if accion == 'eliminar':
            db.execute("DELETE FROM chatbot_aprendido WHERE id=?", (aid,))
        else:
            db.execute("UPDATE chatbot_aprendido SET estado=? WHERE id=?",
                       ('activa' if accion == 'aprobar' else 'rechazada', aid))
        registrar_auditoria(f'Frase aprendida por Alice: {accion}', f'ID {aid}', db=db)
    return redirect(url_for('admin', view='chatbot'))


@app.route('/admin/chatbot/pendiente/<int:mid>/descartar', methods=['POST'])
@permiso_required('chatbot')
def chatbot_descartar(mid):
    """Saca de la lista de pendientes una consulta que no hace falta responder."""
    with get_db() as db:
        db.execute("""UPDATE chatbot_mensajes SET descartado=1 WHERE LOWER(mensaje) =
                      (SELECT LOWER(mensaje) FROM chatbot_mensajes WHERE id=?)""", (mid,))
    return redirect(url_for('admin', view='chatbot'))


# ---------------------------------------------------------------------------
# Manejo de errores
# ---------------------------------------------------------------------------

@app.before_request
def forzar_https():
    """En producción, cualquier petición HTTP se redirige a la misma dirección por HTTPS."""
    if USAR_HTTPS and not request.is_secure:
        return redirect(request.url.replace('http://', 'https://', 1), code=301)


@app.after_request
def agregar_cabeceras_de_seguridad(respuesta):
    """Cabeceras que endurecen el navegador frente a ataques comunes."""
    respuesta.headers.setdefault('X-Content-Type-Options', 'nosniff')
    respuesta.headers.setdefault('X-Frame-Options', 'SAMEORIGIN')  # SAMEORIGIN: la vista previa del panel usa un iframe propio
    respuesta.headers.setdefault('Referrer-Policy', 'strict-origin-when-cross-origin')
    if USAR_HTTPS:
        # El navegador recordará usar siempre HTTPS con este sitio durante un año
        respuesta.headers.setdefault('Strict-Transport-Security', 'max-age=31536000; includeSubDomains')
    return respuesta


@app.route('/api/csrf-token')
def api_csrf_token():
    """Token CSRF actualizado, para que la página lo renueve si venció sin tener que recargarse."""
    return jsonify({'token': generate_csrf()})


@app.before_request
def actualizar_permisos_del_panel():
    """Los permisos del empleado se leen de la base en cada pedido: si el dueño se los cambia, o
    desactiva o borra al usuario, rige de inmediato y no recién cuando vuelve a iniciar sesión."""
    if not session.get('admin_logged_in') or session.get('es_dueno') or request.endpoint == 'static':
        return
    with get_db() as db:
        fila = db.execute("SELECT permisos, activo, es_dueno FROM usuarios WHERE id=?",
                          (session.get('admin_id'),)).fetchone()
    if not fila or not fila['activo']:
        for clave in ('admin_logged_in', 'admin_nombre', 'admin_id', 'es_dueno', 'permisos'):
            session.pop(clave, None)
        return
    session['es_dueno'] = bool(fila['es_dueno'])
    session['permisos'] = permisos_de_usuario(fila)


@app.before_request
def descartar_sesion_de_cliente_inexistente():
    """Si el navegador guarda la sesión de un cliente que ya no está en la base (por ejemplo porque
    se reemplazó o reinició el archivo de datos), la sesión se descarta en vez de provocar errores
    de integridad al agregar al carrito o pagar."""
    cliente_id = session.get('cliente_id')
    if not cliente_id or request.endpoint == 'static':
        return
    with get_db() as db:
        existe = db.execute("SELECT 1 FROM clientes WHERE id=?", (cliente_id,)).fetchone()
    if not existe:
        for clave in ('cliente_id', 'cliente_nombre', 'pago_pendiente'):
            session.pop(clave, None)
        session.modified = True


@app.errorhandler(CSRFError)
def csrf_invalido(e):
    """Token CSRF ausente o inválido (por ejemplo, la sesión venció con la página abierta)."""
    mensaje = 'Tu sesión venció o el formulario no es válido. Volvé a intentarlo.'
    if request.is_json or request.path.startswith('/api/') or '/api/' in request.path:
        return jsonify({'error': mensaje}), 400
    flash(mensaje, 'error')
    return redirect(url_for('admin') if request.path.startswith('/admin') else url_for('index'))


@app.errorhandler(500)
def error_interno(e):
    """Las llamadas de la tienda (/api/) reciben un JSON legible y no una página de error."""
    app.logger.exception("Error interno en %s", request.path)
    if request.path.startswith('/api/'):
        return jsonify({'error': 'Ocurrió un error en el servidor. Probá de nuevo en unos segundos.'}), 500
    return render_template('500.html'), 500


@app.errorhandler(404)
def not_found(e):
    """Página de error 404."""
    return render_template('404.html'), 404


@app.errorhandler(500)
def server_error(e):
    """Página de error 500."""
    return render_template('500.html'), 500


if __name__ == '__main__':
    init_db()
    # Evita que el hilo de backups arranque dos veces por el reloader de Flask
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'true' or not app.debug:
        threading.Thread(target=_backup_scheduler, daemon=True).start()

    # HTTPS en la computadora local (opcional), para probar el sitio con candado:
    #   - SSL_CERT y SSL_KEY: rutas de un certificado propio (por ejemplo, uno generado con openssl)
    #   - LOCAL_SSL=adhoc: certificado temporal automático (requiere "pip install cryptography")
    # En producción NO se usa esto: el certificado lo maneja Caddy o nginx (ver README y carpeta deploy).
    ssl_context = None
    if os.environ.get('SSL_CERT') and os.environ.get('SSL_KEY'):
        ssl_context = (os.environ['SSL_CERT'], os.environ['SSL_KEY'])
    elif os.environ.get('LOCAL_SSL') == 'adhoc':
        ssl_context = 'adhoc'
    app.run(debug=True, ssl_context=ssl_context)
