# Almacén del Valle — Supermercado Online

Proyecto Flask + SQLite: tienda online con carrito, login de clientes, pasarelas
de pago (simuladas) y emisión de facturas en PDF, más un panel de administrador
con dashboard, gráficos y gestión de productos, categorías, pedidos y clientes.

## 1. Requisitos

- Python 3.10 o superior instalado en tu computadora.
- Visual Studio Code con la extensión **Python** (de Microsoft).

## 2. Abrir el proyecto en VS Code

1. Abrí VS Code.
2. `Archivo → Abrir carpeta...` y elegí la carpeta `almacen-del-valle`.

## 3. Crear el entorno virtual

Abrí una terminal dentro de VS Code (`Terminal → Nueva terminal`) y ejecutá:

**Windows:**
```
python -m venv venv
venv\Scripts\activate
```

**Mac / Linux:**
```
python3 -m venv venv
source venv/bin/activate
```

Vas a ver `(venv)` al principio de la línea de la terminal cuando esté activado.

## 4. Instalar las dependencias

```
pip install -r requirements.txt
```

## 5. Ejecutar el proyecto

```
python app.py
```

La terminal va a mostrar algo como:

```
* Running on http://127.0.0.1:5000
```

- **Tienda:** abrí [http://127.0.0.1:5000](http://127.0.0.1:5000)
- **Panel administrador:** abrí [http://127.0.0.1:5000/admin](http://127.0.0.1:5000/admin)
  - Usuario: `admin`
  - Contraseña: `almacen123`

La base de datos SQLite (`data/almacendelvalle.db`) se crea sola la primera vez
que corrés la app, con categorías y productos de ejemplo ya cargados.

Para parar el servidor, volvé a la terminal y presioná `CTRL + C`.

## 6. Cómo probar el flujo de compra completo

1. Entrá a la tienda y agregá productos al carrito (ícono 🧺 arriba a la derecha).
2. Hacé clic en **Finalizar compra** — te va a pedir iniciar sesión.
3. Creá una cuenta nueva en **Crear una cuenta acá** (nombre, correo, DNI y contraseña). Después podés ingresar con tu **correo o tu DNI** en el mismo campo. Las cuentas viejas, que no tenían DNI, pueden cargarlo al finalizar una compra.
4. En el checkout, elegí el método de entrega y una pasarela de pago:
   - **Mercado Pago**: pantalla de pago y aprobación simulada.
   - **Tarjeta de crédito/débito**: formulario completo (número, titular, vencimiento, código, DNI y cuotas). Usá 4111 1111 1111 1111; con 4000 0000 0000 0002 se ve un rechazo.
   - **Transferencia bancaria**: muestra los datos de la cuenta del negocio y queda "pendiente de pago" hasta que el admin use *Confirmar pago*.
   - **Efectivo (Rapipago/Pago Fácil)**: genera el cupón con código de barras y queda pendiente hasta que el admin confirme el pago.
5. Al confirmar, si el pago quedó aprobado, se emite la factura automáticamente
   y podés **descargarla** o **imprimirla** desde la pantalla de confirmación o en **Mis pedidos**.
   - "Descargar" baja el PDF al dispositivo.
   - "Imprimir" abre una vista lista para imprimir y dispara el diálogo de
     impresión del navegador automáticamente (puede que el navegador te pida
     permitir ventanas emergentes la primera vez).

## 7. Panel de administrador — secciones

El menú del panel sigue la misma organización que usás en Electro Centro, adaptada al rubro supermercado:

- **📊 Dashboard** — KPIs, gráficos de ventas por mes y por categoría, pedidos recientes, y un aviso destacado si hay productos sin stock o con stock bajo.
- **🛒 Productos** — alta, edición y baja del catálogo, con código/SKU, unidad de medida (kg, litro, unidad, etc.), descripción, información adicional, **subida de imagen real**, y marca de producto **destacado** (⭐).
- **🗂️ Categorías** — alta y baja de categorías.
- **📢 Ofertas de la Semana** — publicá promociones que aparecen como **banner gigante** arriba de todo en la tienda (con carrusel si hay más de una activa), vinculables a un producto o categoría puntual.
- **📦 Pedidos** — todas las compras online, con cambio de estado y acceso directo a Descargar/Imprimir la factura de cada una.
- **👥 Clientes** — cuentas registradas en la tienda.
- **⚠️ Alertas de Stock** — productos **sin stock** (0 unidades) y con **stock bajo** (10 unidades o menos, configurable en `LOW_STOCK_THRESHOLD`).
- **📑 Reportes** — PDF y Excel de ventas, facturación (mensual/anual), stock bajo, stock general y valorización de inventario.
- **🏷️ Etiquetas de Precio** — generá un PDF con etiquetas listas para pegar en las góndolas (21 por hoja), de forma **global** (todo el catálogo) o **individual** (elegís productos puntuales con checkboxes).
- **💸 Gastos / Compras** — registrá compras de mercadería, sueldos, servicios, alquiler, impuestos y más, con los campos de factura (proveedor, CUIT, razón social, condición de IVA) o de sueldo (nombre, apellido) según corresponda. Incluye reportes mensuales/anuales en PDF y Excel.
- **🏬 Datos del Negocio** — cargá y editá los **horarios de atención**, **direcciones/sucursales**, **teléfonos/WhatsApp/email**, **redes sociales**, y la **configuración de envío** (monto para envío gratis y costo de envío). Todo esto se refleja automáticamente en la tienda.
- **🎟️ Cupones** — creá códigos de descuento (porcentaje o monto fijo), con compra mínima, límite de usos y fecha de vencimiento opcional.
- **📈 Historial de Precios** — mirá todos los cambios de precio de tus productos, con la variación porcentual de cada uno.
- **💾 Backups** — descarga manual inmediata de la base de datos, generación de backup en el servidor con un clic, y un backup automático diario (se conservan los últimos 7).
- **🔐 Usuarios** *(solo visible para el usuario dueño)* — creá cuentas de acceso al panel para empleados, con permisos por sección. Incluye **roles predefinidos** (Cajero, Repositor, Encargado) que marcan los permisos típicos con un clic.
- **🧾 Auditoría** *(solo visible para el usuario dueño)* — registro completo de las últimas 100 acciones realizadas en el panel: **todas** las altas, ediciones y bajas (productos, categorías, ofertas, gastos, horarios, direcciones, contactos, redes, cupones, usuarios, backups, importaciones) quedan registradas con quién y cuándo.

## 8. Otras herramientas del panel

- **Notificaciones de pedidos pendientes**: la sección "Pedidos" del menú muestra un contador en vivo (se actualiza solo cada 20 segundos, sin recargar la página) con la cantidad de pedidos pendientes. Cuando entra un pedido nuevo, el panel reproduce un sonido y muestra un aviso — no hace falta tener la pestaña de Pedidos abierta.
- **Correo electrónico (SMTP)**: en Datos del Negocio configurás tu propio servidor SMTP (por ejemplo Gmail con una "contraseña de aplicación"; puerto 587 con STARTTLS o 465 con SSL). Se usa para dos cosas: **enviarle la factura al cliente** apenas se confirma su pago (con los datos del cliente, el detalle de la compra, subtotal sin IVA, IVA, total y el PDF adjunto) y, si cargás un email de destino, **avisarte de cada pedido nuevo**. Si el pago queda pendiente (transferencia o efectivo), la factura sale cuando marcás el pedido como "pagado" en el panel. Incluye un botón para mandarte un correo de prueba. Los correos se envían en segundo plano; si no configurás nada, o el servidor falla, la compra se completa igual y el error queda en el log.
- **Importar/Exportar productos por Excel**: desde la sección Productos podés exportar todo tu catálogo a un `.xlsx` (útil como plantilla) y volver a importarlo para hacer altas o actualizaciones masivas. La importación empareja productos existentes primero por **código** (`code`) y, si no hay código o no matchea, por **nombre exacto dentro de la misma categoría** — así que no se duplican productos aunque no tengan código cargado. Si el archivo trae un código y el producto encontrado no tenía uno, se lo completa automáticamente para la próxima vez.
- **"Ver como cliente"**: el ícono 👁️ en la barra superior del panel abre una vista previa de la tienda embebida, sin salir del admin, con un botón para alternar entre modo claro y oscuro.

## 9. Lo que ve el cliente — funciones nuevas

- **Página de detalle de producto** (`/producto/<id>`): descripción completa, información adicional, reseñas con calificación promedio, y productos "también compraron".
- **Reseñas y calificaciones**: cualquier cliente logueado puede dejar una reseña de 1 a 5 estrellas con comentario.
- **Favoritos**: botón de corazón en la página de producto; lista completa en "Mis favoritos" desde el menú de usuario.
- **Buscador con autocompletado**: al escribir en la barra de búsqueda aparecen sugerencias en vivo con foto, marca y precio, que llevan directo a la página del producto.
- **Cupones de descuento**: campo de código en el checkout, con validación en tiempo real antes de confirmar la compra. Cupón de ejemplo ya cargado: `BIENVENIDO10` (10% off, compra mínima $5.000).
- **Carrito persistente**: si el cliente inicia sesión, su carrito se guarda en la base de datos y lo recupera al volver a loguearse desde cualquier dispositivo.
- **Seguimiento de pedido**: línea de tiempo con fecha y hora de cada paso (ver sección 9.1), en la confirmación de compra y en "Mis pedidos".

### 9.1 Línea de tiempo del envío

Después de pagar, el cliente ve en la pantalla de confirmación una línea de tiempo que se actualiza sola (cada pocos segundos, sin recargar) a medida que el negocio cambia el estado desde el panel. Los pasos dependen de cómo eligió recibir el pedido:

- **Envío a domicilio**: Pedido recibido → Pago confirmado → Estamos preparando tu envío → Hemos enviado tu pedido → Tu pedido ha sido entregado.
- **Retiro en sucursal**: Pedido recibido → Pago confirmado → Preparando → Listo para retirar → Entregado.

Cada cambio queda guardado con su fecha y hora en la tabla `pedido_estados`. Los pedidos anteriores a esta versión se completan automáticamente al iniciar la aplicación.

### 9.2 Seguridad de los formularios (CSRF)

Todos los formularios POST llevan un token oculto (`csrf_token`) y las llamadas `fetch` del navegador lo envían en la cabecera `X-CSRFToken` (lo agrega `static/js/csrf.js`). Si el token falta o es inválido, el servidor no ejecuta la acción y vuelve a la página con un aviso. Al agregar un formulario nuevo hay que incluir:

```html
<input type="hidden" name="csrf_token" value="{{ csrf_token() }}">
```

### 9.3 Alice, la asistente virtual (chatbot)

En la esquina inferior derecha de la tienda está Alice. Responde sobre horarios, formas de pago, envíos, retiro, estado de pedidos, cupones y precios o stock de productos, y tiene charla básica (saluda, dice quién es, recuerda el nombre del cliente durante la conversación y sigue el hilo: después de hablar de la leche, "¿y cuánto sale?" se refiere a la leche). Corrige errores de tipeo comunes ("horaio" por "horario"). **No usa servicios externos ni inteligencia artificial generativa**: funciona con reglas, con las respuestas que carga el dueño y con lo que aprende de las conversaciones, así que es gratis, predecible y no inventa datos.

**Cómo se entrena (panel de administrador, sección Alice):**

1. **Consultas sin respuesta**: cada pregunta que no supo contestar queda registrada (también las que el cliente marcó con "no me sirvió"). Con el botón para enseñar, la pregunta pasa al formulario y solo falta escribir la respuesta.
2. **Respuestas cargadas**: se definen frases clave (por ejemplo "cambios, devoluciones, mercadería") y la respuesta. Alice elige la que más coincide, ignorando tildes, mayúsculas y plurales.
3. **Probar**: se escribe una pregunta y se ve qué contestaría, sin que quede registrada ni le enseñe nada.

**Cómo aprende sola:**

- **Reformulaciones**: si un cliente pregunta algo que Alice no entiende y en los 5 minutos siguientes lo escribe de otra manera que sí entiende, Alice anota la primera frase como sinónimo de esa respuesta. La frase empieza a usarse cuando **dos clientes distintos** lo confirman (una sola persona repitiendo no alcanza). En "Lo que Alice aprendió sola" se puede aprobar antes, rechazar o borrar cada frase.
- **Valoraciones**: cada respuesta tiene "¿Te sirvió?". Un "no" devuelve la consulta a la lista de pendientes y, si la respuesta venía de una frase aprendida, la rechaza automáticamente.

Alice aprende vocabulario (distintas formas de pedir lo mismo), no inventa respuestas: todo lo que contesta sale de los datos de la tienda o de textos escritos por el dueño. Hay un límite de 20 mensajes por minuto por sesión.

### 9.4 HTTPS / SSL

El certificado SSL no lo genera la aplicación: lo emite el servidor donde se publique, para un dominio real. La aplicación ya queda preparada para trabajar detrás de HTTPS:

- Con la variable de entorno `USE_HTTPS=1` activa cookies `Secure`, redirección de HTTP a HTTPS, cabecera HSTS y las cabeceras de seguridad.
- `SECRET_KEY` y `ADMIN_PASSWORD` se toman de variables de entorno (o del archivo `.env`); en producción hay que definirlas.

**Opción recomendada en un servidor propio (Caddy, certificado gratuito automático de Let's Encrypt):**

```bash
pip install -r requirements.txt
export USE_HTTPS=1 SECRET_KEY="una-clave-larga-y-aleatoria" ADMIN_PASSWORD="otra-clave" SITE_URL="https://tudominio.com.ar"
gunicorn --workers 1 --threads 4 --bind 127.0.0.1:8000 wsgi:app
```

Y en `deploy/Caddyfile` se reemplaza el dominio por el propio; Caddy obtiene y renueva el certificado solo. Hay también un ejemplo con nginx + certbot (`deploy/nginx.conf.example`) y un servicio de systemd (`deploy/almacen.service.example`). Se usa un solo worker porque el respaldo automático corre dentro del proceso.

**Para probar HTTPS en la propia computadora:** `SSL_CERT=cert.pem SSL_KEY=key.pem python app.py`, o `LOCAL_SSL=adhoc python app.py` (requiere `pip install cryptography`; el navegador va a mostrar una advertencia porque el certificado es autofirmado, es normal).

### 9.5 Bloqueo por intentos y recuperación de contraseña

- **Bloqueo**: tres contraseñas incorrectas seguidas bloquean esa cuenta durante **3 minutos**, tanto en la tienda como en el panel. Mientras dura el bloqueo no entra ni con la contraseña correcta. Entrar con el correo o con el DNI cuenta como la misma cuenta. Además hay un tope de 15 fallos por dirección IP, para frenar a quien prueba muchas cuentas. Los valores están al comienzo de la sección "ACCESO" de `app.py` (`MAX_INTENTOS_LOGIN`, `SEGUNDOS_BLOQUEO`).
- **Recuperar la contraseña**: en "¿Olvidaste tu contraseña?" el cliente escribe su correo y recibe un enlace de un solo uso que vence a la hora. En el panel funciona igual con el usuario: el enlace va al email del usuario o, si el dueño no tiene uno, al email de avisos de Datos del Negocio. Cambiar la contraseña por este medio desbloquea la cuenta y envía un aviso de cambio. La pantalla responde siempre lo mismo exista o no el correo, y pedir muchos enlaces seguidos se pausa unos minutos.
- **Requisito**: el envío usa el correo (SMTP) de Datos del Negocio. Si no está configurado, el enlace no se envía (queda un aviso en el log). Probalo con el botón de correo de prueba antes de publicar.
- **En producción** definí `SITE_URL` (por ejemplo `https://tudominio.com.ar`) para que los enlaces de los correos apunten siempre a tu dominio.
- La clave inicial del panel solo se muestra en la pantalla de ingreso mientras `USE_HTTPS` no esté activado. Cambiala antes de publicar el sitio.

**Activación de la cuenta por correo**: si el correo (SMTP) está configurado, al registrarse el cliente recibe un enlace (vale 24 horas, un solo uso) y la cuenta no puede ingresar hasta abrirlo; así se comprueba que el correo ingresado es real. En el login puede pedir que se lo reenvíen (como máximo uno por minuto). Cambiar la contraseña desde el correo también activa la cuenta. Si el correo no está configurado, la cuenta queda activa de inmediato para no trabar las altas. Las cuentas anteriores a esta versión quedan activas, y el panel marca como "sin activar" las que todavía no abrieron el enlace. La configuración de facturación tiene su propia sección en el panel, "Facturación".

### 9.6 Facturación electrónica estilo ARCA (modo demostración)

Cada factura sale con letra, punto de venta, número correlativo, **CAE** y vencimiento del CAE, y un **código QR** que lleva a una página pública de verificación (`/verificar-factura/<número>`) donde se ven los datos de autorización sin información personal del cliente.

- **Formato**: el PDF sigue el diseño de los comprobantes de ARCA (letra y código en el recuadro central, emisor y receptor, detalle con código, unidad y alícuota, importes, CAE, QR y numeración de páginas), con los datos del almacén.
- **Datos del emisor**: en el panel, sección "Facturación": razón social, CUIT (se valida el dígito verificador), domicilio, Ingresos Brutos, inicio de actividades, condición frente al IVA y punto de venta. Cambian solo las facturas nuevas.
- **Responsable Inscripto** emite **Factura A** (a quien tiene CUIT y es Responsable Inscripto, con IVA discriminado por línea) o **Factura B** (a consumidor final u otras condiciones, con la leyenda de Transparencia Fiscal, Ley 27.743). En el checkout el cliente puede pedir Factura A cargando CUIT y razón social.
- **Monotributista** emite **Factura C**, sin IVA.
- **Modo demostración (el único activo hoy)**: el CAE se genera al azar y todos los comprobantes llevan el aviso "sin validez fiscal". Sirve para mostrar el flujo completo: compra, pago, factura con CAE y QR, correo y verificación.
- **Conexión real con ARCA**: todo el sistema ya trabaja con los tres datos que devuelve ARCA (número, CAE, vencimiento). Lo único que falta es reemplazar `solicitar_cae()` en `app.py` por la llamada al web service WSFEv1 (con autenticación WSAA), lo que requiere el certificado digital y el punto de venta habilitado del negocio, probar primero en homologación y definir con el contador la condición frente al IVA. Hasta entonces las facturas **no tienen validez fiscal**.

### 9.7 Pedidos, medios de pago y estados

- **Estados desde el panel**: en Pedidos, cada fila tiene un botón con el siguiente paso: *Confirmar pago* (transferencia o efectivo), *Preparar envío*, *Marcar como enviado* y *Marcar como entregado* (con retiro en el local: *Preparar pedido* y *Listo para retirar*). Cada cambio queda en el historial, aparece en la línea de tiempo del cliente ("Estamos preparando tu envío", "Hemos enviado tu pedido", "Tu pedido ha sido entregado") y le llega un correo si el SMTP está configurado. El selector chico permite corregir el estado a mano.
- **Mercado Pago y tarjetas**: pasan por una pantalla de pago antes de crear el pedido. La tarjeta pide número, titular, vencimiento, código de seguridad, DNI del titular y, en crédito, cuotas; valida número (Luhn), vencimiento, largo del código (4 en American Express) y marca. Nunca se guarda el número ni el código: solo la marca y los últimos 4 dígitos. El cobro es simulado; para ver un rechazo usar una tarjeta terminada en 0002 (por ejemplo 4000 0000 0000 0002). Tarjetas de prueba: 4111 1111 1111 1111 (Visa), 5555 5555 5555 4444 (Mastercard), 3782 822463 10005 (Amex).
- **Transferencia**: el pedido queda pendiente y se muestran los datos de la cuenta del negocio con el importe y la referencia "Pedido #N". Se cargan en el panel, Datos del Negocio, "Datos para transferencias" (CBU de 22 dígitos y alias validados). Mientras no se carguen, el cliente ve datos de ejemplo con un aviso.
- **Efectivo (Rapipago / Pago Fácil)**: se genera un cupón con código de 20 dígitos, código de barras Code 128, vencimiento a las 72 horas y versión imprimible y en PDF. Solo lo ve el cliente dueño del pedido. Cuando el local cobra, se usa *Confirmar pago*.
- **Envío de la factura por correo**: al confirmarse un pago (tarjeta, Mercado Pago o *Confirmar pago* en el panel) la factura en PDF se manda al correo con el que se registró el cliente. Para que salga de verdad hay que configurar el correo en Datos del Negocio, "Correo electrónico" (con Gmail: servidor `smtp.gmail.com`, puerto 587, usuario = tu Gmail y una "contraseña de aplicación"). El botón "Enviar prueba" manda un correo real y muestra el resultado o el motivo exacto del error. En Pedidos, cada factura indica si fue "Enviada por correo" o "No se pudo enviar", y el botón ✉️ la reenvía mostrando el resultado. Cada intento queda en la tabla `correos_enviados`. Si el correo no está configurado, Pedidos muestra un aviso.
- **Si "Enviar prueba" no conecta**: el mensaje indica la causa. Si habla de certificados (frecuente en Mac), ejecutá `pip install -r requirements.txt` (incluye `certifi`) y reiniciá. Si no se puede conectar, probá el puerto 465 (SSL) en vez del 587 o probá desde otra red: algunos proveedores de internet, firewalls y antivirus bloquean el 587.
- **Conexión real**: sigue pendiente. Mercado Pago (Checkout Pro), Rapipago y Pago Fácil requieren las credenciales del negocio; hoy solo cambia el cobro simulado.

## Importante al actualizar a una versión nueva

La base de datos (tus clientes, pedidos y productos) vive en `data/almacendelvalle.db`. **Los zip nuevos ya no la incluyen**, para que no pises la tuya. Para actualizar: descomprimí el zip en una carpeta nueva, copiá tu `data/almacendelvalle.db` (y la carpeta `data/backups`) de la carpeta anterior a la nueva y arrancá desde ahí; las columnas nuevas se agregan solas al abrir. Si no copiás ninguna base, la app crea una nueva con los productos de ejemplo. Antes de actualizar, conviene hacer un backup desde el panel (Backups).

## 10. Sobre las imágenes de productos

Las imágenes **no se suben como archivo**: se guarda únicamente la **URL** que
pegues en el campo "URL de la imagen" al crear o editar un producto (por
ejemplo, un link a una imagen alojada en Google Drive, Imgur, Cloudinary, o
cualquier otro servicio). Esto simplifica el despliegue del proyecto porque no
hay que preocuparse por el almacenamiento de archivos en el servidor. Si el
campo queda vacío, el producto muestra el ícono de su categoría como antes.

## 11. Sobre las pasarelas de pago

Este proyecto **simula** el comportamiento de Mercado Pago, tarjetas de
crédito/débito, transferencia bancaria y pago en efectivo (Rapipago/Pago
Fácil), sin conectarse a ninguna API real ni usar credenciales de producción.
Esto te permite probar todo el flujo de compra de punta a punta.

Para conectar una pasarela real más adelante (por ejemplo Mercado Pago),
tendrías que:
1. Crear una cuenta de vendedor y obtener credenciales (`ACCESS_TOKEN`).
2. Instalar su SDK (`pip install mercadopago`).
3. Reemplazar la función `procesar_pago_simulado()` en `app.py` por una
   llamada real a su API de preferencias de pago / checkout.

## 12. Estructura del proyecto

```
almacen-del-valle/
├── app.py                 # Backend Flask: rutas, lógica, reportes, base de datos
├── requirements.txt
├── data/
│   ├── almacendelvalle.db # Se crea sola al correr la app
│   └── backups/           # Backups automáticos diarios (últimos 7)
├── static/
│   ├── css/
│   │   ├── style.css      # Estilos de la tienda
│   │   └── admin.css      # Estilos del panel administrador
│   └── js/
│       ├── main.js        # Carrito, mega-menú, banner de ofertas, tema (tienda)
│       └── admin.js       # Navegación, gráficos, gastos, backups, negocio, usuarios (admin)
└── templates/
    ├── base.html           # Layout compartido de la tienda (header, mega-menú, footer dinámico)
    ├── index.html          # Página principal (con banner gigante de ofertas)
    ├── login.html / registro.html
    ├── checkout.html       # Selección de pago y confirmación
    ├── pedido_confirmado.html
    ├── mis_pedidos.html
    ├── admin_login.html
    ├── admin.html          # Dashboard + CRUD + Stock + Reportes + Ofertas + Gastos + Backups + Negocio + Usuarios
    ├── factura_imprimir.html  # Vista de impresión de factura (auto-imprime)
    ├── 404.html / 500.html
```

## 13. Usuario dueño y gestión de empleados

Cada empleado ve y puede usar **solo** las secciones que el dueño le marcó: Dashboard, Productos, Categorías, Pedidos, Clientes, Alertas de stock, Reportes, Ofertas, Gastos, Etiquetas, Backups, Datos del Negocio, Cupones, Chatbot (Alice), Historial de precios y Facturación. Sin ningún permiso marcado no ve nada. Los cambios de permisos rigen de inmediato (no hace falta que el empleado vuelva a entrar) y, si se lo desactiva, se le cierra la sesión. Usuarios y Auditoría son solo del dueño. Al actualizar desde una versión anterior, quienes tenían Productos conservan Historial de precios y quienes tenían Datos del Negocio conservan Facturación.


La cuenta `admin` que se crea por defecto queda marcada como **dueño** (`es_dueno`),
que es la única que puede ver la sección **🔐 Usuarios** y crear, editar o
eliminar otras cuentas del panel.

Los permisos que le asignás a cada usuario nuevo **sí se aplican de verdad**:
- El usuario solo ve en el menú lateral las secciones para las que tiene permiso.
- Si intenta entrar a una sección sin permiso (por ejemplo, manipulando la URL
  o forzando un botón oculto), el panel muestra "🔒 No tenés permiso para ver
  esta sección" en vez del contenido real.
- Las rutas del servidor también están protegidas: aunque alguien intente
  llamar directamente a una acción (agregar producto, generar un reporte,
  borrar una categoría, etc.) sin tener el permiso correspondiente, el
  servidor la rechaza.
- El usuario **dueño** siempre tiene acceso a todo, sin excepción.

## 14. Tablas principales de la base de datos

- `categories`, `products` — catálogo de la tienda (incluye imagen como URL, destacado, unidad de medida).
- `clientes` — cuentas de clientes (login público).
- `pedidos`, `pedido_items` — carrito confirmado / historial de compras.
- `pagos` — resultado de cada intento de pago por pasarela.
- `facturas` — una factura por pedido pagado, con IVA discriminado.
- `usuarios` — cuentas del panel de administrador, con permisos y flag de dueño.
- `ofertas_semana` — ofertas que se muestran como banner gigante en la tienda.
- `gastos` — compras de mercadería, sueldos, servicios, impuestos, etc.
- `backup_log` — registro de cada backup generado.
- `horarios_atencion`, `direcciones`, `contactos`, `redes_sociales` — datos del negocio que se muestran en la tienda.
- `configuracion` — clave/valor para ajustes generales (envío gratis desde, costo de envío, configuración SMTP para notificaciones por email).
- `favoritos` — productos guardados por cada cliente.
- `resenas` — reseñas y calificaciones de productos.
- `cupones` — códigos de descuento.
- `carrito_guardado` — carrito persistente por cliente (sobrevive entre sesiones).
- `auditoria` — registro de acciones realizadas en el panel.
- `pedido_estados` — historial de estados de cada pedido, con fecha y hora.
- `chatbot_respuestas` — respuestas cargadas por el dueño para Alice.
- `chatbot_aprendido` — frases que Alice aprendió sola, con su estado y confirmaciones.
- `facturas` — facturas emitidas, con letra, número, CAE y vencimiento.
- `intentos_login` — intentos fallidos de ingreso y bloqueos.
- `recuperacion_clave` — enlaces de recuperación de contraseña (solo su hash).
- `chatbot_mensajes` — preguntas hechas al asistente y si pudo responderlas.

## 15. Reiniciar la base de datos

Si querés borrar todo y volver a los datos de ejemplo, con el servidor
detenido, borrá el archivo `data/almacendelvalle.db` (y los `.db-shm` /
`.db-wal` si existen) y volvé a correr `python app.py`.
