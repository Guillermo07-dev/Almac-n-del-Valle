/* ---------------- THEME ---------------- */
(function initTheme(){
  const saved = localStorage.getItem('av-admin-theme');
  if(saved) document.documentElement.setAttribute('data-theme', saved);
})();
function toggleTheme(){
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  const next = isDark ? 'light' : 'dark';
  html.setAttribute('data-theme', next);
  localStorage.setItem('av-admin-theme', next);
  document.getElementById('themeBtn').textContent = next === 'dark' ? '☀️' : '🌙';
  buildCharts();
}

/* ---------------- NAV ---------------- */
function setView(view, el){
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  el.classList.add('active');
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById('view-' + view).classList.add('active');
  const titles = {
    dashboard: ['Dashboard', 'Resumen general del negocio'],
    productos: ['Productos', 'Gestioná el catálogo de tu tienda'],
    categorias: ['Categorías', 'Organizá las secciones de la tienda'],
    ofertas: ['Ofertas de la Semana', 'Se muestran como banner gigante en la tienda'],
    pedidos: ['Pedidos', 'Seguimiento de compras de clientes'],
    clientes: ['Clientes', 'Cuentas registradas en la tienda'],
    stock: ['Alertas de Stock', 'Productos sin stock o por debajo del mínimo'],
    reportes: ['Reportes', 'Descargá reportes en PDF o Excel'],
    etiquetas: ['Etiquetas de Precio', 'Generá etiquetas listas para imprimir'],
    gastos: ['Gastos / Compras', 'Registrá compras de mercadería, sueldos y más'],
    backups: ['Backups', 'Copias de seguridad de la base de datos'],
    negocio: ['Datos del Negocio', 'Horarios, direcciones, contactos y redes sociales'],
    facturacion: ['Facturación', 'Datos del emisor y comprobantes electrónicos (modo demostración)'],
    cupones: ['Cupones', 'Códigos de descuento para tus clientes'],
    chatbot: ['Alice (chatbot)', 'Entrená a Alice, la asistente virtual de la tienda'],
    historial: ['Historial de Precios', 'Cambios de precio registrados por producto'],
    auditoria: ['Auditoría', 'Registro de acciones realizadas en el panel'],
    usuarios: ['Usuarios', 'Accesos y permisos del panel'],
  };
  document.getElementById('viewTitle').textContent = titles[view][0];
  document.getElementById('viewSub').textContent = titles[view][1];
  document.getElementById('sidebar').classList.remove('open');
  if (view === 'backups') cargarBackups();
  try { sessionStorage.setItem('av-admin-view', view); } catch(e) {}
  /* Las listas se arman en el servidor al cargar la página: si el panel quedó abierto mientras
     alguien se registraba o compraba, no mostraría esos datos hasta recargar. Por eso, al entrar
     a Pedidos o Clientes con la página ya vieja, se vuelve a cargar. */
  if ((view === 'pedidos' || view === 'clientes') && Date.now() - PAGINA_CARGADA > 5000) {
    window.location.href = '/admin?view=' + view;
  }
}
const PAGINA_CARGADA = Date.now();

/* Al recargar (por ejemplo después de guardar algo), el panel vuelve a la sección en la que estabas.
   Una dirección con ?view=nombre tiene prioridad. */
document.addEventListener('DOMContentLoaded', () => {
  let vista = new URLSearchParams(window.location.search).get('view');
  try { vista = vista || sessionStorage.getItem('av-admin-view'); } catch(e) {}
  let item = vista && document.querySelector('.nav-item[data-view="' + vista + '"]');
  if (item && vista !== 'dashboard') { setView(vista, item); return; }
  /* Quien no tiene permiso sobre el Dashboard empieza en la primera sección que sí puede usar */
  if (!document.querySelector('.nav-item[data-view="dashboard"]')) {
    item = document.querySelector('.sb-nav .nav-item[data-view]');
    if (item) setView(item.dataset.view, item);
  } else {
    document.querySelector('.nav-item[data-view="dashboard"]').classList.add('active');
  }
});

/* ---------------- MODALS ---------------- */
function openModal(id){ document.getElementById(id).classList.add('show'); }
function closeModal(id){ document.getElementById(id).classList.remove('show'); }

function editProduct(p){
  document.getElementById('productModalTitle').textContent = 'Editar producto';
  document.getElementById('productForm').action = '/admin/productos/editar/' + p.id;
  document.getElementById('p-code').value = p.code || '';
  document.getElementById('p-name').value = p.name;
  document.getElementById('p-brand').value = p.brand;
  document.getElementById('p-cat').value = p.category_id;
  document.getElementById('p-unidad').value = p.unidad || 'unidad';
  document.getElementById('p-price').value = p.price;
  document.getElementById('p-old-price').value = p.old_price || '';
  document.getElementById('p-stock').value = p.stock;
  document.getElementById('p-description').value = p.description || '';
  document.getElementById('p-info').value = p.info_adicional || '';
  document.getElementById('p-destacado').checked = !!p.destacado;
  document.getElementById('p-active-row').style.display = 'block';
  document.getElementById('p-active').value = p.active ? '1' : '';
  document.getElementById('p-image-url').value = p.image || '';
  updateImagePreview(p.image || '');
  openModal('productModal');
}

function updateImagePreview(url){
  const wrap = document.getElementById('p-image-preview-wrap');
  const img = document.getElementById('p-image-preview');
  if(url){
    img.src = url;
    wrap.style.display = 'block';
  } else {
    wrap.style.display = 'none';
    img.src = '';
  }
}
document.addEventListener('DOMContentLoaded', () => {
  const imgInput = document.getElementById('p-image-url');
  if(imgInput) imgInput.addEventListener('input', () => updateImagePreview(imgInput.value.trim()));
});

async function toggleDestacado(productId, btn){
  const res = await fetch('/admin/productos/' + productId + '/destacar', {method: 'POST'});
  const data = await res.json();
  if(data.ok){
    btn.textContent = data.destacado ? '⭐' : '☆';
    btn.title = data.destacado ? 'Quitar destacado' : 'Destacar';
  }
}

/* Reset del modal de producto a "Nuevo" cuando se abre desde el botón */
document.addEventListener('DOMContentLoaded', () => {
  const newBtn = document.querySelector('[onclick="openModal(\'productModal\')"]');
  if(newBtn){
    newBtn.addEventListener('click', () => {
      document.getElementById('productModalTitle').textContent = 'Nuevo producto';
      document.getElementById('productForm').action = document.getElementById('productForm').dataset.createUrl ||
        document.getElementById('productForm').action;
      document.getElementById('productForm').reset();
      document.getElementById('p-active-row').style.display = 'none';
      updateImagePreview('');
    });
  }
});

/* ---------------- ORDER STATUS ---------------- */
function updateEstado(pedidoId, estado){
  const form = document.getElementById('estadoForm');
  form.action = '/admin/pedidos/' + pedidoId + '/estado';
  form.innerHTML = `<input type="hidden" name="csrf_token" value="${csrfToken()}">` +
                   `<input type="hidden" name="estado" value="${estado}">`;
  form.submit();
}

/* ---------------- CHARTS ---------------- */
let salesChart, catChart;
function buildCharts(){
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const gridColor = isDark ? '#28382F' : '#DEE6DE';
  const textColor = isDark ? '#93A89D' : '#5B6D63';

  const ctx1 = document.getElementById('salesChart');
  if(ctx1){
    if(salesChart) salesChart.destroy();
    salesChart = new Chart(ctx1, {
      type: 'line',
      data: {
        labels: chartData.meses.length ? chartData.meses : ['Sin datos'],
        datasets: [{
          label: 'Ventas ($)',
          data: chartData.ventas.length ? chartData.ventas : [0],
          borderColor: '#1F6F5C', backgroundColor: 'rgba(31,111,92,0.12)',
          tension: .4, fill: true, pointRadius: 3, pointBackgroundColor: '#1F6F5C', borderWidth: 2.5
        }]
      },
      options: {
        plugins: {legend: {display: false}},
        scales: {
          x: {grid: {display: false}, ticks: {color: textColor, font: {family: 'Plus Jakarta Sans', size: 11}}},
          y: {grid: {color: gridColor}, ticks: {color: textColor, font: {family: 'IBM Plex Mono', size: 10}}}
        }
      }
    });
  }

  const ctx2 = document.getElementById('catChart');
  if(ctx2){
    if(catChart) catChart.destroy();
    catChart = new Chart(ctx2, {
      type: 'doughnut',
      data: {
        labels: chartData.cat_labels.length ? chartData.cat_labels : ['Sin ventas aún'],
        datasets: [{
          data: chartData.cat_valores.length ? chartData.cat_valores : [1],
          backgroundColor: ['#1F6F5C', '#F4A62A', '#7B3F61', '#2C8A71', '#DEE6DE', '#C1554A'],
          borderWidth: 0
        }]
      },
      options: {
        cutout: '68%',
        plugins: {legend: {position: 'bottom', labels: {color: textColor, font: {family: 'Plus Jakarta Sans', size: 11}, padding: 14, usePointStyle: true, pointStyle: 'circle'}}}
      }
    });
  }
}

document.addEventListener('DOMContentLoaded', buildCharts);
window.addEventListener('resize', () => {
  document.getElementById('mobileMenuBtn').style.display = window.innerWidth <= 760 ? 'flex' : 'none';
});
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('mobileMenuBtn').style.display = window.innerWidth <= 760 ? 'flex' : 'none';
});

/* ---------------- ETIQUETAS DE PRECIO ---------------- */
function seleccionarTodasEtiquetas(val){
  document.querySelectorAll('.etiqueta-chk').forEach(chk => chk.checked = val);
}

/* ---------------- GASTOS ---------------- */
function actualizarCamposGasto(){
  const cat = document.getElementById('g-cat').value;
  const bloqueFactura = ['Compra mercadería', 'Servicios', 'Alquiler', 'Logística', 'Mantenimiento'];
  const bloqueSueldo = ['Sueldos'];
  const bloqueImpuesto = ['Impuestos'];
  document.getElementById('g-bloque-factura').style.display = bloqueFactura.includes(cat) ? 'block' : 'none';
  document.getElementById('g-bloque-sueldo').style.display = bloqueSueldo.includes(cat) ? 'block' : 'none';
  document.getElementById('g-bloque-impuesto').style.display = bloqueImpuesto.includes(cat) ? 'block' : 'none';
  document.getElementById('g-bloque-general').style.display =
    (!bloqueFactura.includes(cat) && !bloqueSueldo.includes(cat) && !bloqueImpuesto.includes(cat)) ? 'block' : 'none';
}

async function guardarGasto(){
  const cat = document.getElementById('g-cat').value;
  const desc = document.getElementById('g-desc').value.trim();
  if(!desc){ alert('Ingresá una descripción.'); return; }

  let monto = 0, payload = {categoria: cat, descripcion: desc, nota: document.getElementById('g-nota').value.trim()};

  if(cat === 'Sueldos'){
    monto = parseFloat(document.getElementById('g-monto-sueldo').value || 0);
    payload.nombre = document.getElementById('g-nombre').value.trim();
    payload.apellido = document.getElementById('g-apellido').value.trim();
    payload.condicion_iva = document.getElementById('g-condicion-iva-sueldo').value;
  } else if(cat === 'Impuestos'){
    monto = parseFloat(document.getElementById('g-monto-impuesto').value || 0);
    payload.tipo_impuesto = document.getElementById('g-tipo-impuesto').value;
  } else if(['Compra mercadería','Servicios','Alquiler','Logística','Mantenimiento'].includes(cat)){
    monto = parseFloat(document.getElementById('g-monto-factura').value || 0);
    payload.proveedor = document.getElementById('g-proveedor').value.trim();
    payload.cuit = document.getElementById('g-cuit').value.trim();
    payload.razon_social = document.getElementById('g-razon-social').value.trim();
    payload.domicilio_fiscal = document.getElementById('g-domicilio').value.trim();
    payload.condicion_iva = document.getElementById('g-condicion-iva').value;
    payload.nro_factura = document.getElementById('g-nro-factura').value.trim();
    payload.tipo_factura = document.getElementById('g-tipo-factura').value;
  } else {
    monto = parseFloat(document.getElementById('g-monto-general').value || 0);
  }
  if(!monto || monto <= 0){ alert('Ingresá un monto válido.'); return; }
  payload.monto = monto;

  const res = await fetch('/admin/api/gastos/nuevo', {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
  });
  const data = await res.json();
  if(data.ok){
    const msg = document.getElementById('g-msg');
    msg.style.display = 'inline';
    setTimeout(() => msg.style.display = 'none', 2500);
    location.reload();
  }
}

async function eliminarGasto(id){
  if(!confirm('¿Eliminar este gasto?')) return;
  const res = await fetch('/admin/api/gastos/' + id + '/eliminar', {method: 'POST'});
  const data = await res.json();
  if(data.ok){
    const row = document.getElementById('gasto-row-' + id);
    if(row) row.remove();
  }
}

/* ---------------- BACKUPS ---------------- */
async function cargarBackups(){
  const res = await fetch('/admin/api/backups');
  const data = await res.json();
  const tbody = document.getElementById('backupsTableBody');
  if(!data.length){
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-soft);padding:20px">Todavía no hay backups automáticos guardados.</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(b => `
    <tr>
      <td class="mono">${b.filename}</td>
      <td>${b.size_kb} KB</td>
      <td>${b.created}</td>
      <td><a class="mini-btn" href="/admin/backup/descargar/${b.filename}" title="Descargar">⬇️</a></td>
    </tr>
  `).join('');
}

async function generarBackupAhora(){
  const res = await fetch('/admin/backup/ahora', {method: 'POST'});
  const data = await res.json();
  if(data.ok){
    showAdminToast('Backup generado correctamente ✅');
    cargarBackups();
  }
}

/* ---------------- USUARIOS ---------------- */
function nuevoUsuario(){
  document.getElementById('usuarioFormTitle').textContent = 'Nuevo usuario';
  document.getElementById('usuarioForm').action = '/admin/usuarios/agregar';
  document.getElementById('usuarioForm').reset();
  document.getElementById('u-password').required = true;
  document.getElementById('usuario-form-card').style.display = 'block';
}

function editarUsuario(u){
  document.getElementById('usuarioFormTitle').textContent = 'Editar usuario';
  document.getElementById('usuarioForm').action = '/admin/usuarios/editar/' + u.id;
  document.getElementById('u-nombre').value = u.nombre;
  document.getElementById('u-apellido').value = u.apellido;
  document.getElementById('u-usuario').value = u.usuario;
  document.getElementById('u-email').value = u.email || '';
  document.getElementById('u-password').value = '';
  document.getElementById('u-password').required = false;
  const permisos = (u.permisos || '').split(',');
  document.querySelectorAll('.u-permiso-chk').forEach(chk => chk.checked = permisos.includes(chk.value));
  document.getElementById('usuario-form-card').style.display = 'block';
}

function cerrarFormUsuario(){
  document.getElementById('usuario-form-card').style.display = 'none';
}

/* ---------------- TOAST (admin) ---------------- */
function showAdminToast(msg){
  let t = document.getElementById('adminToast');
  if(!t){
    t = document.createElement('div');
    t.id = 'adminToast';
    t.className = 'admin-toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2400);
}

/* ---------------- DATOS DEL NEGOCIO (horarios / direcciones / contactos / redes) ---------------- */
function editarHorario(h){
  document.getElementById('horarioForm').action = '/admin/horarios/editar/' + h.id;
  document.getElementById('h-dias').value = h.dias;
  document.getElementById('h-inicio').value = h.hora_inicio;
  document.getElementById('h-fin').value = h.hora_fin;
  document.getElementById('horarioSubmitBtn').textContent = '💾 Guardar cambios';
  document.getElementById('horarioCancelBtn').style.display = 'inline-flex';
}

function editarDireccion(d){
  document.getElementById('direccionForm').action = '/admin/direcciones/editar/' + d.id;
  document.getElementById('d-nombre').value = d.nombre;
  document.getElementById('d-direccion').value = d.direccion;
  document.getElementById('d-ciudad').value = d.ciudad || '';
  document.getElementById('d-provincia').value = d.provincia || '';
  document.getElementById('direccionSubmitBtn').textContent = '💾 Guardar cambios';
  document.getElementById('direccionCancelBtn').style.display = 'inline-flex';
}

function editarContacto(c){
  document.getElementById('contactoForm').action = '/admin/contactos/editar/' + c.id;
  document.getElementById('c-tipo').value = c.tipo;
  document.getElementById('c-etiqueta').value = c.etiqueta || '';
  document.getElementById('c-valor').value = c.valor;
  document.getElementById('contactoSubmitBtn').textContent = '💾 Guardar cambios';
  document.getElementById('contactoCancelBtn').style.display = 'inline-flex';
}

function editarRed(r){
  document.getElementById('redForm').action = '/admin/redes/editar/' + r.id;
  document.getElementById('r-plataforma').value = r.plataforma;
  document.getElementById('r-url').value = r.url;
  document.getElementById('redSubmitBtn').textContent = '💾 Guardar cambios';
  document.getElementById('redCancelBtn').style.display = 'inline-flex';
}

function cancelarEdicionNegocio(tipo){
  const map = {
    horario: {form: 'horarioForm', btn: 'horarioSubmitBtn', cancel: 'horarioCancelBtn', label: '+ Agregar horario'},
    direccion: {form: 'direccionForm', btn: 'direccionSubmitBtn', cancel: 'direccionCancelBtn', label: '+ Agregar dirección'},
    contacto: {form: 'contactoForm', btn: 'contactoSubmitBtn', cancel: 'contactoCancelBtn', label: '+ Agregar contacto'},
    red: {form: 'redForm', btn: 'redSubmitBtn', cancel: 'redCancelBtn', label: '+ Agregar red'},
  };
  const cfg = map[tipo];
  const form = document.getElementById(cfg.form);
  form.action = form.dataset.createUrl;
  form.reset();
  document.getElementById(cfg.btn).textContent = cfg.label;
  document.getElementById(cfg.cancel).style.display = 'none';
}

/* ---------------- ROLES PREDEFINIDOS (usuarios) ---------------- */
const ROLES_PRESET = {
  cajero: ['pedidos', 'clientes'],
  repositor: ['productos', 'categorias', 'stock'],
  encargado: ['dashboard', 'productos', 'categorias', 'pedidos', 'clientes', 'stock', 'reportes',
              'ofertas', 'gastos', 'etiquetas', 'negocio', 'cupones', 'historial', 'facturacion'],
  ninguno: [],
};
function aplicarRol(rol){
  const permisos = ROLES_PRESET[rol] || [];
  document.querySelectorAll('.u-permiso-chk').forEach(chk => {
    chk.checked = permisos.includes(chk.value);
  });
}

/* ---------------- PREVIEW "VER COMO CLIENTE" ---------------- */
function abrirPreviewTienda(){
  const overlay = document.getElementById('previewOverlay');
  const iframe = document.getElementById('previewIframe');
  if(iframe.src === 'about:blank' || !iframe.src){
    iframe.src = '/';
  }
  overlay.classList.add('show');
}
function cerrarPreviewTienda(){
  document.getElementById('previewOverlay').classList.remove('show');
}
function togglePreviewTheme(){
  const iframe = document.getElementById('previewIframe');
  try{
    const doc = iframe.contentDocument;
    const html = doc.documentElement;
    const isDark = html.getAttribute('data-theme') === 'dark';
    html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  } catch(e){ /* el iframe puede no estar listo todavía */ }
}

/* ---------------- NOTIFICACIONES POR EMAIL ---------------- */
async function probarNotificacion(){
  const msg = document.getElementById('notif-msg');
  msg.style.display = 'inline'; msg.style.color = 'var(--text-soft)'; msg.textContent = 'Enviando…';
  try{
    const res = await fetch('/admin/notificaciones/probar', {method: 'POST'});
    const data = await res.json();
    msg.style.color = data.ok ? 'var(--verde-700)' : 'var(--rojo)';
    msg.textContent = (data.ok ? '✓ ' : '✗ ') + (data.mensaje || 'Sin respuesta del servidor');
  }catch(e){
    msg.style.color = 'var(--rojo)';
    msg.textContent = '✗ No se pudo completar la prueba. Revisá que la aplicación siga en marcha.';
  }
}

/* ---------------- SONIDO + POLLING DE PEDIDOS NUEVOS ---------------- */
function reproducirBeep(){
  try{
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    [0, 0.15].forEach((delay, i) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = i === 0 ? 880 : 1046.5;
      gain.gain.setValueAtTime(0.15, ctx.currentTime + delay);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + delay + 0.25);
      osc.connect(gain).connect(ctx.destination);
      osc.start(ctx.currentTime + delay);
      osc.stop(ctx.currentTime + delay + 0.25);
    });
  } catch(e){ /* el navegador puede bloquear audio sin interacción previa del usuario */ }
}

let pedidosPendientesPrevios = null;
async function chequearPedidosNuevos(){
  try{
    const res = await fetch('/admin/api/notificaciones');
    if(!res.ok) return;
    const data = await res.json();
    if(pedidosPendientesPrevios !== null && data.pendientes > pedidosPendientesPrevios){
      reproducirBeep();
      showAdminToast(`🔔 ¡Nuevo pedido pendiente! (${data.pendientes} en total)`);
    }
    pedidosPendientesPrevios = data.pendientes;
    document.querySelectorAll('[data-view="pedidos"] .nav-badge').forEach(b => {
      b.textContent = data.pendientes;
      b.style.display = data.pendientes > 0 ? '' : 'none';
    });
  } catch(e){ /* sin conexión momentánea, se reintenta en el próximo ciclo */ }
}
document.addEventListener('DOMContentLoaded', () => {
  chequearPedidosNuevos();
  setInterval(chequearPedidosNuevos, 20000);
});


/* ---------------- CHATBOT ---------------- */
async function probarChatbot(){
  const campo = document.getElementById('bot-prueba');
  const caja = document.getElementById('bot-resultado');
  if(!campo.value.trim()) return;
  const res = await fetch('/api/chatbot', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({mensaje: campo.value, prueba: true})
  });
  const data = await res.json();
  caja.style.display = 'block';
  caja.textContent = (data.respuesta || data.error || '') + (data.resuelto === false ? '\n\n(El asistente no supo responder esta consulta)' : '');
}

/* Pasa una consulta sin respuesta al formulario para enseñarle cómo contestarla */
function usarPendiente(boton){
  resetFormBot();
  document.getElementById('bot-pregunta').value = boton.dataset.mensaje;
  document.getElementById('bot-pendiente-id').value = boton.dataset.id;
  document.getElementById('bot-respuesta').focus();
  document.getElementById('botForm').scrollIntoView({behavior: 'smooth', block: 'center'});
}

function editarRespuestaBot(r){
  const form = document.getElementById('botForm');
  form.action = '/admin/chatbot/' + r.id + '/editar';
  document.getElementById('bot-pregunta').value = r.pregunta;
  document.getElementById('bot-claves').value = r.palabras_clave;
  document.getElementById('bot-respuesta').value = r.respuesta;
  document.getElementById('bot-pendiente-id').value = '';
  document.getElementById('bot-form-titulo').textContent = '✏️ Editando una respuesta';
  document.getElementById('bot-submit').textContent = 'Guardar cambios';
  form.scrollIntoView({behavior: 'smooth', block: 'center'});
}

function resetFormBot(){
  const form = document.getElementById('botForm');
  form.reset();
  form.action = form.dataset.createUrl;
  document.getElementById('bot-pendiente-id').value = '';
  document.getElementById('bot-form-titulo').textContent = '🎓 Enseñarle una respuesta';
  document.getElementById('bot-submit').textContent = '+ Guardar respuesta';
}
