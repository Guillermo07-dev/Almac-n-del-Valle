/* ---------------- THEME ---------------- */
(function initTheme(){
  const saved = localStorage.getItem('av-theme');
  if(saved) document.documentElement.setAttribute('data-theme', saved);
})();

function toggleTheme(){
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  const next = isDark ? 'light' : 'dark';
  html.setAttribute('data-theme', next);
  localStorage.setItem('av-theme', next);
  const btn = document.getElementById('themeBtn');
  if(btn) btn.textContent = next === 'dark' ? '☀️' : '🌙';
}
document.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('themeBtn');
  if(btn) btn.textContent = document.documentElement.getAttribute('data-theme') === 'dark' ? '☀️' : '🌙';
});

/* ---------------- MEGA MENU ---------------- */
const SUBCATS = {
  frutasyverduras: ['Frutas', 'Verduras', 'Huevos', 'Frutos secos'],
  frescos: ['Fiambres', 'Quesos', 'Pastas frescas', 'Dulces'],
  carniceria: ['Carne vacuna', 'Carne de cerdo', 'Carne de ave', 'Embutidos'],
  lacteos: ['Leches', 'Yogures', 'Mantecas y margarinas', 'Postres y flanes'],
  panificados: ['Elaboración propia', 'Pan de molde', 'Tortas y tartas', 'Pan dulces'],
  almacen: ['Harinas y leudantes', 'Pastas secas', 'Arroz y legumbres', 'Aceites y vinagres', 'Conservas', 'Snacks'],
  bebidas: ['Aguas y sodas', 'Gaseosas', 'Vinos', 'Cervezas', 'Jugos'],
  limpieza: ['Limpieza de la ropa', 'Limpieza de la cocina', 'Limpieza del baño', 'Insecticidas'],
};

function toggleMegaMenu(force){
  const menu = document.getElementById('megaMenu');
  const overlay = document.getElementById('megaOverlay');
  const trigger = document.getElementById('megaMenuBtn');
  if(!menu) return;
  const shouldOpen = force !== undefined ? force : !menu.classList.contains('open');
  menu.classList.toggle('open', shouldOpen);
  overlay.classList.toggle('open', shouldOpen);
  if(trigger) trigger.classList.toggle('open', shouldOpen);
}

function showSubcats(catId){
  document.querySelectorAll('.mega-cat-item').forEach(el => el.classList.toggle('active', el.dataset.cat === catId));
  const catEl = document.querySelector(`.mega-cat-item[data-cat="${catId}"]`);
  const icon = catEl ? catEl.querySelector('.mci-ico').textContent : '📦';
  const iconBg = catEl ? catEl.querySelector('.mci-ico').style.background : '#EEF6F1';
  const name = catEl ? catEl.querySelector('.mci-name').textContent : '';
  const subcats = SUBCATS[catId] || [];
  const panel = document.getElementById('megaSubcats');
  panel.innerHTML = `
    <div class="subcat-title">
      <span class="sc-ico" style="background:${iconBg}">${icon}</span>
      <h4>${name}</h4>
    </div>
    <div class="subcat-grid">
      ${subcats.map(s => `<a href="#" onclick="setCat('${catId}'); toggleMegaMenu(false); return false;">${s}</a>`).join('')}
    </div>
  `;
}

/* ---------------- FILTERS (home page) ---------------- */
function setCat(catId){
  document.querySelectorAll('.chip').forEach(c => c.classList.toggle('active', c.dataset.cat === catId));
  document.querySelectorAll('#prodGrid .prod-card').forEach(card => {
    const matches = catId === 'all' || card.dataset.cat === catId;
    card.style.display = matches ? '' : 'none';
  });
  const prodSection = document.getElementById('productos');
  if(prodSection) prodSection.scrollIntoView({behavior:'smooth'});
}

const searchInput = document.getElementById('searchInput');
if(searchInput){
  searchInput.addEventListener('input', () => {
    applySearchFilter(searchInput.value);
    debounceSuggestions(searchInput.value);
  });
  document.addEventListener('click', (e) => {
    if(!e.target.closest('.searchbar')){
      const box = document.getElementById('searchSuggestions');
      if(box) box.classList.remove('show');
    }
  });
}
function applySearchFilter(q){
  q = (q || '').toLowerCase();
  document.querySelectorAll('#prodGrid .prod-card').forEach(card => {
    card.style.display = card.dataset.name.includes(q) ? '' : 'none';
  });
}
function filterProductsByName(name){
  if(!searchInput) return;
  searchInput.value = name;
  applySearchFilter(name);
  document.getElementById('productos').scrollIntoView({behavior:'smooth'});
}

/* ---------------- SEARCH AUTOCOMPLETE ---------------- */
let suggestTimer = null;
function debounceSuggestions(q){
  clearTimeout(suggestTimer);
  const box = document.getElementById('searchSuggestions');
  if(!box) return;
  if(!q || q.trim().length < 2){ box.classList.remove('show'); box.innerHTML = ''; return; }
  suggestTimer = setTimeout(() => fetchSuggestions(q.trim()), 250);
}
async function fetchSuggestions(q){
  const box = document.getElementById('searchSuggestions');
  try{
    const res = await fetch('/api/productos/buscar?q=' + encodeURIComponent(q));
    const data = await res.json();
    if(!data.length){ box.classList.remove('show'); box.innerHTML = ''; return; }
    box.innerHTML = data.map(p => `
      <a class="suggestion-item" href="/producto/${p.id}">
        <span class="sg-ico" ${p.image ? `style="background-image:url('${p.image}')"` : ''}>${p.image ? '' : p.icon}</span>
        <span class="sg-info"><b>${p.name}</b><span>${p.brand || ''}</span></span>
        <span class="sg-price mono">$${p.price.toLocaleString('es-AR')}</span>
      </a>
    `).join('');
    box.classList.add('show');
  } catch(e){ box.classList.remove('show'); }
}

/* ---------------- GIANT OFFERS BANNER ---------------- */
let gbIndex = 0, gbTimer = null;
function showBannerSlide(i){
  const slides = document.querySelectorAll('.gb-slide');
  const dots = document.querySelectorAll('.gb-dot');
  if(!slides.length) return;
  gbIndex = (i + slides.length) % slides.length;
  slides.forEach((s, idx) => s.classList.toggle('active', idx === gbIndex));
  dots.forEach((d, idx) => d.classList.toggle('active', idx === gbIndex));
}
function initGiantBanner(){
  const slides = document.querySelectorAll('.gb-slide');
  if(slides.length > 1){
    gbTimer = setInterval(() => showBannerSlide(gbIndex + 1), 6000);
  }
}
document.addEventListener('DOMContentLoaded', initGiantBanner);

/* ---------------- CART (backed by Flask session API) ---------------- */
async function apiCall(url, body, reintento){
  /* Nunca falla en silencio: si el servidor responde con algo que no es JSON (error interno, sesión
     vencida) o no hay conexión, devuelve {error} para que quien llama lo muestre. Si el token CSRF
     venció (la página quedó abierta mucho tiempo o se reinició el servidor), pide uno nuevo y reintenta. */
  try{
    const res = await fetch(url, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(body || {})
    });
    let data = null;
    try{ data = await res.json(); }catch(e){ /* respuesta sin JSON */ }
    if(res.status === 400 && !reintento){
      const nuevo = await fetch('/api/csrf-token').then(r => r.json()).catch(() => null);
      const meta = document.querySelector('meta[name="csrf-token"]');
      if(nuevo && nuevo.token && meta){ meta.content = nuevo.token; return apiCall(url, body, true); }
    }
    if(data) return data;
    console.error('Respuesta inesperada de', url, res.status);
    return {error: 'No pudimos completar la acción (error ' + res.status + '). Recargá la página con Ctrl+F5 e intentá de nuevo.'};
  }catch(e){
    console.error('Falló la llamada a', url, e);
    return {error: 'No hay conexión con el servidor. Revisá que la aplicación siga en marcha.'};
  }
}

function renderCart(data){
  const count = data.count || 0;
  const countEl = document.getElementById('cartCount');
  if(countEl) countEl.textContent = count;

  const itemsEl = document.getElementById('cartItems');
  const footEl = document.getElementById('cartFoot');
  if(!itemsEl || !footEl) return;

  if(!data.items || data.items.length === 0){
    itemsEl.innerHTML = `<div class="cart-empty"><div class="em">🧺</div>Tu canasta está vacía.<br>¡Agregá tus productos favoritos!</div>`;
    footEl.innerHTML = '';
    return;
  }

  itemsEl.innerHTML = data.items.map(it => `
    <div class="cart-item">
      <div class="cart-thumb">${it.icon}</div>
      <div class="cart-item-info">
        <div class="nm">${it.name}</div>
        <div class="price" style="font-size:14px">$${it.line_total.toLocaleString('es-AR')}</div>
        <div class="qty-ctrl">
          <button onclick="changeQty('${it.product_id}', ${it.qty - 1})">–</button>
          <span class="mono">${it.qty}</span>
          <button onclick="changeQty('${it.product_id}', ${it.qty + 1})" ${it.qty >= it.stock ? 'disabled' : ''}>+</button>
        </div>
      </div>
      <button class="rm" onclick="changeQty('${it.product_id}', 0)">Quitar</button>
    </div>
  `).join('');

  footEl.innerHTML = `
    <div class="cart-total-row"><span>Subtotal</span><b>$${data.subtotal.toLocaleString('es-AR')}</b></div>
    <a class="btn-pill btn-primary" href="/checkout">Finalizar compra</a>
  `;
}

async function refreshCart(){
  const res = await fetch('/api/carrito');
  const data = await res.json();
  renderCart(data);
}

function stepQty(productId, delta, maxStock){
  const el = document.getElementById('qty-' + productId);
  if(!el) return;
  let val = parseInt(el.textContent, 10) + delta;
  if(val < 1) val = 1;
  if(maxStock && val > maxStock) val = maxStock;
  el.textContent = val;
}

/* ---------------- FAVORITOS ---------------- */
async function toggleFavorito(productId){
  const res = await fetch('/api/favoritos/toggle', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({product_id: productId})
  });
  if(res.status === 401){
    window.location.href = '/login';
    return;
  }
  const data = await res.json();
  const btn = document.getElementById('favBtn');
  if(btn){
    btn.textContent = data.es_favorito ? '❤️' : '🤍';
    btn.classList.toggle('active', data.es_favorito);
  }
  showToast(data.es_favorito ? 'Agregado a favoritos ❤️' : 'Quitado de favoritos');
}

async function addToCart(productId, btn){
  /* Todo el recorrido va protegido: si algo falla, se avisa en pantalla (y en la consola del navegador)
     en vez de quedarse sin responder. */
  try{
    const qtyEl = document.getElementById('qty-' + productId);
    const qty = qtyEl ? (parseInt(qtyEl.textContent, 10) || 1) : 1;
    const originalLabel = btn ? btn.innerHTML : '';
    const data = await apiCall('/api/carrito/agregar', {product_id: productId, qty});
    if(data.error){ showToast(data.error); return; }
    renderCart(data);
    if(!(data.items || []).some(it => it.product_id === productId)){
      showToast('No hay stock disponible de este producto.');
      return;
    }
    showToast('Producto agregado a la canasta 🧺');
    if(qtyEl) qtyEl.textContent = '1';
    if(btn){
      btn.classList.add('added'); btn.innerHTML = '✓ Agregado';
      setTimeout(() => { btn.classList.remove('added'); btn.innerHTML = originalLabel; }, 900);
    }
  }catch(e){
    console.error('Error al agregar al carrito', e);
    showToast('No se pudo agregar al carrito (' + (e && e.message ? e.message : 'error') + '). Recargá con Ctrl+F5.');
  }
}

async function changeQty(productId, qty){
  const data = await apiCall('/api/carrito/actualizar', {product_id: productId, qty});
  if(data.error){ showToast(data.error); return; }
  renderCart(data);
}

function openCart(){
  document.getElementById('cartDrawer').classList.add('open');
  document.getElementById('overlay').classList.add('show');
  refreshCart();
}
function closeCart(){
  document.getElementById('cartDrawer').classList.remove('open');
  document.getElementById('overlay').classList.remove('show');
}

/* ---------------- TOAST ---------------- */
let toastTimer;
function showToast(msg){
  const t = document.getElementById('toast');
  if(!t) return;
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2200);
}

function subscribeNews(e){
  e.preventDefault();
  showToast('¡Te suscribiste correctamente! 🎉');
  e.target.reset();
  return false;
}

document.addEventListener('DOMContentLoaded', refreshCart);
