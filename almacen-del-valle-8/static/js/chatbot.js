/* Asistente virtual de la tienda. El texto de las respuestas se inserta con textContent
   (nunca como HTML) y los botones solo aceptan enlaces internos del sitio. */

const chatEstado = {iniciado: false};

function toggleChat(){
  const panel = document.getElementById('chatPanel');
  panel.hidden = !panel.hidden;
  if(!panel.hidden){
    if(!chatEstado.iniciado){
      chatEstado.iniciado = true;
      chatAgregarBurbuja('¡Hola! Soy Alice, la asistente de Almacén del Valle. ¿En qué te puedo ayudar?', 'bot');
      chatMostrarSugerencias(['Horarios', 'Costo de envío', 'Medios de pago', 'Estado de mi pedido', 'Ofertas']);
    }
    document.getElementById('chatInput').focus();
  }
}

function chatAgregarBurbuja(texto, quien){
  const caja = document.getElementById('chatMsgs');
  const burbuja = document.createElement('div');
  burbuja.className = 'chat-bubble ' + quien;
  burbuja.textContent = texto;
  caja.appendChild(burbuja);
  caja.scrollTop = caja.scrollHeight;
  return burbuja;
}

function chatAgregarEnlaces(acciones){
  const validas = (acciones || []).filter(a => typeof a.url === 'string' && a.url.startsWith('/'));
  if(!validas.length) return;
  const caja = document.getElementById('chatMsgs');
  const fila = document.createElement('div');
  fila.className = 'chat-links';
  validas.forEach(a => {
    const enlace = document.createElement('a');
    enlace.className = 'chat-link';
    enlace.href = a.url;
    enlace.textContent = a.texto;
    fila.appendChild(enlace);
  });
  caja.appendChild(fila);
  caja.scrollTop = caja.scrollHeight;
}

function chatAgregarValoracion(mensajeId){
  const caja = document.getElementById('chatMsgs');
  const fila = document.createElement('div');
  fila.className = 'chat-rate';
  const pregunta = document.createElement('span');
  pregunta.textContent = '¿Te sirvió?';
  fila.appendChild(pregunta);
  [['👍', true, 'Me sirvió'], ['👎', false, 'No me sirvió']].forEach(([icono, sirvio, titulo]) => {
    const boton = document.createElement('button');
    boton.type = 'button';
    boton.textContent = icono;
    boton.title = titulo;
    boton.setAttribute('aria-label', titulo);
    boton.onclick = async () => {
      fila.textContent = sirvio ? '¡Gracias!' : 'Gracias, voy a aprender de esta.';
      try{
        await fetch('/api/chatbot/valoracion', {
          method: 'POST', headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({mensaje_id: mensajeId, sirvio: sirvio})
        });
      }catch(e){ /* la valoración es opcional */ }
    };
    fila.appendChild(boton);
  });
  caja.appendChild(fila);
  caja.scrollTop = caja.scrollHeight;
}

function chatMostrarSugerencias(lista){
  const contenedor = document.getElementById('chatChips');
  contenedor.replaceChildren();
  (lista || []).forEach(texto => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'chat-chip';
    chip.textContent = texto;
    chip.onclick = () => chatEnviar(texto);
    contenedor.appendChild(chip);
  });
}

async function chatEnviar(texto){
  texto = (texto || '').trim();
  if(!texto) return;
  chatMostrarSugerencias([]);
  chatAgregarBurbuja(texto, 'yo');
  const escribiendo = chatAgregarBurbuja('...', 'bot');
  try{
    const res = await fetch('/api/chatbot', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({mensaje: texto})
    });
    const data = await res.json();
    escribiendo.textContent = data.respuesta || data.error || 'No pude responder en este momento.';
    chatAgregarEnlaces(data.acciones);
    if(data.mensaje_id && data.resuelto) chatAgregarValoracion(data.mensaje_id);
    chatMostrarSugerencias(data.sugerencias);
  }catch(e){
    escribiendo.textContent = 'No tengo conexión en este momento. Probá de nuevo en unos segundos.';
  }
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('chatForm').addEventListener('submit', e => {
    e.preventDefault();
    const campo = document.getElementById('chatInput');
    const texto = campo.value;
    campo.value = '';
    chatEnviar(texto);
  });
});
