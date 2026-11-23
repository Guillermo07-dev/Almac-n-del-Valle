/* Protección CSRF: el servidor entrega un token en la etiqueta <meta name="csrf-token">
   y exige devolverlo en toda petición que modifica datos. */

function csrfToken(){
  const meta = document.querySelector('meta[name="csrf-token"]');
  return meta ? meta.content : '';
}

/* Todas las llamadas fetch() que no son de lectura llevan el token en la cabecera
   X-CSRFToken, así no hay que acordarse de agregarlo en cada una. */
(function protegerFetch(){
  const fetchOriginal = window.fetch;
  window.fetch = function(url, opciones){
    opciones = opciones || {};
    const metodo = (opciones.method || 'GET').toUpperCase();
    if(metodo !== 'GET' && metodo !== 'HEAD'){
      opciones.headers = Object.assign({}, opciones.headers, {'X-CSRFToken': csrfToken()});
    }
    return fetchOriginal.call(window, url, opciones);
  };
})();
