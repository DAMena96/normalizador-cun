/* =========================================================
   RESTAURAR BOTONES DETALLE — FINAL
   No cambia la lógica de distribución.
   Solo garantiza que en la sección 4 de Sin Gestión,
   Interesados y No Contactado esté visible el botón:
   📸 Exportar detalle asesor JPG
   ========================================================= */
(function(){
  function crearBotonDetalle({cardId, buttonId, onclickName}){
    const card = document.getElementById(cardId);
    if(!card) return;

    // Evita duplicar el botón en la cabecera de la sección 4.
    if(document.getElementById(buttonId)) return;

    const wrap = document.createElement('div');
    wrap.className = 'acts detalle-png-actions detalle-jpg-actions';
    wrap.style.marginTop = '10px';
    wrap.innerHTML = `<button id="${buttonId}" class="btn btn-green" onclick="${onclickName}()">📸 Exportar detalle asesor JPG</button>`;

    const p = card.querySelector('p');
    if(p) p.insertAdjacentElement('afterend', wrap);
    else card.insertAdjacentElement('afterbegin', wrap);
  }

  function asegurarBotonesDetalle(){
    crearBotonDetalle({
      cardId: 'sg-detalle-asesor-card',
      buttonId: 'btn-export-detalle-asesor-jpg-sg-final',
      onclickName: 'exportDetalleAsesorSGJPG'
    });

    crearBotonDetalle({
      cardId: 'inter-detalle-asesor-card',
      buttonId: 'btn-export-detalle-asesor-jpg-inter-final',
      onclickName: 'exportDetalleAsesorInterJPG'
    });

    crearBotonDetalle({
      cardId: 'nc-detalle-asesor-card',
      buttonId: 'btn-export-detalle-asesor-jpg-nc-final',
      onclickName: 'exportDetalleAsesorNCJPG'
    });
  }

  // Refuerzo para que no se pierdan por renders posteriores.
  const oldDistribuir = window.distribuir;
  window.distribuir = function(){
    if(typeof oldDistribuir === 'function') oldDistribuir.apply(this, arguments);
    [80, 250, 700].forEach(ms => setTimeout(asegurarBotonesDetalle, ms));
  };

  const oldInter = window.distribuirInteresados;
  window.distribuirInteresados = function(){
    if(typeof oldInter === 'function') oldInter.apply(this, arguments);
    [80, 250, 700].forEach(ms => setTimeout(asegurarBotonesDetalle, ms));
  };

  const oldNC = window.distribuirNC;
  window.distribuirNC = function(){
    if(typeof oldNC === 'function') oldNC.apply(this, arguments);
    [80, 250, 700].forEach(ms => setTimeout(asegurarBotonesDetalle, ms));
  };

  document.addEventListener('DOMContentLoaded', () => {
    [100, 500, 1200, 2200].forEach(ms => setTimeout(asegurarBotonesDetalle, ms));
  });
})();
