/* =========================================================
   PREDICTIVO — BASE SIN GESTIÓN
   Columnas: CONTACTID · Prioridad · email · First_name
             telefono · Programa · number 1 · AgentName

   - CONTACTID  = Contador consecutivo (1, 2, 3…)
   - Prioridad  = 1 (siempre)
   - email      = Correo electrónico validado (solo chars [A-Za-z0-9@._\-,])
   - First_name = Primer nombre de Nombre completo
   - telefono   = Número DEF (10 dígitos normalizados)
   - Programa   = Primero "Programa de interes_", luego "Programa" / PROGRAMA2
   - number 1   = +579 + telefono
   - AgentName  = Email del asesor derivado de "Propietario de Posible Cliente"
   ========================================================= */

var sgPredPage = 1;
var SG_PRED_PAGE_SIZE = 100;
var SG_PRED_COLS = [
  'CONTACTID','Prioridad','email','First_name',
  'telefono','Programa','number 1','AgentName','ValCorreo'
];

function buildSGPredictivoRow(r, counter){
  var tel = String(r['Número DEF'] || '').trim();

  var nombre = String(r['Nombre completo'] || '').trim();
  var primerNombre = 'Sin nombre';
  if(typeof getFirstName === 'function'){
    primerNombre = getFirstName(nombre);
  } else {
    var partes = nombre.split(/\s+/);
    primerNombre = partes[0] || 'Sin nombre';
  }

  var rawEmail = String(r['Correo electrónico'] || '').trim();
  var validEmail = /^[A-Za-z0-9@._\-,]+$/.test(rawEmail) ? rawEmail : '';

  /* Programa: primero "Programa de interes_", luego "Programa", luego PROGRAMA2 — siempre MAYÚSCULAS */
  var rawProg = String(r['Programa de interes_'] || r['Programa'] || r['PROGRAMA2'] || '').trim();
  var programa = (rawProg || 'SIN PROGRAMA').toUpperCase();

  /* AgentName: email del asesor (MAP_ASESOR lookup + derivación de nombre) */
  var agentEmail = '';
  if(typeof getAsesorEmail === 'function'){
    agentEmail = getAsesorEmail(r['Propietario de Posible Cliente']);
  }

  return {
    'CONTACTID' : counter,
    'Prioridad' : 1,
    'email'     : validEmail,
    'First_name': primerNombre,
    'telefono'  : tel,
    'Programa'  : programa,
    'number 1'  : tel ? '+579' + tel : '',
    'AgentName' : agentEmail,
    'ValCorreo' : validEmail ? 'VALIDADO' : 'CORREO INVÁLIDO'
  };
}

function getSGPredictivoFiltrado(){
  return (typeof filteredData !== 'undefined' ? filteredData : [])
    .map(function(r, idx){ return buildSGPredictivoRow(r, idx + 1); });
}

function renderSGPredictivo(){
  var head    = document.getElementById('sg-pred-head');
  var body    = document.getElementById('sg-pred-body');
  var title   = document.getElementById('sg-pred-title');
  var pagInfo = document.getElementById('sg-pred-pag-info');
  if(!head || !body) return;

  var esc = typeof escapeHtml === 'function' ? escapeHtml : function(v){ return String(v ?? ''); };

  head.innerHTML = '<tr>' +
    SG_PRED_COLS.map(function(c){ return '<th>' + esc(c) + '</th>'; }).join('') +
  '</tr>';

  var data  = getSGPredictivoFiltrado();
  var total = data.length;
  var pages = Math.max(1, Math.ceil(total / SG_PRED_PAGE_SIZE));
  sgPredPage = Math.min(sgPredPage, pages);
  var sl = data.slice((sgPredPage - 1) * SG_PRED_PAGE_SIZE, sgPredPage * SG_PRED_PAGE_SIZE);

  if(title)   title.textContent   = total.toLocaleString() + ' registros predictivo';
  if(pagInfo) pagInfo.textContent = 'Página ' + sgPredPage + ' de ' + pages + ' · ' + total.toLocaleString() + ' registros';

  body.innerHTML = sl.map(function(r){
    return '<tr>' + SG_PRED_COLS.map(function(c){
      var v = String(r[c] != null ? r[c] : '');
      return '<td title="' + esc(v) + '">' + esc(v) + '</td>';
    }).join('') + '</tr>';
  }).join('');
}

function changeSGPredPage(n){
  var data  = getSGPredictivoFiltrado();
  var pages = Math.max(1, Math.ceil(data.length / SG_PRED_PAGE_SIZE));
  sgPredPage = Math.min(Math.max(1, sgPredPage + n), pages);
  renderSGPredictivo();
}

function exportSGPredictivo(){
  var rows = getSGPredictivoFiltrado();
  if(!rows.length){ showToast('No hay datos para exportar. Carga primero la base.'); return; }
  var ws = XLSX.utils.json_to_sheet(rows, { header: SG_PRED_COLS });
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Predictivo');
  XLSX.writeFile(wb, 'Predictivo_Sin_Gestion_' + Date.now() + '.xlsx');
  showToast('⬇ ' + rows.length.toLocaleString() + ' registros exportados');
}

function copiarSGPredictivoContactIds(){
  var data = getSGPredictivoFiltrado();
  if(!data.length){ showToast('No hay registros para copiar.'); return; }
  navigator.clipboard
    .writeText(data.map(function(r){ return r['CONTACTID']; }).join('\n'))
    .then(function(){ showToast('📋 ' + data.length.toLocaleString() + ' CONTACTID copiados'); })
    .catch(function(){ showToast('Error al copiar al portapapeles.'); });
}

/* ── Enganche en renderTable para actualizarse con los filtros ──
   Se aplica directamente al cargar el script (después de todos los hooks previos).   */
(function(){
  if(window.__sgPredHookApplied) return;
  window.__sgPredHookApplied = true;
  var _prevRenderTable = window.renderTable;
  if(typeof _prevRenderTable === 'function'){
    window.renderTable = function(){
      _prevRenderTable.apply(this, arguments);
      try{ renderSGPredictivo(); }catch(e){ console.warn('SG Predictivo:', e); }
    };
  }
})();
