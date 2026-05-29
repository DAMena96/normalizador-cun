
/* ════════════════════════════════════════════
   BASE DE INTERESADOS + PREDICTIVO
   ════════════════════════════════════════════ */

let interData = [], interFiltered = [], predictivoData = [];
let interAsesoresList = [], interDistribuidoData = [];
let interSortCol = '', interSortAsc = true, interPage = 1;
let predPage = 1;
const INTER_PAGE_SIZE = 100;

const INTER_BASE_COLS = [
  'ID de registro','Hora de creación','FECHA CREACION','Correo electrónico',
  'Estado de Posible Cliente','Etapa del Registro','Propietario de Posible Cliente',
  'Hora de la última actividad','Nombre completo','Campaña mercadeo','Teléfono',
  'Número de Documento','Programa de interes_','Etiqueta','Canal de gestión del lead',
  'Modificado por','Número de gestiones','Modalidad','Periodo','Creado por',
  'Ultimo nivel de escolaridad','Ciudad de residencia','Hora de modificación',
  'Sub Estado','Sub Estado II','Gestiones del Lead','Total Notas','Fecha ultima nota',
  'utm_campaign','Ciudad utm','utm_content','utm_medium','utm_source','utm_term',
  'conv','Descripción','Descripción actualizacion','Días sin contacto','Programa',
  'PROGRAMA2','CIUDAD','AREA','AREA VALIDA','VAL','SUPERVISOR'
];

const INTER_SHOW_COLS = [
  'ID de registro','Nombre completo','Correo electrónico','Teléfono',
  'PROGRAMA2','CIUDAD','AREA','AREA VALIDA','VAL','SUPERVISOR',
  'FECHA CREACION','Hora de modificación','Campaña mercadeo','Periodo','Creado por'
];

const PRED_COLS = ['CONTACTID','Prioridad','email','First_name','telefono','Programa','number 1','AgentName','ValCorreo'];

const INTER_FILTER_DEFS = [
  {id:'if-area-valida', col:'AREA VALIDA', lbl:'Área válida'},
  {id:'if-supervisor',  col:'SUPERVISOR',  lbl:'Supervisor'},
];

function handleDropInteresados(e){
  e.preventDefault();
  document.getElementById('inter-upload-zone').classList.remove('over');
  loadInteresadosFile(e.dataTransfer.files[0]);
}

function interShowProg(pct, label){
  document.getElementById('inter-progress-wrap').style.display = 'block';
  document.getElementById('inter-progress-fill').style.width = pct + '%';
  document.getElementById('inter-progress-label').textContent = label;
}

function interHideProg(){
  document.getElementById('inter-progress-wrap').style.display = 'none';
  document.getElementById('inter-progress-fill').style.width = '0';
}

async function loadInteresadosFile(file){
  if(!file) return;

  interShowProg(5,'Cargando catálogos JSON…');
  const okCatalogs = await catalogsReady;
  if(!okCatalogs){ interHideProg(); return; }

  interShowProg(15,'Leyendo archivo…');
  const isCSV = /\.csv$/i.test(file.name);
  const reader = new FileReader();

  reader.onload = ev => {
    interShowProg(35, isCSV ? 'Parseando CSV…' : 'Parseando Excel…');
    setTimeout(()=>{
      try{
        const wb = isCSV
          ? XLSX.read(ev.target.result, {type:'string', raw:false})
          : XLSX.read(ev.target.result, {type:'array'});
        const sn = wb.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});

        if(json.length < 2){
          showToast('Sin datos en la base de interesados.');
          interHideProg();
          return;
        }

        const headers = json[0].map(h => String(h).trim());
        interShowProg(60,`Normalizando ${(json.length-1).toLocaleString()} interesados…`);
        setTimeout(()=>{
          try{
            interData = [];
            for(let i=1;i<json.length;i++){
              const row = json[i];
              if(row.every(c => c === '' || c === null || c === undefined)) continue;
              interData.push(normalizeInteresadoRow(row, headers));
            }

            predictivoData = interData.map(buildPredictivoRow);
            interFiltered = [...interData];

            interShowProg(90,'Preparando…');
            setTimeout(()=>{
              buildInterFilterPanel();
              populateInterFilters();
              renderInterKPIs();
              renderPredictivo();
              renderInterTable();

              document.getElementById('inter-upload-zone').classList.add('hidden');
              document.getElementById('inter-kpi-row').classList.remove('hidden');
              document.getElementById('inter-tabs').classList.remove('hidden');
              document.querySelectorAll('.inter-panel').forEach(p => p.classList.remove('hidden'));
              switchInterTab(0);

              interHideProg();
              showToast(`✅ ${interData.length.toLocaleString()} interesados procesados`);
            },80);
          }catch(err){
            console.error(err);
            showToast('Error interesados: ' + err.message);
            interHideProg();
          }
        },50);
      }catch(err){
        console.error(err);
        showToast('Error interesados: ' + err.message);
        interHideProg();
      }
    },30);
  };

  if(isCSV) reader.readAsText(file, 'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function normalizeInteresadoRow(row, headers){
  const hIndex = new Map(headers.map((h, i) => [cleanHeader(h), i]));
  const g = (...names) => {
    for(const name of names){
      const i = hIndex.get(cleanHeader(name));
      if(i !== undefined && row[i] != null) return String(row[i]).trim();
    }
    return '';
  };

  const r = {};
  const baseCols = [
    'ID de registro','Hora de creación','Correo electrónico','Estado de Posible Cliente',
    'Etapa del Registro','Propietario de Posible Cliente','Hora de la última actividad',
    'Nombre completo','Campaña mercadeo','Teléfono','Número de Documento',
    'Programa de interes_','Etiqueta','Canal de gestión del lead','Modificado por',
    'Número de gestiones','Modalidad','Periodo','Creado por','Ultimo nivel de escolaridad',
    'Ciudad de residencia','Hora de modificación','Sub Estado','Sub Estado II',
    'Gestiones del Lead','Total Notas','Fecha ultima nota','utm_campaign','Ciudad utm',
    'utm_content','utm_medium','utm_source','utm_term','conv','Descripción',
    'Descripción actualizacion','Días sin contacto','Programa'
  ];

  baseCols.forEach(c => r[c] = g(c, sinTildes(c)));

  // PROGRAMA para Base Interesados:
  // Primero se arma como en Predictivo:
  // NOMPROPIO(SI(Programa=""; Programa de interes_; Programa))
  const programaPredictivo = getProgramaPredictivoFromRow(r);

  // Luego ese Programa se cruza con programas.json para normalizar.
  const programaNormalizado = lookup(MAP_PROGRAMA, programaPredictivo, programaPredictivo || 'SIN PROGRAMA') || 'SIN PROGRAMA';
  r['PROGRAMA'] = programaPredictivo;
  r['PROGRAMA2'] = programaNormalizado;

  // AREA se cruza desde areas.json usando el programa normalizado.
  const area = cleanCatalogValue(lookup(MAP_AREA, programaNormalizado, 'OTRO') || 'OTRO');
  r['AREA'] = area;

  // Ciudad: si Ciudad utm está vacía usa Ciudad de residencia; si es 0/vacía => SIN CIUDAD; BUSCARV VAL!L:M; MAYUSC
  const ciudadSrc = emptyOrZero(r['Ciudad utm']) ? r['Ciudad de residencia'] : r['Ciudad utm'];
  const ciudadKey = emptyOrZero(ciudadSrc) ? 'SIN CIUDAD' : ciudadSrc;
  r['CIUDAD'] = String(lookup(MAP_CIUDAD, ciudadKey, ciudadKey || 'SIN CIUDAD')).toUpperCase();

  // VAL: BUSCARV(Programa normalizado; areas.json; 2; FALSO)
  // En este módulo VAL es el área original del programa: MIXTO / PRESENCIAL / VIRTUAL / POSGRADO / OTRO.
  r['VAL'] = area;

  // Área válida:
  // SI(AREA="MIXTO";
  //    SI(conv<>""; SI(conv contiene "V";"VIRTUAL";"PRESENCIAL");
  //       SI(Periodo contiene "V";"VIRTUAL";"PRESENCIAL"));
  //    AREA)
  // AREA viene de areas.json. Esta regla solo aplica cuando AREA = MIXTO.
  r['AREA VALIDA'] = calcularAreaValida(r['VAL'], r['conv'], r['Periodo']);

  // Supervisor: igual que Sin Gestión → supervisorSimpleLookup(email_asesor)
  const _interAsesorEmail = (typeof getAsesorEmail === 'function')
    ? getAsesorEmail(r['Propietario de Posible Cliente'])
    : String(r['Propietario de Posible Cliente'] || '').trim();
  r['SUPERVISOR'] = (typeof supervisorSimpleLookup === 'function')
    ? supervisorSimpleLookup(_interAsesorEmail)
    : '';

  // Hora creación truncada: fecha sin hora
  r['FECHA CREACION'] = truncDate(r['Hora de creación']);

  return r;
}


function getProgramaPredictivoFromRow(r){
  const rawPrograma = String(r['Programa'] || '').trim();
  const rawInteres = String(r['Programa de interes_'] || '').trim();
  const src = rawPrograma ? rawPrograma : rawInteres;
  return toTitle(src || 'Sin programa');
}

/* ── AgentName desde "Propietario de Posible Cliente" ────────────────────────
   Estrategia en dos pasos:
   1. MAP_ASESOR lookup (asesores.json):
      - Si existe un mapeo y el valor resultante es DISTINTO al original
        → devolver exactamente ese valor (email, "Retirado", cualquier categoría).
      - Si el valor mapeado es IGUAL al original (sin mapeo explícito) → paso 2.
      - Si no se encontró en el mapa → paso 2.
   2. Derivación del nombre (Vinculaciones / CUNdigitales sin email en asesores.json):
      - Si el campo ya trae un email inline → extraerlo.
      - Todo mayúsculas → entrada institucional → ''.
      - Quitar prefijo conocido; con las palabras restantes construir email:
          1 pal → nombre@cun.edu.co
          2–3 pal → nombre_apellido1@cun.edu.co
          4+ pal → nombre_apellido1@cun.edu.co (salteando segundo nombre)
*/
function getAsesorEmail(propietario){
  if(!propietario) return '';
  var s = String(propietario).trim();
  if(!s) return '';

  // ── Paso 1: MAP_ASESOR lookup ────────────────────────────────────────────
  if(typeof lookup === 'function' && typeof MAP_ASESOR !== 'undefined'){
    var mapped = lookup(MAP_ASESOR, s, null);
    if(mapped !== null && mapped !== undefined){
      var mappedStr = String(mapped).trim();
      // Resultado distinto al original → mapeo explícito → devolver tal cual
      // (email, "Retirado", categoría de proceso, etc.)
      if(mappedStr !== s) return mappedStr;
      // Resultado igual al original (self-map):
      //   • Si tiene prefijo de nombre personal (Vinculaciones, Contact…) → paso 2 (derivar email)
      //   • Si NO tiene prefijo → es una entidad conocida (CUN Digital, Telecampus…) → devolver tal cual
      var hasPersonPrefix = /^(?:Vinculaciones|Vinculacion|Contact|Contatc|CUNdigitales?)\s/i.test(s);
      if(!hasPersonPrefix) return mappedStr;
      // Tiene prefijo → caer a derivación de nombre
    }
  }

  // ── Paso 2: Derivación del nombre ────────────────────────────────────────
  // Si ya contiene un email directamente
  var em = s.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  if(em) return em[0].toLowerCase();

  // Todo mayúsculas → institucional
  if(s === s.toUpperCase() && /^[A-Z0-9\s\-_.]+$/.test(s)) return '';

  // Quitar prefijo conocido (orden de mayor a menor especificidad)
  var prefixRE = /^(?:Vinculaciones\s+(?:Regionales|Santa\s+Marta)?\s*|Vinculacion\s+Regional\s+|CUNdigitales?\s+|Contact\s+|Contatc\s+)/i;
  var name = s.replace(prefixRE, '').trim();

  // Sin prefijo reconocido → no generamos email (evita Telecampus, Proceso, etc.)
  if(!name || name === s) return '';

  var words = name.split(/\s+/).filter(function(w){ return w.length > 0; });
  if(!words.length) return '';

  var diacRE = new RegExp('[\\u0300-\\u036f]', 'g');
  var norm = function(w){
    return w.normalize('NFD').replace(diacRE, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  };

  var fn = norm(words[0]);
  if(!fn) return '';

  var ln = '';
  if(words.length === 1){
    return fn + '@cun.edu.co';
  } else if(words.length <= 3){
    ln = norm(words[1]);          // nombre apellido1 [apellido2]
  } else {
    ln = norm(words[2]);          // nombre 2doNombre apellido1 [apellido2]
  }

  return ln ? fn + '_' + ln + '@cun.edu.co' : fn + '@cun.edu.co';
}

function buildPredictivoRow(r, counter){
  var tel = normalizarTelefono(r['Teléfono']);
  var rawEmail = String(r['Correo electrónico'] || '').trim();
  var validEmail = /^[A-Za-z0-9@._\-,]+$/.test(rawEmail) ? rawEmail : '';
  var programa = (String(r['Programa de interes_'] || r['Programa'] || '').trim() || 'SIN PROGRAMA').toUpperCase();

  return {
    'CONTACTID' : counter,
    'Prioridad' : 1,
    'email'     : validEmail,
    'First_name': getFirstName(r['Nombre completo']),
    'telefono'  : tel,
    'Programa'  : programa,
    'number 1'  : tel ? '+579' + tel : '',
    'AgentName' : getAsesorEmail(r['Propietario de Posible Cliente']),
    'ValCorreo' : validEmail ? 'VALIDADO' : 'CORREO INVÁLIDO'
  };
}


function cleanCatalogValue(v){
  const s = String(v ?? '').trim();
  if(!s) return '-';
  return s.toUpperCase();
}

function calcularAreaValida(area, conv, periodo){
  const a = String(area || '').toUpperCase();

  // Esta regla solo aplica cuando el AREA viene como MIXTO.
  // Si conv tiene valor, se revisa conv.
  // Ejemplo: conv = 2023C => PRESENCIAL porque no contiene V.
  // Si conv está vacío, se revisa Periodo.
  // Ejemplo: Periodo = 26V03 => VIRTUAL.
  if(a === 'MIXTO'){
    const c = String(conv || '').trim().toUpperCase();
    const p = String(periodo || '').trim().toUpperCase();

    if(c){
      return c.includes('V') ? 'VIRTUAL' : 'PRESENCIAL';
    }

    if(p){
      return p.includes('V') ? 'VIRTUAL' : 'PRESENCIAL';
    }

    return 'PRESENCIAL';
  }

  return cleanCatalogValue(area || 'OTRO');
}

function truncDate(value){
  if(!value) return '';
  const s = String(value).trim();

  // Excel puede traer fechas tipo 44927.75
  if(/^\d+(\.\d+)?$/.test(s)){
    const n = Number(s);
    if(n > 25000 && n < 60000){
      const d = new Date(Math.round((n - 25569) * 86400 * 1000));
      return d.toISOString().slice(0,10);
    }
  }

  // 2025-03-18 14:30 → 2025-03-18
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if(m){
    const y = m[1], mo = String(m[2]).padStart(2,'0'), da = String(m[3]).padStart(2,'0');
    return `${y}-${mo}-${da}`;
  }

  // 18/03/2025 14:30 → 2025-03-18
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if(m){
    const da = String(m[1]).padStart(2,'0');
    const mo = String(m[2]).padStart(2,'0');
    let y = String(m[3]);
    if(y.length === 2) y = '20' + y;
    return `${y}-${mo}-${da}`;
  }

  const d = new Date(s);
  if(!isNaN(d)) return d.toISOString().slice(0,10);

  return s.split(' ')[0] || s;
}

function getFirstName(fullName){
  let s = String(fullName || '').trim();
  if(!s) return 'Sin nombre';

  // Si empieza con caracteres raros o no latinos, similar al control CODIGO>=127
  if(s.charCodeAt(0) >= 127) return 'Sin nombre';

  s = toTitle(s);
  const m = s.match(/[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+/);
  return m ? m[0] : 'Sin nombre';
}

function toTitle(s){
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g,' ')
    .trim()
    .replace(/\b([a-záéíóúñ])/g, c => c.toUpperCase());
}

function sinTildes(s){
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

function emptyOrZero(v){
  const s = String(v ?? '').trim();
  return s === '' || s === '0';
}

function switchInterTab(n){
  document.querySelectorAll('#inter-tabs .tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('#interesados-module .inter-panel').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n === 2) updateInterDistInfo();
  if(n === 3) updateInterCariInfo();
  if(n === 4) updateInterNotasInfo();
}

function buildInterFilterPanel(){
  const row = document.getElementById('inter-filter-row');
  row.innerHTML = '';

  INTER_FILTER_DEFS.forEach(f => {
    const div = document.createElement('div');
    div.className = 'fg';
    div.innerHTML = `
      <label>${f.lbl}<span class="clr" onclick="clearOneInter('${f.id}')">✖</span></label>
      <div class="ms-wrap">
        <button class="ms-btn" onclick="toggleMS('${f.id}',event)">
          <span class="st" id="txt-${f.id}">Todas</span>
          <span style="opacity:.4;font-size:.62rem;flex-shrink:0;margin-left:2px">▼</span>
        </button>
        <div class="ms-dd" id="${f.id}">
          <div class="ms-sb"><input type="text" placeholder="Buscar…" oninput="filterOpts('${f.id}',this.value)"></div>
          <div class="ms-all" onclick="toggleAllInter('${f.id}')">☑ Sel/Des todo</div>
          <div class="ms-list" id="list-${f.id}"></div>
        </div>
      </div>`;
    row.appendChild(div);
  });

  /* ── Rango de fechas por Hora de modificación ── */
  const dateDiv = document.createElement('div');
  dateDiv.className = 'fg';
  dateDiv.style.minWidth = '280px';
  dateDiv.innerHTML = `
    <label>Hora de modificación<span class="clr" onclick="clearInterDateFilter()">✖</span></label>
    <div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap;">
      <input type="date" id="if-fecha-desde"
             style="flex:1;min-width:120px;padding:5px 7px;border:1.5px solid #c5cfd9;border-radius:6px;font-size:.82rem;"
             oninput="applyInterFilters()" title="Desde">
      <span style="color:#888;font-size:.8rem;flex-shrink:0;">–</span>
      <input type="date" id="if-fecha-hasta"
             style="flex:1;min-width:120px;padding:5px 7px;border:1.5px solid #c5cfd9;border-radius:6px;font-size:.82rem;"
             oninput="applyInterFilters()" title="Hasta">
    </div>`;
  row.appendChild(dateDiv);

  const txt = document.createElement('div');
  txt.className = 'fg';
  txt.style.minWidth = '190px';
  txt.innerHTML = `<label>Buscar nombre/correo/tel/ID</label>
    <input type="text" id="if-buscar" placeholder="Escribe para buscar…" oninput="applyInterFilters()">`;
  row.appendChild(txt);

  const clrb = document.createElement('div');
  clrb.className = 'fg';
  clrb.style.flex = '0';
  clrb.style.minWidth = 'auto';
  clrb.innerHTML = `<label>&nbsp;</label><button class="btn btn-gray" onclick="clearInterFilters()">✖ Limpiar</button>`;
  row.appendChild(clrb);
}

function clearInterDateFilter(){
  const d = document.getElementById('if-fecha-desde');
  const h = document.getElementById('if-fecha-hasta');
  if(d) d.value = '';
  if(h) h.value = '';
  applyInterFilters();
}

function populateInterFilters(){
  INTER_FILTER_DEFS.forEach(f => {
    const vals = [...new Set(interData.map(r => String(r[f.col] || '').trim()).filter(v => v && v !== '-'))]
      .sort((a,b)=>a.localeCompare(b));
    document.getElementById('list-' + f.id).innerHTML = vals.map(v =>
      `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(v)}" onchange="onCheckInter('${f.id}')"> <span>${escapeHtml(v)}</span></label>`
    ).join('');
    updateInterFilterText(f.id);
  });
}

function onCheckInter(id){
  applyInterFilters();
  updateInterFilterText(id);
  closeAll();
}

function toggleAllInter(id){
  const list = document.getElementById('list-' + id);
  const checks = [...list.querySelectorAll('input[type=checkbox]')];
  const all = checks.every(c => c.checked);
  checks.forEach(c => c.checked = !all);
  applyInterFilters();
  updateInterFilterText(id);
  closeAll();
}

function clearOneInter(id){
  document.querySelectorAll('#list-' + id + ' input').forEach(c => c.checked = false);
  applyInterFilters();
  updateInterFilterText(id);
  closeAll();
}

function clearInterFilters(){
  INTER_FILTER_DEFS.forEach(f => {
    document.querySelectorAll('#list-' + f.id + ' input').forEach(c => c.checked = false);
    updateInterFilterText(f.id);
  });
  const buscar = document.getElementById('if-buscar');
  if(buscar) buscar.value = '';
  const desde = document.getElementById('if-fecha-desde');
  const hasta  = document.getElementById('if-fecha-hasta');
  if(desde) desde.value = '';
  if(hasta)  hasta.value  = '';
  applyInterFilters();
}

function updateInterFilterText(id){
  const checks = [...document.querySelectorAll('#list-' + id + ' input:checked')];
  const txt = document.getElementById('txt-' + id);
  if(!txt) return;
  txt.textContent = checks.length ? (checks.length === 1 ? checks[0].value : `${checks.length} seleccionados`) : 'Todas';
}

function applyInterFilters(){
  const q     = String(document.getElementById('if-buscar')?.value || '').toLowerCase().trim();
  const desde = document.getElementById('if-fecha-desde')?.value || '';
  const hasta  = document.getElementById('if-fecha-hasta')?.value  || '';

  interFiltered = interData.filter(r => {
    for(const f of INTER_FILTER_DEFS){
      const vals = [...document.querySelectorAll('#list-' + f.id + ' input:checked')].map(c => c.value);
      updateInterFilterText(f.id);
      if(vals.length && !vals.includes(String(r[f.col] || ''))) return false;
    }

    /* Rango de fechas por Hora de modificación */
    if(desde || hasta){
      const modDate = truncDate(r['Hora de modificación'] || '');
      if(desde && modDate < desde) return false;
      if(hasta  && modDate > hasta) return false;
    }

    if(q){
      const hay = [
        r['ID de registro'], r['Nombre completo'], r['Correo electrónico'], r['Teléfono'],
        r['PROGRAMA2'], r['CIUDAD'], r['AREA VALIDA'], r['SUPERVISOR']
      ].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }

    return true;
  });

  interPage = 1;
  predPage = 1;
  renderInterKPIs();
  renderPredictivo();
  renderInterTable();
  updateInterDistInfo();
  if(typeof updateInterCariInfo  === 'function') updateInterCariInfo();
  if(typeof updateInterNotasInfo === 'function') updateInterNotasInfo();
}

function renderInterKPIs(){
  const d = interFiltered.length ? interFiltered : interData;
  document.getElementById('ik-total').textContent = interData.length.toLocaleString();
  document.getElementById('ik-pred').textContent = predictivoData.length.toLocaleString();
  document.getElementById('ik-virt').textContent = d.filter(r => r['AREA VALIDA'] === 'VIRTUAL').length.toLocaleString();
  document.getElementById('ik-pres').textContent = d.filter(r => r['AREA VALIDA'] === 'PRESENCIAL').length.toLocaleString();
  document.getElementById('ik-sin').textContent = d.filter(r => ['SIN PROGRAMA',''].includes(r['PROGRAMA2']) || ['SIN CIUDAD',''].includes(r['CIUDAD'])).length.toLocaleString();
}

function getPredictivoFiltrado(){
  var real = interFiltered.map(function(r, idx){ return buildPredictivoRow(r, idx + 2); });
  var first = typeof buildPredFirstRow === 'function' ? [buildPredFirstRow(real)] : [];
  return first.concat(real);
}

function renderPredictivo(){
  document.getElementById('pred-head').innerHTML = '<tr>' + PRED_COLS.map(c => `<th>${c}</th>`).join('') + '</tr>';

  const data = getPredictivoFiltrado();
  const total = data.length;
  const pages = Math.max(1, Math.ceil(total / INTER_PAGE_SIZE));
  predPage = Math.min(predPage, pages);
  const sl = data.slice((predPage-1)*INTER_PAGE_SIZE, predPage*INTER_PAGE_SIZE);

  document.getElementById('pred-title').textContent = `${total.toLocaleString()} registros predictivo`;
  document.getElementById('pred-body').innerHTML = sl.map(r =>
    '<tr>' + PRED_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>'
  ).join('');

  document.getElementById('pred-pag-info').textContent = `Página ${predPage} de ${pages}`;
}

function renderInterTable(){
  document.getElementById('inter-head').innerHTML = '<tr>' +
    INTER_SHOW_COLS.map(c => `<th onclick="sortInterBy('${c}')">${c} <span style="opacity:.35">${interSortCol===c?(interSortAsc?'▲':'▼'):'⇅'}</span></th>`).join('') +
    '</tr>';

  let data = [...interFiltered];
  if(interSortCol){
    data.sort((a,b) => {
      const va = String(a[interSortCol] || ''), vb = String(b[interSortCol] || '');
      return interSortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    });
  }

  const total = data.length;
  const pages = Math.max(1, Math.ceil(total / INTER_PAGE_SIZE));
  interPage = Math.min(interPage, pages);
  const sl = data.slice((interPage-1)*INTER_PAGE_SIZE, interPage*INTER_PAGE_SIZE);

  document.getElementById('inter-title').textContent = `${total.toLocaleString()} interesados normalizados`;
  document.getElementById('inter-body').innerHTML = sl.map(r =>
    '<tr>' + INTER_SHOW_COLS.map(c => {
      const v = String(r[c] ?? '');
      return `<td title="${escapeHtml(v)}">${getInterBadge(c,v) || escapeHtml(v)}</td>`;
    }).join('') + '</tr>'
  ).join('');

  document.getElementById('inter-pag-info').textContent = `Página ${interPage} de ${pages}`;
}

function getInterBadge(col, v){
  if(col === 'AREA VALIDA'){
    if(v === 'VIRTUAL') return `<span class="badge vir">VIRTUAL</span>`;
    if(v === 'PRESENCIAL') return `<span class="badge pre">PRESENCIAL</span>`;
    if(v === 'POSGRADO') return `<span class="badge pos">POSGRADO</span>`;
  }
  return '';
}

function sortInterBy(c){
  if(interSortCol === c) interSortAsc = !interSortAsc;
  else { interSortCol = c; interSortAsc = true; }
  renderInterTable();
}

function changeInterPage(n){
  const pages = Math.max(1, Math.ceil(interFiltered.length / INTER_PAGE_SIZE));
  interPage = Math.min(Math.max(1, interPage + n), pages);
  renderInterTable();
}

function changePredPage(n){
  const pages = Math.max(1, Math.ceil(getPredictivoFiltrado().length / INTER_PAGE_SIZE));
  predPage = Math.min(Math.max(1, predPage + n), pages);
  renderPredictivo();
}

function exportInterRows(rows, name, cols){
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    cols.forEach(c => o[c] = r[c] !== undefined ? r[c] : '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Interesados');
  XLSX.writeFile(wb, name);
  showToast(`⬇ ${rows.length.toLocaleString()} filas exportadas`);
}

function exportPredictivo(){
  exportInterRows(getPredictivoFiltrado(), 'Predictivo_Interesados_' + Date.now() + '.xlsx', PRED_COLS);
}

/* ── Descarga CSV con BOM UTF-8 para correcta apertura en Excel ── */
function downloadCSVFile(csv, filename){
  var blob = new Blob(['﻿' + csv], {type:'text/csv;charset=utf-8;'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function exportPredictivoCSV(){
  var rows = getPredictivoFiltrado();
  if(!rows.length){ showToast('No hay datos para exportar.'); return; }
  var ws = XLSX.utils.json_to_sheet(rows, {header: PRED_COLS});
  downloadCSVFile(XLSX.utils.sheet_to_csv(ws), 'Predictivo_Interesados_' + Date.now() + '.csv');
  showToast('⬇ CSV exportado: ' + rows.length.toLocaleString() + ' registros');
}

/* ── JPG agrupado por AgentName — función compartida entre los 3 módulos ── */
function exportPredictivoAgentJPG(rows, moduleLabel){
  if(!rows || !rows.length){ showToast('No hay datos para exportar.'); return; }

  // Agrupar por AgentName y contar
  var counts = {};
  rows.forEach(function(r){
    var agent = String(r['AgentName'] || '').trim() || 'Sin asignar';
    counts[agent] = (counts[agent] || 0) + 1;
  });
  var agentRows = Object.keys(counts)
    .map(function(a){ return {AgentName: a, Leads: counts[a]}; })
    .sort(function(a,b){ return b.Leads - a.Leads || a.AgentName.localeCompare(b.AgentName); });

  var total     = rows.length;
  var generated = new Date().toLocaleString('es-CO',{
    year:'numeric', month:'numeric', day:'numeric',
    hour:'numeric', minute:'2-digit', second:'2-digit'
  });

  var scale    = 4;
  var width    = 1200;
  var margin   = 34;
  var titleH   = 116;
  var summaryH = 86;   // tarjeta resumen (solo pág 1)
  var tableHdrH= 38;   // encabezado de columnas
  var rowH     = 34;
  var footerH  = 60;
  var MAX_PAGE_H = 2400; // px lógicos máx por página

  // Espacio fijo pág 1: header + tarjeta + padding + encabezado tabla + footer
  var fixedH1 = titleH + summaryH + 24 + tableHdrH + footerH + 24;
  // Espacio fijo págs siguientes: header + padding + encabezado tabla + footer
  var fixedHN = titleH + 24 + tableHdrH + footerH + 24;

  // ── Dividir agentRows en páginas ──
  var pages    = [];
  var curPage  = [];
  var curH     = fixedH1;

  agentRows.forEach(function(r){
    var fixed = curPage.length === 0 && pages.length === 0 ? fixedH1 : fixedHN;
    if(curPage.length > 0 && curH + rowH > MAX_PAGE_H){
      pages.push(curPage);
      curPage = [r];
      curH    = fixedHN + rowH;
    } else {
      curPage.push(r);
      curH += rowH;
    }
  });
  if(curPage.length) pages.push(curPage);

  var multiPage  = pages.length > 1;
  var moduleName = String(moduleLabel).replace(/\s+/g,'_');

  // ── Renderizar y descargar cada página ──
  function renderPage(pageRows, pageIdx){
    var isFirst = pageIdx === 0;
    var height  = isFirst
      ? Math.max(460, fixedH1 + pageRows.length * rowH)
      : Math.max(300, fixedHN + pageRows.length * rowH);

    var canvas = document.createElement('canvas');
    canvas.width  = width * scale;
    canvas.height = height * scale;
    var ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);

    function cutText(text, x, yy, maxW){
      var s = String(text || '');
      while(ctx.measureText(s).width > maxW && s.length > 4) s = s.slice(0, -2);
      if(s !== String(text || '')) s += '…';
      ctx.fillText(s, x, yy);
    }

    // Fondo
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);

    // Header
    ctx.fillStyle = '#0C2340';
    ctx.fillRect(0, 0, width, titleH);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 30px Segoe UI, Arial, sans-serif';
    var pageLabel = multiPage ? '  (' + (pageIdx+1) + ' / ' + pages.length + ')' : '';
    ctx.fillText('Predictivo por AgentName' + pageLabel, margin, 54);
    ctx.font = '15px Segoe UI, Arial, sans-serif';
    ctx.fillText(moduleLabel + '  ·  Generado: ' + generated, margin, 88);

    var y = titleH + 24;

    // Tarjeta resumen (solo pág 1)
    if(isFirst){
      ctx.fillStyle = '#f4f8fb';
      ctx.fillRect(margin, y, width - margin*2, 58);
      ctx.fillStyle = '#0C2340';
      ctx.font = 'bold 26px Segoe UI, Arial, sans-serif';
      ctx.fillText(total.toLocaleString(), margin + 22, y + 37);
      ctx.font = '14px Segoe UI, Arial, sans-serif';
      ctx.fillStyle = '#58677a';
      ctx.fillText('leads  ·  ' + agentRows.length.toLocaleString() + ' agentes', margin + 170, y + 36);
      y += summaryH + 14;
    }

    // Encabezado tabla
    ctx.fillStyle = '#1B365D';
    ctx.fillRect(margin, y, width - margin*2, tableHdrH);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
    ctx.fillText('AgentName', margin + 12, y + 25);
    ctx.fillText('Leads', width - margin - 70, y + 25);
    y += tableHdrH;

    // Filas
    pageRows.forEach(function(r, i){
      ctx.fillStyle = i % 2 ? '#ffffff' : '#f8fafc';
      ctx.fillRect(margin, y, width - margin*2, rowH);
      ctx.strokeStyle = '#e8eef5';
      ctx.beginPath();
      ctx.moveTo(margin, y + rowH); ctx.lineTo(width - margin, y + rowH);
      ctx.stroke();
      ctx.fillStyle = '#2c3e50';
      ctx.font = '14px Segoe UI, Arial, sans-serif';
      cutText(r.AgentName, margin + 12, y + 22, 920);
      ctx.fillStyle = '#0C2340';
      ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
      ctx.fillText(r.Leads.toLocaleString(), width - margin - 60, y + 22);
      y += rowH;
    });

    // Footer
    ctx.fillStyle = '#898D8D';
    ctx.font = '12px Segoe UI, Arial, sans-serif';
    ctx.fillText('Generado desde App Normalizador Contact CUN  ·  ' + generated, margin, height - 18);

    var suffix = multiPage ? '_P' + (pageIdx + 1) : '';
    var link   = document.createElement('a');
    link.href  = canvas.toDataURL('image/jpeg', 0.98);
    link.download = 'Predictivo_AgentName_' + moduleName + suffix + '.jpg';
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
  }

  pages.forEach(function(pg, idx){ setTimeout(function(){ renderPage(pg, idx); }, idx * 400); });

  var msg = multiPage
    ? '📸 Exportando ' + pages.length + ' imágenes (P1…P' + pages.length + ')  ·  ' + agentRows.length + ' agentes'
    : '📸 JPG exportado  ·  ' + agentRows.length.toLocaleString() + ' agentes';
  showToast(msg);
}

function exportPredictivoJPG(){
  exportPredictivoAgentJPG(getPredictivoFiltrado(), 'Base Interesados');
}

function exportInterFiltered(){
  exportInterRows(interFiltered, 'Base_Interesados_Filtrada_' + Date.now() + '.xlsx', INTER_BASE_COLS);
}

function exportInterAll(){
  exportInterRows(interData, 'Base_Interesados_Completa_' + Date.now() + '.xlsx', INTER_BASE_COLS);
}

function copiarPredictivoContactIds(){
  const data = getPredictivoFiltrado();
  navigator.clipboard.writeText(data.map(r => r['CONTACTID']).join('\n'))
    .then(() => showToast(`📋 ${data.length.toLocaleString()} CONTACTID copiados`));
}




/* ════════════════════════════════════════════
   DISTRIBUCIÓN — BASE DE INTERESADOS
   ════════════════════════════════════════════ */

function updateInterDistInfo(){
  const el = document.getElementById('inter-dist-info');
  if(el) el.textContent = `${interFiltered.length.toLocaleString()} leads listos para distribuir (filtro activo)`;
}

function dropInterAsesores(e){
  e.preventDefault();
  loadInterAsesores(e.dataTransfer.files[0]);
}

function loadInterAsesores(file){
  if(!file) return;

  const reader = new FileReader();
  reader.onload = ev => {
    try{
      const wb = XLSX.read(ev.target.result,{type:'array'});
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const col = Object.keys(json[0] || {}).find(k => /correo|email/i.test(k));

      if(!col){
        showToast('No se encontró columna Correo o Email');
        return;
      }

      interAsesoresList = [...new Set(json.map(r => String(r[col]).trim()).filter(Boolean))];
      renderInterAsesoresTags();
      showToast(`✅ ${interAsesoresList.length} asesores cargados`);
    }catch(err){
      showToast('Error asesores: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
}

function renderInterAsesoresTags(){
  const box = document.getElementById('inter-asesores-tags');
  if(!box) return;

  box.innerHTML =
    `<p style="font-size:.77rem;color:#666;margin-bottom:4px"><strong>${interAsesoresList.length} asesores:</strong></p>` +
    interAsesoresList.map((a,i) =>
      `<span class="tag">${escapeHtml(a)}<span class="rm" onclick="rmInterAsesor(${i})">×</span></span>`
    ).join('');
}

function rmInterAsesor(i){
  interAsesoresList.splice(i,1);
  renderInterAsesoresTags();
}

function getInterLeadId(row){
  if(!row) return '';

  const direct = row['ID de registro'];
  if(direct !== undefined && direct !== null && String(direct).trim() !== '') {
    return String(direct).trim();
  }

  const possibleNames = ['Id de registro','ID Registro','ID','Id','Record ID','Record Id','record_id','Lead ID','Lead Id'];

  for(const name of possibleNames){
    if(row[name] !== undefined && row[name] !== null && String(row[name]).trim() !== '') {
      return String(row[name]).trim();
    }
  }

  for(const key of Object.keys(row)){
    const clean = typeof cleanHeader === 'function'
      ? cleanHeader(key)
      : String(key).toUpperCase().replace(/[^A-Z0-9]/g,'');

    if(['IDDEREGISTRO','IDREGISTRO','RECORDID','LEADID','ID'].includes(clean)){
      const value = row[key];
      if(value !== undefined && value !== null && String(value).trim() !== '') {
        return String(value).trim();
      }
    }
  }

  return '';
}

function distribuirInteresados(){
  if(!interFiltered.length){
    showToast('No hay leads para distribuir.');
    return;
  }

  if(!interAsesoresList.length){
    showToast('Carga primero el Excel de asesores.');
    return;
  }

  const leadsOrdenados = ordenarPorHoraCreacionInterParaDistribuir(interFiltered);
  const n = interAsesoresList.length;
  interDistribuidoData = leadsOrdenados.map((lead, i) => {
    const asesor = interAsesoresList[i % n];
    return asignarSupervisorSimple({ 'ID de registro': getInterLeadId(lead) }, asesor);
  });

  renderInterDistPreview();
  showInterDistStats();
  renderDetalleAsesorInter();
  setTimeout(renderDetalleAsesorInter, 100);
  showToast(`✅ ${interDistribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}

function renderInterDistPreview(){
  const wrap = document.getElementById('inter-dist-preview');
  const tbl = document.getElementById('inter-dist-tbl');
  if(!wrap || !tbl) return;

  wrap.style.display = 'block';
  const preview = interDistribuidoData.slice(0,300);

  tbl.innerHTML =
    '<thead><tr><th>ID de registro</th><th>CORREO ASESOR ASIGNADO</th></tr></thead><tbody>' +
    preview.map(r =>
      `<tr><td>${escapeHtml(String(r['ID de registro'] || ''))}</td><td>${escapeHtml(String(r['CORREO ASESOR ASIGNADO'] || ''))}</td></tr>`
    ).join('') +
    '</tbody>';

  if(interDistribuidoData.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 registros de ${interDistribuidoData.length.toLocaleString()}</caption>`;
  }
}

function showInterDistStats(){
  const box = document.getElementById('inter-stat-box');
  if(!box) return;

  const counts = {};
  interDistribuidoData.forEach(r => {
    const a = r['CORREO ASESOR ASIGNADO'];
    counts[a] = (counts[a] || 0) + 1;
  });

  box.style.display = 'block';
  box.innerHTML = '<strong>Resumen de asignación:</strong><br>' +
    Object.entries(counts).map(([a,c]) => `<span class="dist-chip">${escapeHtml(a)}: <b>${c}</b></span>`).join('') +
    htmlDetalleVisible(interDistribuidoData,'Detalle_Asesor_Interesados');
}

function exportarDistInteresados(){
  if(!interDistribuidoData.length){
    showToast('Primero debes distribuir los leads.');
    return;
  }

  const rows = interDistribuidoData.map(r => ({
    'ID de registro': r['ID de registro'],
    'CORREO ASESOR ASIGNADO': r['CORREO ASESOR ASIGNADO']
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Distribucion');
  XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' INTER' + (typeof getInterFileLabel==='function'?getInterFileLabel():'') + '.xlsx');
  showToast(`⬇ ${interDistribuidoData.length.toLocaleString()} asignaciones exportadas`);
}

function copiarInterIDs(){
  navigator.clipboard.writeText(interFiltered.map(r => getInterLeadId(r)).join('\n'))
    .then(() => showToast(`📋 ${interFiltered.length.toLocaleString()} IDs copiados`));
}

function copiarInterDistIDs(){
  const data = interDistribuidoData.length ? interDistribuidoData : interFiltered.map(r => ({'ID de registro': getInterLeadId(r)}));
  navigator.clipboard.writeText(data.map(r => r['ID de registro']).join('\n'))
    .then(() => showToast(`📋 ${data.length.toLocaleString()} IDs copiados`));
}


// Refuerzo filtros interesados: cerrar cualquier filtro abierto al hacer clic fuera.
document.addEventListener('click', function(e){
  if(!e.target.closest('.ms-wrap')){
    document.querySelectorAll('.ms-dd.open').forEach(d => d.classList.remove('open'));
    const ov = document.getElementById('ov');
    if(ov) ov.classList.remove('on');
  }
});


function parseFechaDistribucionInter(v){
  if(!v) return 0;
  const s = String(v).trim();

  if(/^\d+(\.\d+)?$/.test(s)){
    const n = Number(s);
    if(n > 25000 && n < 60000) return Math.round((n - 25569) * 86400 * 1000);
  }

  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T](\d{1,2}):?(\d{1,2})?:?(\d{1,2})?)?/);
  if(m){
    return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)).getTime();
  }

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):?(\d{1,2})?:?(\d{1,2})?)?/);
  if(m){
    let y = String(m[3]);
    if(y.length === 2) y = '20' + y;
    return new Date(Number(y), Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)).getTime();
  }

  const d = new Date(s);
  return isNaN(d) ? 0 : d.getTime();
}

function ordenarPorHoraCreacionInterParaDistribuir(rows){
  return [...rows].sort((a,b) => {
    const fa = parseFechaDistribucionInter(a['Hora de creación'] || a['Hora de creacion'] || a['Created Time']);
    const fb = parseFechaDistribucionInter(b['Hora de creación'] || b['Hora de creacion'] || b['Created Time']);
    return fa - fb;
  });
}


/* === OVERRIDE FINAL: Detalle asesor por mentor — Base Interesados === */
function getMentorInter(row){
  const asesor = row['CORREO ASESOR ASIGNADO'] || '';
  return String(row['MENTOR'] || mentorDeAsesor(asesor) || 'Sin mentor').trim() || 'Sin mentor';
}

function renderDetalleAsesorInter(){
  const card = document.getElementById('inter-detalle-asesor-card');
  const box = document.getElementById('inter-detalle-asesor');
  const sel = document.getElementById('inter-mentor-filter');

  if(!card || !box || !sel) return;

  card.style.display = 'block';
  card.classList.add('detalle-visible');
  card.classList.remove('detalle-placeholder');

  const data = typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [];
  if(!data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el resumen por mentor y asesor.</div>';
    return;
  }

  const mentores = [...new Set(data.map(getMentorInter))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  const actual = sel.value || '';
  sel.innerHTML = '<option value="">Todos</option>' + mentores.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  if(mentores.includes(actual)) sel.value = actual;

  const filtro = sel.value;
  const rows = filtro ? data.filter(r => getMentorInter(r) === filtro) : data;

  const grouped = {};
  rows.forEach(r => {
    const mentor = getMentorInter(r);
    const asesor = r['CORREO ASESOR ASIGNADO'] || 'Sin asesor';
    if(!grouped[mentor]) grouped[mentor] = {};
    grouped[mentor][asesor] = (grouped[mentor][asesor] || 0) + 1;
  });

  const totalAsesores = new Set(rows.map(r => r['CORREO ASESOR ASIGNADO'] || 'Sin asesor')).size;
  const totalLeads = rows.length;

  box.innerHTML = `
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${Object.keys(grouped).length.toLocaleString()}</strong><span>Mentores</span></div>
    </div>
  ` + Object.entries(grouped).map(([mentor, asesores]) => {
    const total = Object.values(asesores).reduce((a,b)=>a+b,0);
    const cards = Object.entries(asesores).sort((a,b)=>b[1]-a[1]).map(([asesor,c]) =>
      `<div class="detalle-asesor-item"><span><b>${escapeHtml(asesor)}</b><small>${escapeHtml(areaDeAsesor(asesor))}</small></span><strong>${c.toLocaleString()}</strong></div>`
    ).join('');
    return `<div class="detalle-mentor-card">
      <div class="detalle-mentor-head"><span>👤 ${escapeHtml(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
      <div class="detalle-asesor-grid">${cards}</div>
    </div>`;
  }).join('');
}


/* === DETALLE ASESOR FUNCIONAL — Base Interesados === */
function getMentorInter(row){
  return String(row['MENTOR'] || mentorDeAsesor(row['CORREO ASESOR ASIGNADO']) || 'Sin supervisor').trim();
}

function renderDetalleAsesorInter(){
  renderDetalleAsesorGenerico({
    data: (typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []),
    cardId: 'inter-detalle-asesor-card',
    boxId: 'inter-detalle-asesor',
    selectId: 'inter-mentor-filter',
    getMentor: getMentorInter,
    renderAgain: 'renderDetalleAsesorInter'
  });
}


/* === OVERRIDE FINAL DISTRIBUIR INTERESADOS === */
function distribuirInteresados(){
  if(!interFiltered.length){showToast('No hay leads para distribuir.');return;}
  if(!interAsesoresList.length){showToast('Carga primero el Excel de asesores.');return;}
  const leadsOrdenados = ordenarPorHoraCreacionInterParaDistribuir(interFiltered);
  const n = interAsesoresList.length;
  interDistribuidoData = leadsOrdenados.map((lead,i)=>asignarSupervisorSimple({'ID de registro':getInterLeadId(lead)}, interAsesoresList[i % n]));
  renderInterDistPreview();
  showInterDistStats();
  renderDetalleAsesorInter();
  showToast(`✅ ${interDistribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}


/* ════════════════════════════════════════════
   CARI AI + NOTAS — BASE INTERESADOS
   ════════════════════════════════════════════ */

const INTER_CARI_COLS = ['numero_telefono','nombre_aspirante','carrera_interes','identificacion','correo_electronico','telefono_adicional','periodo','campania'];
const INTER_NOTAS_COLS = ['ID','Notes'];

const INTER_CARI_FIXED_ROWS = [
  {
    'numero_telefono': '573332322810',
    'nombre_aspirante': 'Esteban',
    'carrera_interes': 'especializacion en contratacion estatal',
    'identificacion': 'Sin documento',
    'correo_electronico': 'Prueba@gmail.com',
    'telefono_adicional': '573332322810',
    'periodo': 'Virtual/pregrado',
    'campania': 'Organico'
  },
  {
    'numero_telefono': '573134268590',
    'nombre_aspirante': 'Diguar',
    'carrera_interes': 'ingenieria de sistemas',
    'identificacion': 'Sin documento',
    'correo_electronico': 'Prueba@gmail.com',
    'telefono_adicional': '573134268590',
    'periodo': 'Virtual/pregrado',
    'campania': 'Organico'
  }
];

const INTER_NOTA_FIJA = 'Se intenta contacto con el aspirante via CARI AI';

function updateInterCariInfo(){
  const el = document.getElementById('inter-cari-info');
  if(el) el.textContent = interFiltered.length.toLocaleString() + ' registros listos para CARI AI (filtro activo)';
}

function updateInterNotasInfo(){
  const el = document.getElementById('inter-notas-info');
  if(el) el.textContent = interFiltered.length.toLocaleString() + ' notas listas para generar (filtro activo)';
}

function getInterCariPhone(row){
  const tel = normalizarTelefono(row['Telefono'] || row['Teléfono'] || '');
  if(!tel) return '';
  return tel.startsWith('57') ? tel : '57' + tel;
}

function getInterCariPeriodo(row){
  const area = String(row['AREA VALIDA'] || '').trim().toUpperCase();
  return area === 'POSGRADO' ? 'Virtual/Especializacion' : 'Virtual/pregrado';
}

function buildInterCariRows(){
  const rows = interFiltered.map(function(r){
    const phone = getInterCariPhone(r);
    const doc   = String(r['Número de Documento'] || r['Numero de Documento'] || '').trim();
    return {
      'numero_telefono'   : phone,
      'nombre_aspirante'  : r['Nombre completo'] || '',
      'carrera_interes'   : String(r['PROGRAMA NORMALIZADO'] || r['Programa de interes_'] || r['Programa'] || '').toLowerCase(),
      'identificacion'    : doc || 'Sin documento',
      'correo_electronico': r['Correo electrónico'] || r['Correo electronico'] || '',
      'telefono_adicional': phone,
      'periodo'           : getInterCariPeriodo(r),
      'campania'          : 'Organico'
    };
  });
  return INTER_CARI_FIXED_ROWS.concat(rows);
}

function buildInterNotasRows(){
  return interFiltered.map(function(r){
    return { 'ID': getInterLeadId(r), 'Notes': INTER_NOTA_FIJA };
  });
}

function renderInterCari(){
  const rows = buildInterCariRows();
  const wrap = document.getElementById('inter-cari-preview');
  const tbl  = document.getElementById('inter-cari-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0, 300);
  tbl.innerHTML = '<thead><tr>' + INTER_CARI_COLS.map(function(c){ return '<th>' + c + '</th>'; }).join('') + '</tr></thead><tbody>' +
    preview.map(function(r){ return '<tr>' + INTER_CARI_COLS.map(function(c){
      var v = escapeHtml(String(r[c] != null ? r[c] : ''));
      return '<td title="' + v + '">' + v + '</td>';
    }).join('') + '</tr>'; }).join('') + '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += '<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 registros de ' + rows.length.toLocaleString() + '</caption>';
  }
  updateInterCariInfo();
}

function renderInterNotas(){
  const rows = buildInterNotasRows();
  const wrap = document.getElementById('inter-notas-preview');
  const tbl  = document.getElementById('inter-notas-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0, 300);
  tbl.innerHTML = '<thead><tr>' + INTER_NOTAS_COLS.map(function(c){ return '<th>' + c + '</th>'; }).join('') + '</tr></thead><tbody>' +
    preview.map(function(r){ return '<tr>' + INTER_NOTAS_COLS.map(function(c){
      var v = escapeHtml(String(r[c] != null ? r[c] : ''));
      return '<td title="' + v + '">' + v + '</td>';
    }).join('') + '</tr>'; }).join('') + '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += '<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 notas de ' + rows.length.toLocaleString() + '</caption>';
  }
  updateInterNotasInfo();
}

function exportInterCari(){
  const rows = buildInterCariRows();
  if(!rows.length){ showToast('No hay registros CARI AI para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(function(r){
    const o = {};
    INTER_CARI_COLS.forEach(function(c){ o[c] = r[c] != null ? r[c] : ''; });
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CARI AI');
  var stamp = typeof getDateStamp === 'function' ? getDateStamp() : '';
  XLSX.writeFile(wb, 'INTERESADOS_' + stamp + '.xlsx');
  showToast('⬇ ' + rows.length.toLocaleString() + ' registros CARI AI exportados');
}

function exportInterNotas(){
  const rows = buildInterNotasRows();
  if(!rows.length){ showToast('No hay notas para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(function(r){
    const o = {};
    INTER_NOTAS_COLS.forEach(function(c){ o[c] = r[c] != null ? r[c] : ''; });
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  var stamp     = typeof getDateStamp     === 'function' ? getDateStamp()     : '';
  var fileLabel = typeof getInterFileLabel === 'function' ? getInterFileLabel() : '';
  XLSX.writeFile(wb, 'ASIGNACION ' + stamp + ' INTER' + fileLabel + '.xlsx');
  showToast('⬇ ' + rows.length.toLocaleString() + ' notas exportadas');
}

function copiarInterNotasIDs(){
  const rows = buildInterNotasRows();
  navigator.clipboard.writeText(rows.map(function(r){ return r['ID']; }).join('\n'))
    .then(function(){ showToast('📋 ' + rows.length.toLocaleString() + ' IDs copiados'); });
}
