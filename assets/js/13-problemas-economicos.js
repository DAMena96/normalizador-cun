
/* ════════════════════════════════════════════
   BASE PROBLEMAS ECONÓMICOS + PREDICTIVO + DISTRIBUCIÓN
   Misma lógica de Base No Contactado
   ════════════════════════════════════════════ */

let peData = [], peFiltered = [], pePredictivoData = [];
let peAsesoresList = [], peDistribuidoData = [];
let peSortCol = '', peSortAsc = true, pePage = 1, pePredPage = 1;
const PE_PAGE_SIZE = 100;

const PE_BASE_COLS = [
  'ID de registro','Hora de creación','FECHA CREACION','Correo electrónico',
  'Estado de Posible Cliente','Etapa del Registro','Propietario de Posible Cliente',
  'Hora de la última actividad','Nombre completo','Campaña mercadeo','Teléfono',
  'Número de Documento','Programa de interes_','Etiqueta','Canal de gestión del lead',
  'Modificado por','Número de gestiones','Modalidad','Periodo','Creado por',
  'Ultimo nivel de escolaridad','Ciudad de residencia','Hora de modificación',
  'Sub Estado','Sub Estado II','Gestiones del Lead','Total Notas','Fecha ultima nota',
  'utm_campaign','Ciudad utm','utm_content','utm_medium','utm_source','utm_term',
  'conv','Descripción','Descripción actualizacion','Días sin contacto','Programa',
  'PROGRAMA NORMALIZADO','CIUDAD','AREA VALIDA','VAL','SUPERVISOR'
];

const PE_SHOW_COLS = [
  'ID de registro','Nombre completo','Correo electrónico','Teléfono',
  'PROGRAMA NORMALIZADO','CIUDAD','AREA VALIDA','VAL','SUPERVISOR',
  'FECHA CREACION','Hora de modificación','Campaña mercadeo','Periodo','Creado por'
];

const PE_PRED_COLS = ['CONTACTID','Prioridad','email','First_name','telefono','Programa','number 1','AgentName','ValCorreo'];

const PE_FILTER_DEFS = [
  {id:'pe-area-valida', col:'AREA VALIDA',  lbl:'Área válida'},
  {id:'pe-supervisor',  col:'SUPERVISOR',   lbl:'Supervisor'},
];

function handleDropPE(e){
  e.preventDefault();
  document.getElementById('pe-upload-zone').classList.remove('over');
  loadPEFile(e.dataTransfer.files[0]);
}

function peShowProg(pct,label){
  document.getElementById('pe-progress-wrap').style.display='block';
  document.getElementById('pe-progress-fill').style.width=pct+'%';
  document.getElementById('pe-progress-label').textContent=label;
}
function peHideProg(){
  document.getElementById('pe-progress-wrap').style.display='none';
  document.getElementById('pe-progress-fill').style.width='0';
}

async function loadPEFile(file){
  if(!file) return;
  peShowProg(5,'Cargando catálogos JSON…');
  const okCatalogs = await catalogsReady;
  if(!okCatalogs){ peHideProg(); return; }

  peShowProg(15,'Leyendo archivo…');
  const isCSV = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onload = ev => {
    peShowProg(35, isCSV ? 'Parseando CSV…' : 'Parseando Excel…');
    setTimeout(()=>{
      try{
        const wb = isCSV
          ? XLSX.read(ev.target.result, {type:'string', raw:false})
          : XLSX.read(ev.target.result, {type:'array'});
        const sn = wb.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
        if(json.length < 2){ showToast('Sin datos en la base problemas económicos.'); peHideProg(); return; }
        const headers = json[0].map(h => String(h).trim());

        peShowProg(60,`Normalizando ${(json.length-1).toLocaleString()} registros…`);
        setTimeout(()=>{
          try{
            peData = [];
            for(let i=1;i<json.length;i++){
              const row = json[i];
              if(row.every(c => c === '' || c === null || c === undefined)) continue;
              peData.push(normalizePERow(row, headers));
            }

            pePredictivoData = peData.map(buildPEPredictivoRow);
            peFiltered = [...peData];

            peShowProg(90,'Preparando…');
            setTimeout(()=>{
              buildPEFilterPanel();
              populatePEFilters();
              renderPEKPIs();
              renderPEPredictivo();
              renderPETable();

              document.getElementById('pe-upload-zone').classList.add('hidden');
              document.getElementById('pe-kpi-row').classList.remove('hidden');
              document.getElementById('pe-tabs').classList.remove('hidden');
              document.querySelectorAll('.pe-panel').forEach(p => p.classList.remove('hidden'));
              switchPETab(0);

              peHideProg();
              showToast(`✅ ${peData.length.toLocaleString()} registros problemas económicos procesados`);
            },80);
          }catch(err){
            console.error(err);
            showToast('Error Problemas Económicos: ' + err.message);
            peHideProg();
          }
        },50);
      }catch(err){
        console.error(err);
        showToast('Error Problemas Económicos: ' + err.message);
        peHideProg();
      }
    },30);
  };
  if(isCSV) reader.readAsText(file, 'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function normalizePERow(row, headers){
  const hIndex = new Map(headers.map((h,i)=>[cleanHeader(h),i]));
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
  baseCols.forEach(c => r[c] = g(c, sinTildesNC(c)));

  const programaPredictivo = getPEProgramaPredictivoFromRow(r);
  const programaNormalizado = cleanNCCatalog(lookup(MAP_PROGRAMA, programaPredictivo, programaPredictivo || 'SIN PROGRAMA') || 'SIN PROGRAMA');
  r['Programa'] = programaPredictivo;
  r['PROGRAMA NORMALIZADO'] = programaNormalizado;

  r['VAL'] = cleanNCCatalog(lookup(MAP_AREA, programaNormalizado, 'OTRO') || 'OTRO');

  const ciudadSrc = emptyOrZeroNC(r['Ciudad utm']) ? r['Ciudad de residencia'] : r['Ciudad utm'];
  const ciudadKey = emptyOrZeroNC(ciudadSrc) ? 'SIN CIUDAD' : ciudadSrc;
  r['CIUDAD'] = cleanNCCatalog(lookup(MAP_CIUDAD, ciudadKey, ciudadKey || 'SIN CIUDAD'));

  r['AREA VALIDA'] = calcularNCAreaValida(r['VAL'], r['conv'], r['Periodo']);

  // Supervisor: igual que Sin Gestión → supervisorSimpleLookup(email_asesor)
  const _peAsesorEmail = (typeof getAsesorEmail === 'function')
    ? getAsesorEmail(r['Propietario de Posible Cliente'])
    : String(r['Propietario de Posible Cliente'] || '').trim();
  r['SUPERVISOR'] = (typeof supervisorSimpleLookup === 'function')
    ? supervisorSimpleLookup(_peAsesorEmail)
    : '';
  r['FECHA CREACION'] = truncNCDate(r['Hora de creación']);

  return r;
}

function getPEProgramaPredictivoFromRow(r){
  const rawPrograma = String(r['Programa'] || '').trim();
  const rawInteres  = String(r['Programa de interes_'] || '').trim();
  const src = rawPrograma ? rawPrograma : rawInteres;
  return toNCTitle(src || 'Sin programa');
}

function buildPEPredictivoRow(r, counter){
  const tel = normalizarTelefono(r['Teléfono']);
  const rawEmail = String(r['Correo electrónico'] || '').trim();
  const validEmail = /^[A-Za-z0-9@._\-,]+$/.test(rawEmail) ? rawEmail : '';
  const programa = (String(r['Programa de interes_'] || r['Programa'] || '').trim() || 'SIN PROGRAMA').toUpperCase();
  return {
    'CONTACTID' : counter,
    'Prioridad' : 1,
    'email'     : validEmail,
    'First_name': getNCFirstName(r['Nombre completo']),
    'telefono'  : tel,
    'Programa'  : programa,
    'number 1'  : tel ? '+579' + tel : '',
    'AgentName' : typeof getAsesorEmail === 'function' ? getAsesorEmail(r['Propietario de Posible Cliente']) : '',
    'ValCorreo' : validEmail ? 'VALIDADO' : 'CORREO INVÁLIDO'
  };
}

function switchPETab(n){
  document.querySelectorAll('#pe-tabs .tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('#problemas-economicos-module .pe-panel').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n === 2) updatePEDistInfo();
  if(n === 3) updatePECariInfo();
  if(n === 4) updatePENotasInfo();
}

function buildPEFilterPanel(){
  const row = document.getElementById('pe-filter-row');
  row.innerHTML = '';
  PE_FILTER_DEFS.forEach(f=>{
    const div = document.createElement('div');
    div.className = 'fg';
    div.innerHTML = `
      <label>${f.lbl}<span class="clr" onclick="clearOnePE('${f.id}')">✖</span></label>
      <div class="ms-wrap">
        <button class="ms-btn" onclick="toggleMS('${f.id}',event)">
          <span class="st" id="txt-${f.id}">Todas</span>
          <span style="opacity:.4;font-size:.62rem;flex-shrink:0;margin-left:2px">▼</span>
        </button>
        <div class="ms-dd" id="${f.id}">
          <div class="ms-sb"><input type="text" placeholder="Buscar…" oninput="filterOpts('${f.id}',this.value)"></div>
          <div class="ms-all" onclick="toggleAllPE('${f.id}')">☑ Sel/Des todo</div>
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
    <label>Hora de modificación<span class="clr" onclick="clearPEDateFilter()">✖</span></label>
    <div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap;">
      <input type="date" id="pe-fecha-desde"
             style="flex:1;min-width:120px;padding:5px 7px;border:1.5px solid #c5cfd9;border-radius:6px;font-size:.82rem;"
             oninput="applyPEFilters()" title="Desde">
      <span style="color:#888;font-size:.8rem;flex-shrink:0;">–</span>
      <input type="date" id="pe-fecha-hasta"
             style="flex:1;min-width:120px;padding:5px 7px;border:1.5px solid #c5cfd9;border-radius:6px;font-size:.82rem;"
             oninput="applyPEFilters()" title="Hasta">
    </div>`;
  row.appendChild(dateDiv);

  const txt = document.createElement('div');
  txt.className='fg'; txt.style.minWidth='190px';
  txt.innerHTML = `<label>Buscar nombre/correo/tel/ID</label><input type="text" id="pe-buscar" placeholder="Escribe para buscar…" oninput="applyPEFilters()">`;
  row.appendChild(txt);
  const clrb = document.createElement('div');
  clrb.className='fg'; clrb.style.flex='0'; clrb.style.minWidth='auto';
  clrb.innerHTML = `<label>&nbsp;</label><button class="btn btn-gray" onclick="clearPEFilters()">✖ Limpiar</button>`;
  row.appendChild(clrb);
}

function clearPEDateFilter(){
  const d = document.getElementById('pe-fecha-desde');
  const h = document.getElementById('pe-fecha-hasta');
  if(d) d.value = '';
  if(h) h.value = '';
  applyPEFilters();
}

function populatePEFilters(){
  PE_FILTER_DEFS.forEach(f=>{
    const vals = [...new Set(peData.map(r => String(r[f.col] || '').trim()).filter(v => v && v !== '-'))].sort((a,b)=>a.localeCompare(b));
    document.getElementById('list-'+f.id).innerHTML = vals.map(v =>
      `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(v)}" onchange="onCheckPE('${f.id}')"> <span>${escapeHtml(v)}</span></label>`
    ).join('');
    updatePEFilterText(f.id);
  });
}

function onCheckPE(id){ applyPEFilters(); updatePEFilterText(id); closeAll(); }
function toggleAllPE(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input')];
  const all=checks.every(c=>c.checked);
  checks.forEach(c=>c.checked=!all);
  applyPEFilters(); updatePEFilterText(id); closeAll();
}
function clearOnePE(id){
  document.querySelectorAll('#list-'+id+' input').forEach(c=>c.checked=false);
  applyPEFilters(); updatePEFilterText(id); closeAll();
}
function clearPEFilters(){
  PE_FILTER_DEFS.forEach(f=>{
    document.querySelectorAll('#list-'+f.id+' input').forEach(c=>c.checked=false);
    updatePEFilterText(f.id);
  });
  const b=document.getElementById('pe-buscar'); if(b)b.value='';
  const desde=document.getElementById('pe-fecha-desde'); if(desde)desde.value='';
  const hasta=document.getElementById('pe-fecha-hasta'); if(hasta)hasta.value='';
  applyPEFilters();
}
function updatePEFilterText(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input:checked')];
  const txt=document.getElementById('txt-'+id);
  if(txt) txt.textContent = checks.length ? (checks.length===1 ? checks[0].value : `${checks.length} seleccionados`) : 'Todas';
}

function applyPEFilters(){
  const q     = String(document.getElementById('pe-buscar')?.value || '').toLowerCase().trim();
  const desde = document.getElementById('pe-fecha-desde')?.value || '';
  const hasta  = document.getElementById('pe-fecha-hasta')?.value  || '';

  peFiltered = peData.filter(r=>{
    for(const f of PE_FILTER_DEFS){
      const vals=[...document.querySelectorAll('#list-'+f.id+' input:checked')].map(c=>c.value);
      updatePEFilterText(f.id);
      if(vals.length && !vals.includes(String(r[f.col] || ''))) return false;
    }
    /* Rango de fechas por Hora de modificación */
    if(desde || hasta){
      const modDate = truncNCDate(r['Hora de modificación'] || '');
      if(desde && modDate < desde) return false;
      if(hasta  && modDate > hasta) return false;
    }
    if(q){
      const hay=[r['ID de registro'],r['Nombre completo'],r['Correo electrónico'],r['Teléfono'],r['PROGRAMA NORMALIZADO'],r['CIUDAD'],r['AREA VALIDA'],r['SUPERVISOR']].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });
  pePage=1; pePredPage=1;
  renderPEKPIs(); renderPEPredictivo(); renderPETable(); updatePEDistInfo(); updatePECariInfo(); updatePENotasInfo();
}

function renderPEKPIs(){
  const d = peFiltered.length ? peFiltered : peData;
  document.getElementById('pk-total').textContent = peData.length.toLocaleString();
  document.getElementById('pk-pred').textContent  = pePredictivoData.length.toLocaleString();
  document.getElementById('pk-virt').textContent  = d.filter(r=>r['AREA VALIDA']==='VIRTUAL').length.toLocaleString();
  document.getElementById('pk-pres').textContent  = d.filter(r=>r['AREA VALIDA']==='PRESENCIAL').length.toLocaleString();
  const posEl = document.getElementById('pk-pos');
  if(posEl) posEl.textContent = d.filter(r=>r['AREA VALIDA']==='POSGRADO').length.toLocaleString();
  document.getElementById('pk-sin').textContent   = d.filter(r=>['SIN PROGRAMA',''].includes(r['PROGRAMA NORMALIZADO']) || ['SIN CIUDAD',''].includes(r['CIUDAD'])).length.toLocaleString();
}

function getPEPredictivoFiltrado(){ return peFiltered.map(function(r, idx){ return buildPEPredictivoRow(r, idx + 1); }); }

function renderPEPredictivo(){
  document.getElementById('pe-pred-head').innerHTML = '<tr>'+PE_PRED_COLS.map(c=>`<th>${c}</th>`).join('')+'</tr>';
  const data = getPEPredictivoFiltrado();
  const total=data.length, pages=Math.max(1,Math.ceil(total/PE_PAGE_SIZE));
  pePredPage=Math.min(pePredPage,pages);
  const sl=data.slice((pePredPage-1)*PE_PAGE_SIZE,pePredPage*PE_PAGE_SIZE);
  document.getElementById('pe-pred-title').textContent = `${total.toLocaleString()} registros predictivo`;
  document.getElementById('pe-pred-body').innerHTML = sl.map(r=>'<tr>'+PE_PRED_COLS.map(c=>`<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('')+'</tr>').join('');
  document.getElementById('pe-pred-pag-info').textContent = `Página ${pePredPage} de ${pages}`;
}

function renderPETable(){
  document.getElementById('pe-head').innerHTML = '<tr>'+PE_SHOW_COLS.map(c=>`<th onclick="sortPEBy('${c}')">${c} <span style="opacity:.35">${peSortCol===c?(peSortAsc?'▲':'▼'):'⇅'}</span></th>`).join('')+'</tr>';
  let data=[...peFiltered];
  if(peSortCol){
    data.sort((a,b)=>{
      const va=String(a[peSortCol]||''), vb=String(b[peSortCol]||'');
      return peSortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    });
  }
  const total=data.length, pages=Math.max(1,Math.ceil(total/PE_PAGE_SIZE));
  pePage=Math.min(pePage,pages);
  const sl=data.slice((pePage-1)*PE_PAGE_SIZE,pePage*PE_PAGE_SIZE);
  document.getElementById('pe-title').textContent = `${total.toLocaleString()} registros problemas económicos normalizados`;
  document.getElementById('pe-body').innerHTML = sl.map(r=>'<tr>'+PE_SHOW_COLS.map(c=>{
    const v=String(r[c] ?? '');
    return `<td title="${escapeHtml(v)}">${getPEBadge(c,v) || escapeHtml(v)}</td>`;
  }).join('')+'</tr>').join('');
  document.getElementById('pe-pag-info').textContent = `Página ${pePage} de ${pages}`;
}
function getPEBadge(col,v){
  if(col==='AREA VALIDA' || col==='VAL'){
    if(v==='VIRTUAL')    return `<span class="badge vir">VIRTUAL</span>`;
    if(v==='PRESENCIAL') return `<span class="badge pre">PRESENCIAL</span>`;
    if(v==='POSGRADO')   return `<span class="badge pos">POSGRADO</span>`;
    if(v==='MIXTO')      return `<span class="badge mix">MIXTO</span>`;
  }
  return '';
}
function sortPEBy(c){ if(peSortCol===c)peSortAsc=!peSortAsc; else{peSortCol=c;peSortAsc=true;} renderPETable(); }
function changePEPage(n){ const pages=Math.max(1,Math.ceil(peFiltered.length/PE_PAGE_SIZE)); pePage=Math.min(Math.max(1,pePage+n),pages); renderPETable(); }
function changePEPredPage(n){ const pages=Math.max(1,Math.ceil(getPEPredictivoFiltrado().length/PE_PAGE_SIZE)); pePredPage=Math.min(Math.max(1,pePredPage+n),pages); renderPEPredictivo(); }

function exportPERows(rows,name,cols){
  const ws=XLSX.utils.json_to_sheet(rows.map(r=>{const o={}; cols.forEach(c=>o[c]=r[c]!==undefined?r[c]:''); return o;}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'ProblemasEconomicos');
  XLSX.writeFile(wb,name);
  showToast(`⬇ ${rows.length.toLocaleString()} filas exportadas`);
}
function exportPEPredictivo(){ exportPERows(getPEPredictivoFiltrado(),'Predictivo_Problemas_Economicos_'+Date.now()+'.xlsx',PE_PRED_COLS); }

function exportPEPredictivoCSV(){
  var rows = getPEPredictivoFiltrado();
  if(!rows.length){ showToast('No hay datos para exportar.'); return; }
  var ws = XLSX.utils.json_to_sheet(rows, {header: PE_PRED_COLS});
  if(typeof downloadCSVFile === 'function'){
    downloadCSVFile(XLSX.utils.sheet_to_csv(ws), 'Predictivo_Problemas_Economicos_' + Date.now() + '.csv');
  }
  showToast('⬇ CSV exportado: ' + rows.length.toLocaleString() + ' registros');
}

function exportPEPredictivoJPG(){
  if(typeof exportPredictivoAgentJPG === 'function'){
    exportPredictivoAgentJPG(getPEPredictivoFiltrado(), 'Problemas Economicos');
  }
}

function exportPEFiltered(){ exportPERows(peFiltered,'Base_Problemas_Economicos_Filtrada_'+Date.now()+'.xlsx',PE_BASE_COLS); }
function exportPEAll(){ exportPERows(peData,'Base_Problemas_Economicos_Completa_'+Date.now()+'.xlsx',PE_BASE_COLS); }
function copiarPEPredictivoContactIds(){ const data=getPEPredictivoFiltrado(); navigator.clipboard.writeText(data.map(r=>r['CONTACTID']).join('\n')).then(()=>showToast(`📋 ${data.length.toLocaleString()} CONTACTID copiados`)); }
function copiarPEIDs(){ navigator.clipboard.writeText(peFiltered.map(r=>getPELeadId(r)).join('\n')).then(()=>showToast(`📋 ${peFiltered.length.toLocaleString()} IDs copiados`)); }

function updatePEDistInfo(){
  const el=document.getElementById('pe-dist-info');
  if(el) el.textContent = `${peFiltered.length.toLocaleString()} leads listos para distribuir (filtro activo)`;
}
function dropPEAsesores(e){ e.preventDefault(); loadPEAsesores(e.dataTransfer.files[0]); }
function loadPEAsesores(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:'array'});
      const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const col=Object.keys(json[0]||{}).find(k=>/correo|email/i.test(k));
      if(!col){showToast('No se encontró columna Correo o Email');return;}
      peAsesoresList=[...new Set(json.map(r=>String(r[col]).trim()).filter(Boolean))];
      renderPEAsesoresTags();
      showToast(`✅ ${peAsesoresList.length} asesores cargados`);
    }catch(err){ showToast('Error asesores: '+err.message); }
  };
  reader.readAsArrayBuffer(file);
}
function renderPEAsesoresTags(){
  const box=document.getElementById('pe-asesores-tags');
  if(!box)return;
  box.innerHTML=`<p style="font-size:.77rem;color:#666;margin-bottom:4px"><strong>${peAsesoresList.length} asesores:</strong></p>`+
    peAsesoresList.map((a,i)=>`<span class="tag">${escapeHtml(a)}<span class="rm" onclick="rmPEAsesor(${i})">×</span></span>`).join('');
}
function rmPEAsesor(i){ peAsesoresList.splice(i,1); renderPEAsesoresTags(); }
function getPELeadId(row){
  if(!row)return '';
  if(row['ID de registro']!==undefined && row['ID de registro']!==null && String(row['ID de registro']).trim()!=='') return String(row['ID de registro']).trim();
  for(const key of Object.keys(row)){
    const clean=typeof cleanHeader==='function'?cleanHeader(key):String(key).toUpperCase().replace(/[^A-Z0-9]/g,'');
    if(['IDDEREGISTRO','IDREGISTRO','RECORDID','LEADID','ID'].includes(clean)){
      const value=row[key]; if(value!==undefined&&value!==null&&String(value).trim()!=='') return String(value).trim();
    }
  }
  return '';
}

function distribuirPE(){
  if(!peFiltered.length){showToast('No hay leads para distribuir.');return;}
  if(!peAsesoresList.length){showToast('Carga primero el Excel de asesores.');return;}
  const leadsOrdenados = ordenarPorHoraCreacionNCParaDistribuir(peFiltered);
  const n = peAsesoresList.length;
  peDistribuidoData = leadsOrdenados.map((lead,i)=>asignarSupervisorSimple({'ID de registro':getPELeadId(lead)}, peAsesoresList[i % n]));
  renderPEDistPreview();
  showPEDistStats();
  renderDetalleAsesorPE();
  showToast(`✅ ${peDistribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}
function renderPEDistPreview(){
  const wrap=document.getElementById('pe-dist-preview'), tbl=document.getElementById('pe-dist-tbl');
  if(!wrap||!tbl)return;
  wrap.style.display='block';
  const preview=peDistribuidoData.slice(0,300);
  tbl.innerHTML='<thead><tr><th>ID de registro</th><th>CORREO ASESOR ASIGNADO</th></tr></thead><tbody>'+
    preview.map(r=>`<tr><td>${escapeHtml(String(r['ID de registro']||''))}</td><td>${escapeHtml(String(r['CORREO ASESOR ASIGNADO']||''))}</td></tr>`).join('')+'</tbody>';
}
function showPEDistStats(){
  const box=document.getElementById('pe-stat-box');
  if(!box)return;
  const counts={};
  peDistribuidoData.forEach(r=>{const a=r['CORREO ASESOR ASIGNADO']; counts[a]=(counts[a]||0)+1;});
  box.style.display='block';
  box.innerHTML='<strong>Resumen de asignación:</strong><br>'+Object.entries(counts).map(([a,c])=>`<span class="dist-chip">${escapeHtml(a)}: <b>${c}</b></span>`).join('');
}
function exportarDistPE(){
  if(!peDistribuidoData.length){showToast('Primero debes distribuir los leads.');return;}
  const rows = peDistribuidoData.map(r => ({
    'ID de registro': r['ID de registro'],
    'CORREO ASESOR ASIGNADO': r['CORREO ASESOR ASIGNADO']
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Distribucion');
  XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' PROBLEMAS' + (typeof getPEFileLabel==='function'?getPEFileLabel():'') + '.xlsx');
  showToast(`⬇ ${peDistribuidoData.length.toLocaleString()} asignaciones exportadas`);
}
function copiarPEDistIDs(){
  const data=peDistribuidoData.length?peDistribuidoData:peFiltered.map(r=>({'ID de registro':getPELeadId(r)}));
  navigator.clipboard.writeText(data.map(r=>r['ID de registro']).join('\n')).then(()=>showToast(`📋 ${data.length.toLocaleString()} IDs copiados`));
}


/* ════════════════════════════════════════════
   CARI AI + NOTAS — PROBLEMAS ECONÓMICOS
   ════════════════════════════════════════════ */

const PE_CARI_COLS = ['numero_telefono','nombre_aspirante','carrera_interes','identificacion','correo_electronico','telefono_adicional','periodo','campania'];
const PE_NOTAS_COLS = ['ID','Notes'];
const PE_NOTA_FIJA  = 'Se intenta contacto con el aspirante via CARI AI';

function updatePECariInfo(){
  const el = document.getElementById('pe-cari-info');
  if(el) el.textContent = `${peFiltered.length.toLocaleString()} registros listos para CARI AI (filtro activo)`;
}
function updatePENotasInfo(){
  const el = document.getElementById('pe-notas-info');
  if(el) el.textContent = `${peFiltered.length.toLocaleString()} notas listas para generar (filtro activo)`;
}

function buildPECariRows(){
  const rows = peFiltered.map(r => {
    const phone = (function(row){
      const tel = normalizarTelefono(row['Teléfono']);
      if(!tel) return '';
      return tel.startsWith('57') ? tel : '57' + tel;
    })(r);
    const area  = String(r['AREA VALIDA'] || '').trim().toUpperCase();
    const doc   = String(r['Número de Documento'] || '').trim();
    return {
      'numero_telefono'    : phone,
      'nombre_aspirante'   : r['Nombre completo'] || '',
      'carrera_interes'    : String(r['PROGRAMA NORMALIZADO'] || '').toLowerCase(),
      'identificacion'     : doc || 'Sin documento',
      'correo_electronico' : r['Correo electrónico'] || '',
      'telefono_adicional' : phone,
      'periodo'            : area === 'POSGRADO' ? 'Virtual/Especialización' : 'Virtual/pregrado',
      'campania'           : 'Organico'
    };
  });
  /* Filas fijas obligatorias */
  return [
    {
      'numero_telefono':'573332322810','nombre_aspirante':'Esteban',
      'carrera_interes':'especializacion en contratacion estatal','identificacion':'Sin documento',
      'correo_electronico':'Prueba@gmail.com','telefono_adicional':'573332322810',
      'periodo':'Virtual/pregrado','campania':'Organico'
    },
    {
      'numero_telefono':'573134268590','nombre_aspirante':'Diguar',
      'carrera_interes':'ingenieria de sistemas','identificacion':'Sin documento',
      'correo_electronico':'Prueba@gmail.com','telefono_adicional':'573134268590',
      'periodo':'Virtual/pregrado','campania':'Organico'
    },
    ...rows
  ];
}

function buildPENotasRows(){
  return peFiltered.map(r => ({
    'ID'   : getPELeadId(r),
    'Notes': PE_NOTA_FIJA
  }));
}

function renderPECari(){
  const rows = buildPECariRows();
  const wrap = document.getElementById('pe-cari-preview');
  const tbl  = document.getElementById('pe-cari-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + PE_CARI_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + PE_CARI_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 registros de ${rows.length.toLocaleString()}</caption>`;
  }
  updatePECariInfo();
}

function renderPENotas(){
  const rows = buildPENotasRows();
  const wrap = document.getElementById('pe-notas-preview');
  const tbl  = document.getElementById('pe-notas-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + PE_NOTAS_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + PE_NOTAS_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de ${rows.length.toLocaleString()} notas</caption>`;
  }
  updatePENotasInfo();
}

function exportPECari(){
  const rows = buildPECariRows();
  if(!rows.length){ showToast('No hay registros CARI AI para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    PE_CARI_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CARI AI');
  XLSX.writeFile(wb, 'PROBLEMAS_ECONOMICOS_' + (typeof getDateStamp==='function'?getDateStamp():'') + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} registros CARI AI exportados`);
}

function exportPENotas(){
  const rows = buildPENotasRows();
  if(!rows.length){ showToast('No hay notas para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    PE_NOTAS_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' PROBLEMAS' + (typeof getPEFileLabel==='function'?getPEFileLabel():'') + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} notas exportadas`);
}

function copiarPENotasIDs(){
  const rows = buildPENotasRows();
  navigator.clipboard.writeText(rows.map(r => r['ID']).join('\n'))
    .then(() => showToast(`📋 ${rows.length.toLocaleString()} IDs copiados`));
}


/* === DETALLE ASESOR — Problemas Económicos === */
function getMentorPE(row){
  return String(row['MENTOR'] || mentorDeAsesor(row['CORREO ASESOR ASIGNADO']) || 'Sin supervisor').trim();
}

function renderDetalleAsesorPE(){
  if(typeof renderDetalleAsesorGenerico === 'function'){
    renderDetalleAsesorGenerico({
      data: (typeof peDistribuidoData !== 'undefined' ? peDistribuidoData : []),
      cardId: 'pe-detalle-asesor-card',
      boxId:  'pe-detalle-asesor',
      selectId: 'pe-mentor-filter',
      getMentor: getMentorPE,
      renderAgain: 'renderDetalleAsesorPE'
    });
  }
}
