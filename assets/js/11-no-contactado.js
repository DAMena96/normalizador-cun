
/* ════════════════════════════════════════════
   BASE NO CONTACTADO + PREDICTIVO + DISTRIBUCIÓN
   Misma lógica de Base Interesados
   ════════════════════════════════════════════ */

let ncData = [], ncFiltered = [], ncPredictivoData = [];
let ncAsesoresList = [], ncDistribuidoData = [];
let ncSortCol = '', ncSortAsc = true, ncPage = 1, ncPredPage = 1;
const NC_PAGE_SIZE = 100;

const NC_BASE_COLS = [
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

const NC_SHOW_COLS = [
  'ID de registro','Nombre completo','Correo electrónico','Teléfono',
  'PROGRAMA NORMALIZADO','CIUDAD','AREA VALIDA','VAL','SUPERVISOR',
  'FECHA CREACION','Hora de modificación','Campaña mercadeo','Periodo','Creado por'
];

const NC_PRED_COLS = ['CONTACTID','Prioridad','email','First_name','telefono','Programa','number 1','AgentName','ValCorreo'];

const NC_FILTER_DEFS = [
  {id:'nc-area-valida', col:'AREA VALIDA', lbl:'Área válida'},
  {id:'nc-val', col:'VAL', lbl:'VAL'},
  {id:'nc-fecha-crea', col:'FECHA CREACION', lbl:'Fecha creación'}
];

function handleDropNC(e){
  e.preventDefault();
  document.getElementById('nc-upload-zone').classList.remove('over');
  loadNCFile(e.dataTransfer.files[0]);
}

function ncShowProg(pct,label){
  document.getElementById('nc-progress-wrap').style.display='block';
  document.getElementById('nc-progress-fill').style.width=pct+'%';
  document.getElementById('nc-progress-label').textContent=label;
}
function ncHideProg(){
  document.getElementById('nc-progress-wrap').style.display='none';
  document.getElementById('nc-progress-fill').style.width='0';
}

async function loadNCFile(file){
  if(!file) return;
  ncShowProg(5,'Cargando catálogos JSON…');
  const okCatalogs = await catalogsReady;
  if(!okCatalogs){ ncHideProg(); return; }

  ncShowProg(15,'Leyendo archivo…');
  const isCSV = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onload = ev => {
    ncShowProg(35, isCSV ? 'Parseando CSV…' : 'Parseando Excel…');
    setTimeout(()=>{
      try{
        const wb = isCSV
          ? XLSX.read(ev.target.result, {type:'string', raw:false})
          : XLSX.read(ev.target.result, {type:'array'});
        const sn = wb.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
        if(json.length < 2){ showToast('Sin datos en la base no contactado.'); ncHideProg(); return; }
        const headers = json[0].map(h => String(h).trim());

        ncShowProg(60,`Normalizando ${(json.length-1).toLocaleString()} registros…`);
        setTimeout(()=>{
          try{
            ncData = [];
            for(let i=1;i<json.length;i++){
              const row = json[i];
              if(row.every(c => c === '' || c === null || c === undefined)) continue;
              ncData.push(normalizeNCRow(row, headers));
            }

            ncPredictivoData = ncData.map(buildNCPredictivoRow);
            ncFiltered = [...ncData];

            ncShowProg(90,'Preparando…');
            setTimeout(()=>{
              buildNCFilterPanel();
              populateNCFilters();
              renderNCKPIs();
              renderNCPredictivo();
              renderNCTable();

              document.getElementById('nc-upload-zone').classList.add('hidden');
              document.getElementById('nc-kpi-row').classList.remove('hidden');
              document.getElementById('nc-tabs').classList.remove('hidden');
              document.querySelectorAll('.nc-panel').forEach(p => p.classList.remove('hidden'));
              switchNCTab(0);

              ncHideProg();
              showToast(`✅ ${ncData.length.toLocaleString()} registros no contactado procesados`);
            },80);
          }catch(err){
            console.error(err);
            showToast('Error No Contactado: ' + err.message);
            ncHideProg();
          }
        },50);
      }catch(err){
        console.error(err);
        showToast('Error No Contactado: ' + err.message);
        ncHideProg();
      }
    },30);
  };
  if(isCSV) reader.readAsText(file, 'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function normalizeNCRow(row, headers){
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

  const programaPredictivo = getNCProgramaPredictivoFromRow(r);
  const programaNormalizado = cleanNCCatalog(lookup(MAP_PROGRAMA, programaPredictivo, programaPredictivo || 'SIN PROGRAMA') || 'SIN PROGRAMA');
  r['Programa'] = programaPredictivo;
  r['PROGRAMA NORMALIZADO'] = programaNormalizado;

  // VAL: se cruza el programa normalizado contra areas.json.
  // Resultado esperado: MIXTO / PRESENCIAL / VIRTUAL / POSGRADO / OTRO.
  r['VAL'] = cleanNCCatalog(lookup(MAP_AREA, programaNormalizado, 'OTRO') || 'OTRO');

  const ciudadSrc = emptyOrZeroNC(r['Ciudad utm']) ? r['Ciudad de residencia'] : r['Ciudad utm'];
  const ciudadKey = emptyOrZeroNC(ciudadSrc) ? 'SIN CIUDAD' : ciudadSrc;
  r['CIUDAD'] = cleanNCCatalog(lookup(MAP_CIUDAD, ciudadKey, ciudadKey || 'SIN CIUDAD'));

  r['AREA VALIDA'] = calcularNCAreaValida(r['VAL'], r['conv'], r['Periodo']);

  r['SUPERVISOR'] = lookup(MAP_ASESOR, r['Ciudad de residencia'], r['CIUDAD'] || '-');
  r['FECHA CREACION'] = truncNCDate(r['Hora de creación']);

  return r;
}

function getNCProgramaPredictivoFromRow(r){
  const rawPrograma = String(r['Programa'] || '').trim();
  const rawInteres = String(r['Programa de interes_'] || '').trim();
  const src = rawPrograma ? rawPrograma : rawInteres;
  return toNCTitle(src || 'Sin programa');
}

function buildNCPredictivoRow(r, counter){
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

function calcularNCAreaValida(val, conv, periodo){
  const a = String(val || '').toUpperCase();
  if(a === 'MIXTO'){
    const c = String(conv || '').trim().toUpperCase();
    const p = String(periodo || '').trim().toUpperCase();
    if(c) return c.includes('V') ? 'VIRTUAL' : 'PRESENCIAL';
    if(p) return p.includes('V') ? 'VIRTUAL' : 'PRESENCIAL';
    return 'PRESENCIAL';
  }
  return cleanNCCatalog(val || 'OTRO');
}

function cleanNCCatalog(v){
  const s = String(v ?? '').trim();
  if(!s) return '-';
  return s.toUpperCase();
}
function emptyOrZeroNC(v){
  const s = String(v ?? '').trim();
  return s === '' || s === '0';
}
function sinTildesNC(s){
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function toNCTitle(s){
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/\b([a-záéíóúñ])/g, c => c.toUpperCase());
}
function getNCFirstName(fullName){
  let s = String(fullName || '').trim();
  if(!s) return 'Sin nombre';
  if(s.charCodeAt(0) >= 127) return 'Sin nombre';
  s = toNCTitle(s);
  const m = s.match(/[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+/);
  return m ? m[0] : 'Sin nombre';
}
function truncNCDate(value){
  if(!value) return '';
  const s = String(value).trim();
  if(/^\d+(\.\d+)?$/.test(s)){
    const n = Number(s);
    if(n > 25000 && n < 60000){
      const d = new Date(Math.round((n - 25569) * 86400 * 1000));
      return d.toISOString().slice(0,10);
    }
  }
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if(m) return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if(m){
    let y = String(m[3]); if(y.length === 2) y = '20'+y;
    return `${y}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  }
  const d = new Date(s);
  if(!isNaN(d)) return d.toISOString().slice(0,10);
  return s.split(' ')[0] || s;
}

function switchNCTab(n){
  document.querySelectorAll('#nc-tabs .tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('#no-contactado-module .nc-panel').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n === 2) updateNCDistInfo();
  if(n === 3) updateNCCariInfo();
  if(n === 4) updateNCNotasInfo();
}

function buildNCFilterPanel(){
  const row = document.getElementById('nc-filter-row');
  row.innerHTML = '';
  NC_FILTER_DEFS.forEach(f=>{
    const div = document.createElement('div');
    div.className = 'fg';
    div.innerHTML = `
      <label>${f.lbl}<span class="clr" onclick="clearOneNC('${f.id}')">✖</span></label>
      <div class="ms-wrap">
        <button class="ms-btn" onclick="toggleMS('${f.id}',event)">
          <span class="st" id="txt-${f.id}">Todas</span>
          <span style="opacity:.4;font-size:.62rem;flex-shrink:0;margin-left:2px">▼</span>
        </button>
        <div class="ms-dd" id="${f.id}">
          <div class="ms-sb"><input type="text" placeholder="Buscar…" oninput="filterOpts('${f.id}',this.value)"></div>
          <div class="ms-all" onclick="toggleAllNC('${f.id}')">☑ Sel/Des todo</div>
          <div class="ms-list" id="list-${f.id}"></div>
        </div>
      </div>`;
    row.appendChild(div);
  });
  const txt = document.createElement('div');
  txt.className='fg'; txt.style.minWidth='190px';
  txt.innerHTML = `<label>Buscar nombre/correo/tel/ID</label><input type="text" id="nc-buscar" placeholder="Escribe para buscar…" oninput="applyNCFilters()">`;
  row.appendChild(txt);
  const clrb = document.createElement('div');
  clrb.className='fg'; clrb.style.flex='0'; clrb.style.minWidth='auto';
  clrb.innerHTML = `<label>&nbsp;</label><button class="btn btn-gray" onclick="clearNCFilters()">✖ Limpiar</button>`;
  row.appendChild(clrb);
}

function populateNCFilters(){
  NC_FILTER_DEFS.forEach(f=>{
    const vals = [...new Set(ncData.map(r => String(r[f.col] || '').trim()).filter(v => v && v !== '-'))].sort((a,b)=>a.localeCompare(b));
    document.getElementById('list-'+f.id).innerHTML = vals.map(v =>
      `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(v)}" onchange="onCheckNC('${f.id}')"> <span>${escapeHtml(v)}</span></label>`
    ).join('');
    updateNCFilterText(f.id);
  });
}

function onCheckNC(id){ applyNCFilters(); updateNCFilterText(id); closeAll(); }
function toggleAllNC(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input')];
  const all=checks.every(c=>c.checked);
  checks.forEach(c=>c.checked=!all);
  applyNCFilters(); updateNCFilterText(id); closeAll();
}
function clearOneNC(id){
  document.querySelectorAll('#list-'+id+' input').forEach(c=>c.checked=false);
  applyNCFilters(); updateNCFilterText(id); closeAll();
}
function clearNCFilters(){
  NC_FILTER_DEFS.forEach(f=>{
    document.querySelectorAll('#list-'+f.id+' input').forEach(c=>c.checked=false);
    updateNCFilterText(f.id);
  });
  const b=document.getElementById('nc-buscar'); if(b)b.value='';
  applyNCFilters();
}
function updateNCFilterText(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input:checked')];
  const txt=document.getElementById('txt-'+id);
  if(txt) txt.textContent = checks.length ? (checks.length===1 ? checks[0].value : `${checks.length} seleccionados`) : 'Todas';
}

function applyNCFilters(){
  const q = String(document.getElementById('nc-buscar')?.value || '').toLowerCase().trim();
  ncFiltered = ncData.filter(r=>{
    for(const f of NC_FILTER_DEFS){
      const vals=[...document.querySelectorAll('#list-'+f.id+' input:checked')].map(c=>c.value);
      updateNCFilterText(f.id);
      if(vals.length && !vals.includes(String(r[f.col] || ''))) return false;
    }
    if(q){
      const hay=[r['ID de registro'],r['Nombre completo'],r['Correo electrónico'],r['Teléfono'],r['PROGRAMA NORMALIZADO'],r['CIUDAD'],r['AREA VALIDA'],r['VAL']].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });
  ncPage=1; ncPredPage=1;
  renderNCKPIs(); renderNCPredictivo(); renderNCTable(); updateNCDistInfo(); updateNCCariInfo(); updateNCNotasInfo();
}

function renderNCKPIs(){
  const d = ncFiltered.length ? ncFiltered : ncData;
  document.getElementById('nk-total').textContent = ncData.length.toLocaleString();
  document.getElementById('nk-pred').textContent = ncPredictivoData.length.toLocaleString();
  document.getElementById('nk-virt').textContent = d.filter(r=>r['AREA VALIDA']==='VIRTUAL').length.toLocaleString();
  document.getElementById('nk-pres').textContent = d.filter(r=>r['AREA VALIDA']==='PRESENCIAL').length.toLocaleString();
  document.getElementById('nk-sin').textContent = d.filter(r=>['SIN PROGRAMA',''].includes(r['PROGRAMA NORMALIZADO']) || ['SIN CIUDAD',''].includes(r['CIUDAD'])).length.toLocaleString();
}

function getNCPredictivoFiltrado(){ return ncFiltered.map(function(r, idx){ return buildNCPredictivoRow(r, idx + 1); }); }

function renderNCPredictivo(){
  document.getElementById('nc-pred-head').innerHTML = '<tr>'+NC_PRED_COLS.map(c=>`<th>${c}</th>`).join('')+'</tr>';
  const data = getNCPredictivoFiltrado();
  const total=data.length, pages=Math.max(1,Math.ceil(total/NC_PAGE_SIZE));
  ncPredPage=Math.min(ncPredPage,pages);
  const sl=data.slice((ncPredPage-1)*NC_PAGE_SIZE,ncPredPage*NC_PAGE_SIZE);
  document.getElementById('nc-pred-title').textContent = `${total.toLocaleString()} registros predictivo`;
  document.getElementById('nc-pred-body').innerHTML = sl.map(r=>'<tr>'+NC_PRED_COLS.map(c=>`<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('')+'</tr>').join('');
  document.getElementById('nc-pred-pag-info').textContent = `Página ${ncPredPage} de ${pages}`;
}

function renderNCTable(){
  document.getElementById('nc-head').innerHTML = '<tr>'+NC_SHOW_COLS.map(c=>`<th onclick="sortNCBy('${c}')">${c} <span style="opacity:.35">${ncSortCol===c?(ncSortAsc?'▲':'▼'):'⇅'}</span></th>`).join('')+'</tr>';
  let data=[...ncFiltered];
  if(ncSortCol){
    data.sort((a,b)=>{
      const va=String(a[ncSortCol]||''), vb=String(b[ncSortCol]||'');
      return ncSortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    });
  }
  const total=data.length, pages=Math.max(1,Math.ceil(total/NC_PAGE_SIZE));
  ncPage=Math.min(ncPage,pages);
  const sl=data.slice((ncPage-1)*NC_PAGE_SIZE,ncPage*NC_PAGE_SIZE);
  document.getElementById('nc-title').textContent = `${total.toLocaleString()} registros no contactado normalizados`;
  document.getElementById('nc-body').innerHTML = sl.map(r=>'<tr>'+NC_SHOW_COLS.map(c=>{
    const v=String(r[c] ?? '');
    return `<td title="${escapeHtml(v)}">${getNCBadge(c,v) || escapeHtml(v)}</td>`;
  }).join('')+'</tr>').join('');
  document.getElementById('nc-pag-info').textContent = `Página ${ncPage} de ${pages}`;
}
function getNCBadge(col,v){
  if(col==='AREA VALIDA' || col==='VAL'){
    if(v==='VIRTUAL') return `<span class="badge vir">VIRTUAL</span>`;
    if(v==='PRESENCIAL') return `<span class="badge pre">PRESENCIAL</span>`;
    if(v==='POSGRADO') return `<span class="badge pos">POSGRADO</span>`;
    if(v==='MIXTO') return `<span class="badge mix">MIXTO</span>`;
  }
  return '';
}
function sortNCBy(c){ if(ncSortCol===c)ncSortAsc=!ncSortAsc; else{ncSortCol=c;ncSortAsc=true;} renderNCTable(); }
function changeNCPage(n){ const pages=Math.max(1,Math.ceil(ncFiltered.length/NC_PAGE_SIZE)); ncPage=Math.min(Math.max(1,ncPage+n),pages); renderNCTable(); }
function changeNCPredPage(n){ const pages=Math.max(1,Math.ceil(getNCPredictivoFiltrado().length/NC_PAGE_SIZE)); ncPredPage=Math.min(Math.max(1,ncPredPage+n),pages); renderNCPredictivo(); }

function exportNCRows(rows,name,cols){
  const ws=XLSX.utils.json_to_sheet(rows.map(r=>{const o={}; cols.forEach(c=>o[c]=r[c]!==undefined?r[c]:''); return o;}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'NoContactado');
  XLSX.writeFile(wb,name);
  showToast(`⬇ ${rows.length.toLocaleString()} filas exportadas`);
}
function exportNCPredictivo(){ exportNCRows(getNCPredictivoFiltrado(),'Predictivo_No_Contactado_'+Date.now()+'.xlsx',NC_PRED_COLS); }
function exportNCFiltered(){ exportNCRows(ncFiltered,'Base_No_Contactado_Filtrada_'+Date.now()+'.xlsx',NC_BASE_COLS); }
function exportNCAll(){ exportNCRows(ncData,'Base_No_Contactado_Completa_'+Date.now()+'.xlsx',NC_BASE_COLS); }
function copiarNCPredictivoContactIds(){ const data=getNCPredictivoFiltrado(); navigator.clipboard.writeText(data.map(r=>r['CONTACTID']).join('\n')).then(()=>showToast(`📋 ${data.length.toLocaleString()} CONTACTID copiados`)); }
function copiarNCIDs(){ navigator.clipboard.writeText(ncFiltered.map(r=>getNCLeadId(r)).join('\n')).then(()=>showToast(`📋 ${ncFiltered.length.toLocaleString()} IDs copiados`)); }

function updateNCDistInfo(){
  const el=document.getElementById('nc-dist-info');
  if(el) el.textContent = `${ncFiltered.length.toLocaleString()} leads listos para distribuir (filtro activo)`;
}
function dropNCAsesores(e){ e.preventDefault(); loadNCAsesores(e.dataTransfer.files[0]); }
function loadNCAsesores(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:'array'});
      const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const col=Object.keys(json[0]||{}).find(k=>/correo|email/i.test(k));
      if(!col){showToast('No se encontró columna Correo o Email');return;}
      ncAsesoresList=[...new Set(json.map(r=>String(r[col]).trim()).filter(Boolean))];
      renderNCAsesoresTags();
      showToast(`✅ ${ncAsesoresList.length} asesores cargados`);
    }catch(err){ showToast('Error asesores: '+err.message); }
  };
  reader.readAsArrayBuffer(file);
}
function renderNCAsesoresTags(){
  const box=document.getElementById('nc-asesores-tags');
  if(!box)return;
  box.innerHTML=`<p style="font-size:.77rem;color:#666;margin-bottom:4px"><strong>${ncAsesoresList.length} asesores:</strong></p>`+
    ncAsesoresList.map((a,i)=>`<span class="tag">${escapeHtml(a)}<span class="rm" onclick="rmNCAsesor(${i})">×</span></span>`).join('');
}
function rmNCAsesor(i){ ncAsesoresList.splice(i,1); renderNCAsesoresTags(); }
function getNCLeadId(row){
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
function distribuirNC(){
  if(!ncFiltered.length){showToast('No hay leads para distribuir.');return;}
  if(!ncAsesoresList.length){showToast('Carga primero el Excel de asesores.');return;}

  const leadsOrdenados = ordenarPorHoraCreacionNCParaDistribuir(ncFiltered);
  const n = ncAsesoresList.length;
  ncDistribuidoData = leadsOrdenados.map((lead, i) => {
    const asesor = ncAsesoresList[i % n];
    return asignarSupervisorSimple({ 'ID de registro': getNCLeadId(lead) }, asesor);
  });

  renderNCDistPreview();
  showNCDistStats();
  renderDetalleAsesorNC();
  setTimeout(renderDetalleAsesorNC, 100);
  showToast(`✅ ${ncDistribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}
function renderNCDistPreview(){
  const wrap=document.getElementById('nc-dist-preview'), tbl=document.getElementById('nc-dist-tbl');
  if(!wrap||!tbl)return;
  wrap.style.display='block';
  const preview=ncDistribuidoData.slice(0,300);
  tbl.innerHTML='<thead><tr><th>ID de registro</th><th>CORREO ASESOR ASIGNADO</th></tr></thead><tbody>'+
    preview.map(r=>`<tr><td>${escapeHtml(String(r['ID de registro']||''))}</td><td>${escapeHtml(String(r['CORREO ASESOR ASIGNADO']||''))}</td></tr>`).join('')+'</tbody>';
}
function showNCDistStats(){
  const box=document.getElementById('nc-stat-box');
  if(!box)return;
  const counts={};
  ncDistribuidoData.forEach(r=>{const a=r['CORREO ASESOR ASIGNADO']; counts[a]=(counts[a]||0)+1;});
  box.style.display='block';
  box.innerHTML='<strong>Resumen de asignación:</strong><br>'+Object.entries(counts).map(([a,c])=>`<span class="dist-chip">${escapeHtml(a)}: <b>${c}</b></span>`).join('')+htmlDetalleVisible(ncDistribuidoData,'Detalle_Asesor_No_Contactado');
}
function exportarDistNC(){
  if(!ncDistribuidoData.length){showToast('Primero debes distribuir los leads.');return;}
  const rows = ncDistribuidoData.map(r => ({
    'ID de registro': r['ID de registro'],
    'CORREO ASESOR ASIGNADO': r['CORREO ASESOR ASIGNADO']
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Distribucion');
  XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' NOCONTAC' + (typeof getNCFileLabel==='function'?getNCFileLabel():'') + '.xlsx');
  showToast(`⬇ ${ncDistribuidoData.length.toLocaleString()} asignaciones exportadas`);
}
function copiarNCDistIDs(){
  const data=ncDistribuidoData.length?ncDistribuidoData:ncFiltered.map(r=>({'ID de registro':getNCLeadId(r)}));
  navigator.clipboard.writeText(data.map(r=>r['ID de registro']).join('\n')).then(()=>showToast(`📋 ${data.length.toLocaleString()} IDs copiados`));
}


/* ════════════════════════════════════════════
   CARI AI + NOTAS — BASE NO CONTACTADO
   ════════════════════════════════════════════ */

const NC_CARI_COLS = ['numero_telefono','nombre_aspirante','carrera_interes','identificacion','correo_electronico','telefono_adicional','periodo','campania'];
const NC_NOTAS_COLS = ['ID','Notes'];

const NC_CARI_FIXED_ROWS = [
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

const NC_NOTA_FIJA = 'Se intenta contacto con el aspirante via CARI AI';

function updateNCCariInfo(){
  const el = document.getElementById('nc-cari-info');
  if(el) el.textContent = `${ncFiltered.length.toLocaleString()} registros listos para CARI AI (filtro activo)`;
}

function updateNCNotasInfo(){
  const el = document.getElementById('nc-notas-info');
  if(el) el.textContent = `${ncFiltered.length.toLocaleString()} notas listas para generar (filtro activo)`;
}

function getCariPhone(row){
  const tel = normalizarTelefono(row['Teléfono']);
  if(!tel) return '';
  return tel.startsWith('57') ? tel : '57' + tel;
}

function getCariPeriodo(row){
  const area = String(row['AREA VALIDA'] || '').trim().toUpperCase();
  return area === 'POSGRADO' ? 'Virtual/Especialización' : 'Virtual/pregrado';
}

function buildNCCariRows(){
  const rows = ncFiltered.map(r => {
    const phone = getCariPhone(r);
    const doc = String(r['Número de Documento'] || '').trim();

    return {
      'numero_telefono': phone,
      'nombre_aspirante': r['Nombre completo'] || '',
      'carrera_interes': String(r['PROGRAMA NORMALIZADO'] || '').toLowerCase(),
      'identificacion': doc || 'Sin documento',
      'correo_electronico': r['Correo electrónico'] || '',
      'telefono_adicional': phone,
      'periodo': getCariPeriodo(r),
      'campania': 'Organico'
    };
  });

  // Estas dos filas siempre deben ir en CARI AI, sin importar los filtros.
  return [...NC_CARI_FIXED_ROWS, ...rows];
}

function buildNCNotasRows(){
  return ncFiltered.map(r => ({
    'ID': getNCLeadId(r),
    'Notes': NC_NOTA_FIJA
  }));
}

function renderNCCari(){
  const rows = buildNCCariRows();
  const wrap = document.getElementById('nc-cari-preview');
  const tbl = document.getElementById('nc-cari-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + NC_CARI_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + NC_CARI_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 registros de ${rows.length.toLocaleString()}</caption>`;
  }
  updateNCCariInfo();
}

function renderNCNotas(){
  const rows = buildNCNotasRows();
  const wrap = document.getElementById('nc-notas-preview');
  const tbl = document.getElementById('nc-notas-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + NC_NOTAS_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + NC_NOTAS_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 notas de ${rows.length.toLocaleString()}</caption>`;
  }
  updateNCNotasInfo();
}

function exportNCCari(){
  const rows = buildNCCariRows();
  if(!rows.length){ showToast('No hay registros CARI AI para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    NC_CARI_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CARI AI');
  XLSX.writeFile(wb, 'NO CONTACTADOS_' + (typeof getDateStamp==='function'?getDateStamp():'') + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} registros CARI AI exportados`);
}

function exportNCNotas(){
  const rows = buildNCNotasRows();
  if(!rows.length){ showToast('No hay notas para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    NC_NOTAS_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' NOCONTAC' + (typeof getNCFileLabel==='function'?getNCFileLabel():'') + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} notas exportadas`);
}

function copiarNCNotasIDs(){
  const rows = buildNCNotasRows();
  navigator.clipboard.writeText(rows.map(r => r['ID']).join('\n'))
    .then(() => showToast(`📋 ${rows.length.toLocaleString()} IDs copiados`));
}


function parseFechaDistribucionNC(v){
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

function ordenarPorHoraCreacionNCParaDistribuir(rows){
  return [...rows].sort((a,b) => {
    const fa = parseFechaDistribucionNC(a['Hora de creación'] || a['Hora de creacion'] || a['Created Time']);
    const fb = parseFechaDistribucionNC(b['Hora de creación'] || b['Hora de creacion'] || b['Created Time']);
    return fa - fb;
  });
}


/* === OVERRIDE FINAL: Detalle asesor por mentor — Base No Contactado === */
function getMentorNC(row){
  const asesor = row['CORREO ASESOR ASIGNADO'] || '';
  return String(row['MENTOR'] || mentorDeAsesor(asesor) || 'Sin mentor').trim() || 'Sin mentor';
}

function renderDetalleAsesorNC(){
  const card = document.getElementById('nc-detalle-asesor-card');
  const box = document.getElementById('nc-detalle-asesor');
  const sel = document.getElementById('nc-mentor-filter');

  if(!card || !box || !sel) return;

  card.style.display = 'block';
  card.classList.add('detalle-visible');
  card.classList.remove('detalle-placeholder');

  const data = typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [];
  if(!data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el resumen por mentor y asesor.</div>';
    return;
  }

  const mentores = [...new Set(data.map(getMentorNC))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  const actual = sel.value || '';
  sel.innerHTML = '<option value="">Todos</option>' + mentores.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  if(mentores.includes(actual)) sel.value = actual;

  const filtro = sel.value;
  const rows = filtro ? data.filter(r => getMentorNC(r) === filtro) : data;

  const grouped = {};
  rows.forEach(r => {
    const mentor = getMentorNC(r);
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


/* === DETALLE ASESOR FUNCIONAL — Base No Contactado === */
function getMentorNC(row){
  return String(row['MENTOR'] || mentorDeAsesor(row['CORREO ASESOR ASIGNADO']) || 'Sin supervisor').trim();
}

function renderDetalleAsesorNC(){
  renderDetalleAsesorGenerico({
    data: (typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []),
    cardId: 'nc-detalle-asesor-card',
    boxId: 'nc-detalle-asesor',
    selectId: 'nc-mentor-filter',
    getMentor: getMentorNC,
    renderAgain: 'renderDetalleAsesorNC'
  });
}


/* === OVERRIDE FINAL DISTRIBUIR NO CONTACTADO === */
function distribuirNC(){
  if(!ncFiltered.length){showToast('No hay leads para distribuir.');return;}
  if(!ncAsesoresList.length){showToast('Carga primero el Excel de asesores.');return;}
  const leadsOrdenados = ordenarPorHoraCreacionNCParaDistribuir(ncFiltered);
  const n = ncAsesoresList.length;
  ncDistribuidoData = leadsOrdenados.map((lead,i)=>asignarSupervisorSimple({'ID de registro':getNCLeadId(lead)}, ncAsesoresList[i % n]));
  renderNCDistPreview();
  showNCDistStats();
  renderDetalleAsesorNC();
  showToast(`✅ ${ncDistribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}
