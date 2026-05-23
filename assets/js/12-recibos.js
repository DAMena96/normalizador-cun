
/* ════════════════════════════════════════════
   RECIBOS — SOLO CARI AI Y NOTAS
   ════════════════════════════════════════════ */

let recData = [], recFiltered = [], recCariPage = 1;
const REC_PAGE_SIZE = 100;

const REC_CARI_COLS = ['numero_telefono','nombre_aspirante','carrera_interes','identificacion','correo_electronico','telefono_adicional','periodo','campania'];
const REC_NOTAS_COLS = ['ID','Notes'];
const REC_NOTA_FIJA = 'Se intenta contacto con el aspirante via CARI AI';

const REC_CARI_FIXED_ROWS = [
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

const REC_FILTER_DEFS = [
  {id:'rec-val', col:'VAL', lbl:'VAL'},
  {id:'rec-telefono', col:'VALIDACION TELEFONO', lbl:'Validación teléfono'},
  {id:'rec-periodo', col:'Periodo CARI', lbl:'Periodo CARI'}
];

function handleDropRecibos(e){
  e.preventDefault();
  document.getElementById('rec-upload-zone').classList.remove('over');
  loadRecibosFile(e.dataTransfer.files[0]);
}

function recShowProg(pct,label){
  document.getElementById('rec-progress-wrap').style.display='block';
  document.getElementById('rec-progress-fill').style.width=pct+'%';
  document.getElementById('rec-progress-label').textContent=label;
}
function recHideProg(){
  document.getElementById('rec-progress-wrap').style.display='none';
  document.getElementById('rec-progress-fill').style.width='0';
}

async function loadRecibosFile(file){
  if(!file) return;
  recShowProg(5,'Cargando catálogos JSON…');
  const okCatalogs = await catalogsReady;
  if(!okCatalogs){ recHideProg(); return; }

  recShowProg(15,'Leyendo archivo…');
  const reader = new FileReader();
  reader.onload = ev => {
    try{
      recShowProg(35,'Parseando Excel…');
      const wb = XLSX.read(ev.target.result,{type:'array'});
      const sn = wb.SheetNames[0];
      const json = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
      if(json.length < 2){ showToast('Sin datos en la base de recibos.'); recHideProg(); return; }
      const headers = json[0].map(h => String(h).trim());

      recData = [];
      for(let i=1;i<json.length;i++){
        const row = json[i];
        if(row.every(c => c === '' || c === null || c === undefined)) continue;
        recData.push(normalizeReciboRow(row, headers));
      }

      recFiltered = [...recData];
      buildRecFilterPanel();
      populateRecFilters();
      renderRecKPIs();
      renderRecCari();

      document.getElementById('rec-upload-zone').classList.add('hidden');
      document.getElementById('rec-kpi-row').classList.remove('hidden');
      document.getElementById('rec-tabs').classList.remove('hidden');
      document.querySelectorAll('.rec-panel').forEach(p => p.classList.remove('hidden'));
      switchRecTab(0);

      recHideProg();
      showToast(`✅ ${recData.length.toLocaleString()} recibos procesados`);
    }catch(err){
      console.error(err);
      showToast('Error Recibos: ' + err.message);
      recHideProg();
    }
  };
  reader.readAsArrayBuffer(file);
}

function normalizeReciboRow(row, headers){
  const hIndex = new Map(headers.map((h,i)=>[cleanHeader(h),i]));
  const g = (...names) => {
    for(const name of names){
      const i = hIndex.get(cleanHeader(name));
      if(i !== undefined && row[i] != null) return String(row[i]).trim();
    }
    return '';
  };

  const r = {};
  const cols = [
    'ID de registro','Hora de creación','Correo electrónico personal','Estado del interesado',
    'Etapa del Registro','Propietario de Interesado','Nombre completo','Campaña mercadeo',
    'Teléfono','Número de Documento','Programa de interes','Correo institucional',
    'Número de gestiones','Modalidad','Periodo','Tipo de homologación','Sub.estado',
    'Sub.estado II','Hora de la última actividad','Estado a tipificar','Creado por',
    'Ciudad de residencia','Fecha última nota','utm_campaign','Programa de interés_',
    'Fecha pago oportuno','Hora de modificación','Programa'
  ];
  cols.forEach(c => r[c] = g(c, sinTildesRec(c)));

  const progSrc = r['Programa'] || r['Programa de interes'] || r['Programa de interés_'] || 'Sin programa';
  const programaNormalizado = cleanRecCatalog(lookup(MAP_PROGRAMA, toRecTitle(progSrc), progSrc || 'SIN PROGRAMA'));
  r['PROGRAMA NORMALIZADO'] = programaNormalizado;
  r['VAL'] = cleanRecCatalog(lookup(MAP_AREA, programaNormalizado, 'OTRO') || 'OTRO');
  r['VALIDACION TELEFONO'] = validarTelefonoRecibo(r['Teléfono']);
  r['TELEFONO CARI'] = getRecPhone(r);
  r['Periodo CARI'] = getRecPeriodoCari(r);

  return r;
}

function validarTelefonoRecibo(tel){
  const n = normalizarTelefono(tel);
  if(!n || n === '0') return 'Sin número';
  if(n.length < 10) return 'Número errado';
  const dos = Number(n.slice(0,2));
  if(dos < 30) return 'Número errado';
  if(dos > 35) return 'Número errado';
  return 'Ok';
}

function getRecPhone(row){
  const n = normalizarTelefono(row['Teléfono']);
  if(!n) return '';
  return n.startsWith('57') ? n : '57' + n;
}

function getRecPeriodoCari(row){
  const val = String(row['VAL'] || '').trim().toUpperCase();
  return val === 'POSGRADO' ? 'Virtual/Especialización' : 'Virtual/pregrado';
}

function buildRecCariRows(){
  const rows = recFiltered.map(r => {
    const phone = getRecPhone(r);
    const doc = String(r['Número de Documento'] || '').trim();
    return {
      'numero_telefono': phone,
      'nombre_aspirante': r['Nombre completo'] || '',
      'carrera_interes': String(r['PROGRAMA NORMALIZADO'] || '').toLowerCase(),
      'identificacion': doc || 'Sin documento',
      'correo_electronico': r['Correo electrónico personal'] || '',
      'telefono_adicional': phone,
      'periodo': getRecPeriodoCari(r),
      'campania': r['Campaña mercadeo'] || ''
    };
  });
  return [...REC_CARI_FIXED_ROWS, ...rows];
}

function buildRecNotasRows(){
  return recFiltered.map(r => ({
    'ID': getRecLeadId(r),
    'Notes': REC_NOTA_FIJA
  }));
}

function switchRecTab(n){
  document.querySelectorAll('#rec-tabs .tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('#recibos-module .rec-panel').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n === 0) renderRecCari();
  if(n === 1) updateRecNotasInfo();
}

function buildRecFilterPanel(){
  const row = document.getElementById('rec-filter-row');
  row.innerHTML = '';
  REC_FILTER_DEFS.forEach(f=>{
    const div = document.createElement('div');
    div.className='fg';
    div.innerHTML = `
      <label>${f.lbl}<span class="clr" onclick="clearOneRec('${f.id}')">✖</span></label>
      <div class="ms-wrap">
        <button class="ms-btn" onclick="toggleMS('${f.id}',event)">
          <span class="st" id="txt-${f.id}">Todas</span>
          <span style="opacity:.4;font-size:.62rem;flex-shrink:0;margin-left:2px">▼</span>
        </button>
        <div class="ms-dd" id="${f.id}">
          <div class="ms-sb"><input type="text" placeholder="Buscar…" oninput="filterOpts('${f.id}',this.value)"></div>
          <div class="ms-all" onclick="toggleAllRec('${f.id}')">☑ Sel/Des todo</div>
          <div class="ms-list" id="list-${f.id}"></div>
        </div>
      </div>`;
    row.appendChild(div);
  });

  const txt = document.createElement('div');
  txt.className='fg'; txt.style.minWidth='190px';
  txt.innerHTML = `<label>Buscar nombre/correo/tel/ID</label><input type="text" id="rec-buscar" placeholder="Escribe para buscar…" oninput="applyRecFilters()">`;
  row.appendChild(txt);

  const clrb = document.createElement('div');
  clrb.className='fg'; clrb.style.flex='0'; clrb.style.minWidth='auto';
  clrb.innerHTML = `<label>&nbsp;</label><button class="btn btn-gray" onclick="clearRecFilters()">✖ Limpiar</button>`;
  row.appendChild(clrb);
}

function populateRecFilters(){
  REC_FILTER_DEFS.forEach(f=>{
    const vals = [...new Set(recData.map(r => String(r[f.col] || '').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    document.getElementById('list-'+f.id).innerHTML = vals.map(v =>
      `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(v)}" onchange="onCheckRec('${f.id}')"> <span>${escapeHtml(v)}</span></label>`
    ).join('');
    updateRecFilterText(f.id);
  });
}

function onCheckRec(id){ applyRecFilters(); updateRecFilterText(id); closeAll(); }
function toggleAllRec(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input')];
  const all=checks.every(c=>c.checked);
  checks.forEach(c=>c.checked=!all);
  applyRecFilters(); updateRecFilterText(id); closeAll();
}
function clearOneRec(id){
  document.querySelectorAll('#list-'+id+' input').forEach(c=>c.checked=false);
  applyRecFilters(); updateRecFilterText(id); closeAll();
}
function clearRecFilters(){
  REC_FILTER_DEFS.forEach(f=>{
    document.querySelectorAll('#list-'+f.id+' input').forEach(c=>c.checked=false);
    updateRecFilterText(f.id);
  });
  const b=document.getElementById('rec-buscar'); if(b)b.value='';
  applyRecFilters();
}
function updateRecFilterText(id){
  const checks=[...document.querySelectorAll('#list-'+id+' input:checked')];
  const txt=document.getElementById('txt-'+id);
  if(txt) txt.textContent = checks.length ? (checks.length===1 ? checks[0].value : `${checks.length} seleccionados`) : 'Todas';
}

function applyRecFilters(){
  const q = String(document.getElementById('rec-buscar')?.value || '').toLowerCase().trim();
  recFiltered = recData.filter(r=>{
    for(const f of REC_FILTER_DEFS){
      const vals=[...document.querySelectorAll('#list-'+f.id+' input:checked')].map(c=>c.value);
      updateRecFilterText(f.id);
      if(vals.length && !vals.includes(String(r[f.col] || ''))) return false;
    }
    if(q){
      const hay=[r['ID de registro'],r['Nombre completo'],r['Correo electrónico personal'],r['Teléfono'],r['PROGRAMA NORMALIZADO'],r['VAL']].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });
  recCariPage=1;
  renderRecKPIs(); renderRecCari(); updateRecNotasInfo();
}

function renderRecKPIs(){
  document.getElementById('rk-total').textContent = recData.length.toLocaleString();
  document.getElementById('rk-ok').textContent = recFiltered.filter(r=>r['VALIDACION TELEFONO']==='Ok').length.toLocaleString();
  document.getElementById('rk-err').textContent = recFiltered.filter(r=>r['VALIDACION TELEFONO']!=='Ok').length.toLocaleString();
  document.getElementById('rk-cari').textContent = buildRecCariRows().length.toLocaleString();
}

function renderRecCari(){
  const rows = buildRecCariRows();
  const tbl = document.getElementById('rec-cari-tbl');
  if(!tbl) return;
  const pages = Math.max(1, Math.ceil(rows.length / REC_PAGE_SIZE));
  recCariPage = Math.min(recCariPage, pages);
  const sl = rows.slice((recCariPage-1)*REC_PAGE_SIZE, recCariPage*REC_PAGE_SIZE);
  document.getElementById('rec-cari-title').textContent = `${rows.length.toLocaleString()} registros CARI AI`;
  tbl.innerHTML = '<thead><tr>' + REC_CARI_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    sl.map(r => '<tr>' + REC_CARI_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  document.getElementById('rec-cari-pag-info').textContent = `Página ${recCariPage} de ${pages}`;
  renderRecKPIs();
}

function changeRecCariPage(n){
  const pages = Math.max(1, Math.ceil(buildRecCariRows().length / REC_PAGE_SIZE));
  recCariPage = Math.min(Math.max(1, recCariPage + n), pages);
  renderRecCari();
}

function renderRecNotas(){
  const rows = buildRecNotasRows();
  const wrap = document.getElementById('rec-notas-preview');
  const tbl = document.getElementById('rec-notas-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + REC_NOTAS_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + REC_NOTAS_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  updateRecNotasInfo();
}

function updateRecNotasInfo(){
  const el = document.getElementById('rec-notas-info');
  if(el) el.textContent = `${recFiltered.length.toLocaleString()} notas listas para generar (filtro activo)`;
}

function exportRecCari(){
  const rows = buildRecCariRows();
  if(!rows.length){ showToast('No hay registros CARI AI para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r=>{
    const o={}; REC_CARI_COLS.forEach(c=>o[c]=r[c]??''); return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CARI AI');
  XLSX.writeFile(wb, 'Recibos_CARI_AI_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} registros CARI AI exportados`);
}

function exportRecNotas(){
  const rows = buildRecNotasRows();
  if(!rows.length){ showToast('No hay notas para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r=>{
    const o={}; REC_NOTAS_COLS.forEach(c=>o[c]=r[c]??''); return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  XLSX.writeFile(wb, 'Recibos_Notas_CARI_AI_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} notas exportadas`);
}

function copiarRecNotasIDs(){
  const rows = buildRecNotasRows();
  navigator.clipboard.writeText(rows.map(r => r['ID']).join('\n'))
    .then(() => showToast(`📋 ${rows.length.toLocaleString()} IDs copiados`));
}

function getRecLeadId(row){
  if(!row) return '';
  const direct = row['ID de registro'];
  if(direct !== undefined && direct !== null && String(direct).trim() !== '') return String(direct).trim();
  return '';
}

function sinTildesRec(s){
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function toRecTitle(s){
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/\b([a-záéíóúñ])/g, c => c.toUpperCase());
}
function cleanRecCatalog(v){
  const s = String(v ?? '').trim();
  if(!s) return '-';
  return s.toUpperCase();
}
