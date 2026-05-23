/* ════════════════════════════════════════════
   TABS
   ════════════════════════════════════════════ */
function switchTab(n){
  document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('.tab-panel').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n===1)
    document.getElementById('dist-info').textContent=
      `${filteredData.length.toLocaleString()} leads listos para distribuir (filtro activo)`;
}


/* ════════════════════════════════════════════
   MULTISELECT
   ════════════════════════════════════════════ */
function buildFilterPanel(){
  const row=document.getElementById('filter-row');
  row.innerHTML='';
  FILTER_DEFS.forEach(f=>{
    const div=document.createElement('div');
    div.className='fg';
    div.innerHTML=`
      <label>${f.lbl}<span class="clr" onclick="clearOne('${f.id}')">✖</span></label>
      <div class="ms-wrap">
        <button class="ms-btn" onclick="toggleMS('${f.id}',event)">
          <span class="st" id="txt-${f.id}">Todas</span>
          <span style="opacity:.4;font-size:.62rem;flex-shrink:0;margin-left:2px">▼</span>
        </button>
        <div class="ms-dd" id="${f.id}">
          <div class="ms-sb"><input type="text" placeholder="Buscar…" oninput="filterOpts('${f.id}',this.value)"></div>
          <div class="ms-all" onclick="toggleAll('${f.id}')">☑ Sel/Des todo</div>
          <div class="ms-list" id="list-${f.id}"></div>
        </div>
      </div>`;
    row.appendChild(div);
  });
  const txt=document.createElement('div');
  txt.className='fg'; txt.style.minWidth='165px';
  txt.innerHTML=`<label>Buscar nombre/correo/tel/ID</label>
    <input type="text" id="f-buscar" placeholder="Escribe para buscar…" oninput="applyFilters()">`;
  row.appendChild(txt);
  const clrb=document.createElement('div');
  clrb.className='fg'; clrb.style.flex='0'; clrb.style.minWidth='auto';
  clrb.innerHTML=`<label>&nbsp;</label><button class="btn btn-gray" onclick="clearFilters()">✖ Limpiar</button>`;
  row.appendChild(clrb);
}

function populateFilters(){
  FILTER_DEFS.forEach(f=>{
    const vals=[...new Set(allData.map(r=>String(r[f.col]||'')).filter(v=>v!==''))].sort((a,b)=>a.localeCompare(b));
    document.getElementById('list-'+f.id).innerHTML=vals.map(v=>
      `<label class="ms-opt"><input type="checkbox" value="${v.replace(/"/g,'&quot;')}" onchange="onCheck('${f.id}')"><span>${v}</span></label>`
    ).join('');
  });
}

function toggleMS(id,e){
  e.stopPropagation();
  const dd=document.getElementById(id);
  const was=dd.classList.contains('open');
  closeAll();
  if(!was){dd.classList.add('open');document.getElementById('ov').classList.add('on');}
}
function closeAll(){
  document.querySelectorAll('.ms-dd.open').forEach(d=>d.classList.remove('open'));
  const ov=document.getElementById('ov');
  if(ov)ov.classList.remove('on');
}
function filterOpts(id,q){
  document.querySelectorAll(`#list-${id} .ms-opt`).forEach(o=>{
    o.style.display=o.querySelector('span').textContent.toLowerCase().includes(q.toLowerCase())?'':'none';
  });
}
function toggleAll(id){
  const checks=[...document.querySelectorAll(`#list-${id} input`)];
  const all=checks.every(c=>c.checked);
  msSel[id].clear();
  checks.forEach(c=>{c.checked=!all;if(!all)msSel[id].add(c.value);});
  updateLabel(id);applyFilters();
}
function onCheck(id){
  msSel[id].clear();
  document.querySelectorAll(`#list-${id} input`).forEach(c=>{if(c.checked)msSel[id].add(c.value);});
  updateLabel(id);applyFilters();
}
function updateLabel(id){
  const s=msSel[id],el=document.getElementById('txt-'+id);
  if(!el)return;
  el.textContent=s.size===0?'Todas':s.size===1?[...s][0]:`${s.size} seleccionados`;
}
function clearOne(id){
  msSel[id].clear();
  document.querySelectorAll(`#list-${id} input`).forEach(c=>c.checked=false);
  updateLabel(id);applyFilters();
}
function clearFilters(){
  FILTER_DEFS.forEach(f=>clearOne(f.id));
  const fb=document.getElementById('f-buscar');
  if(fb)fb.value='';
  applyFilters();
}


/* ════════════════════════════════════════════
   FILTRAR
   ════════════════════════════════════════════ */
function applyFilters(){
  const fb=(document.getElementById('f-buscar')?.value||'').toLowerCase();
  filteredData=allData.filter(r=>{
    for(const f of FILTER_DEFS){
      if(msSel[f.id].size>0&&!msSel[f.id].has(String(r[f.col]||'')))return false;
    }
    if(fb){
      const s=((r['Nombre completo']||'')+(r['CORREO ELECTRONICO DEF']||'')+
               (r['Número DEF']||'')+(r['ID de registro']||'')).toLowerCase();
      if(!s.includes(fb))return false;
    }
    return true;
  });
  currentPage=1;renderKPIs();renderTable();
}


