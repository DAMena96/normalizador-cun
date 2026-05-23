/* ════════════════════════════════════════════
   DISTRIBUCIÓN
   Vista y exportación con:
   - ID de registro
   - CORREO ASESOR ASIGNADO
   - MENTOR
   ════════════════════════════════════════════ */

const DIST_COLS = ['ID de registro', 'CORREO ASESOR ASIGNADO', 'MENTOR'];
const DIST_SHOW = ['ID de registro', 'CORREO ASESOR ASIGNADO', 'MENTOR'];

function dropAs(e){
  e.preventDefault();
  loadAsesores(e.dataTransfer.files[0]);
}

function loadAsesores(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:'array'});
      const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const col=Object.keys(json[0]||{}).find(k=>/correo|email/i.test(k));
      if(!col){showToast('No se encontró columna Correo o Email');return;}
      asesoresList=[...new Set(json.map(r=>String(r[col]).trim()).filter(Boolean))];
      renderTags();
      showToast(`✅ ${asesoresList.length} asesores cargados`);
    }catch(err){
      showToast('Error asesores: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderTags(){
  document.getElementById('asesores-tags').innerHTML=
    `<p style="font-size:.77rem;color:#666;margin-bottom:4px"><strong>${asesoresList.length} asesores:</strong></p>`+
    asesoresList.map((a,i)=>`<span class="tag">${escapeHtml(a)}<span class="rm" onclick="rmAsesor(${i})">×</span></span>`).join('');
}

function rmAsesor(i){
  asesoresList.splice(i,1);
  renderTags();
}

function getLeadId(row){
  if(!row) return '';

  // 1) Caso normal del normalizador
  const direct = row['ID de registro'];
  if(direct !== undefined && direct !== null && String(direct).trim() !== '') {
    return String(direct).trim();
  }

  // 2) Fallback por nombres comunes
  const possibleNames = [
    'Id de registro',
    'ID Registro',
    'ID',
    'Id',
    'Record ID',
    'Record Id',
    'record_id',
    'Lead ID',
    'Lead Id'
  ];

  for(const name of possibleNames){
    if(row[name] !== undefined && row[name] !== null && String(row[name]).trim() !== '') {
      return String(row[name]).trim();
    }
  }

  // 3) Fallback flexible usando encabezados limpios
  const wanted = [
    'IDDEREGISTRO',
    'IDREGISTRO',
    'RECORDID',
    'LEADID',
    'ID'
  ];

  for(const key of Object.keys(row)){
    const clean = typeof cleanHeader === 'function'
      ? cleanHeader(key)
      : String(key).toUpperCase().replace(/[^A-Z0-9]/g,'');

    if(wanted.includes(clean)){
      const value = row[key];
      if(value !== undefined && value !== null && String(value).trim() !== '') {
        return String(value).trim();
      }
    }
  }

  return '';
}

function getDistRow(row){
  const asesor = row?.['CORREO ASESOR ASIGNADO'] || '';
  const mentor = row?.['MENTOR'] || row?.['SUPERVISOR'] || (typeof mentorDeAsesor === 'function' ? mentorDeAsesor(asesor) : '');

  return {
    'ID de registro': getLeadId(row),
    'CORREO ASESOR ASIGNADO': asesor,
    'MENTOR': mentor || ''
  };
}

function distribuir(){
  if(!filteredData.length){showToast('No hay leads cargados');return;}
  if(!asesoresList.length){showToast('⚠️ Debes cargar el Excel de asesores primero');return;}

  const leadsOrdenados = ordenarPorHoraCreacionParaDistribuir(filteredData);
  const n = asesoresList.length;
  distribuidoData = leadsOrdenados.map((lead, i) => {
    const asesor = asesoresList[i % n];
    return asignarSupervisorSimple({ ...lead, 'ID de registro': getLeadId(lead) }, asesor);
  });

  const res = {};
  distribuidoData.forEach(r=>{
    const a = r['CORREO ASESOR ASIGNADO'];
    res[a] = (res[a] || 0) + 1;
  });

  const sb=document.getElementById('stat-box');
  sb.style.display='block';
  sb.innerHTML=`<strong>📊 ${filteredData.length.toLocaleString()} leads · ${n} asesores:</strong><br><br>`+
    Object.entries(res).map(([a,c])=>`<span class="dist-chip">${escapeHtml(a)}: <b>${c.toLocaleString()}</b></span>`).join('')+
    htmlDetalleVisible(distribuidoData,'Detalle_Asesor_Sin_Gestion');

  renderDistribucionPreview();
  renderDetalleAsesorSG();
  setTimeout(renderDetalleAsesorSG, 100);
  showToast(`✅ ${distribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}

function renderDistribucionPreview(){
  const rows = distribuidoData.slice(0,300).map(getDistRow);

  const head='<thead style="background:var(--azul);color:white;position:sticky;top:0"><tr>'+ 
    DIST_SHOW.map(c=>`<th style="padding:7px 9px;white-space:nowrap;font-size:.72rem">${escapeHtml(c)}</th>`).join('')+
    '</tr></thead>';

  const body='<tbody>'+rows.map((r,i)=>
    `<tr style="${i%2?'background:#f8fafc':''}">`+
    DIST_SHOW.map(c=>{
      const v = r[c] || '';
      return `<td style="padding:5px 9px;white-space:nowrap;font-size:.72rem;border-bottom:1px solid #eef1f5;max-width:360px;overflow:hidden;text-overflow:ellipsis" title="${escapeHtml(v)}">${escapeHtml(v)}</td>`;
    }).join('')+
    '</tr>'
  ).join('')+'</tbody>';

  const prev=document.getElementById('dist-preview');
  prev.style.display='block';
  prev.innerHTML=`<table id="dist-tbl" style="width:100%;border-collapse:collapse">${head}${body}</table>`+
    `<div style="padding:8px 12px;font-size:.77rem;color:#888;border-top:1px solid var(--borde)">Mostrando ${Math.min(300,distribuidoData.length)} de ${distribuidoData.length.toLocaleString()} · exporta para ver todo</div>`;
  if(typeof renderDetalleAsesorSG === 'function') setTimeout(renderDetalleAsesorSG, 50);
}

function exportarDist(){
  if(!distribuidoData.length){showToast('Primero distribuye los leads');return;}

  const rows = distribuidoData.map(getDistRow);

  // Validación rápida para evitar exportar archivo vacío de IDs sin avisar.
  const sinId = rows.filter(r => !r['ID de registro']).length;
  if(sinId > 0){
    console.warn(`Hay ${sinId} registros sin ID de registro. Revisa el encabezado del Excel original.`);
  }

  const ws=XLSX.utils.json_to_sheet(rows, { header: DIST_COLS });
  const wb2=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2,ws,'Distribuido');
  XLSX.writeFile(wb2,'Leads_Distribuidos_ID_Asesor_Mentor_'+Date.now()+'.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} leads exportados con mentor`);
}

function escapeHtml(value=''){
  return String(value)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

function showToast(msg){
  const t=document.getElementById('toast');
  t.textContent=msg;
  t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'),3500);
}


function parseFechaDistribucion(v){
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

function ordenarPorHoraCreacionParaDistribuir(rows){
  return [...rows].sort((a,b) => {
    const fa = parseFechaDistribucion(a['Hora de creación'] || a['Hora de creacion'] || a['Created Time']);
    const fb = parseFechaDistribucion(b['Hora de creación'] || b['Hora de creacion'] || b['Created Time']);
    return fa - fb;
  });
}


/* Override switchTab para incluir CARI AI y Notas en Base Sin Gestión */
function switchTab(n){
  document.querySelectorAll('#tabs .tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('#tp-0,#tp-1,#tp-2,#tp-3').forEach((p,i)=>p.classList.toggle('active',i===n));
  if(n===1){
    const el = document.getElementById('dist-info');
    if(el) el.textContent = `${filteredData.length.toLocaleString()} leads listos para distribuir (filtro activo)`;
  }
  if(n===2) updateSGCariInfo();
  if(n===3) updateSGNotasInfo();
}

/* CARI AI + NOTAS — BASE SIN GESTIÓN */
const SG_CARI_COLS = ['numero_telefono','nombre_aspirante','carrera_interes','identificacion','correo_electronico','telefono_adicional','periodo','campania'];
const SG_NOTAS_COLS = ['ID','Notes'];
const SG_NOTA_FIJA = 'Se intenta contacto con el aspirante via CARI AI';

const SG_CARI_FIXED_ROWS = [
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

function updateSGCariInfo(){
  const el = document.getElementById('sg-cari-info');
  if(el) el.textContent = `${filteredData.length.toLocaleString()} registros listos para CARI AI (filtro activo)`;
}

function updateSGNotasInfo(){
  const el = document.getElementById('sg-notas-info');
  if(el) el.textContent = `${filteredData.length.toLocaleString()} notas listas para generar (filtro activo)`;
}

function getSGCariPhone(row){
  const tel = normalizarTelefono(row['Número DEF'] || row['Teléfono'] || row['Teléfono Continua']);
  if(!tel) return '';
  return tel.startsWith('57') ? tel : '57' + tel;
}

function getSGCariPeriodo(row){
  const area = String(row['VALIDACION FINAL'] || row['AREA PROGRAMA'] || '').trim().toUpperCase();
  return area === 'POSGRADO' ? 'Virtual/Especialización' : 'Virtual/pregrado';
}

function getSGName(fullName){
  return String(fullName || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/\b([a-záéíóúñ])/g, c => c.toUpperCase());
}

function buildSGCariRows(){
  const rows = filteredData.map(r => {
    const phone = getSGCariPhone(r);
    const doc = String(r['Número de Documento'] || '').trim();
    return {
      'numero_telefono': phone,
      'nombre_aspirante': getSGName(r['Nombre completo']),
      'carrera_interes': String(r['PROGRAMA2'] || r['Programa'] || '').toLowerCase(),
      'identificacion': doc || 'Sin documento',
      'correo_electronico': r['CORREO ELECTRONICO DEF'] || r['Correo electrónico'] || '',
      'telefono_adicional': phone,
      'periodo': getSGCariPeriodo(r),
      'campania': 'Organico'
    };
  });
  return [...SG_CARI_FIXED_ROWS, ...rows];
}

function buildSGNotasRows(){
  return filteredData.map(r => ({
    'ID': getLeadId(r),
    'Notes': SG_NOTA_FIJA
  }));
}

function renderSGCari(){
  const rows = buildSGCariRows();
  const wrap = document.getElementById('sg-cari-preview');
  const tbl = document.getElementById('sg-cari-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + SG_CARI_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + SG_CARI_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 registros de ${rows.length.toLocaleString()}</caption>`;
  }
  updateSGCariInfo();
}

function renderSGNotas(){
  const rows = buildSGNotasRows();
  const wrap = document.getElementById('sg-notas-preview');
  const tbl = document.getElementById('sg-notas-tbl');
  if(!wrap || !tbl) return;
  wrap.style.display = 'block';
  const preview = rows.slice(0,300);
  tbl.innerHTML = '<thead><tr>' + SG_NOTAS_COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>' +
    preview.map(r => '<tr>' + SG_NOTAS_COLS.map(c => `<td title="${escapeHtml(String(r[c] ?? ''))}">${escapeHtml(String(r[c] ?? ''))}</td>`).join('') + '</tr>').join('') +
    '</tbody>';
  if(rows.length > 300){
    tbl.innerHTML += `<caption style="caption-side:bottom;padding:8px;color:#777">Vista previa de 300 notas de ${rows.length.toLocaleString()}</caption>`;
  }
  updateSGNotasInfo();
}

function exportSGCari(){
  const rows = buildSGCariRows();
  if(!rows.length){ showToast('No hay registros CARI AI para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    SG_CARI_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'CARI AI');
  XLSX.writeFile(wb, 'Sin_Gestion_CARI_AI_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} registros CARI AI exportados`);
}

function exportSGNotas(){
  const rows = buildSGNotasRows();
  if(!rows.length){ showToast('No hay notas para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows.map(r => {
    const o = {};
    SG_NOTAS_COLS.forEach(c => o[c] = r[c] ?? '');
    return o;
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Notas');
  XLSX.writeFile(wb, 'Sin_Gestion_Notas_CARI_AI_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} notas exportadas`);
}

function copiarSGNotasIDs(){
  const rows = buildSGNotasRows();
  navigator.clipboard.writeText(rows.map(r => r['ID']).join('\n'))
    .then(() => showToast(`📋 ${rows.length.toLocaleString()} IDs copiados`));
}


/* === OVERRIDE FINAL: Detalle asesor por mentor — Base Sin Gestión === */
function getMentorSG(row){
  const asesor = row['CORREO ASESOR ASIGNADO'] || '';
  return String(row['MENTOR'] || mentorDeAsesor(asesor) || 'Sin mentor').trim() || 'Sin mentor';
}

function renderDetalleAsesorSG(){
  const card = document.getElementById('sg-detalle-asesor-card');
  const box = document.getElementById('sg-detalle-asesor');
  const sel = document.getElementById('sg-mentor-filter');

  if(!card || !box || !sel) return;

  card.style.display = 'block';
  card.classList.add('detalle-visible');
  card.classList.remove('detalle-placeholder');

  if(!window.distribuidoData && typeof distribuidoData === 'undefined') return;
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  if(!data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el resumen por mentor y asesor.</div>';
    return;
  }

  const mentores = [...new Set(data.map(getMentorSG))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  const actual = sel.value || '';
  sel.innerHTML = '<option value="">Todos</option>' + mentores.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  if(mentores.includes(actual)) sel.value = actual;

  const filtro = sel.value;
  const rows = filtro ? data.filter(r => getMentorSG(r) === filtro) : data;

  const grouped = {};
  rows.forEach(r => {
    const mentor = getMentorSG(r);
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


/* Render único para Detalle asesor por supervisor */
function renderDetalleAsesorGenerico(cfg){
  const card = document.getElementById(cfg.cardId);
  const box = document.getElementById(cfg.boxId);
  const sel = document.getElementById(cfg.selectId);
  if(!card || !box || !sel) return;

  card.style.display = 'block';
  card.classList.add('detalle-visible');
  card.classList.remove('detalle-placeholder');

  const data = Array.isArray(cfg.data) ? cfg.data : [];
  if(!data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el resumen por supervisor y asesor.</div>';
    return;
  }

  const supervisores = [...new Set(data.map(cfg.getMentor))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  const actual = sel.value || '';
  sel.innerHTML = '<option value="">Todos</option>' + supervisores.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  if(supervisores.includes(actual)) sel.value = actual;

  const filtro = sel.value;
  const rows = filtro ? data.filter(r => cfg.getMentor(r) === filtro) : data;

  const grouped = {};
  rows.forEach(r => {
    const supervisor = cfg.getMentor(r);
    const asesor = r['CORREO ASESOR ASIGNADO'] || 'Sin asesor';
    if(!grouped[supervisor]) grouped[supervisor] = {};
    if(!grouped[supervisor][asesor]){
      grouped[supervisor][asesor] = {
        total: 0,
        area: r['AREA ASESOR'] || areaDeAsesor(asesor),
        grupo: r['GRUPO ASESOR'] || grupoDeAsesor(asesor)
      };
    }
    grouped[supervisor][asesor].total++;
  });

  const totalAsesores = new Set(rows.map(r => r['CORREO ASESOR ASIGNADO'] || 'Sin asesor')).size;

  box.innerHTML = `
    <div class="detalle-summary">
      <div><strong>${rows.length.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${Object.keys(grouped).length.toLocaleString()}</strong><span>Supervisores</span></div>
    </div>
  ` + Object.entries(grouped).map(([supervisor, asesores]) => {
    const totalSupervisor = Object.values(asesores).reduce((a,b)=>a+b.total,0);
    const cards = Object.entries(asesores)
      .sort((a,b)=>b[1].total-a[1].total)
      .map(([asesor,info]) =>
        `<div class="detalle-asesor-item">
          <span>
            <b>${escapeHtml(asesor)}</b>
            <small>${escapeHtml(info.area || 'Sin área')}${info.grupo ? ' · ' + escapeHtml(info.grupo) : ''}</small>
          </span>
          <strong>${info.total.toLocaleString()}</strong>
        </div>`
      ).join('');

    return `<div class="detalle-mentor-card">
      <div class="detalle-mentor-head"><span>👤 ${escapeHtml(supervisor)}</span><b>${totalSupervisor.toLocaleString()} leads</b></div>
      <div class="detalle-asesor-grid">${cards}</div>
    </div>`;
  }).join('');
}


/* === DETALLE ASESOR FUNCIONAL — Base Sin Gestión === */
function getMentorSG(row){
  return String(row['MENTOR'] || mentorDeAsesor(row['CORREO ASESOR ASIGNADO']) || 'Sin supervisor').trim();
}

function renderDetalleAsesorSG(){
  renderDetalleAsesorGenerico({
    data: (typeof distribuidoData !== 'undefined' ? distribuidoData : []),
    cardId: 'sg-detalle-asesor-card',
    boxId: 'sg-detalle-asesor',
    selectId: 'sg-mentor-filter',
    getMentor: getMentorSG,
    renderAgain: 'renderDetalleAsesorSG'
  });
}


/* === SUPERVISOR + DETALLE INTEGRADO DEFINITIVO === */
const SUPERVISOR_EMBEBIDO = {"Vinculaciones Carlos Andres Burgos Carvajal": "Otro", "Vinculaciones LAURA CAMILA CASTILLO SEPULVEDA": "Otro", "Vinculaciones Erika Ramirez Rico": "Otro", "Vinculaciones Laura Alejandra Guzman Gutierrez": "Otro", "Vinculaciones Margarita Rosa Navarro Garzon": "Otro", "Vinculaciones Maria Alejandra Vergara Chavez": "Otro", "Vinculaciones Maria Claudia Salas Morales": "Otro", "Vinculaciones Regionales Diego Villota": "Otro", "Vinculaciones Carlos Andres Rodriguez Maturana": "Otro", "Vinculaciones Jorge Mauricio Salgado Viloria": "Otro", "Vinculaciones Regionales": "Otro", "Vinculaciones Regionales MARIA FERNANDA JIMENEZ": "Otro", "Vinculaciones Regionales Yenny Sanchez": "Otro", "Vinculaciones Sindy Alexandra Gomez Herrera": "Otro", "Vinculaciones Maria Paula Prada Coronado": "Otro", "Vinculaciones Regionales Diana Cortes": "Otro", "Vinculaciones Regionales Johana Gomez": "Otro", "Vinculaciones Regionales Erick Carmona": "Otro", "Vinculaciones Carlos Alberto Arias Higuera": "Otro", "Vinculaciones  JAMBERT ANDRÉS MALOTT MUÑOZ": "Otro", "Vinculaciones ANA MILENA PATIÑO PINEDA": "Otro", "Vinculaciones Luz Dary Ramirez Quevedo": "Otro", "Vinculaciones Regionales Angie Borja": "Otro", "Vinculaciones Regionales Jose Viatela": "Otro", "Vinculaciones CAROLINA ROCIO VILORIA PEDREROS": "Otro", "Vinculaciones Regionales Eduardo Paez": "Otro", "Vinculaciones Regionales Katrin Steffi Garzón Avila": "Otro", "Vinculaciones Regionales yuri_pulido@cun.edu.co": "Otro", "Vinculaciones Jeisson Fernando Martinez Sissa": "Otro", "Vinculaciones Jeffer Sebastian Vargas Jimenez": "Otro", "Vinculaciones Sebastian Rodriguez Hernandez": "Otro", "Vinculaciones Diego Fernando Cruz Briceño": "Otro", "Vinculaciones Regionales Laura Montaña": "Otro", "Vinculaciones Andres Beltran": "Otro", "Vinculaciones Regionales Paula Sanchez": "Otro", "Vinculaciones Regionales Angie Moreno": "Otro", "Vinculaciones Sebastian Orlando Cabrera Baez": "Otro", "Vinculaciones Regionales Julian Guaido": "Otro", "Vinculaciones Marianne Ginette Morales Pupo": "Otro", "Vinculaciones Regionales Karen Hernandez": "Otro", "Vinculaciones Regionales Carolina Rosero": "Otro", "Vinculaciones Regionales Cesar Parra": "Otro", "Vinculaciones Regionales Camilo Gualteros": "Otro", "Vinculaciones dairon felipe tellez zambrano": "Otro", "Vinculaciones Maria Fernanda Urbano Cortes": "Otro", "Vinculaciones Alexander Enrique Penaranda Polo": "Otro", "Vinculaciones Regionales Estefany Armijos": "Otro", "Vinculaciones Erick Mauricio Carmona Vergara": "Otro", "Vinculacion Regional Vega Medina": "Otro", "Vinculaciones Manuel Jose Garcia Pabon": "Otro", "Vinculaciones KATHERINE LILIANA GONZALEZ RUBIO ANDRADE": "Otro", "Vinculaciones Fabio Andres Miranda Sequea": "Otro", "Vinculaciones NATALIA REMOLINA HERRERA": "Otro", "Vinculaciones Santa Marta Karen Neira": "Otro", "Vinculaciones Oscar Ivan Torres Gomez": "Otro", "Vinculaciones Juan Camilo Pulido Fagua": "Otro", "Vinculaciones Lina Marcela Paez Cantero": "Otro", "Vinculaciones Victor Hugo Tobon Pérez": "Otro", "Vinculaciones KEVIN LARA ORTEGA": "Otro", "Vinculaciones Regionales Maria Baquero": "Otro", "Proceso Telecampus": "Otro", "TELECAMPUS KENNEDY": "Otro", "TELECAMPUS SUBA": "Otro", "TELECAMPUS MOSQUERA": "Otro", "TELECAMPUS MEDELLIN": "Otro", "TELECAMPUS ENGATIVA": "Otro", "TELECAMPUS TOLIMA": "Otro", "TELECAMPUS BOSA": "Otro", "TELECAMPUS IBAGUE": "Otro", "TELECAMPUS NEIVA": "Otro", "TELECAMPUS BIZEN": "Otro", "Telecampus Chia": "Otro", "TELECAMPUS POLIMODAL": "Otro", "TELECAMPUS SANTA MARTA": "Otro", "TELECAMPUS SINCELEJO": "Otro", "TELECAMPUS TUNJA": "Otro", "Telecampus Telecampus Bucaramanga": "Otro", "TELECAMPUS SOTAVENTO": "Otro", "Telecampus armenia@telecampus.co": "Otro", "TELECAMPUS ELITE": "Otro", "TELECAMPUS CAVES": "Otro", "Telecampus Funza (Facatativá Y Cota). Mauricio Alvarez Diaz": "Otro", "TELECAMPUS LA UNIÓN": "Otro", "TELECAMPUS GAVES": "Otro", "Telecampus Apartadó": "Otro", "pasto@telecampus.co pasto@telecampus.co": "Otro", "Ejecutiva de Cuenta": "Otro", "Proceso Camilo Rodriguez": "Otro", "Bachillersitario": "Otro", "BE": "Otro", "Proceso Regional": "Otro", "jenifer_rodriguez@cun.edu.co": "Retiro", "zharick_olmos@cun.edu.co": "Retiro", "allison_campos@cun.edu.co": "Alejandra", "Retirado": "Retiro", "hasel_moncada@cun.edu.co": "Retiro", "cristian_garciab@cun.edu.co": "Retiro", "jairo_amaya@cun.edu.co": "Retiro", "alexandra_gonzalez@cun.edu.co": "Lorena", "Angie_barretop@cun.edu.co": "Posgrado", "Aura_hipolito@cun.edu.co": "Alejandra", "brayan_rodriguez@cun.edu.co": "Retiro", "bryan_ortiz@cun.edu.co": "Alejandra", "cindy_jerezr@cun.edu.co": "Alejandra", "cristian_pinilla@cun.edu.co": "Alejandra", "cristian_uribe@cun.edu.co": "Retiro", "CUN Digital": "Reasignar", "daniel_pinzonp@cun.edu.co": "Alejandra", "dayanna_gonzalez@cun.edu.co": "Retiro", "diana_proanof@cun.edu.co": "Lili", "evelin_mahecha@cun.edu.co": "Lili", "jameson_rodriguez@cun.edu.co": "Lorena", "Jannis_castellanos@cun.edu.co": "Alejandra", "javier_ramirez@cun.edu.co": "Lili", "jennifer_avila@cun.edu.co": "Lorena", "juan_cardonav@cun.edu.co": "Retiro", "lady_aguirre@cun.edu.co": "Retiro", "lilia_vasquez@cun.edu.co": "Lili", "luisa_velandia@cun.edu.co": "Posgrado", "luz_carrenor@cun.edu.co": "Lili", "marjorie_villadiego@cun.edu.co": "Retiro", "nancy_ome@cun.edu.co": "Posgrado", "nohora_ruiz@cun.edu.co": "Alejandra", "Para Reasignar": "Reasignar", "ricardo_alvarezro@cun.edu.co": "Lili", "sandra_gelvez@cun.edu.co": "Lorena", "santiago_cajamarca@cun.edu.co": "Retiro", "steven_aldana@cun.edu.co": "Retiro", "viviana_calderon@cun.edu.co": "Lili", "karol_ibagon@cun.edu.co": "Retiro", "juan_rodriguezroa@cun.edu.co": "Lorena", "lesslly_acero@cun.edu.co": "Retiro", "sergio_beltran@cun.edu.co": "Retiro", "sindy_rios@cun.edu.co": "Retiro", "katherine_herrera@cun.edu.co": "Retiro", "angela_galeano@cun.edu.co": "Retiro", "alejandra_merchan@cun.edu.co": "Retiro", "deimer_meneses@cun.edu.co": "Retiro", "johanna_barreto@cun.edu.co": "Retiro", "Laura_murillo@cun.edu.co": "Retiro", "karen_penagos@cun.edu.co": "Alejandra", "katherine_latorre@cun.edu.co": "Retiro", "yoharlis_gomez@cun.edu.co": "Retiro", "claudia_ballen@cun.edu.co": "Lili", "laura_ramirezp@cun.edu.co": "Retiro", "elizabeth_villarreal@cun.edu.co": "Lili", "pablo_buitrago@cun.edu.co": "Retiro", "johan_ramirez@cun.edu.co": "Retiro", "laura_ruizba@cun.edu.co": "Retiro", "edison_torreses@cun.edu.co": "Lili", "santiago_buitrago@cun.edu.co": "Lili", "jeison_gamba@cun.edu.co": "Lorena", "elifeth_romero@cun.edu.co": "Lili", "Wilson_duartet@cun.edu.co": "Retiro", "jhans_manzanares@cun.edu.co": "Lili", "cristhian_delgado@cun.edu.co": "Retiro", "elena_rojas@cun.edu.co": "Alejandra", "Jhon_rodriguez@cun.edu.co": "Retiro", "nicolas_fandino@cun.edu.co": "Retiro", "ashley_narino@cun.edu.co": "Retiro", "nicolas_rodriguezd@cun.edu.co": "Retiro", "ana_rayo@cun.edu.co": "Retiro", "kevin_mateus@cun.edu.co": "Retiro", "juan_bautistasa@cun.edu.co": "Retiro", "leonel_sepulvedasa@cun.edu.co": "Retiro", "jose_ospinobe@cun.edu.co": "Retiro", "kimberly_torove@cun.edu.co": "Retiro", "juan_mirandabe@cun.edu.co": "Retiro", "sebastian_usaquenq@cun.edu.co": "Retiro", "luisa_contreraso@cun.edu.co": "Retiro", "andry_loaizac@cun.edu.co": "Retiro", "yeisson_sierrag@cun.edu.co": "Retiro", "pedro_mendozap@cun.edu.co": "Retiro", "natalia_nietoa@cun.edu.co": "Retiro", "nicole_hidalgo@cun.edu.co": "Retiro", "juan_avilap@cun.edu.co": "Lorena", "nicol_ramosv@cun.edu.co": "Alejandra", "dayana_rodriguezh@cun.edu.co": "Lorena", "edwar_pena@cun.edu.co": "Retiro", "gloria_burgos@cun.edu.co": "Retiro", "wilmar_loaiza@cun.edu.co": "Retiro", "edson_bosiga@cun.edu.co": "Retiro", "jaider_gonzalez@cun.edu.co": "Retiro", "jhonatan_reyes@cun.edu.co": "Retiro", "ricardo_arevalo@cun.edu.co": "reasignar", "Comunicaciones Mercadeo": "Retiro", "nestor_obandom@cun.edu.co": "Retiro", "yeimi_bustosb@cun.edu.co": "Lili", "nestor_conrado@cun.edu.co": "Alejandra", "bertha_ramos@cun.edu.co": "Retiro", "jasmin_trianag@cun.edu.co": "Retiro", "nicoll_malpica@cun.edu.co": "Retiro", "cristian_gonzalezm@cun.edu.co": "Retiro", "Cristian_rodriguezr@cun.edu.co": "Retiro", "nicol_moyano@cun.edu.co": "Retiro", "jose_daza@cun.edu.co": "Retiro", "einer_carreno@cun.edu.co": "Retiro", "brayan_cespedesp@cun.edu.co": "Lorena", "caterine_chaparro@cun.edu.co": "Retiro", "anny_vargas@cun.edu.co": "Posgrado", "miguel_fuentes@cun.edu.co": "Retiro", "andrea_aroca@cun.edu.co": "Posgrado", "danna_molano@cun.edu.co": "Posgrado", "jose_jimenezro@cun.edu.co": "Retiro", "valentina_saavedras@cun.edu.co": "Retiro", "camilo_collazos@cun.edu.co": "Retiro", "nicolas_beltranp@cun.edu.co": "Retiro", "nestor_huertas@cun.edu.co": "Retiro", "Juan_diazma@cun.edu.co": "Retiro", "juan_avilag@cun.edu.co": "Retiro", "asdrubal_sotop@cun.edu.co": "Retiro", "stefania_castanedae@cun.edu.co": "Retiro", "cristian_herreno@cun.edu.co": "Retiro", "Vinculaciones Catherine Andrea Cano Llanos": "Otro", "Vinculaciones Lina Paola Hernandez Palacios": "Otro", "Vinculaciones Mary Andrea Garcia Llerena": "Otro", "Vinculaciones Jessika Paola Moreno Guerrero": "Otro", "Vinculaciones Diana Camila Novoa Rodriguez": "Otro", "Vinculaciones Laura Juliana Orduz Landinez": "Otro", "Vinculaciones Juan Sebastian Valenciano Martinez": "Otro", "Vinculaciones Juan Carlos Herrera Cañate": "Otro", "Vinculaciones Karen Yulieth Potosí Silva": "Otro", "Vinculaciones Edwin Julian Valdes Mendes": "Otro", "michaell_murillo@cun.edu.co": "Retiro", "franklim_sanabria@cun.edu.co": "Retiro", "carlos_betancur@cun.edu.co": "Retiro", "edith_rangel@cun.edu.co": "Retiro", "leydy_alba@cun.edu.co": "Retiro", "laura_buitragot@cun.edu.co": "Retiro", "gonzalo_montes@cun.edu.co": "Retiro", "lisseth_montilla@cun.edu.co": "Alejandra", "eric_benavides@cun.edu.co": "Retiro", "alexander_hortua@cun.edu.co": "Retiro", "ginna_pulido@cun.edu.co": "Retiro", "hector_quiroga@cun.edu.co": "Lorena", "fredy_galeano@cun.edu.co": "Retiro", "yuri_torres@cun.edu.co": "Lorena", "carlos_pinzonp@cun.edu.co": "Retiro", "sara_fuentesv@cun.edu.co": "Lili", "nelly_ramosg@cun.edu.co": "Retiro", "mairyn_rueda@cun.edu.co": "Retiro", "pablo_ramirezm@cun.edu.co": "Retiro", "juan_cuellarb@cun.edu.co": "Retiro", "luis_plazasl@cun.edu.co": "Retiro", "sandra_tangarife@cun.edu.co": "Retiro", "lizbeth_dominguez@cun.edu.co": "Retiro", "luis_ricaurtem@cun.edu.co": "Retiro", "feider_granados@cun.edu.co": "Lili", "julian_ortizb@cun.edu.co": "Retiro", "yeferson_bolivars@cun.edu.co": "Retiro", "julieth_sabogal@cun.edu.co": "Retiro", "anderxon_rodriguez@cun.edu.co": "Retiro", "juan_suarezr@cun.edu.co": "Lorena", "nicole_hernandeza@cun.edu.co": "Retiro", "juan_pinzond@cun.edu.co": "Retiro", "angie_leal@cun.edu.co": "Retiro", "juan_leonal@cun.edu.co": "Retiro", "juan_mendozaza@cun.edu.co": "Retiro", "ingrid_becerra@cun.edu.co": "Retiro", "erik_riano@cun.edu.co": "Alejandra", "francisco_martinez@cun.edu.co": "Retiro", "adriana_farfanro@cun.edu.co": "Retiro", "jirineth_castrova@cun.edu.co": "Retiro", "julian_rojasro@cun.edu.co": "Retiro", "yader_fernandez@cun.edu.co": "Retiro", "brayan_sapuy@cun.edu.co": "Retiro", "sahian_camino@cun.edu.co": "Retiro", "karol_castaneda@cun.edu.co": "Retiro", "yeimi_rivera@cun.edu.co": "Retiro", "luis_hernandezc@cun.edu.co": "Retiro", "juan_ardillaos@cun.edu.co": "Retiro", "marivel_rodriguez@cun.edu.co": "Retiro", "sara_vargas@cun.edu.co": "Retiro", "jesus_rincon@cun.edu.co": "Retiro", "saray_barrera@cun.edu.co": "Retiro", "jhon_avilaq@cun.edu.co": "Retiro", "kevin_parada@cun.edu.co": "Lili", "edwin_torres@cun.edu.co": "Retiro", "ivan_burgosc@cun.edu.co": "Retiro", "isis_martinezta@cun.edu.co": "Retiro", "brian_solanoes@cun.edu.co": "Retiro", "angelo_bernalre@cun.edu.co": "Retiro", "karen_quijanolo@cun.edu.co": "Alejandra", "maria_samacamo@cun.edu.co": "Retiro", "johan_morenome@cun.edu.co": "Retiro", "juan_lopezbe@cun.edu.co": "Retiro", "kevin_diaz@cun.edu.co": "Retiro", "orlenis_sanchez@cun.edu.co": "Retiro", "dayana_galvez@cun.edu.co": "Lorena", "nelly_fletcher@cun.edu.co": "Lorena", "anderson_rubiano@cun.edu.co": "Posgrado", "leidy_moralesc@cun.edu.co": "Retiro", "brandon_mora@cun.edu.co": "Alejandra", "carmen_porras@cun.edu.co": "Retiro", "brandon_lara@cun.edu.co": "Alejandra", "evelyt_marin@cun.edu.co": "Retiro", "rafael_patino@cun.edu.co": "Retiro", "francisco_torresp@cun.edu.co": "Retiro", "sebastian_patino@cun.edu.co": "Lili", "joyce_marriaga@cun.edu.co": "Retiro", "ana_gordillo@cun.edu.co": "Retiro", "sulay_chirivi@cun.edu.co": "Lorena", "david_gallegos@cun.edu.co": "Retiro", "juan_florezro@cun.edu.co": "Lili", "jonathan_gonzalezba@cun.edu.co": "Lorena", "carolina_bernal@cun.edu.co": "Retiro", "jenny_munoz@cun.edu.co": "Retiro", "maria_berrio@cun.edu.co": "Retiro", "albert_cantor@cun.edu.co": "Retiro", "pablo_penuelame@cun.edu.co": "Retiro", "yisel_ariasri@cun.edu.co": "Retiro", "cesar_avendanoya@cun.edu.co": "Retiro", "monica_foreroar@cun.edu.co": "Retiro", "laurenth_sotomo@cun.edu.co": "Retiro", "gary_jimenezpa@cun.edu.co": "Lorena", "gabriela_ordonezro@cun.edu.co": "Posgrado", "nicolas_betancourt@cun.edu.co": "Posgrado", "juan_candamilso@cun.edu.co": "Posgrado", "michael_mesahe@cun.edu.co": "Posgrado", "juan_delvalle@cun.edu.co": "Posgrado", "angel_avilasa@cun.edu.co": "Posgrado", "kimberly_castiblanco@cun.edu.co": "Posgrado", "karen_gongora@cun.edu.co": "Posgrado", "jonathan_alarcon@cun.edu.co": "Lili", "lenis_carvajal@cun.edu.co": "Presencial", "angela_caipa@cun.edu.co": "Retiro", "cesar_gomezd@cun.edu.co": "Retiro", "javier_noguera@cun.edu.co": "Retiro", "laura_gonzalezme@cun.edu.co": "Retiro", "jhonatan_beltran@cun.edu.co": "Presencial", "diana_beltrante@cun.edu.co": "Retiro", "santiago_ordonez@cun.edu.co": "Lili", "lina_osorio@cun.edu.co": "Retiro", "erika_guerrero@cun.edu.co": "Retiro", "henry_sanchez@cun.edu.co": "Lorena", "angie_bernal@cun.edu.co": "Retiro", "eilen_hernandez@cun.edu.co": "Retiro", " andres_lopezh@cun.edu.co": "Retiro", "manuel_garcia@cun.edu.co": "Lili", "paula_casasb@cun.edu.co": "Julio", "juan_ortizl@cun.edu.co": "Retiro", "kevin_montenegro@cun.edu.co": "Lorena", "freddy_guerrero@cun.edu.co": "Retiro", "carlos_vargasc@cun.edu.co": "Retiro", "karen_arevalo@cun.edu.co": "Retiro", "william_suarez@cun.edu.co": "Retiro", "nicolas_vidal@cun.edu.co": "Retiro", "shirley_villarreal@cun.edu.co": "Posgrado", "maryuri_ramirez@cun.edu.co": "Retiro", "brandon_cerinza@cun.edu.co": "Retiro", "juan_buitragov@cun.edu.co": "Posgrado", "valentina_torres@cun.edu.co": "Alejandra", "fabian_jimenezc@cun.edu.co": "Posgrado", "juan_rigueros@cun.edu.co": "Retiro", "eyleen_cuspoca@cun.edu.co": "Retiro", "yenifer_huerfano@cun.edu.co": "Retiro", "paula_perez@cun.edu.co": "Presencial", "julieth_trujillo@cun.edu.co": "Retiro", "rosmary_molina@cun.edu.co": "Presencial", "laura_amayar@cun.edu.co": "Presencial", "tania_cuellar@cun.edu.co": "Presencial", "yecid_orjuela@cun.edu.co": "Presencial", "jeimmy_bernal@cun.edu.co": "Posgrado", "laura_martinezv@cun.edu.co": "Presencial", "jenny_mosquera@cun.edu.co": "Retiro", "karol_ortiz@cun.edu.co": "Presencial", "cesar_rodriguezba@cun.edu.co": "Presencial", "nicol_anzola@cun.edu.co": "Retiro", "leidy_montanas@cun.edu.co": "Retiro", "maria_contrerasc@cun.edu.co": "Retiro", "juan_mancerac@cun.edu.co": "Presencial", "ruben_cruz@cun.edu.co": "Presencial", "fernanda_garcia@cun.edu.co": "Alejandra", "jostin_gomez@cun.edu.co": "Alejandra", "jonnathan_poveda@cun.edu.co": "Alejandra", "joan_perez@cun.edu.co": "Alejandra", "laura_zamudio@cun.edu.co": "Retiro", "jose_galeano@cun.edu.co": "Alejandra", "marlon_martinez@cun.edu.co": "Retiro", "cristian_sanchez@cun.edu.co": "Alejandra", "derly_arguellez@cun.edu.co": "Retiro", "brandon_hernandezt@cun.edu.co": "Retiro", "karen_quintero@cun.edu.co": "Presencial", "yeymy_ninog@cun.edu.co": "Retiro", "sarays_bolivarm@cun.edu.co": "Retiro", "karen_giraldoz@cun.edu.co": "Retiro", "jean_mantillac@cun.edu.co": "Presencial", "erik_monterof@cun.edu.co": "Presencial", "maria_castilloc@cun.edu.co": "Retiro", "michael_fernandezr@cun.edu.co": "Retiro", "jhon_hurtador@cun.edu.co": "Retiro", "santiago_povedav@cun.edu.co": "Retiro", "ingrid_hurtadob@cun.edu.co": "Retiro", "geraldine_pinzons@cun.edu.co": "Retiro", "heyner_balaguera@cun.edu.co": "Retiro", "daniela_velasquezv@cun.edu.co": "Retiro", "jhon_vanegasm@cun.edu.co": "Presencial", "carlos_chiappe@cun.edu.co": "Retiro", "leidy_ortiza@cun.edu.co": "Retiro", "angie_salazarr@cun.edu.co": "Retiro", "jean_aguilart@cun.edu.co": "Posgrado", "santiago_gomezl@cun.edu.co": "Posgrado", "andres_bolivarn@cun.edu.co": "Retiro", "shirly_garciag@cun.edu.co": "Retiro", "maicol_vivasc@cun.edu.co": "Retiro", "juan_duranr@cun.edu.co": "Posgrado", "lisbet_rojash@cun.edu.co": "Retiro", "diego_rodriguezb@cun.edu.co": "Posgrado", "angi_hernandezt@cun.edu.co": "Retiro", "paula_ortizm@cun.edu.co": "Lili", "miguel_garciaf@cun.edu.co": "Lili", "andrea_alvarezr@cun.edu.co": "Lorena", "jenny_vegag@cun.edu.co": "Lorena", "karen_torresa@cun.edu.co": "Lili", "ingrid_rativa@cun.edu.co": "Lorena", "braillin_fuenteso@cun.edu.co": "OJT Virtual", "karen_acevedoa@cun.edu.co": "OJT Virtual", "lizeth_benavidesr@cun.edu.co": "OJT Virtual", "kevin_ramirez@cun.edu.co": "OJT Virtual", "kevin_chavarria@cun.edu.co": "OJT Virtual", "jose_pinzon@cun.edu.co": "OJT Virtual", "rody_navarrete@cun.edu.co": "OJT Virtual", "anderson_pinilla@cun.edu.co": "OJT Virtual", "laura_coronadou@cun.edu.co": "OJT Virtual", "jhon_arenas@cun.edu.co": "OJT Virtual", "kevin_bautista@cun.edu.co": "OJT Virtual", "jose_benavides@cun.edu.co": "OJT Virtual"};

function normSupervisorEmbebido(v=''){
  return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
}
function emailSupervisorEmbebido(v=''){
  const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].toLowerCase() : '';
}
function supervisorEmbebido(asesor){
  const raw = String(asesor || '').trim();
  const email = emailSupervisorEmbebido(raw);
  if(email && SUPERVISOR_EMBEBIDO[email] !== undefined) return String(SUPERVISOR_EMBEBIDO[email] || 'Sin supervisor').trim() || 'Sin supervisor';

  const norm = normSupervisorEmbebido(raw);
  const key = Object.keys(SUPERVISOR_EMBEBIDO).find(k => normSupervisorEmbebido(k) === norm || emailSupervisorEmbebido(k) === email);
  if(key) return String(SUPERVISOR_EMBEBIDO[key] || 'Sin supervisor').trim() || 'Sin supervisor';

  return 'Sin supervisor';
}
function enriquecerAsignacionSupervisor(row, asesor){
  const sup = supervisorEmbebido(asesor);
  return {
    ...row,
    'CORREO ASESOR ASIGNADO': asesor,
    'SUPERVISOR': sup,
    'MENTOR': sup
  };
}
function detalleRowsSupervisor(data){
  const map = {};
  (data || []).forEach(r => {
    const asesor = r['CORREO ASESOR ASIGNADO'] || '';
    if(!asesor) return;
    const supervisor = r['SUPERVISOR'] || r['MENTOR'] || supervisorEmbebido(asesor);
    const key = supervisor + '||' + asesor;
    if(!map[key]) map[key] = {Supervisor: supervisor, Asesor: asesor, 'Cantidad leads': 0};
    map[key]['Cantidad leads']++;
  });
  return Object.values(map).sort((a,b)=>String(a.Supervisor).localeCompare(String(b.Supervisor)) || b['Cantidad leads']-a['Cantidad leads']);
}
function pintarDetalleSupervisor(data, boxId, cardId, fileName){
  const box = document.getElementById(boxId);
  const card = document.getElementById(cardId);
  if(!box) return;

  if(card){
    card.style.display = 'block';
    card.classList.add('detalle-visible');
    card.classList.remove('detalle-placeholder');
  }

  const rows = detalleRowsSupervisor(data);
  window.__detalleSupervisorExport = window.__detalleSupervisorExport || {};
  window.__detalleSupervisorExport[boxId] = {rows, fileName};

  if(!rows.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por supervisor y asesor.</div>';
    return;
  }

  const totalLeads = rows.reduce((a,b)=>a+b['Cantidad leads'],0);
  const totalAsesores = new Set(rows.map(r=>r.Asesor)).size;
  const totalSupervisores = new Set(rows.map(r=>r.Supervisor)).size;
  const grouped = {};
  rows.forEach(r=>{ if(!grouped[r.Supervisor]) grouped[r.Supervisor]=[]; grouped[r.Supervisor].push(r); });

  box.innerHTML = `
    <div class="detalle-actions">
      <button class="btn btn-green" onclick="exportDetalleSupervisor('${boxId}')">⬇ Exportar detalle asesor</button>
    </div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalSupervisores.toLocaleString()}</strong><span>Supervisores</span></div>
    </div>
    ${Object.entries(grouped).map(([sup, asesores])=>{
      const totalSup = asesores.reduce((a,b)=>a+b['Cantidad leads'],0);
      return `
        <div class="detalle-mentor-card">
          <div class="detalle-mentor-head"><span>👤 ${escapeHtml(sup)}</span><b>${totalSup.toLocaleString()} leads</b></div>
          <div class="detalle-asesor-grid">
            ${asesores.map(r=>`
              <div class="detalle-asesor-item">
                <span><b>${escapeHtml(r.Asesor)}</b></span>
                <strong>${r['Cantidad leads'].toLocaleString()}</strong>
              </div>`).join('')}
          </div>
        </div>`;
    }).join('')}
  `;
}
function exportDetalleSupervisor(boxId){
  const data = window.__detalleSupervisorExport?.[boxId]?.rows || [];
  const name = window.__detalleSupervisorExport?.[boxId]?.fileName || 'Detalle_Asesor';
  if(!data.length){ showToast('No hay detalle asesor para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Detalle asesor');
  XLSX.writeFile(wb, name + '_' + Date.now() + '.xlsx');
}
function renderDetalleAsesorSG(){ pintarDetalleSupervisor(typeof distribuidoData !== 'undefined' ? distribuidoData : [], 'sg-detalle-asesor', 'sg-detalle-asesor-card', 'Detalle_Asesor_Sin_Gestion'); }
function renderDetalleAsesorInter(){ pintarDetalleSupervisor(typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [], 'inter-detalle-asesor', 'inter-detalle-asesor-card', 'Detalle_Asesor_Interesados'); }
function renderDetalleAsesorNC(){ pintarDetalleSupervisor(typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [], 'nc-detalle-asesor', 'nc-detalle-asesor-card', 'Detalle_Asesor_No_Contactado'); }


/* === DETALLE VISIBLE DENTRO DEL RESUMEN DE DISTRIBUCIÓN === */
const SUPERVISOR_VISIBLE = {"Vinculaciones Carlos Andres Burgos Carvajal": "Otro", "Vinculaciones LAURA CAMILA CASTILLO SEPULVEDA": "Otro", "Vinculaciones Erika Ramirez Rico": "Otro", "Vinculaciones Laura Alejandra Guzman Gutierrez": "Otro", "Vinculaciones Margarita Rosa Navarro Garzon": "Otro", "Vinculaciones Maria Alejandra Vergara Chavez": "Otro", "Vinculaciones Maria Claudia Salas Morales": "Otro", "Vinculaciones Regionales Diego Villota": "Otro", "Vinculaciones Carlos Andres Rodriguez Maturana": "Otro", "Vinculaciones Jorge Mauricio Salgado Viloria": "Otro", "Vinculaciones Regionales": "Otro", "Vinculaciones Regionales MARIA FERNANDA JIMENEZ": "Otro", "Vinculaciones Regionales Yenny Sanchez": "Otro", "Vinculaciones Sindy Alexandra Gomez Herrera": "Otro", "Vinculaciones Maria Paula Prada Coronado": "Otro", "Vinculaciones Regionales Diana Cortes": "Otro", "Vinculaciones Regionales Johana Gomez": "Otro", "Vinculaciones Regionales Erick Carmona": "Otro", "Vinculaciones Carlos Alberto Arias Higuera": "Otro", "Vinculaciones  JAMBERT ANDRÉS MALOTT MUÑOZ": "Otro", "Vinculaciones ANA MILENA PATIÑO PINEDA": "Otro", "Vinculaciones Luz Dary Ramirez Quevedo": "Otro", "Vinculaciones Regionales Angie Borja": "Otro", "Vinculaciones Regionales Jose Viatela": "Otro", "Vinculaciones CAROLINA ROCIO VILORIA PEDREROS": "Otro", "Vinculaciones Regionales Eduardo Paez": "Otro", "Vinculaciones Regionales Katrin Steffi Garzón Avila": "Otro", "Vinculaciones Regionales yuri_pulido@cun.edu.co": "Otro", "Vinculaciones Jeisson Fernando Martinez Sissa": "Otro", "Vinculaciones Jeffer Sebastian Vargas Jimenez": "Otro", "Vinculaciones Sebastian Rodriguez Hernandez": "Otro", "Vinculaciones Diego Fernando Cruz Briceño": "Otro", "Vinculaciones Regionales Laura Montaña": "Otro", "Vinculaciones Andres Beltran": "Otro", "Vinculaciones Regionales Paula Sanchez": "Otro", "Vinculaciones Regionales Angie Moreno": "Otro", "Vinculaciones Sebastian Orlando Cabrera Baez": "Otro", "Vinculaciones Regionales Julian Guaido": "Otro", "Vinculaciones Marianne Ginette Morales Pupo": "Otro", "Vinculaciones Regionales Karen Hernandez": "Otro", "Vinculaciones Regionales Carolina Rosero": "Otro", "Vinculaciones Regionales Cesar Parra": "Otro", "Vinculaciones Regionales Camilo Gualteros": "Otro", "Vinculaciones dairon felipe tellez zambrano": "Otro", "Vinculaciones Maria Fernanda Urbano Cortes": "Otro", "Vinculaciones Alexander Enrique Penaranda Polo": "Otro", "Vinculaciones Regionales Estefany Armijos": "Otro", "Vinculaciones Erick Mauricio Carmona Vergara": "Otro", "Vinculacion Regional Vega Medina": "Otro", "Vinculaciones Manuel Jose Garcia Pabon": "Otro", "Vinculaciones KATHERINE LILIANA GONZALEZ RUBIO ANDRADE": "Otro", "Vinculaciones Fabio Andres Miranda Sequea": "Otro", "Vinculaciones NATALIA REMOLINA HERRERA": "Otro", "Vinculaciones Santa Marta Karen Neira": "Otro", "Vinculaciones Oscar Ivan Torres Gomez": "Otro", "Vinculaciones Juan Camilo Pulido Fagua": "Otro", "Vinculaciones Lina Marcela Paez Cantero": "Otro", "Vinculaciones Victor Hugo Tobon Pérez": "Otro", "Vinculaciones KEVIN LARA ORTEGA": "Otro", "Vinculaciones Regionales Maria Baquero": "Otro", "Proceso Telecampus": "Otro", "TELECAMPUS KENNEDY": "Otro", "TELECAMPUS SUBA": "Otro", "TELECAMPUS MOSQUERA": "Otro", "TELECAMPUS MEDELLIN": "Otro", "TELECAMPUS ENGATIVA": "Otro", "TELECAMPUS TOLIMA": "Otro", "TELECAMPUS BOSA": "Otro", "TELECAMPUS IBAGUE": "Otro", "TELECAMPUS NEIVA": "Otro", "TELECAMPUS BIZEN": "Otro", "Telecampus Chia": "Otro", "TELECAMPUS POLIMODAL": "Otro", "TELECAMPUS SANTA MARTA": "Otro", "TELECAMPUS SINCELEJO": "Otro", "TELECAMPUS TUNJA": "Otro", "Telecampus Telecampus Bucaramanga": "Otro", "TELECAMPUS SOTAVENTO": "Otro", "Telecampus armenia@telecampus.co": "Otro", "TELECAMPUS ELITE": "Otro", "TELECAMPUS CAVES": "Otro", "Telecampus Funza (Facatativá Y Cota). Mauricio Alvarez Diaz": "Otro", "TELECAMPUS LA UNIÓN": "Otro", "TELECAMPUS GAVES": "Otro", "Telecampus Apartadó": "Otro", "pasto@telecampus.co pasto@telecampus.co": "Otro", "Ejecutiva de Cuenta": "Otro", "Proceso Camilo Rodriguez": "Otro", "Bachillersitario": "Otro", "BE": "Otro", "Proceso Regional": "Otro", "jenifer_rodriguez@cun.edu.co": "Retiro", "zharick_olmos@cun.edu.co": "Retiro", "allison_campos@cun.edu.co": "Alejandra", "Retirado": "Retiro", "hasel_moncada@cun.edu.co": "Retiro", "cristian_garciab@cun.edu.co": "Retiro", "jairo_amaya@cun.edu.co": "Retiro", "alexandra_gonzalez@cun.edu.co": "Lorena", "Angie_barretop@cun.edu.co": "Posgrado", "Aura_hipolito@cun.edu.co": "Alejandra", "brayan_rodriguez@cun.edu.co": "Retiro", "bryan_ortiz@cun.edu.co": "Alejandra", "cindy_jerezr@cun.edu.co": "Alejandra", "cristian_pinilla@cun.edu.co": "Alejandra", "cristian_uribe@cun.edu.co": "Retiro", "CUN Digital": "Reasignar", "daniel_pinzonp@cun.edu.co": "Alejandra", "dayanna_gonzalez@cun.edu.co": "Retiro", "diana_proanof@cun.edu.co": "Lili", "evelin_mahecha@cun.edu.co": "Lili", "jameson_rodriguez@cun.edu.co": "Lorena", "Jannis_castellanos@cun.edu.co": "Alejandra", "javier_ramirez@cun.edu.co": "Lili", "jennifer_avila@cun.edu.co": "Lorena", "juan_cardonav@cun.edu.co": "Retiro", "lady_aguirre@cun.edu.co": "Retiro", "lilia_vasquez@cun.edu.co": "Lili", "luisa_velandia@cun.edu.co": "Posgrado", "luz_carrenor@cun.edu.co": "Lili", "marjorie_villadiego@cun.edu.co": "Retiro", "nancy_ome@cun.edu.co": "Posgrado", "nohora_ruiz@cun.edu.co": "Alejandra", "Para Reasignar": "Reasignar", "ricardo_alvarezro@cun.edu.co": "Lili", "sandra_gelvez@cun.edu.co": "Lorena", "santiago_cajamarca@cun.edu.co": "Retiro", "steven_aldana@cun.edu.co": "Retiro", "viviana_calderon@cun.edu.co": "Lili", "karol_ibagon@cun.edu.co": "Retiro", "juan_rodriguezroa@cun.edu.co": "Lorena", "lesslly_acero@cun.edu.co": "Retiro", "sergio_beltran@cun.edu.co": "Retiro", "sindy_rios@cun.edu.co": "Retiro", "katherine_herrera@cun.edu.co": "Retiro", "angela_galeano@cun.edu.co": "Retiro", "alejandra_merchan@cun.edu.co": "Retiro", "deimer_meneses@cun.edu.co": "Retiro", "johanna_barreto@cun.edu.co": "Retiro", "Laura_murillo@cun.edu.co": "Retiro", "karen_penagos@cun.edu.co": "Alejandra", "katherine_latorre@cun.edu.co": "Retiro", "yoharlis_gomez@cun.edu.co": "Retiro", "claudia_ballen@cun.edu.co": "Lili", "laura_ramirezp@cun.edu.co": "Retiro", "elizabeth_villarreal@cun.edu.co": "Lili", "pablo_buitrago@cun.edu.co": "Retiro", "johan_ramirez@cun.edu.co": "Retiro", "laura_ruizba@cun.edu.co": "Retiro", "edison_torreses@cun.edu.co": "Lili", "santiago_buitrago@cun.edu.co": "Lili", "jeison_gamba@cun.edu.co": "Lorena", "elifeth_romero@cun.edu.co": "Lili", "Wilson_duartet@cun.edu.co": "Retiro", "jhans_manzanares@cun.edu.co": "Lili", "cristhian_delgado@cun.edu.co": "Retiro", "elena_rojas@cun.edu.co": "Alejandra", "Jhon_rodriguez@cun.edu.co": "Retiro", "nicolas_fandino@cun.edu.co": "Retiro", "ashley_narino@cun.edu.co": "Retiro", "nicolas_rodriguezd@cun.edu.co": "Retiro", "ana_rayo@cun.edu.co": "Retiro", "kevin_mateus@cun.edu.co": "Retiro", "juan_bautistasa@cun.edu.co": "Retiro", "leonel_sepulvedasa@cun.edu.co": "Retiro", "jose_ospinobe@cun.edu.co": "Retiro", "kimberly_torove@cun.edu.co": "Retiro", "juan_mirandabe@cun.edu.co": "Retiro", "sebastian_usaquenq@cun.edu.co": "Retiro", "luisa_contreraso@cun.edu.co": "Retiro", "andry_loaizac@cun.edu.co": "Retiro", "yeisson_sierrag@cun.edu.co": "Retiro", "pedro_mendozap@cun.edu.co": "Retiro", "natalia_nietoa@cun.edu.co": "Retiro", "nicole_hidalgo@cun.edu.co": "Retiro", "juan_avilap@cun.edu.co": "Lorena", "nicol_ramosv@cun.edu.co": "Alejandra", "dayana_rodriguezh@cun.edu.co": "Lorena", "edwar_pena@cun.edu.co": "Retiro", "gloria_burgos@cun.edu.co": "Retiro", "wilmar_loaiza@cun.edu.co": "Retiro", "edson_bosiga@cun.edu.co": "Retiro", "jaider_gonzalez@cun.edu.co": "Retiro", "jhonatan_reyes@cun.edu.co": "Retiro", "ricardo_arevalo@cun.edu.co": "reasignar", "Comunicaciones Mercadeo": "Retiro", "nestor_obandom@cun.edu.co": "Retiro", "yeimi_bustosb@cun.edu.co": "Lili", "nestor_conrado@cun.edu.co": "Alejandra", "bertha_ramos@cun.edu.co": "Retiro", "jasmin_trianag@cun.edu.co": "Retiro", "nicoll_malpica@cun.edu.co": "Retiro", "cristian_gonzalezm@cun.edu.co": "Retiro", "Cristian_rodriguezr@cun.edu.co": "Retiro", "nicol_moyano@cun.edu.co": "Retiro", "jose_daza@cun.edu.co": "Retiro", "einer_carreno@cun.edu.co": "Retiro", "brayan_cespedesp@cun.edu.co": "Lorena", "caterine_chaparro@cun.edu.co": "Retiro", "anny_vargas@cun.edu.co": "Posgrado", "miguel_fuentes@cun.edu.co": "Retiro", "andrea_aroca@cun.edu.co": "Posgrado", "danna_molano@cun.edu.co": "Posgrado", "jose_jimenezro@cun.edu.co": "Retiro", "valentina_saavedras@cun.edu.co": "Retiro", "camilo_collazos@cun.edu.co": "Retiro", "nicolas_beltranp@cun.edu.co": "Retiro", "nestor_huertas@cun.edu.co": "Retiro", "Juan_diazma@cun.edu.co": "Retiro", "juan_avilag@cun.edu.co": "Retiro", "asdrubal_sotop@cun.edu.co": "Retiro", "stefania_castanedae@cun.edu.co": "Retiro", "cristian_herreno@cun.edu.co": "Retiro", "Vinculaciones Catherine Andrea Cano Llanos": "Otro", "Vinculaciones Lina Paola Hernandez Palacios": "Otro", "Vinculaciones Mary Andrea Garcia Llerena": "Otro", "Vinculaciones Jessika Paola Moreno Guerrero": "Otro", "Vinculaciones Diana Camila Novoa Rodriguez": "Otro", "Vinculaciones Laura Juliana Orduz Landinez": "Otro", "Vinculaciones Juan Sebastian Valenciano Martinez": "Otro", "Vinculaciones Juan Carlos Herrera Cañate": "Otro", "Vinculaciones Karen Yulieth Potosí Silva": "Otro", "Vinculaciones Edwin Julian Valdes Mendes": "Otro", "michaell_murillo@cun.edu.co": "Retiro", "franklim_sanabria@cun.edu.co": "Retiro", "carlos_betancur@cun.edu.co": "Retiro", "edith_rangel@cun.edu.co": "Retiro", "leydy_alba@cun.edu.co": "Retiro", "laura_buitragot@cun.edu.co": "Retiro", "gonzalo_montes@cun.edu.co": "Retiro", "lisseth_montilla@cun.edu.co": "Alejandra", "eric_benavides@cun.edu.co": "Retiro", "alexander_hortua@cun.edu.co": "Retiro", "ginna_pulido@cun.edu.co": "Retiro", "hector_quiroga@cun.edu.co": "Lorena", "fredy_galeano@cun.edu.co": "Retiro", "yuri_torres@cun.edu.co": "Lorena", "carlos_pinzonp@cun.edu.co": "Retiro", "sara_fuentesv@cun.edu.co": "Lili", "nelly_ramosg@cun.edu.co": "Retiro", "mairyn_rueda@cun.edu.co": "Retiro", "pablo_ramirezm@cun.edu.co": "Retiro", "juan_cuellarb@cun.edu.co": "Retiro", "luis_plazasl@cun.edu.co": "Retiro", "sandra_tangarife@cun.edu.co": "Retiro", "lizbeth_dominguez@cun.edu.co": "Retiro", "luis_ricaurtem@cun.edu.co": "Retiro", "feider_granados@cun.edu.co": "Lili", "julian_ortizb@cun.edu.co": "Retiro", "yeferson_bolivars@cun.edu.co": "Retiro", "julieth_sabogal@cun.edu.co": "Retiro", "anderxon_rodriguez@cun.edu.co": "Retiro", "juan_suarezr@cun.edu.co": "Lorena", "nicole_hernandeza@cun.edu.co": "Retiro", "juan_pinzond@cun.edu.co": "Retiro", "angie_leal@cun.edu.co": "Retiro", "juan_leonal@cun.edu.co": "Retiro", "juan_mendozaza@cun.edu.co": "Retiro", "ingrid_becerra@cun.edu.co": "Retiro", "erik_riano@cun.edu.co": "Alejandra", "francisco_martinez@cun.edu.co": "Retiro", "adriana_farfanro@cun.edu.co": "Retiro", "jirineth_castrova@cun.edu.co": "Retiro", "julian_rojasro@cun.edu.co": "Retiro", "yader_fernandez@cun.edu.co": "Retiro", "brayan_sapuy@cun.edu.co": "Retiro", "sahian_camino@cun.edu.co": "Retiro", "karol_castaneda@cun.edu.co": "Retiro", "yeimi_rivera@cun.edu.co": "Retiro", "luis_hernandezc@cun.edu.co": "Retiro", "juan_ardillaos@cun.edu.co": "Retiro", "marivel_rodriguez@cun.edu.co": "Retiro", "sara_vargas@cun.edu.co": "Retiro", "jesus_rincon@cun.edu.co": "Retiro", "saray_barrera@cun.edu.co": "Retiro", "jhon_avilaq@cun.edu.co": "Retiro", "kevin_parada@cun.edu.co": "Lili", "edwin_torres@cun.edu.co": "Retiro", "ivan_burgosc@cun.edu.co": "Retiro", "isis_martinezta@cun.edu.co": "Retiro", "brian_solanoes@cun.edu.co": "Retiro", "angelo_bernalre@cun.edu.co": "Retiro", "karen_quijanolo@cun.edu.co": "Alejandra", "maria_samacamo@cun.edu.co": "Retiro", "johan_morenome@cun.edu.co": "Retiro", "juan_lopezbe@cun.edu.co": "Retiro", "kevin_diaz@cun.edu.co": "Retiro", "orlenis_sanchez@cun.edu.co": "Retiro", "dayana_galvez@cun.edu.co": "Lorena", "nelly_fletcher@cun.edu.co": "Lorena", "anderson_rubiano@cun.edu.co": "Posgrado", "leidy_moralesc@cun.edu.co": "Retiro", "brandon_mora@cun.edu.co": "Alejandra", "carmen_porras@cun.edu.co": "Retiro", "brandon_lara@cun.edu.co": "Alejandra", "evelyt_marin@cun.edu.co": "Retiro", "rafael_patino@cun.edu.co": "Retiro", "francisco_torresp@cun.edu.co": "Retiro", "sebastian_patino@cun.edu.co": "Lili", "joyce_marriaga@cun.edu.co": "Retiro", "ana_gordillo@cun.edu.co": "Retiro", "sulay_chirivi@cun.edu.co": "Lorena", "david_gallegos@cun.edu.co": "Retiro", "juan_florezro@cun.edu.co": "Lili", "jonathan_gonzalezba@cun.edu.co": "Lorena", "carolina_bernal@cun.edu.co": "Retiro", "jenny_munoz@cun.edu.co": "Retiro", "maria_berrio@cun.edu.co": "Retiro", "albert_cantor@cun.edu.co": "Retiro", "pablo_penuelame@cun.edu.co": "Retiro", "yisel_ariasri@cun.edu.co": "Retiro", "cesar_avendanoya@cun.edu.co": "Retiro", "monica_foreroar@cun.edu.co": "Retiro", "laurenth_sotomo@cun.edu.co": "Retiro", "gary_jimenezpa@cun.edu.co": "Lorena", "gabriela_ordonezro@cun.edu.co": "Posgrado", "nicolas_betancourt@cun.edu.co": "Posgrado", "juan_candamilso@cun.edu.co": "Posgrado", "michael_mesahe@cun.edu.co": "Posgrado", "juan_delvalle@cun.edu.co": "Posgrado", "angel_avilasa@cun.edu.co": "Posgrado", "kimberly_castiblanco@cun.edu.co": "Posgrado", "karen_gongora@cun.edu.co": "Posgrado", "jonathan_alarcon@cun.edu.co": "Lili", "lenis_carvajal@cun.edu.co": "Presencial", "angela_caipa@cun.edu.co": "Retiro", "cesar_gomezd@cun.edu.co": "Retiro", "javier_noguera@cun.edu.co": "Retiro", "laura_gonzalezme@cun.edu.co": "Retiro", "jhonatan_beltran@cun.edu.co": "Presencial", "diana_beltrante@cun.edu.co": "Retiro", "santiago_ordonez@cun.edu.co": "Lili", "lina_osorio@cun.edu.co": "Retiro", "erika_guerrero@cun.edu.co": "Retiro", "henry_sanchez@cun.edu.co": "Lorena", "angie_bernal@cun.edu.co": "Retiro", "eilen_hernandez@cun.edu.co": "Retiro", " andres_lopezh@cun.edu.co": "Retiro", "manuel_garcia@cun.edu.co": "Lili", "paula_casasb@cun.edu.co": "Julio", "juan_ortizl@cun.edu.co": "Retiro", "kevin_montenegro@cun.edu.co": "Lorena", "freddy_guerrero@cun.edu.co": "Retiro", "carlos_vargasc@cun.edu.co": "Retiro", "karen_arevalo@cun.edu.co": "Retiro", "william_suarez@cun.edu.co": "Retiro", "nicolas_vidal@cun.edu.co": "Retiro", "shirley_villarreal@cun.edu.co": "Posgrado", "maryuri_ramirez@cun.edu.co": "Retiro", "brandon_cerinza@cun.edu.co": "Retiro", "juan_buitragov@cun.edu.co": "Posgrado", "valentina_torres@cun.edu.co": "Alejandra", "fabian_jimenezc@cun.edu.co": "Posgrado", "juan_rigueros@cun.edu.co": "Retiro", "eyleen_cuspoca@cun.edu.co": "Retiro", "yenifer_huerfano@cun.edu.co": "Retiro", "paula_perez@cun.edu.co": "Presencial", "julieth_trujillo@cun.edu.co": "Retiro", "rosmary_molina@cun.edu.co": "Presencial", "laura_amayar@cun.edu.co": "Presencial", "tania_cuellar@cun.edu.co": "Presencial", "yecid_orjuela@cun.edu.co": "Presencial", "jeimmy_bernal@cun.edu.co": "Posgrado", "laura_martinezv@cun.edu.co": "Presencial", "jenny_mosquera@cun.edu.co": "Retiro", "karol_ortiz@cun.edu.co": "Presencial", "cesar_rodriguezba@cun.edu.co": "Presencial", "nicol_anzola@cun.edu.co": "Retiro", "leidy_montanas@cun.edu.co": "Retiro", "maria_contrerasc@cun.edu.co": "Retiro", "juan_mancerac@cun.edu.co": "Presencial", "ruben_cruz@cun.edu.co": "Presencial", "fernanda_garcia@cun.edu.co": "Alejandra", "jostin_gomez@cun.edu.co": "Alejandra", "jonnathan_poveda@cun.edu.co": "Alejandra", "joan_perez@cun.edu.co": "Alejandra", "laura_zamudio@cun.edu.co": "Retiro", "jose_galeano@cun.edu.co": "Alejandra", "marlon_martinez@cun.edu.co": "Retiro", "cristian_sanchez@cun.edu.co": "Alejandra", "derly_arguellez@cun.edu.co": "Retiro", "brandon_hernandezt@cun.edu.co": "Retiro", "karen_quintero@cun.edu.co": "Presencial", "yeymy_ninog@cun.edu.co": "Retiro", "sarays_bolivarm@cun.edu.co": "Retiro", "karen_giraldoz@cun.edu.co": "Retiro", "jean_mantillac@cun.edu.co": "Presencial", "erik_monterof@cun.edu.co": "Presencial", "maria_castilloc@cun.edu.co": "Retiro", "michael_fernandezr@cun.edu.co": "Retiro", "jhon_hurtador@cun.edu.co": "Retiro", "santiago_povedav@cun.edu.co": "Retiro", "ingrid_hurtadob@cun.edu.co": "Retiro", "geraldine_pinzons@cun.edu.co": "Retiro", "heyner_balaguera@cun.edu.co": "Retiro", "daniela_velasquezv@cun.edu.co": "Retiro", "jhon_vanegasm@cun.edu.co": "Presencial", "carlos_chiappe@cun.edu.co": "Retiro", "leidy_ortiza@cun.edu.co": "Retiro", "angie_salazarr@cun.edu.co": "Retiro", "jean_aguilart@cun.edu.co": "Posgrado", "santiago_gomezl@cun.edu.co": "Posgrado", "andres_bolivarn@cun.edu.co": "Retiro", "shirly_garciag@cun.edu.co": "Retiro", "maicol_vivasc@cun.edu.co": "Retiro", "juan_duranr@cun.edu.co": "Posgrado", "lisbet_rojash@cun.edu.co": "Retiro", "diego_rodriguezb@cun.edu.co": "Posgrado", "angi_hernandezt@cun.edu.co": "Retiro", "paula_ortizm@cun.edu.co": "Lili", "miguel_garciaf@cun.edu.co": "Lili", "andrea_alvarezr@cun.edu.co": "Lorena", "jenny_vegag@cun.edu.co": "Lorena", "karen_torresa@cun.edu.co": "Lili", "ingrid_rativa@cun.edu.co": "Lorena", "braillin_fuenteso@cun.edu.co": "OJT Virtual", "karen_acevedoa@cun.edu.co": "OJT Virtual", "lizeth_benavidesr@cun.edu.co": "OJT Virtual", "kevin_ramirez@cun.edu.co": "OJT Virtual", "kevin_chavarria@cun.edu.co": "OJT Virtual", "jose_pinzon@cun.edu.co": "OJT Virtual", "rody_navarrete@cun.edu.co": "OJT Virtual", "anderson_pinilla@cun.edu.co": "OJT Virtual", "laura_coronadou@cun.edu.co": "OJT Virtual", "jhon_arenas@cun.edu.co": "OJT Virtual", "kevin_bautista@cun.edu.co": "OJT Virtual", "jose_benavides@cun.edu.co": "OJT Virtual"};
function normSupervisorVisible(v=''){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();}
function emailSupervisorVisible(v=''){const m=String(v||'').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);return m?m[0].toLowerCase():'';}
function supervisorVisible(asesor){
  const raw=String(asesor||'').trim();
  const email=emailSupervisorVisible(raw);
  if(email && SUPERVISOR_VISIBLE[email]!==undefined) return String(SUPERVISOR_VISIBLE[email]||'Sin supervisor').trim()||'Sin supervisor';
  const norm=normSupervisorVisible(raw);
  const key=Object.keys(SUPERVISOR_VISIBLE).find(k=>normSupervisorVisible(k)===norm || emailSupervisorVisible(k)===email);
  return key ? String(SUPERVISOR_VISIBLE[key]||'Sin supervisor').trim()||'Sin supervisor' : 'Sin supervisor';
}
function detalleRowsVisible(data){
  const map={};
  (data||[]).forEach(r=>{
    const asesor=r['CORREO ASESOR ASIGNADO']||'';
    if(!asesor)return;
    const supervisor=r['SUPERVISOR']||r['MENTOR']||supervisorVisible(asesor);
    const key=supervisor+'||'+asesor;
    if(!map[key]) map[key]={Supervisor:supervisor,Asesor:asesor,'Cantidad leads':0};
    map[key]['Cantidad leads']++;
  });
  return Object.values(map).sort((a,b)=>String(a.Supervisor).localeCompare(String(b.Supervisor))||b['Cantidad leads']-a['Cantidad leads']);
}
function htmlDetalleVisible(data, exportKey){
  const rows=detalleRowsVisible(data);
  window.__detalleVisibleExport=window.__detalleVisibleExport||{};
  window.__detalleVisibleExport[exportKey]=rows;
  if(!rows.length) return '<div class="detalle-empty">No hay detalle para mostrar todavía.</div>';
  const totalLeads=rows.reduce((a,b)=>a+b['Cantidad leads'],0);
  const totalAsesores=new Set(rows.map(r=>r.Asesor)).size;
  const totalSup=new Set(rows.map(r=>r.Supervisor)).size;
  const grouped={};
  rows.forEach(r=>{if(!grouped[r.Supervisor])grouped[r.Supervisor]=[];grouped[r.Supervisor].push(r);});
  return `
    <div class="detalle-inline">
      <div class="detalle-inline-head">
        <h3>4&nbsp; Detalle asesor por supervisor</h3>
        <button class="btn btn-green" onclick="exportDetalleVisible('${exportKey}')">⬇ Exportar detalle asesor</button>
      </div>
      <div class="detalle-summary">
        <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
        <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
        <div><strong>${totalSup.toLocaleString()}</strong><span>Supervisores</span></div>
      </div>
      ${Object.entries(grouped).map(([sup,asesores])=>{
        const t=asesores.reduce((a,b)=>a+b['Cantidad leads'],0);
        return `<div class="detalle-mentor-card">
          <div class="detalle-mentor-head"><span>👤 ${escapeHtml(sup)}</span><b>${t.toLocaleString()} leads</b></div>
          <div class="detalle-asesor-grid">${asesores.map(r=>`<div class="detalle-asesor-item"><span><b>${escapeHtml(r.Asesor)}</b></span><strong>${r['Cantidad leads'].toLocaleString()}</strong></div>`).join('')}</div>
        </div>`;
      }).join('')}
    </div>`;
}
function exportDetalleVisible(key){
  const data=window.__detalleVisibleExport?.[key]||[];
  if(!data.length){showToast('No hay detalle asesor para exportar.');return;}
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Detalle asesor');
  XLSX.writeFile(key+'_'+Date.now()+'.xlsx');
}
function enriquecerSupervisorVisible(row, asesor){const s=supervisorVisible(asesor);return {...row,'CORREO ASESOR ASIGNADO':asesor,'SUPERVISOR':s,'MENTOR':s};}

const SUPERVISOR_SIMPLE_MAP = {"Vinculaciones Carlos Andres Burgos Carvajal": "Otro", "Vinculaciones LAURA CAMILA CASTILLO SEPULVEDA": "Otro", "Vinculaciones Erika Ramirez Rico": "Otro", "Vinculaciones Laura Alejandra Guzman Gutierrez": "Otro", "Vinculaciones Margarita Rosa Navarro Garzon": "Otro", "Vinculaciones Maria Alejandra Vergara Chavez": "Otro", "Vinculaciones Maria Claudia Salas Morales": "Otro", "Vinculaciones Regionales Diego Villota": "Otro", "Vinculaciones Carlos Andres Rodriguez Maturana": "Otro", "Vinculaciones Jorge Mauricio Salgado Viloria": "Otro", "Vinculaciones Regionales": "Otro", "Vinculaciones Regionales MARIA FERNANDA JIMENEZ": "Otro", "Vinculaciones Regionales Yenny Sanchez": "Otro", "Vinculaciones Sindy Alexandra Gomez Herrera": "Otro", "Vinculaciones Maria Paula Prada Coronado": "Otro", "Vinculaciones Regionales Diana Cortes": "Otro", "Vinculaciones Regionales Johana Gomez": "Otro", "Vinculaciones Regionales Erick Carmona": "Otro", "Vinculaciones Carlos Alberto Arias Higuera": "Otro", "Vinculaciones  JAMBERT ANDRÉS MALOTT MUÑOZ": "Otro", "Vinculaciones ANA MILENA PATIÑO PINEDA": "Otro", "Vinculaciones Luz Dary Ramirez Quevedo": "Otro", "Vinculaciones Regionales Angie Borja": "Otro", "Vinculaciones Regionales Jose Viatela": "Otro", "Vinculaciones CAROLINA ROCIO VILORIA PEDREROS": "Otro", "Vinculaciones Regionales Eduardo Paez": "Otro", "Vinculaciones Regionales Katrin Steffi Garzón Avila": "Otro", "Vinculaciones Regionales yuri_pulido@cun.edu.co": "Otro", "Vinculaciones Jeisson Fernando Martinez Sissa": "Otro", "Vinculaciones Jeffer Sebastian Vargas Jimenez": "Otro", "Vinculaciones Sebastian Rodriguez Hernandez": "Otro", "Vinculaciones Diego Fernando Cruz Briceño": "Otro", "Vinculaciones Regionales Laura Montaña": "Otro", "Vinculaciones Andres Beltran": "Otro", "Vinculaciones Regionales Paula Sanchez": "Otro", "Vinculaciones Regionales Angie Moreno": "Otro", "Vinculaciones Sebastian Orlando Cabrera Baez": "Otro", "Vinculaciones Regionales Julian Guaido": "Otro", "Vinculaciones Marianne Ginette Morales Pupo": "Otro", "Vinculaciones Regionales Karen Hernandez": "Otro", "Vinculaciones Regionales Carolina Rosero": "Otro", "Vinculaciones Regionales Cesar Parra": "Otro", "Vinculaciones Regionales Camilo Gualteros": "Otro", "Vinculaciones dairon felipe tellez zambrano": "Otro", "Vinculaciones Maria Fernanda Urbano Cortes": "Otro", "Vinculaciones Alexander Enrique Penaranda Polo": "Otro", "Vinculaciones Regionales Estefany Armijos": "Otro", "Vinculaciones Erick Mauricio Carmona Vergara": "Otro", "Vinculacion Regional Vega Medina": "Otro", "Vinculaciones Manuel Jose Garcia Pabon": "Otro", "Vinculaciones KATHERINE LILIANA GONZALEZ RUBIO ANDRADE": "Otro", "Vinculaciones Fabio Andres Miranda Sequea": "Otro", "Vinculaciones NATALIA REMOLINA HERRERA": "Otro", "Vinculaciones Santa Marta Karen Neira": "Otro", "Vinculaciones Oscar Ivan Torres Gomez": "Otro", "Vinculaciones Juan Camilo Pulido Fagua": "Otro", "Vinculaciones Lina Marcela Paez Cantero": "Otro", "Vinculaciones Victor Hugo Tobon Pérez": "Otro", "Vinculaciones KEVIN LARA ORTEGA": "Otro", "Vinculaciones Regionales Maria Baquero": "Otro", "Proceso Telecampus": "Otro", "TELECAMPUS KENNEDY": "Otro", "TELECAMPUS SUBA": "Otro", "TELECAMPUS MOSQUERA": "Otro", "TELECAMPUS MEDELLIN": "Otro", "TELECAMPUS ENGATIVA": "Otro", "TELECAMPUS TOLIMA": "Otro", "TELECAMPUS BOSA": "Otro", "TELECAMPUS IBAGUE": "Otro", "TELECAMPUS NEIVA": "Otro", "TELECAMPUS BIZEN": "Otro", "Telecampus Chia": "Otro", "TELECAMPUS POLIMODAL": "Otro", "TELECAMPUS SANTA MARTA": "Otro", "TELECAMPUS SINCELEJO": "Otro", "TELECAMPUS TUNJA": "Otro", "Telecampus Telecampus Bucaramanga": "Otro", "TELECAMPUS SOTAVENTO": "Otro", "Telecampus armenia@telecampus.co": "Otro", "TELECAMPUS ELITE": "Otro", "TELECAMPUS CAVES": "Otro", "Telecampus Funza (Facatativá Y Cota). Mauricio Alvarez Diaz": "Otro", "TELECAMPUS LA UNIÓN": "Otro", "TELECAMPUS GAVES": "Otro", "Telecampus Apartadó": "Otro", "pasto@telecampus.co pasto@telecampus.co": "Otro", "Ejecutiva de Cuenta": "Otro", "Proceso Camilo Rodriguez": "Otro", "Bachillersitario": "Otro", "BE": "Otro", "Proceso Regional": "Otro", "jenifer_rodriguez@cun.edu.co": "Retiro", "zharick_olmos@cun.edu.co": "Retiro", "allison_campos@cun.edu.co": "Alejandra", "Retirado": "Retiro", "hasel_moncada@cun.edu.co": "Retiro", "cristian_garciab@cun.edu.co": "Retiro", "jairo_amaya@cun.edu.co": "Retiro", "alexandra_gonzalez@cun.edu.co": "Lorena", "Angie_barretop@cun.edu.co": "Posgrado", "Aura_hipolito@cun.edu.co": "Alejandra", "brayan_rodriguez@cun.edu.co": "Retiro", "bryan_ortiz@cun.edu.co": "Alejandra", "cindy_jerezr@cun.edu.co": "Alejandra", "cristian_pinilla@cun.edu.co": "Alejandra", "cristian_uribe@cun.edu.co": "Retiro", "CUN Digital": "Reasignar", "daniel_pinzonp@cun.edu.co": "Alejandra", "dayanna_gonzalez@cun.edu.co": "Retiro", "diana_proanof@cun.edu.co": "Lili", "evelin_mahecha@cun.edu.co": "Lili", "jameson_rodriguez@cun.edu.co": "Lorena", "Jannis_castellanos@cun.edu.co": "Alejandra", "javier_ramirez@cun.edu.co": "Lili", "jennifer_avila@cun.edu.co": "Lorena", "juan_cardonav@cun.edu.co": "Retiro", "lady_aguirre@cun.edu.co": "Retiro", "lilia_vasquez@cun.edu.co": "Lili", "luisa_velandia@cun.edu.co": "Posgrado", "luz_carrenor@cun.edu.co": "Lili", "marjorie_villadiego@cun.edu.co": "Retiro", "nancy_ome@cun.edu.co": "Posgrado", "nohora_ruiz@cun.edu.co": "Alejandra", "Para Reasignar": "Reasignar", "ricardo_alvarezro@cun.edu.co": "Lili", "sandra_gelvez@cun.edu.co": "Lorena", "santiago_cajamarca@cun.edu.co": "Retiro", "steven_aldana@cun.edu.co": "Retiro", "viviana_calderon@cun.edu.co": "Lili", "karol_ibagon@cun.edu.co": "Retiro", "juan_rodriguezroa@cun.edu.co": "Lorena", "lesslly_acero@cun.edu.co": "Retiro", "sergio_beltran@cun.edu.co": "Retiro", "sindy_rios@cun.edu.co": "Retiro", "katherine_herrera@cun.edu.co": "Retiro", "angela_galeano@cun.edu.co": "Retiro", "alejandra_merchan@cun.edu.co": "Retiro", "deimer_meneses@cun.edu.co": "Retiro", "johanna_barreto@cun.edu.co": "Retiro", "Laura_murillo@cun.edu.co": "Retiro", "karen_penagos@cun.edu.co": "Alejandra", "katherine_latorre@cun.edu.co": "Retiro", "yoharlis_gomez@cun.edu.co": "Retiro", "claudia_ballen@cun.edu.co": "Lili", "laura_ramirezp@cun.edu.co": "Retiro", "elizabeth_villarreal@cun.edu.co": "Lili", "pablo_buitrago@cun.edu.co": "Retiro", "johan_ramirez@cun.edu.co": "Retiro", "laura_ruizba@cun.edu.co": "Retiro", "edison_torreses@cun.edu.co": "Lili", "santiago_buitrago@cun.edu.co": "Lili", "jeison_gamba@cun.edu.co": "Lorena", "elifeth_romero@cun.edu.co": "Lili", "Wilson_duartet@cun.edu.co": "Retiro", "jhans_manzanares@cun.edu.co": "Lili", "cristhian_delgado@cun.edu.co": "Retiro", "elena_rojas@cun.edu.co": "Alejandra", "Jhon_rodriguez@cun.edu.co": "Retiro", "nicolas_fandino@cun.edu.co": "Retiro", "ashley_narino@cun.edu.co": "Retiro", "nicolas_rodriguezd@cun.edu.co": "Retiro", "ana_rayo@cun.edu.co": "Retiro", "kevin_mateus@cun.edu.co": "Retiro", "juan_bautistasa@cun.edu.co": "Retiro", "leonel_sepulvedasa@cun.edu.co": "Retiro", "jose_ospinobe@cun.edu.co": "Retiro", "kimberly_torove@cun.edu.co": "Retiro", "juan_mirandabe@cun.edu.co": "Retiro", "sebastian_usaquenq@cun.edu.co": "Retiro", "luisa_contreraso@cun.edu.co": "Retiro", "andry_loaizac@cun.edu.co": "Retiro", "yeisson_sierrag@cun.edu.co": "Retiro", "pedro_mendozap@cun.edu.co": "Retiro", "natalia_nietoa@cun.edu.co": "Retiro", "nicole_hidalgo@cun.edu.co": "Retiro", "juan_avilap@cun.edu.co": "Lorena", "nicol_ramosv@cun.edu.co": "Alejandra", "dayana_rodriguezh@cun.edu.co": "Lorena", "edwar_pena@cun.edu.co": "Retiro", "gloria_burgos@cun.edu.co": "Retiro", "wilmar_loaiza@cun.edu.co": "Retiro", "edson_bosiga@cun.edu.co": "Retiro", "jaider_gonzalez@cun.edu.co": "Retiro", "jhonatan_reyes@cun.edu.co": "Retiro", "ricardo_arevalo@cun.edu.co": "reasignar", "Comunicaciones Mercadeo": "Retiro", "nestor_obandom@cun.edu.co": "Retiro", "yeimi_bustosb@cun.edu.co": "Lili", "nestor_conrado@cun.edu.co": "Alejandra", "bertha_ramos@cun.edu.co": "Retiro", "jasmin_trianag@cun.edu.co": "Retiro", "nicoll_malpica@cun.edu.co": "Retiro", "cristian_gonzalezm@cun.edu.co": "Retiro", "Cristian_rodriguezr@cun.edu.co": "Retiro", "nicol_moyano@cun.edu.co": "Retiro", "jose_daza@cun.edu.co": "Retiro", "einer_carreno@cun.edu.co": "Retiro", "brayan_cespedesp@cun.edu.co": "Lorena", "caterine_chaparro@cun.edu.co": "Retiro", "anny_vargas@cun.edu.co": "Posgrado", "miguel_fuentes@cun.edu.co": "Retiro", "andrea_aroca@cun.edu.co": "Posgrado", "danna_molano@cun.edu.co": "Posgrado", "jose_jimenezro@cun.edu.co": "Retiro", "valentina_saavedras@cun.edu.co": "Retiro", "camilo_collazos@cun.edu.co": "Retiro", "nicolas_beltranp@cun.edu.co": "Retiro", "nestor_huertas@cun.edu.co": "Retiro", "Juan_diazma@cun.edu.co": "Retiro", "juan_avilag@cun.edu.co": "Retiro", "asdrubal_sotop@cun.edu.co": "Retiro", "stefania_castanedae@cun.edu.co": "Retiro", "cristian_herreno@cun.edu.co": "Retiro", "Vinculaciones Catherine Andrea Cano Llanos": "Otro", "Vinculaciones Lina Paola Hernandez Palacios": "Otro", "Vinculaciones Mary Andrea Garcia Llerena": "Otro", "Vinculaciones Jessika Paola Moreno Guerrero": "Otro", "Vinculaciones Diana Camila Novoa Rodriguez": "Otro", "Vinculaciones Laura Juliana Orduz Landinez": "Otro", "Vinculaciones Juan Sebastian Valenciano Martinez": "Otro", "Vinculaciones Juan Carlos Herrera Cañate": "Otro", "Vinculaciones Karen Yulieth Potosí Silva": "Otro", "Vinculaciones Edwin Julian Valdes Mendes": "Otro", "michaell_murillo@cun.edu.co": "Retiro", "franklim_sanabria@cun.edu.co": "Retiro", "carlos_betancur@cun.edu.co": "Retiro", "edith_rangel@cun.edu.co": "Retiro", "leydy_alba@cun.edu.co": "Retiro", "laura_buitragot@cun.edu.co": "Retiro", "gonzalo_montes@cun.edu.co": "Retiro", "lisseth_montilla@cun.edu.co": "Alejandra", "eric_benavides@cun.edu.co": "Retiro", "alexander_hortua@cun.edu.co": "Retiro", "ginna_pulido@cun.edu.co": "Retiro", "hector_quiroga@cun.edu.co": "Lorena", "fredy_galeano@cun.edu.co": "Retiro", "yuri_torres@cun.edu.co": "Lorena", "carlos_pinzonp@cun.edu.co": "Retiro", "sara_fuentesv@cun.edu.co": "Lili", "nelly_ramosg@cun.edu.co": "Retiro", "mairyn_rueda@cun.edu.co": "Retiro", "pablo_ramirezm@cun.edu.co": "Retiro", "juan_cuellarb@cun.edu.co": "Retiro", "luis_plazasl@cun.edu.co": "Retiro", "sandra_tangarife@cun.edu.co": "Retiro", "lizbeth_dominguez@cun.edu.co": "Retiro", "luis_ricaurtem@cun.edu.co": "Retiro", "feider_granados@cun.edu.co": "Lili", "julian_ortizb@cun.edu.co": "Retiro", "yeferson_bolivars@cun.edu.co": "Retiro", "julieth_sabogal@cun.edu.co": "Retiro", "anderxon_rodriguez@cun.edu.co": "Retiro", "juan_suarezr@cun.edu.co": "Lorena", "nicole_hernandeza@cun.edu.co": "Retiro", "juan_pinzond@cun.edu.co": "Retiro", "angie_leal@cun.edu.co": "Retiro", "juan_leonal@cun.edu.co": "Retiro", "juan_mendozaza@cun.edu.co": "Retiro", "ingrid_becerra@cun.edu.co": "Retiro", "erik_riano@cun.edu.co": "Alejandra", "francisco_martinez@cun.edu.co": "Retiro", "adriana_farfanro@cun.edu.co": "Retiro", "jirineth_castrova@cun.edu.co": "Retiro", "julian_rojasro@cun.edu.co": "Retiro", "yader_fernandez@cun.edu.co": "Retiro", "brayan_sapuy@cun.edu.co": "Retiro", "sahian_camino@cun.edu.co": "Retiro", "karol_castaneda@cun.edu.co": "Retiro", "yeimi_rivera@cun.edu.co": "Retiro", "luis_hernandezc@cun.edu.co": "Retiro", "juan_ardillaos@cun.edu.co": "Retiro", "marivel_rodriguez@cun.edu.co": "Retiro", "sara_vargas@cun.edu.co": "Retiro", "jesus_rincon@cun.edu.co": "Retiro", "saray_barrera@cun.edu.co": "Retiro", "jhon_avilaq@cun.edu.co": "Retiro", "kevin_parada@cun.edu.co": "Lili", "edwin_torres@cun.edu.co": "Retiro", "ivan_burgosc@cun.edu.co": "Retiro", "isis_martinezta@cun.edu.co": "Retiro", "brian_solanoes@cun.edu.co": "Retiro", "angelo_bernalre@cun.edu.co": "Retiro", "karen_quijanolo@cun.edu.co": "Alejandra", "maria_samacamo@cun.edu.co": "Retiro", "johan_morenome@cun.edu.co": "Retiro", "juan_lopezbe@cun.edu.co": "Retiro", "kevin_diaz@cun.edu.co": "Retiro", "orlenis_sanchez@cun.edu.co": "Retiro", "dayana_galvez@cun.edu.co": "Lorena", "nelly_fletcher@cun.edu.co": "Lorena", "anderson_rubiano@cun.edu.co": "Posgrado", "leidy_moralesc@cun.edu.co": "Retiro", "brandon_mora@cun.edu.co": "Alejandra", "carmen_porras@cun.edu.co": "Retiro", "brandon_lara@cun.edu.co": "Alejandra", "evelyt_marin@cun.edu.co": "Retiro", "rafael_patino@cun.edu.co": "Retiro", "francisco_torresp@cun.edu.co": "Retiro", "sebastian_patino@cun.edu.co": "Lili", "joyce_marriaga@cun.edu.co": "Retiro", "ana_gordillo@cun.edu.co": "Retiro", "sulay_chirivi@cun.edu.co": "Lorena", "david_gallegos@cun.edu.co": "Retiro", "juan_florezro@cun.edu.co": "Lili", "jonathan_gonzalezba@cun.edu.co": "Lorena", "carolina_bernal@cun.edu.co": "Retiro", "jenny_munoz@cun.edu.co": "Retiro", "maria_berrio@cun.edu.co": "Retiro", "albert_cantor@cun.edu.co": "Retiro", "pablo_penuelame@cun.edu.co": "Retiro", "yisel_ariasri@cun.edu.co": "Retiro", "cesar_avendanoya@cun.edu.co": "Retiro", "monica_foreroar@cun.edu.co": "Retiro", "laurenth_sotomo@cun.edu.co": "Retiro", "gary_jimenezpa@cun.edu.co": "Lorena", "gabriela_ordonezro@cun.edu.co": "Posgrado", "nicolas_betancourt@cun.edu.co": "Posgrado", "juan_candamilso@cun.edu.co": "Posgrado", "michael_mesahe@cun.edu.co": "Posgrado", "juan_delvalle@cun.edu.co": "Posgrado", "angel_avilasa@cun.edu.co": "Posgrado", "kimberly_castiblanco@cun.edu.co": "Posgrado", "karen_gongora@cun.edu.co": "Posgrado", "jonathan_alarcon@cun.edu.co": "Lili", "lenis_carvajal@cun.edu.co": "Presencial", "angela_caipa@cun.edu.co": "Retiro", "cesar_gomezd@cun.edu.co": "Retiro", "javier_noguera@cun.edu.co": "Retiro", "laura_gonzalezme@cun.edu.co": "Retiro", "jhonatan_beltran@cun.edu.co": "Presencial", "diana_beltrante@cun.edu.co": "Retiro", "santiago_ordonez@cun.edu.co": "Lili", "lina_osorio@cun.edu.co": "Retiro", "erika_guerrero@cun.edu.co": "Retiro", "henry_sanchez@cun.edu.co": "Lorena", "angie_bernal@cun.edu.co": "Retiro", "eilen_hernandez@cun.edu.co": "Retiro", " andres_lopezh@cun.edu.co": "Retiro", "manuel_garcia@cun.edu.co": "Lili", "paula_casasb@cun.edu.co": "Julio", "juan_ortizl@cun.edu.co": "Retiro", "kevin_montenegro@cun.edu.co": "Lorena", "freddy_guerrero@cun.edu.co": "Retiro", "carlos_vargasc@cun.edu.co": "Retiro", "karen_arevalo@cun.edu.co": "Retiro", "william_suarez@cun.edu.co": "Retiro", "nicolas_vidal@cun.edu.co": "Retiro", "shirley_villarreal@cun.edu.co": "Posgrado", "maryuri_ramirez@cun.edu.co": "Retiro", "brandon_cerinza@cun.edu.co": "Retiro", "juan_buitragov@cun.edu.co": "Posgrado", "valentina_torres@cun.edu.co": "Alejandra", "fabian_jimenezc@cun.edu.co": "Posgrado", "juan_rigueros@cun.edu.co": "Retiro", "eyleen_cuspoca@cun.edu.co": "Retiro", "yenifer_huerfano@cun.edu.co": "Retiro", "paula_perez@cun.edu.co": "Presencial", "julieth_trujillo@cun.edu.co": "Retiro", "rosmary_molina@cun.edu.co": "Presencial", "laura_amayar@cun.edu.co": "Presencial", "tania_cuellar@cun.edu.co": "Presencial", "yecid_orjuela@cun.edu.co": "Presencial", "jeimmy_bernal@cun.edu.co": "Posgrado", "laura_martinezv@cun.edu.co": "Presencial", "jenny_mosquera@cun.edu.co": "Retiro", "karol_ortiz@cun.edu.co": "Presencial", "cesar_rodriguezba@cun.edu.co": "Presencial", "nicol_anzola@cun.edu.co": "Retiro", "leidy_montanas@cun.edu.co": "Retiro", "maria_contrerasc@cun.edu.co": "Retiro", "juan_mancerac@cun.edu.co": "Presencial", "ruben_cruz@cun.edu.co": "Presencial", "fernanda_garcia@cun.edu.co": "Alejandra", "jostin_gomez@cun.edu.co": "Alejandra", "jonnathan_poveda@cun.edu.co": "Alejandra", "joan_perez@cun.edu.co": "Alejandra", "laura_zamudio@cun.edu.co": "Retiro", "jose_galeano@cun.edu.co": "Alejandra", "marlon_martinez@cun.edu.co": "Retiro", "cristian_sanchez@cun.edu.co": "Alejandra", "derly_arguellez@cun.edu.co": "Retiro", "brandon_hernandezt@cun.edu.co": "Retiro", "karen_quintero@cun.edu.co": "Presencial", "yeymy_ninog@cun.edu.co": "Retiro", "sarays_bolivarm@cun.edu.co": "Retiro", "karen_giraldoz@cun.edu.co": "Retiro", "jean_mantillac@cun.edu.co": "Presencial", "erik_monterof@cun.edu.co": "Presencial", "maria_castilloc@cun.edu.co": "Retiro", "michael_fernandezr@cun.edu.co": "Retiro", "jhon_hurtador@cun.edu.co": "Retiro", "santiago_povedav@cun.edu.co": "Retiro", "ingrid_hurtadob@cun.edu.co": "Retiro", "geraldine_pinzons@cun.edu.co": "Retiro", "heyner_balaguera@cun.edu.co": "Retiro", "daniela_velasquezv@cun.edu.co": "Retiro", "jhon_vanegasm@cun.edu.co": "Presencial", "carlos_chiappe@cun.edu.co": "Retiro", "leidy_ortiza@cun.edu.co": "Retiro", "angie_salazarr@cun.edu.co": "Retiro", "jean_aguilart@cun.edu.co": "Posgrado", "santiago_gomezl@cun.edu.co": "Posgrado", "andres_bolivarn@cun.edu.co": "Retiro", "shirly_garciag@cun.edu.co": "Retiro", "maicol_vivasc@cun.edu.co": "Retiro", "juan_duranr@cun.edu.co": "Posgrado", "lisbet_rojash@cun.edu.co": "Retiro", "diego_rodriguezb@cun.edu.co": "Posgrado", "angi_hernandezt@cun.edu.co": "Retiro", "paula_ortizm@cun.edu.co": "Lili", "miguel_garciaf@cun.edu.co": "Lili", "andrea_alvarezr@cun.edu.co": "Lorena", "jenny_vegag@cun.edu.co": "Lorena", "karen_torresa@cun.edu.co": "Lili", "ingrid_rativa@cun.edu.co": "Lorena", "braillin_fuenteso@cun.edu.co": "OJT Virtual", "karen_acevedoa@cun.edu.co": "OJT Virtual", "lizeth_benavidesr@cun.edu.co": "OJT Virtual", "kevin_ramirez@cun.edu.co": "OJT Virtual", "kevin_chavarria@cun.edu.co": "OJT Virtual", "jose_pinzon@cun.edu.co": "OJT Virtual", "rody_navarrete@cun.edu.co": "OJT Virtual", "anderson_pinilla@cun.edu.co": "OJT Virtual", "laura_coronadou@cun.edu.co": "OJT Virtual", "jhon_arenas@cun.edu.co": "OJT Virtual", "kevin_bautista@cun.edu.co": "OJT Virtual", "jose_benavides@cun.edu.co": "OJT Virtual"};

/* === DETALLE POR SUPERVISOR USANDO SOLO Supervisor_simple === */
function normSupSimple(v=''){
  return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
}
function emailSupSimple(v=''){
  const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].toLowerCase() : '';
}
function supervisorSimple(asesor){
  const raw = String(asesor || '').trim();
  const email = emailSupSimple(raw);
  if(email && SUPERVISOR_SIMPLE_MAP[email] !== undefined){
    return String(SUPERVISOR_SIMPLE_MAP[email] || 'Sin supervisor').trim() || 'Sin supervisor';
  }
  const norm = normSupSimple(raw);
  const key = Object.keys(SUPERVISOR_SIMPLE_MAP).find(k => normSupSimple(k) === norm || emailSupSimple(k) === email);
  if(key) return String(SUPERVISOR_SIMPLE_MAP[key] || 'Sin supervisor').trim() || 'Sin supervisor';
  return 'Sin supervisor';
}
function asignarSupervisorSimple(row, asesor){
  const supervisor = supervisorSimple(asesor);
  return {...row, 'CORREO ASESOR ASIGNADO': asesor, 'SUPERVISOR': supervisor, 'MENTOR': supervisor};
}
function detalleRowsSimple(data){
  const map = {};
  (data || []).forEach(r => {
    const asesor = r['CORREO ASESOR ASIGNADO'] || '';
    if(!asesor) return;
    const supervisor = r['SUPERVISOR'] || r['MENTOR'] || supervisorSimple(asesor);
    const key = supervisor + '||' + asesor;
    if(!map[key]) map[key] = {Supervisor: supervisor, Asesor: asesor, 'Cantidad leads': 0};
    map[key]['Cantidad leads']++;
  });
  return Object.values(map).sort((a,b)=>String(a.Supervisor).localeCompare(String(b.Supervisor)) || b['Cantidad leads'] - a['Cantidad leads']);
}
function pintarDetalleSimple(data, boxId, exportName){
  const box = document.getElementById(boxId);
  if(!box) return;
  const rows = detalleRowsSimple(data);
  window.__detalleSimpleExport = window.__detalleSimpleExport || {};
  window.__detalleSimpleExport[boxId] = {rows, exportName};
  if(!rows.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por supervisor y asesor.</div>';
    return;
  }
  const totalLeads = rows.reduce((a,b)=>a+b['Cantidad leads'],0);
  const totalAsesores = new Set(rows.map(r=>r.Asesor)).size;
  const totalSup = new Set(rows.map(r=>r.Supervisor)).size;
  const grouped = {};
  rows.forEach(r=>{ if(!grouped[r.Supervisor]) grouped[r.Supervisor]=[]; grouped[r.Supervisor].push(r); });
  box.innerHTML = `
    <div class="detalle-actions"><button class="btn btn-green" onclick="exportDetalleSimple('${boxId}')">⬇ Exportar detalle asesor</button></div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalSup.toLocaleString()}</strong><span>Supervisores</span></div>
    </div>
    ${Object.entries(grouped).map(([sup, asesores])=>{
      const total = asesores.reduce((a,b)=>a+b['Cantidad leads'],0);
      return `<div class="detalle-mentor-card">
        <div class="detalle-mentor-head"><span>👤 ${escapeHtml(sup)}</span><b>${total.toLocaleString()} leads</b></div>
        <div class="detalle-asesor-grid">
          ${asesores.map(r=>`<div class="detalle-asesor-item"><span><b>${escapeHtml(r.Asesor)}</b></span><strong>${r['Cantidad leads'].toLocaleString()}</strong></div>`).join('')}
        </div>
      </div>`;
    }).join('')}
  `;
}
function exportDetalleSimple(boxId){
  const item = window.__detalleSimpleExport?.[boxId];
  const rows = item?.rows || [];
  if(!rows.length){ showToast('No hay detalle asesor para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Detalle asesor');
  XLSX.writeFile((item.exportName || 'Detalle_Asesor') + '_' + Date.now() + '.xlsx');
}
function renderDetalleAsesorSG(){ pintarDetalleSimple(typeof distribuidoData !== 'undefined' ? distribuidoData : [], 'sg-detalle-asesor', 'Detalle_Asesor_Sin_Gestion'); }
function renderDetalleAsesorInter(){ pintarDetalleSimple(typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [], 'inter-detalle-asesor', 'Detalle_Asesor_Interesados'); }
function renderDetalleAsesorNC(){ pintarDetalleSimple(typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [], 'nc-detalle-asesor', 'Detalle_Asesor_No_Contactado'); }


/* === OVERRIDE FINAL DISTRIBUIR SIN GESTIÓN === */
function distribuir(){
  if(!filteredData.length){showToast('No hay leads cargados');return;}
  if(!asesoresList.length){showToast('⚠️ Debes cargar el Excel de asesores primero');return;}
  const leadsOrdenados = ordenarPorHoraCreacionParaDistribuir(filteredData);
  const n = asesoresList.length;
  distribuidoData = leadsOrdenados.map((lead,i)=>asignarSupervisorSimple({...lead,'ID de registro':getLeadId(lead)}, asesoresList[i % n]));
  const res = {};
  distribuidoData.forEach(r=>{const a=r['CORREO ASESOR ASIGNADO']; res[a]=(res[a]||0)+1;});
  const sb=document.getElementById('stat-box');
  if(sb){
    sb.style.display='block';
    sb.innerHTML=`<strong>📊 ${filteredData.length.toLocaleString()} leads · ${n} asesores:</strong><br><br>`+
      Object.entries(res).map(([a,c])=>`<span class="dist-chip">${escapeHtml(a)}: <b>${c.toLocaleString()}</b></span>`).join('');
  }
  renderDistribucionPreview();
  renderDetalleAsesorSG();
  showToast(`✅ ${distribuidoData.length.toLocaleString()} leads distribuidos mezclando antiguos y nuevos`);
}

const SUPERVISOR_SIMPLE_KV = {"Vinculaciones Carlos Andres Burgos Carvajal": "Otro", "Vinculaciones LAURA CAMILA CASTILLO SEPULVEDA": "Otro", "Vinculaciones Erika Ramirez Rico": "Otro", "Vinculaciones Laura Alejandra Guzman Gutierrez": "Otro", "Vinculaciones Margarita Rosa Navarro Garzon": "Otro", "Vinculaciones Maria Alejandra Vergara Chavez": "Otro", "Vinculaciones Maria Claudia Salas Morales": "Otro", "Vinculaciones Regionales Diego Villota": "Otro", "Vinculaciones Carlos Andres Rodriguez Maturana": "Otro", "Vinculaciones Jorge Mauricio Salgado Viloria": "Otro", "Vinculaciones Regionales": "Otro", "Vinculaciones Regionales MARIA FERNANDA JIMENEZ": "Otro", "Vinculaciones Regionales Yenny Sanchez": "Otro", "Vinculaciones Sindy Alexandra Gomez Herrera": "Otro", "Vinculaciones Maria Paula Prada Coronado": "Otro", "Vinculaciones Regionales Diana Cortes": "Otro", "Vinculaciones Regionales Johana Gomez": "Otro", "Vinculaciones Regionales Erick Carmona": "Otro", "Vinculaciones Carlos Alberto Arias Higuera": "Otro", "Vinculaciones  JAMBERT ANDRÉS MALOTT MUÑOZ": "Otro", "Vinculaciones ANA MILENA PATIÑO PINEDA": "Otro", "Vinculaciones Luz Dary Ramirez Quevedo": "Otro", "Vinculaciones Regionales Angie Borja": "Otro", "Vinculaciones Regionales Jose Viatela": "Otro", "Vinculaciones CAROLINA ROCIO VILORIA PEDREROS": "Otro", "Vinculaciones Regionales Eduardo Paez": "Otro", "Vinculaciones Regionales Katrin Steffi Garzón Avila": "Otro", "Vinculaciones Regionales yuri_pulido@cun.edu.co": "Otro", "Vinculaciones Jeisson Fernando Martinez Sissa": "Otro", "Vinculaciones Jeffer Sebastian Vargas Jimenez": "Otro", "Vinculaciones Sebastian Rodriguez Hernandez": "Otro", "Vinculaciones Diego Fernando Cruz Briceño": "Otro", "Vinculaciones Regionales Laura Montaña": "Otro", "Vinculaciones Andres Beltran": "Otro", "Vinculaciones Regionales Paula Sanchez": "Otro", "Vinculaciones Regionales Angie Moreno": "Otro", "Vinculaciones Sebastian Orlando Cabrera Baez": "Otro", "Vinculaciones Regionales Julian Guaido": "Otro", "Vinculaciones Marianne Ginette Morales Pupo": "Otro", "Vinculaciones Regionales Karen Hernandez": "Otro", "Vinculaciones Regionales Carolina Rosero": "Otro", "Vinculaciones Regionales Cesar Parra": "Otro", "Vinculaciones Regionales Camilo Gualteros": "Otro", "Vinculaciones dairon felipe tellez zambrano": "Otro", "Vinculaciones Maria Fernanda Urbano Cortes": "Otro", "Vinculaciones Alexander Enrique Penaranda Polo": "Otro", "Vinculaciones Regionales Estefany Armijos": "Otro", "Vinculaciones Erick Mauricio Carmona Vergara": "Otro", "Vinculacion Regional Vega Medina": "Otro", "Vinculaciones Manuel Jose Garcia Pabon": "Otro", "Vinculaciones KATHERINE LILIANA GONZALEZ RUBIO ANDRADE": "Otro", "Vinculaciones Fabio Andres Miranda Sequea": "Otro", "Vinculaciones NATALIA REMOLINA HERRERA": "Otro", "Vinculaciones Santa Marta Karen Neira": "Otro", "Vinculaciones Oscar Ivan Torres Gomez": "Otro", "Vinculaciones Juan Camilo Pulido Fagua": "Otro", "Vinculaciones Lina Marcela Paez Cantero": "Otro", "Vinculaciones Victor Hugo Tobon Pérez": "Otro", "Vinculaciones KEVIN LARA ORTEGA": "Otro", "Vinculaciones Regionales Maria Baquero": "Otro", "Proceso Telecampus": "Otro", "TELECAMPUS KENNEDY": "Otro", "TELECAMPUS SUBA": "Otro", "TELECAMPUS MOSQUERA": "Otro", "TELECAMPUS MEDELLIN": "Otro", "TELECAMPUS ENGATIVA": "Otro", "TELECAMPUS TOLIMA": "Otro", "TELECAMPUS BOSA": "Otro", "TELECAMPUS IBAGUE": "Otro", "TELECAMPUS NEIVA": "Otro", "TELECAMPUS BIZEN": "Otro", "Telecampus Chia": "Otro", "TELECAMPUS POLIMODAL": "Otro", "TELECAMPUS SANTA MARTA": "Otro", "TELECAMPUS SINCELEJO": "Otro", "TELECAMPUS TUNJA": "Otro", "Telecampus Telecampus Bucaramanga": "Otro", "TELECAMPUS SOTAVENTO": "Otro", "Telecampus armenia@telecampus.co": "Otro", "TELECAMPUS ELITE": "Otro", "TELECAMPUS CAVES": "Otro", "Telecampus Funza (Facatativá Y Cota). Mauricio Alvarez Diaz": "Otro", "TELECAMPUS LA UNIÓN": "Otro", "TELECAMPUS GAVES": "Otro", "Telecampus Apartadó": "Otro", "pasto@telecampus.co pasto@telecampus.co": "Otro", "Ejecutiva de Cuenta": "Otro", "Proceso Camilo Rodriguez": "Otro", "Bachillersitario": "Otro", "BE": "Otro", "Proceso Regional": "Otro", "jenifer_rodriguez@cun.edu.co": "Retiro", "zharick_olmos@cun.edu.co": "Retiro", "allison_campos@cun.edu.co": "Alejandra", "Retirado": "Retiro", "hasel_moncada@cun.edu.co": "Retiro", "cristian_garciab@cun.edu.co": "Retiro", "jairo_amaya@cun.edu.co": "Retiro", "alexandra_gonzalez@cun.edu.co": "Lorena", "Angie_barretop@cun.edu.co": "Posgrado", "Aura_hipolito@cun.edu.co": "Alejandra", "brayan_rodriguez@cun.edu.co": "Retiro", "bryan_ortiz@cun.edu.co": "Alejandra", "cindy_jerezr@cun.edu.co": "Alejandra", "cristian_pinilla@cun.edu.co": "Alejandra", "cristian_uribe@cun.edu.co": "Retiro", "CUN Digital": "Reasignar", "daniel_pinzonp@cun.edu.co": "Alejandra", "dayanna_gonzalez@cun.edu.co": "Retiro", "diana_proanof@cun.edu.co": "Lili", "evelin_mahecha@cun.edu.co": "Lili", "jameson_rodriguez@cun.edu.co": "Lorena", "Jannis_castellanos@cun.edu.co": "Alejandra", "javier_ramirez@cun.edu.co": "Lili", "jennifer_avila@cun.edu.co": "Lorena", "juan_cardonav@cun.edu.co": "Retiro", "lady_aguirre@cun.edu.co": "Retiro", "lilia_vasquez@cun.edu.co": "Lili", "luisa_velandia@cun.edu.co": "Posgrado", "luz_carrenor@cun.edu.co": "Lili", "marjorie_villadiego@cun.edu.co": "Retiro", "nancy_ome@cun.edu.co": "Posgrado", "nohora_ruiz@cun.edu.co": "Alejandra", "Para Reasignar": "Reasignar", "ricardo_alvarezro@cun.edu.co": "Lili", "sandra_gelvez@cun.edu.co": "Lorena", "santiago_cajamarca@cun.edu.co": "Retiro", "steven_aldana@cun.edu.co": "Retiro", "viviana_calderon@cun.edu.co": "Lili", "karol_ibagon@cun.edu.co": "Retiro", "juan_rodriguezroa@cun.edu.co": "Lorena", "lesslly_acero@cun.edu.co": "Retiro", "sergio_beltran@cun.edu.co": "Retiro", "sindy_rios@cun.edu.co": "Retiro", "katherine_herrera@cun.edu.co": "Retiro", "angela_galeano@cun.edu.co": "Retiro", "alejandra_merchan@cun.edu.co": "Retiro", "deimer_meneses@cun.edu.co": "Retiro", "johanna_barreto@cun.edu.co": "Retiro", "Laura_murillo@cun.edu.co": "Retiro", "karen_penagos@cun.edu.co": "Alejandra", "katherine_latorre@cun.edu.co": "Retiro", "yoharlis_gomez@cun.edu.co": "Retiro", "claudia_ballen@cun.edu.co": "Lili", "laura_ramirezp@cun.edu.co": "Retiro", "elizabeth_villarreal@cun.edu.co": "Lili", "pablo_buitrago@cun.edu.co": "Retiro", "johan_ramirez@cun.edu.co": "Retiro", "laura_ruizba@cun.edu.co": "Retiro", "edison_torreses@cun.edu.co": "Lili", "santiago_buitrago@cun.edu.co": "Lili", "jeison_gamba@cun.edu.co": "Lorena", "elifeth_romero@cun.edu.co": "Lili", "Wilson_duartet@cun.edu.co": "Retiro", "jhans_manzanares@cun.edu.co": "Lili", "cristhian_delgado@cun.edu.co": "Retiro", "elena_rojas@cun.edu.co": "Alejandra", "Jhon_rodriguez@cun.edu.co": "Retiro", "nicolas_fandino@cun.edu.co": "Retiro", "ashley_narino@cun.edu.co": "Retiro", "nicolas_rodriguezd@cun.edu.co": "Retiro", "ana_rayo@cun.edu.co": "Retiro", "kevin_mateus@cun.edu.co": "Retiro", "juan_bautistasa@cun.edu.co": "Retiro", "leonel_sepulvedasa@cun.edu.co": "Retiro", "jose_ospinobe@cun.edu.co": "Retiro", "kimberly_torove@cun.edu.co": "Retiro", "juan_mirandabe@cun.edu.co": "Retiro", "sebastian_usaquenq@cun.edu.co": "Retiro", "luisa_contreraso@cun.edu.co": "Retiro", "andry_loaizac@cun.edu.co": "Retiro", "yeisson_sierrag@cun.edu.co": "Retiro", "pedro_mendozap@cun.edu.co": "Retiro", "natalia_nietoa@cun.edu.co": "Retiro", "nicole_hidalgo@cun.edu.co": "Retiro", "juan_avilap@cun.edu.co": "Lorena", "nicol_ramosv@cun.edu.co": "Alejandra", "dayana_rodriguezh@cun.edu.co": "Lorena", "edwar_pena@cun.edu.co": "Retiro", "gloria_burgos@cun.edu.co": "Retiro", "wilmar_loaiza@cun.edu.co": "Retiro", "edson_bosiga@cun.edu.co": "Retiro", "jaider_gonzalez@cun.edu.co": "Retiro", "jhonatan_reyes@cun.edu.co": "Retiro", "ricardo_arevalo@cun.edu.co": "reasignar", "Comunicaciones Mercadeo": "Retiro", "nestor_obandom@cun.edu.co": "Retiro", "yeimi_bustosb@cun.edu.co": "Lili", "nestor_conrado@cun.edu.co": "Alejandra", "bertha_ramos@cun.edu.co": "Retiro", "jasmin_trianag@cun.edu.co": "Retiro", "nicoll_malpica@cun.edu.co": "Retiro", "cristian_gonzalezm@cun.edu.co": "Retiro", "Cristian_rodriguezr@cun.edu.co": "Retiro", "nicol_moyano@cun.edu.co": "Retiro", "jose_daza@cun.edu.co": "Retiro", "einer_carreno@cun.edu.co": "Retiro", "brayan_cespedesp@cun.edu.co": "Lorena", "caterine_chaparro@cun.edu.co": "Retiro", "anny_vargas@cun.edu.co": "Posgrado", "miguel_fuentes@cun.edu.co": "Retiro", "andrea_aroca@cun.edu.co": "Posgrado", "danna_molano@cun.edu.co": "Posgrado", "jose_jimenezro@cun.edu.co": "Retiro", "valentina_saavedras@cun.edu.co": "Retiro", "camilo_collazos@cun.edu.co": "Retiro", "nicolas_beltranp@cun.edu.co": "Retiro", "nestor_huertas@cun.edu.co": "Retiro", "Juan_diazma@cun.edu.co": "Retiro", "juan_avilag@cun.edu.co": "Retiro", "asdrubal_sotop@cun.edu.co": "Retiro", "stefania_castanedae@cun.edu.co": "Retiro", "cristian_herreno@cun.edu.co": "Retiro", "Vinculaciones Catherine Andrea Cano Llanos": "Otro", "Vinculaciones Lina Paola Hernandez Palacios": "Otro", "Vinculaciones Mary Andrea Garcia Llerena": "Otro", "Vinculaciones Jessika Paola Moreno Guerrero": "Otro", "Vinculaciones Diana Camila Novoa Rodriguez": "Otro", "Vinculaciones Laura Juliana Orduz Landinez": "Otro", "Vinculaciones Juan Sebastian Valenciano Martinez": "Otro", "Vinculaciones Juan Carlos Herrera Cañate": "Otro", "Vinculaciones Karen Yulieth Potosí Silva": "Otro", "Vinculaciones Edwin Julian Valdes Mendes": "Otro", "michaell_murillo@cun.edu.co": "Retiro", "franklim_sanabria@cun.edu.co": "Retiro", "carlos_betancur@cun.edu.co": "Retiro", "edith_rangel@cun.edu.co": "Retiro", "leydy_alba@cun.edu.co": "Retiro", "laura_buitragot@cun.edu.co": "Retiro", "gonzalo_montes@cun.edu.co": "Retiro", "lisseth_montilla@cun.edu.co": "Alejandra", "eric_benavides@cun.edu.co": "Retiro", "alexander_hortua@cun.edu.co": "Retiro", "ginna_pulido@cun.edu.co": "Retiro", "hector_quiroga@cun.edu.co": "Lorena", "fredy_galeano@cun.edu.co": "Retiro", "yuri_torres@cun.edu.co": "Lorena", "carlos_pinzonp@cun.edu.co": "Retiro", "sara_fuentesv@cun.edu.co": "Lili", "nelly_ramosg@cun.edu.co": "Retiro", "mairyn_rueda@cun.edu.co": "Retiro", "pablo_ramirezm@cun.edu.co": "Retiro", "juan_cuellarb@cun.edu.co": "Retiro", "luis_plazasl@cun.edu.co": "Retiro", "sandra_tangarife@cun.edu.co": "Retiro", "lizbeth_dominguez@cun.edu.co": "Retiro", "luis_ricaurtem@cun.edu.co": "Retiro", "feider_granados@cun.edu.co": "Lili", "julian_ortizb@cun.edu.co": "Retiro", "yeferson_bolivars@cun.edu.co": "Retiro", "julieth_sabogal@cun.edu.co": "Retiro", "anderxon_rodriguez@cun.edu.co": "Retiro", "juan_suarezr@cun.edu.co": "Lorena", "nicole_hernandeza@cun.edu.co": "Retiro", "juan_pinzond@cun.edu.co": "Retiro", "angie_leal@cun.edu.co": "Retiro", "juan_leonal@cun.edu.co": "Retiro", "juan_mendozaza@cun.edu.co": "Retiro", "ingrid_becerra@cun.edu.co": "Retiro", "erik_riano@cun.edu.co": "Alejandra", "francisco_martinez@cun.edu.co": "Retiro", "adriana_farfanro@cun.edu.co": "Retiro", "jirineth_castrova@cun.edu.co": "Retiro", "julian_rojasro@cun.edu.co": "Retiro", "yader_fernandez@cun.edu.co": "Retiro", "brayan_sapuy@cun.edu.co": "Retiro", "sahian_camino@cun.edu.co": "Retiro", "karol_castaneda@cun.edu.co": "Retiro", "yeimi_rivera@cun.edu.co": "Retiro", "luis_hernandezc@cun.edu.co": "Retiro", "juan_ardillaos@cun.edu.co": "Retiro", "marivel_rodriguez@cun.edu.co": "Retiro", "sara_vargas@cun.edu.co": "Retiro", "jesus_rincon@cun.edu.co": "Retiro", "saray_barrera@cun.edu.co": "Retiro", "jhon_avilaq@cun.edu.co": "Retiro", "kevin_parada@cun.edu.co": "Lili", "edwin_torres@cun.edu.co": "Retiro", "ivan_burgosc@cun.edu.co": "Retiro", "isis_martinezta@cun.edu.co": "Retiro", "brian_solanoes@cun.edu.co": "Retiro", "angelo_bernalre@cun.edu.co": "Retiro", "karen_quijanolo@cun.edu.co": "Alejandra", "maria_samacamo@cun.edu.co": "Retiro", "johan_morenome@cun.edu.co": "Retiro", "juan_lopezbe@cun.edu.co": "Retiro", "kevin_diaz@cun.edu.co": "Retiro", "orlenis_sanchez@cun.edu.co": "Retiro", "dayana_galvez@cun.edu.co": "Lorena", "nelly_fletcher@cun.edu.co": "Lorena", "anderson_rubiano@cun.edu.co": "Posgrado", "leidy_moralesc@cun.edu.co": "Retiro", "brandon_mora@cun.edu.co": "Alejandra", "carmen_porras@cun.edu.co": "Retiro", "brandon_lara@cun.edu.co": "Alejandra", "evelyt_marin@cun.edu.co": "Retiro", "rafael_patino@cun.edu.co": "Retiro", "francisco_torresp@cun.edu.co": "Retiro", "sebastian_patino@cun.edu.co": "Lili", "joyce_marriaga@cun.edu.co": "Retiro", "ana_gordillo@cun.edu.co": "Retiro", "sulay_chirivi@cun.edu.co": "Lorena", "david_gallegos@cun.edu.co": "Retiro", "juan_florezro@cun.edu.co": "Lili", "jonathan_gonzalezba@cun.edu.co": "Lorena", "carolina_bernal@cun.edu.co": "Retiro", "jenny_munoz@cun.edu.co": "Retiro", "maria_berrio@cun.edu.co": "Retiro", "albert_cantor@cun.edu.co": "Retiro", "pablo_penuelame@cun.edu.co": "Retiro", "yisel_ariasri@cun.edu.co": "Retiro", "cesar_avendanoya@cun.edu.co": "Retiro", "monica_foreroar@cun.edu.co": "Retiro", "laurenth_sotomo@cun.edu.co": "Retiro", "gary_jimenezpa@cun.edu.co": "Lorena", "gabriela_ordonezro@cun.edu.co": "Posgrado", "nicolas_betancourt@cun.edu.co": "Posgrado", "juan_candamilso@cun.edu.co": "Posgrado", "michael_mesahe@cun.edu.co": "Posgrado", "juan_delvalle@cun.edu.co": "Posgrado", "angel_avilasa@cun.edu.co": "Posgrado", "kimberly_castiblanco@cun.edu.co": "Posgrado", "karen_gongora@cun.edu.co": "Posgrado", "jonathan_alarcon@cun.edu.co": "Lili", "lenis_carvajal@cun.edu.co": "Presencial", "angela_caipa@cun.edu.co": "Retiro", "cesar_gomezd@cun.edu.co": "Retiro", "javier_noguera@cun.edu.co": "Retiro", "laura_gonzalezme@cun.edu.co": "Retiro", "jhonatan_beltran@cun.edu.co": "Presencial", "diana_beltrante@cun.edu.co": "Retiro", "santiago_ordonez@cun.edu.co": "Lili", "lina_osorio@cun.edu.co": "Retiro", "erika_guerrero@cun.edu.co": "Retiro", "henry_sanchez@cun.edu.co": "Lorena", "angie_bernal@cun.edu.co": "Retiro", "eilen_hernandez@cun.edu.co": "Retiro", " andres_lopezh@cun.edu.co": "Retiro", "manuel_garcia@cun.edu.co": "Lili", "paula_casasb@cun.edu.co": "Julio", "juan_ortizl@cun.edu.co": "Retiro", "kevin_montenegro@cun.edu.co": "Lorena", "freddy_guerrero@cun.edu.co": "Retiro", "carlos_vargasc@cun.edu.co": "Retiro", "karen_arevalo@cun.edu.co": "Retiro", "william_suarez@cun.edu.co": "Retiro", "nicolas_vidal@cun.edu.co": "Retiro", "shirley_villarreal@cun.edu.co": "Posgrado", "maryuri_ramirez@cun.edu.co": "Retiro", "brandon_cerinza@cun.edu.co": "Retiro", "juan_buitragov@cun.edu.co": "Posgrado", "valentina_torres@cun.edu.co": "Alejandra", "fabian_jimenezc@cun.edu.co": "Posgrado", "juan_rigueros@cun.edu.co": "Retiro", "eyleen_cuspoca@cun.edu.co": "Retiro", "yenifer_huerfano@cun.edu.co": "Retiro", "paula_perez@cun.edu.co": "Presencial", "julieth_trujillo@cun.edu.co": "Retiro", "rosmary_molina@cun.edu.co": "Presencial", "laura_amayar@cun.edu.co": "Presencial", "tania_cuellar@cun.edu.co": "Presencial", "yecid_orjuela@cun.edu.co": "Presencial", "jeimmy_bernal@cun.edu.co": "Posgrado", "laura_martinezv@cun.edu.co": "Presencial", "jenny_mosquera@cun.edu.co": "Retiro", "karol_ortiz@cun.edu.co": "Presencial", "cesar_rodriguezba@cun.edu.co": "Presencial", "nicol_anzola@cun.edu.co": "Retiro", "leidy_montanas@cun.edu.co": "Retiro", "maria_contrerasc@cun.edu.co": "Retiro", "juan_mancerac@cun.edu.co": "Presencial", "ruben_cruz@cun.edu.co": "Presencial", "fernanda_garcia@cun.edu.co": "Alejandra", "jostin_gomez@cun.edu.co": "Alejandra", "jonnathan_poveda@cun.edu.co": "Alejandra", "joan_perez@cun.edu.co": "Alejandra", "laura_zamudio@cun.edu.co": "Retiro", "jose_galeano@cun.edu.co": "Alejandra", "marlon_martinez@cun.edu.co": "Retiro", "cristian_sanchez@cun.edu.co": "Alejandra", "derly_arguellez@cun.edu.co": "Retiro", "brandon_hernandezt@cun.edu.co": "Retiro", "karen_quintero@cun.edu.co": "Presencial", "yeymy_ninog@cun.edu.co": "Retiro", "sarays_bolivarm@cun.edu.co": "Retiro", "karen_giraldoz@cun.edu.co": "Retiro", "jean_mantillac@cun.edu.co": "Presencial", "erik_monterof@cun.edu.co": "Presencial", "maria_castilloc@cun.edu.co": "Retiro", "michael_fernandezr@cun.edu.co": "Retiro", "jhon_hurtador@cun.edu.co": "Retiro", "santiago_povedav@cun.edu.co": "Retiro", "ingrid_hurtadob@cun.edu.co": "Retiro", "geraldine_pinzons@cun.edu.co": "Retiro", "heyner_balaguera@cun.edu.co": "Retiro", "daniela_velasquezv@cun.edu.co": "Retiro", "jhon_vanegasm@cun.edu.co": "Presencial", "carlos_chiappe@cun.edu.co": "Retiro", "leidy_ortiza@cun.edu.co": "Retiro", "angie_salazarr@cun.edu.co": "Retiro", "jean_aguilart@cun.edu.co": "Posgrado", "santiago_gomezl@cun.edu.co": "Posgrado", "andres_bolivarn@cun.edu.co": "Retiro", "shirly_garciag@cun.edu.co": "Retiro", "maicol_vivasc@cun.edu.co": "Retiro", "juan_duranr@cun.edu.co": "Posgrado", "lisbet_rojash@cun.edu.co": "Retiro", "diego_rodriguezb@cun.edu.co": "Posgrado", "angi_hernandezt@cun.edu.co": "Retiro", "paula_ortizm@cun.edu.co": "Lili", "miguel_garciaf@cun.edu.co": "Lili", "andrea_alvarezr@cun.edu.co": "Lorena", "jenny_vegag@cun.edu.co": "Lorena", "karen_torresa@cun.edu.co": "Lili", "ingrid_rativa@cun.edu.co": "Lorena", "braillin_fuenteso@cun.edu.co": "OJT Virtual", "karen_acevedoa@cun.edu.co": "OJT Virtual", "lizeth_benavidesr@cun.edu.co": "OJT Virtual", "kevin_ramirez@cun.edu.co": "OJT Virtual", "kevin_chavarria@cun.edu.co": "OJT Virtual", "jose_pinzon@cun.edu.co": "OJT Virtual", "rody_navarrete@cun.edu.co": "OJT Virtual", "anderson_pinilla@cun.edu.co": "OJT Virtual", "laura_coronadou@cun.edu.co": "OJT Virtual", "jhon_arenas@cun.edu.co": "OJT Virtual", "kevin_bautista@cun.edu.co": "OJT Virtual", "jose_benavides@cun.edu.co": "OJT Virtual"};

/* === SUPERVISOR_SIMPLE FINAL: IZQUIERDA = ASESOR/CORREO, DERECHA = SUPERVISOR === */
function supSimpleNorm(v=''){
  return String(v || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .trim()
    .toLowerCase()
    .replace(/\s+/g,' ');
}

function supSimpleEmail(v=''){
  const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].toLowerCase() : '';
}

function supervisorSimpleLookup(asesor){
  const raw = String(asesor || '').trim();
  if(!raw) return 'Sin supervisor';

  // 1) Exacto
  if(Object.prototype.hasOwnProperty.call(SUPERVISOR_SIMPLE_KV, raw)){
    return String(SUPERVISOR_SIMPLE_KV[raw] || 'Sin supervisor').trim() || 'Sin supervisor';
  }

  // 2) Por correo dentro del texto
  const email = supSimpleEmail(raw);
  if(email){
    for(const [k,v] of Object.entries(SUPERVISOR_SIMPLE_KV)){
      const ke = supSimpleEmail(k);
      if(ke && ke === email){
        return String(v || 'Sin supervisor').trim() || 'Sin supervisor';
      }
    }
  }

  // 3) Normalizado sin tildes y case-insensitive
  const norm = supSimpleNorm(raw);
  for(const [k,v] of Object.entries(SUPERVISOR_SIMPLE_KV)){
    if(supSimpleNorm(k) === norm){
      return String(v || 'Sin supervisor').trim() || 'Sin supervisor';
    }
  }

  return 'Sin supervisor';
}

/* Alias para compatibilidad con el resto de archivos */
function supervisorSimple(asesor){ return supervisorSimpleLookup(asesor); }
function supervisorEmbebido(asesor){ return supervisorSimpleLookup(asesor); }
function supervisorPorAsesorDOM(asesor){ return supervisorSimpleLookup(asesor); }
function supervisorUniversal(asesor){ return supervisorSimpleLookup(asesor); }
function mentorDeAsesor(asesor){ return supervisorSimpleLookup(asesor); }
function areaDeAsesor(){ return ''; }
function grupoDeAsesor(){ return ''; }

function asignarSupervisorSimple(row, asesor){
  const sup = supervisorSimpleLookup(asesor);
  return {
    ...row,
    'CORREO ASESOR ASIGNADO': asesor,
    'SUPERVISOR': sup,
    'MENTOR': sup
  };
}

function detalleRowsSimple(data){
  const map = {};
  (data || []).forEach(r => {
    const asesor = r['CORREO ASESOR ASIGNADO'] || r['Asesor'] || '';
    if(!asesor) return;

    const supervisor = r['SUPERVISOR'] || r['MENTOR'] || supervisorSimpleLookup(asesor);
    const key = supervisor + '||' + asesor;

    if(!map[key]){
      map[key] = {Supervisor: supervisor, Asesor: asesor, 'Cantidad leads': 0};
    }
    map[key]['Cantidad leads']++;
  });

  return Object.values(map).sort((a,b) =>
    String(a.Supervisor).localeCompare(String(b.Supervisor)) ||
    b['Cantidad leads'] - a['Cantidad leads']
  );
}

function pintarDetalleSimple(data, boxId, exportName){
  const box = document.getElementById(boxId);
  if(!box) return;

  const rows = detalleRowsSimple(data);
  window.__detalleSimpleExport = window.__detalleSimpleExport || {};
  window.__detalleSimpleExport[boxId] = {rows, exportName};

  if(!rows.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por supervisor y asesor.</div>';
    return;
  }

  const totalLeads = rows.reduce((a,b)=>a+b['Cantidad leads'],0);
  const totalAsesores = new Set(rows.map(r=>r.Asesor)).size;
  const totalSup = new Set(rows.map(r=>r.Supervisor)).size;
  const grouped = {};
  rows.forEach(r => {
    if(!grouped[r.Supervisor]) grouped[r.Supervisor] = [];
    grouped[r.Supervisor].push(r);
  });

  box.innerHTML = `
    <div class="detalle-actions">
      <button class="btn btn-green" onclick="exportDetalleSimple('${boxId}')">⬇ Exportar detalle asesor</button>
    </div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalSup.toLocaleString()}</strong><span>Supervisores</span></div>
    </div>
    ${Object.entries(grouped).map(([sup, asesores]) => {
      const total = asesores.reduce((a,b)=>a+b['Cantidad leads'],0);
      return `<div class="detalle-mentor-card">
        <div class="detalle-mentor-head"><span>👤 ${escapeHtml(sup)}</span><b>${total.toLocaleString()} leads</b></div>
        <div class="detalle-asesor-grid">
          ${asesores.map(r => `<div class="detalle-asesor-item"><span><b>${escapeHtml(r.Asesor)}</b></span><strong>${r['Cantidad leads'].toLocaleString()}</strong></div>`).join('')}
        </div>
      </div>`;
    }).join('')}
  `;
}

function exportDetalleSimple(boxId){
  const item = window.__detalleSimpleExport?.[boxId];
  const rows = item?.rows || [];
  if(!rows.length){ showToast('No hay detalle asesor para exportar.'); return; }

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Detalle asesor');
  XLSX.writeFile((item.exportName || 'Detalle_Asesor') + '_' + Date.now() + '.xlsx');
}

function renderDetalleAsesorSG(){ pintarDetalleSimple(typeof distribuidoData !== 'undefined' ? distribuidoData : [], 'sg-detalle-asesor', 'Detalle_Asesor_Sin_Gestion'); }
function renderDetalleAsesorInter(){ pintarDetalleSimple(typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [], 'inter-detalle-asesor', 'Detalle_Asesor_Interesados'); }
function renderDetalleAsesorNC(){ pintarDetalleSimple(typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [], 'nc-detalle-asesor', 'Detalle_Asesor_No_Contactado'); }


/* === EXPORT DETALLE SIMPLE FIX === */
function exportDetalleSimple(boxId){
  const item = window.__detalleSimpleExport?.[boxId];
  const rows = item?.rows || [];
  if(!rows.length){
    showToast('No hay detalle asesor para exportar.');
    return;
  }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Detalle asesor');
  XLSX.writeFile((item.exportName || 'Detalle_Asesor') + '_' + Date.now() + '.xlsx');
}


/* === OVERRIDE FINAL: SOLO Supervisor_simple.json ===
   Formato:
   izquierda = asesor/correo
   derecha = supervisor
*/
function supervisorSimpleLookup(asesor){
  const raw = String(asesor || '').trim();
  if(!raw) return 'Sin supervisor';

  if(Object.prototype.hasOwnProperty.call(SUPERVISOR_SIMPLE_KV, raw)){
    return String(SUPERVISOR_SIMPLE_KV[raw] || 'Sin supervisor').trim() || 'Sin supervisor';
  }

  const email = supSimpleEmail(raw);
  if(email){
    for(const [k,v] of Object.entries(SUPERVISOR_SIMPLE_KV)){
      if(supSimpleEmail(k) === email){
        return String(v || 'Sin supervisor').trim() || 'Sin supervisor';
      }
    }
  }

  const norm = supSimpleNorm(raw);
  for(const [k,v] of Object.entries(SUPERVISOR_SIMPLE_KV)){
    if(supSimpleNorm(k) === norm){
      return String(v || 'Sin supervisor').trim() || 'Sin supervisor';
    }
  }

  return 'Sin supervisor';
}

function supervisorSimple(asesor){ return supervisorSimpleLookup(asesor); }
function mentorDeAsesor(asesor){ return supervisorSimpleLookup(asesor); }

function asignarSupervisorSimple(row, asesor){
  const sup = supervisorSimpleLookup(asesor);
  return {
    ...row,
    'CORREO ASESOR ASIGNADO': asesor,
    'SUPERVISOR': sup,
    'MENTOR': sup
  };
}

/* =========================================================
   FIX SOLICITADO — SIN GESTIÓN
   - La tabla de Distribución muestra: ID, CORREO ASESOR ASIGNADO, MENTOR
   - El Excel exportado incluye MENTOR
   - El Detalle asesor agrupa por MENTOR
   ========================================================= */
function mentorAsignadoFinal(asesor, row){
  const directo = String(row?.['MENTOR'] || row?.['SUPERVISOR'] || '').trim();
  if(directo) return directo;
  if(typeof supervisorSimpleLookup === 'function') return supervisorSimpleLookup(asesor);
  if(typeof mentorDeAsesor === 'function') return mentorDeAsesor(asesor);
  return 'Sin mentor';
}

function getDistRow(row){
  const asesor = String(row?.['CORREO ASESOR ASIGNADO'] || '').trim();
  return {
    'ID de registro': getLeadId(row),
    'CORREO ASESOR ASIGNADO': asesor,
    'MENTOR': mentorAsignadoFinal(asesor, row)
  };
}

function renderDistribucionPreview(){
  const rows = distribuidoData.slice(0,300).map(getDistRow);
  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO', 'MENTOR'];

  const head = '<thead style="background:var(--azul);color:white;position:sticky;top:0"><tr>'+
    cols.map(c=>`<th style="padding:7px 9px;white-space:nowrap;font-size:.72rem">${escapeHtml(c)}</th>`).join('')+
    '</tr></thead>';

  const body = '<tbody>'+rows.map((r,i)=>
    `<tr style="${i%2?'background:#f8fafc':''}">`+
    cols.map(c=>{
      const v = r[c] || '';
      return `<td style="padding:5px 9px;white-space:nowrap;font-size:.72rem;border-bottom:1px solid #eef1f5;max-width:360px;overflow:hidden;text-overflow:ellipsis" title="${escapeHtml(v)}">${escapeHtml(v)}</td>`;
    }).join('')+
    '</tr>'
  ).join('')+'</tbody>';

  const prev = document.getElementById('dist-preview');
  if(!prev) return;
  prev.style.display = 'block';
  prev.innerHTML = `<table id="dist-tbl" style="width:100%;border-collapse:collapse">${head}${body}</table>`+
    `<div style="padding:8px 12px;font-size:.77rem;color:#888;border-top:1px solid var(--borde)">Mostrando ${Math.min(300,distribuidoData.length)} de ${distribuidoData.length.toLocaleString()} · exporta para ver todo</div>`;
}

function exportarDist(){
  if(!distribuidoData.length){showToast('Primero distribuye los leads');return;}
  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO', 'MENTOR'];
  const rows = distribuidoData.map(getDistRow);
  const sinId = rows.filter(r => !r['ID de registro']).length;
  if(sinId > 0){
    console.warn(`Hay ${sinId} registros sin ID de registro. Revisa el encabezado del Excel original.`);
  }
  const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Distribuido');
  XLSX.writeFile(wb2, 'Leads_Distribuidos_ID_Asesor_Mentor_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} leads exportados con MENTOR`);
}

function pintarDetallePorMentorSG(data){
  const card = document.getElementById('sg-detalle-asesor-card');
  const box = document.getElementById('sg-detalle-asesor');
  if(!box) return;
  if(card) card.style.display = 'block';

  if(!data || !data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por mentor y asesor.</div>';
    return;
  }

  const grouped = {};
  data.forEach(r=>{
    const asesor = String(r['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
    const mentor = mentorAsignadoFinal(asesor, r) || 'Sin mentor';
    if(!grouped[mentor]) grouped[mentor] = {};
    grouped[mentor][asesor] = (grouped[mentor][asesor] || 0) + 1;
  });

  const rowsExport = [];
  Object.entries(grouped).forEach(([mentor, asesores])=>{
    Object.entries(asesores).forEach(([asesor, cantidad])=>{
      rowsExport.push({ MENTOR: mentor, 'CORREO ASESOR ASIGNADO': asesor, 'Cantidad leads': cantidad });
    });
  });
  window.__detalleSimpleExport = window.__detalleSimpleExport || {};
  window.__detalleSimpleExport['sg-detalle-asesor'] = { rows: rowsExport, exportName: 'Detalle_Asesor_Por_Mentor_Sin_Gestion' };

  const totalLeads = data.length;
  const totalAsesores = new Set(data.map(r => r['CORREO ASESOR ASIGNADO'] || 'Sin asesor')).size;
  const totalMentores = Object.keys(grouped).length;

  box.innerHTML = `
    <div class="detalle-actions"><button class="btn btn-green" onclick="exportDetalleSimple('sg-detalle-asesor')">⬇ Exportar detalle asesor</button></div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalMentores.toLocaleString()}</strong><span>Mentores</span></div>
    </div>
    ${Object.entries(grouped).sort((a,b)=>a[0].localeCompare(b[0])).map(([mentor, asesores])=>{
      const total = Object.values(asesores).reduce((a,b)=>a+b,0);
      const cards = Object.entries(asesores).sort((a,b)=>b[1]-a[1]).map(([asesor,c]) =>
        `<div class="detalle-asesor-item"><span><b>${escapeHtml(asesor)}</b></span><strong>${c.toLocaleString()}</strong></div>`
      ).join('');
      return `<div class="detalle-mentor-card">
        <div class="detalle-mentor-head"><span>👤 MENTOR: ${escapeHtml(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
        <div class="detalle-asesor-grid">${cards}</div>
      </div>`;
    }).join('')}
  `;
}

function renderDetalleAsesorSG(){
  pintarDetallePorMentorSG(typeof distribuidoData !== 'undefined' ? distribuidoData : []);
}

function distribuir(){
  if(!filteredData.length){showToast('No hay leads cargados');return;}
  if(!asesoresList.length){showToast('⚠️ Debes cargar el Excel de asesores primero');return;}

  const leadsOrdenados = ordenarPorHoraCreacionParaDistribuir(filteredData);
  const n = asesoresList.length;
  distribuidoData = leadsOrdenados.map((lead, i) => {
    const asesor = asesoresList[i % n];
    const mentor = mentorAsignadoFinal(asesor, null);
    return {
      ...lead,
      'ID de registro': getLeadId(lead),
      'CORREO ASESOR ASIGNADO': asesor,
      'MENTOR': mentor,
      'SUPERVISOR': mentor
    };
  });

  const res = {};
  distribuidoData.forEach(r=>{
    const a = r['CORREO ASESOR ASIGNADO'];
    res[a] = (res[a] || 0) + 1;
  });

  const sb = document.getElementById('stat-box');
  if(sb){
    sb.style.display = 'block';
    sb.innerHTML = `<strong>📊 ${filteredData.length.toLocaleString()} leads · ${n} asesores:</strong><br><br>`+
      Object.entries(res).map(([a,c])=>`<span class="dist-chip">${escapeHtml(a)}: <b>${c.toLocaleString()}</b></span>`).join('');
  }

  renderDistribucionPreview();
  renderDetalleAsesorSG();
  showToast(`✅ ${distribuidoData.length.toLocaleString()} leads distribuidos con MENTOR`);
}

/* =========================================================
   AJUSTE FINAL — SIN GESTIÓN
   - La tabla Distribución vuelve a mostrar solo ID y CORREO ASESOR ASIGNADO
   - El Excel de distribución vuelve a exportar solo ID y CORREO ASESOR ASIGNADO
   - El Detalle asesor sigue agrupando por MENTOR
   - El botón exporta el Detalle asesor como PNG
   ========================================================= */
function getDistRow(row){
  return {
    'ID de registro': getLeadId(row),
    'CORREO ASESOR ASIGNADO': String(row?.['CORREO ASESOR ASIGNADO'] || '').trim()
  };
}

function renderDistribucionPreview(){
  const rows = distribuidoData.slice(0,300).map(getDistRow);
  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];

  const head = '<thead style="background:var(--azul);color:white;position:sticky;top:0"><tr>'+ 
    cols.map(c=>`<th style="padding:7px 9px;white-space:nowrap;font-size:.72rem">${escapeHtml(c)}</th>`).join('')+
    '</tr></thead>';

  const body = '<tbody>'+rows.map((r,i)=>
    `<tr style="${i%2?'background:#f8fafc':''}">`+
    cols.map(c=>{
      const v = r[c] || '';
      return `<td style="padding:5px 9px;white-space:nowrap;font-size:.72rem;border-bottom:1px solid #eef1f5;max-width:360px;overflow:hidden;text-overflow:ellipsis" title="${escapeHtml(v)}">${escapeHtml(v)}</td>`;
    }).join('')+
    '</tr>'
  ).join('')+'</tbody>';

  const prev = document.getElementById('dist-preview');
  if(!prev) return;
  prev.style.display = 'block';
  prev.innerHTML = `<table id="dist-tbl" style="width:100%;border-collapse:collapse">${head}${body}</table>`+
    `<div style="padding:8px 12px;font-size:.77rem;color:#888;border-top:1px solid var(--borde)">Mostrando ${Math.min(300,distribuidoData.length)} de ${distribuidoData.length.toLocaleString()} · exporta para ver todo</div>`;
}

function exportarDist(){
  if(!distribuidoData.length){showToast('Primero distribuye los leads');return;}
  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
  const rows = distribuidoData.map(getDistRow);
  const sinId = rows.filter(r => !r['ID de registro']).length;
  if(sinId > 0){
    console.warn(`Hay ${sinId} registros sin ID de registro. Revisa el encabezado del Excel original.`);
  }
  const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Distribuido');
  XLSX.writeFile(wb2, 'Leads_Distribuidos_ID_Asesor_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} leads exportados`);
}

function getDetalleMentorRowsSG(data){
  const grouped = {};
  (data || []).forEach(r=>{
    const asesor = String(r['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
    const mentor = mentorAsignadoFinal(asesor, r) || 'Sin mentor';
    if(!grouped[mentor]) grouped[mentor] = {};
    grouped[mentor][asesor] = (grouped[mentor][asesor] || 0) + 1;
  });

  const rows = [];
  Object.entries(grouped).forEach(([mentor, asesores])=>{
    Object.entries(asesores).forEach(([asesor, cantidad])=>{
      rows.push({ MENTOR: mentor, 'CORREO ASESOR ASIGNADO': asesor, 'Cantidad leads': cantidad });
    });
  });
  return rows.sort((a,b)=>
    String(a.MENTOR).localeCompare(String(b.MENTOR)) ||
    Number(b['Cantidad leads']) - Number(a['Cantidad leads']) ||
    String(a['CORREO ASESOR ASIGNADO']).localeCompare(String(b['CORREO ASESOR ASIGNADO']))
  );
}

function exportDetalleAsesorPNG(){
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  const rows = getDetalleMentorRowsSG(data);
  if(!rows.length){ showToast('No hay detalle asesor para exportar.'); return; }

  const totalLeads = data.length;
  const totalAsesores = new Set(data.map(r => r['CORREO ASESOR ASIGNADO'] || 'Sin asesor')).size;
  const totalMentores = new Set(rows.map(r => r.MENTOR)).size;

  const scale = 2;
  const width = 1200;
  const margin = 36;
  const titleH = 78;
  const summaryH = 86;
  const mentorH = 42;
  const rowH = 34;
  const footerH = 36;

  let height = titleH + summaryH + footerH + margin;
  let currentMentor = '';
  rows.forEach(r => {
    if(r.MENTOR !== currentMentor){ height += mentorH; currentMentor = r.MENTOR; }
    height += rowH;
  });
  height = Math.max(height, 420);

  const canvas = document.createElement('canvas');
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  ctx.fillStyle = '#0C2340';
  ctx.fillRect(0, 0, width, titleH);
  ctx.fillStyle = '#ffffff';
  ctx.font = 'bold 28px Segoe UI, Arial, sans-serif';
  ctx.fillText('Detalle asesor por MENTOR', margin, 46);
  ctx.font = '14px Segoe UI, Arial, sans-serif';
  ctx.fillText('Base Sin Gestión', width - 180, 46);

  const cardW = (width - margin*2 - 28) / 3;
  const cardY = titleH + 24;
  const cards = [
    ['Leads distribuidos', totalLeads.toLocaleString()],
    ['Asesores con leads', totalAsesores.toLocaleString()],
    ['Mentores', totalMentores.toLocaleString()]
  ];
  cards.forEach(([label, value], i) => {
    const x = margin + i * (cardW + 14);
    ctx.fillStyle = '#f4f8fb';
    ctx.fillRect(x, cardY, cardW, 62);
    ctx.strokeStyle = '#d1d9e0';
    ctx.strokeRect(x, cardY, cardW, 62);
    ctx.fillStyle = '#0C2340';
    ctx.font = 'bold 24px Segoe UI, Arial, sans-serif';
    ctx.fillText(value, x + 18, cardY + 30);
    ctx.fillStyle = '#667085';
    ctx.font = '13px Segoe UI, Arial, sans-serif';
    ctx.fillText(label, x + 18, cardY + 50);
  });

  function drawText(text, x, y, maxWidth){
    text = String(text || '');
    if(ctx.measureText(text).width <= maxWidth){ ctx.fillText(text, x, y); return; }
    let out = text;
    while(out.length && ctx.measureText(out + '…').width > maxWidth){ out = out.slice(0, -1); }
    ctx.fillText(out + '…', x, y);
  }

  let y = titleH + summaryH + 24;
  currentMentor = '';
  rows.forEach((r, idx) => {
    if(r.MENTOR !== currentMentor){
      currentMentor = r.MENTOR;
      const mentorTotal = rows.filter(x => x.MENTOR === currentMentor).reduce((a,b)=>a+Number(b['Cantidad leads']||0),0);
      ctx.fillStyle = '#1B365D';
      ctx.fillRect(margin, y, width - margin*2, 34);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
      drawText('MENTOR: ' + currentMentor, margin + 14, y + 22, 760);
      ctx.font = 'bold 14px Segoe UI, Arial, sans-serif';
      ctx.fillText(mentorTotal.toLocaleString() + ' leads', width - margin - 110, y + 22);
      y += mentorH;
    }

    ctx.fillStyle = idx % 2 ? '#ffffff' : '#f8fafc';
    ctx.fillRect(margin, y - 5, width - margin*2, rowH);
    ctx.strokeStyle = '#edf1f5';
    ctx.beginPath();
    ctx.moveTo(margin, y + rowH - 5);
    ctx.lineTo(width - margin, y + rowH - 5);
    ctx.stroke();

    ctx.fillStyle = '#2c3e50';
    ctx.font = '14px Segoe UI, Arial, sans-serif';
    drawText(r['CORREO ASESOR ASIGNADO'], margin + 16, y + 17, 860);
    ctx.fillStyle = '#0C2340';
    ctx.font = 'bold 16px Segoe UI, Arial, sans-serif';
    ctx.fillText(String(r['Cantidad leads']).toLocaleString(), width - margin - 80, y + 18);
    y += rowH;
  });

  ctx.fillStyle = '#898D8D';
  ctx.font = '12px Segoe UI, Arial, sans-serif';
  ctx.fillText('Generado desde App Normalizador Contact CUN', margin, height - 18);

  const link = document.createElement('a');
  link.href = canvas.toDataURL('image/png');
  link.download = 'Detalle_Asesor_Por_Mentor_Sin_Gestion_' + Date.now() + '.png';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('📸 Detalle asesor exportado en PNG');
}

function pintarDetallePorMentorSG(data){
  const card = document.getElementById('sg-detalle-asesor-card');
  const box = document.getElementById('sg-detalle-asesor');
  if(!box) return;
  if(card) card.style.display = 'block';

  if(!data || !data.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por mentor y asesor.</div>';
    return;
  }

  const grouped = {};
  data.forEach(r=>{
    const asesor = String(r['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
    const mentor = mentorAsignadoFinal(asesor, r) || 'Sin mentor';
    if(!grouped[mentor]) grouped[mentor] = {};
    grouped[mentor][asesor] = (grouped[mentor][asesor] || 0) + 1;
  });

  const rowsExport = getDetalleMentorRowsSG(data);
  window.__detalleSimpleExport = window.__detalleSimpleExport || {};
  window.__detalleSimpleExport['sg-detalle-asesor'] = { rows: rowsExport, exportName: 'Detalle_Asesor_Por_Mentor_Sin_Gestion' };

  const totalLeads = data.length;
  const totalAsesores = new Set(data.map(r => r['CORREO ASESOR ASIGNADO'] || 'Sin asesor')).size;
  const totalMentores = Object.keys(grouped).length;

  box.innerHTML = `
    <div class="detalle-actions"><button class="btn btn-green" onclick="exportDetalleAsesorPNG()">📸 Exportar detalle asesor PNG</button></div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalMentores.toLocaleString()}</strong><span>Mentores</span></div>
    </div>
    ${Object.entries(grouped).sort((a,b)=>a[0].localeCompare(b[0])).map(([mentor, asesores])=>{
      const total = Object.values(asesores).reduce((a,b)=>a+b,0);
      const cards = Object.entries(asesores).sort((a,b)=>b[1]-a[1]).map(([asesor,c]) =>
        `<div class="detalle-asesor-item"><span><b>${escapeHtml(asesor)}</b></span><strong>${c.toLocaleString()}</strong></div>`
      ).join('');
      return `<div class="detalle-mentor-card">
        <div class="detalle-mentor-head"><span>👤 MENTOR: ${escapeHtml(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
        <div class="detalle-asesor-grid">${cards}</div>
      </div>`;
    }).join('')}
  `;
}

function renderDetalleAsesorSG(){
  pintarDetallePorMentorSG(typeof distribuidoData !== 'undefined' ? distribuidoData : []);
}
