/* =========================================================
   AJUSTES FINALES — BASE SIN GESTIÓN
   1) Tabla Distribución: solo ID de registro + CORREO ASESOR ASIGNADO.
   2) Detalle asesor: agrupado por MENTOR.
   3) Botón: Exportar detalle asesor PNG.
   4) Resumen superior en Filtros y Normalización:
      conteo por Propietario de Posible Cliente original, mostrado como Asesor.
   ========================================================= */

function sgTextoSeguro(v=''){
  return String(v ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

function sgEmail(v=''){
  const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].toLowerCase() : '';
}

function sgNorm(v=''){
  return String(v || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .trim()
    .toLowerCase();
}

function sgMentorPorAsesor(asesor, row){
  const directo = String(row?.MENTOR || row?.SUPERVISOR || '').trim();
  if(directo) return directo;

  if(typeof mentorAsignadoFinal === 'function'){
    const m = String(mentorAsignadoFinal(asesor, row) || '').trim();
    if(m) return m;
  }
  if(typeof mentorDeAsesor === 'function'){
    const m = String(mentorDeAsesor(asesor) || '').trim();
    if(m) return m;
  }

  // Fallback con los mapas que existen en diferentes versiones del proyecto.
  const email = sgEmail(asesor);
  const norm = sgNorm(asesor);
  const maps = [
    window.SUP_DIRECT_MAP,
    window.SUP_DOM_MAP,
    window.SUP_FINAL_MAP,
    window.SUP_SIMPLE_MAP,
    window.SUPERVISOR_MAP_FINAL,
    window.MAP_SUPERVISOR
  ].filter(Boolean);

  for(const map of maps){
    const info = (email && map[email]) || (norm && map[norm]) || map[asesor];
    if(!info) continue;
    if(typeof info === 'string') return info;
    if(info.supervisor) return info.supervisor;
    if(info.mentor) return info.mentor;
    if(info.grupo) return info.grupo;
  }
  return 'Sin mentor';
}

function sgDistRow(row){
  return {
    'ID de registro': typeof getLeadId === 'function' ? getLeadId(row) : String(row?.['ID de registro'] || '').trim(),
    'CORREO ASESOR ASIGNADO': String(row?.['CORREO ASESOR ASIGNADO'] || '').trim()
  };
}

function renderDistribucionPreview(){
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
  const rows = data.slice(0,300).map(sgDistRow);
  const prev = document.getElementById('dist-preview');
  if(!prev) return;

  const head = '<thead><tr>' + cols.map(c => `<th>${sgTextoSeguro(c)}</th>`).join('') + '</tr></thead>';
  const body = '<tbody>' + rows.map(r =>
    '<tr>' + cols.map(c => `<td title="${sgTextoSeguro(r[c])}">${sgTextoSeguro(r[c])}</td>`).join('') + '</tr>'
  ).join('') + '</tbody>';

  prev.style.display = 'block';
  prev.innerHTML = `
    <table id="dist-tbl" class="dist-simple-table">${head}${body}</table>
    <div class="dist-preview-note">Mostrando ${Math.min(300,data.length).toLocaleString()} de ${data.length.toLocaleString()} · exporta para ver todo</div>
  `;
}

function exportarDist(){
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  if(!data.length){ showToast('Primero distribuye los leads'); return; }

  const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
  const rows = data.map(sgDistRow);
  const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Distribuido');
  XLSX.writeFile(wb, 'Leads_Distribuidos_ID_Asesor_' + Date.now() + '.xlsx');
  showToast(`⬇ ${rows.length.toLocaleString()} leads exportados`);
}

function sgDetalleRowsPorMentor(data){
  const grouped = {};
  (data || []).forEach(r => {
    const asesor = String(r['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
    const mentor = sgMentorPorAsesor(asesor, r) || 'Sin mentor';
    if(!grouped[mentor]) grouped[mentor] = {};
    grouped[mentor][asesor] = (grouped[mentor][asesor] || 0) + 1;
  });

  const rows = [];
  Object.entries(grouped).forEach(([mentor, asesores]) => {
    Object.entries(asesores).forEach(([asesor, cantidad]) => {
      rows.push({ MENTOR: mentor, Asesor: asesor, 'Cantidad leads': cantidad });
    });
  });

  return rows.sort((a,b) =>
    String(a.MENTOR).localeCompare(String(b.MENTOR)) ||
    Number(b['Cantidad leads']) - Number(a['Cantidad leads']) ||
    String(a.Asesor).localeCompare(String(b.Asesor))
  );
}

function renderDetalleAsesorSG(){
  const card = document.getElementById('sg-detalle-asesor-card');
  const box = document.getElementById('sg-detalle-asesor');
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  if(!box) return;
  if(card){
    card.style.display = 'block';
    card.classList.add('detalle-visible');
  }

  const rows = sgDetalleRowsPorMentor(data);
  if(!rows.length){
    box.innerHTML = '<div class="detalle-empty">Primero distribuye los leads para ver el detalle por mentor y asesor.</div>';
    return;
  }

  const totalLeads = data.length;
  const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
  const totalMentores = new Set(rows.map(r => r.MENTOR)).size;
  const grouped = {};
  rows.forEach(r => {
    if(!grouped[r.MENTOR]) grouped[r.MENTOR] = [];
    grouped[r.MENTOR].push(r);
  });

  box.innerHTML = `
    <div class="detalle-actions">
      <button class="btn btn-green" onclick="exportDetalleAsesorPNG()">📸 Exportar detalle asesor PNG</button>
    </div>
    <div class="detalle-summary">
      <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
      <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
      <div><strong>${totalMentores.toLocaleString()}</strong><span>Mentores</span></div>
    </div>
    ${Object.entries(grouped).map(([mentor, asesores]) => {
      const total = asesores.reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
      return `<div class="detalle-mentor-card">
        <div class="detalle-mentor-head"><span>👤 MENTOR: ${sgTextoSeguro(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
        <div class="detalle-asesor-grid">
          ${asesores.map(r => `<div class="detalle-asesor-item"><span><b>${sgTextoSeguro(r.Asesor)}</b></span><strong>${Number(r['Cantidad leads']).toLocaleString()}</strong></div>`).join('')}
        </div>
      </div>`;
    }).join('')}
  `;
}

function exportDetalleAsesorPNG(){
  const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
  const rows = sgDetalleRowsPorMentor(data);
  if(!rows.length){ showToast('No hay detalle asesor para exportar.'); return; }

  const totalLeads = data.length;
  const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
  const totalMentores = new Set(rows.map(r => r.MENTOR)).size;

  const generated = new Date().toLocaleString('es-CO', {
    year:'numeric', month:'numeric', day:'numeric',
    hour:'numeric', minute:'2-digit', second:'2-digit'
  });

  const scale = 2;
  const width = 1200;
  const margin = 36;
  const titleH = 106;
  const summaryH = 92;
  const mentorH = 44;
  const rowH = 34;
  const footerH = 40;
  let height = titleH + summaryH + footerH + margin;
  let lastMentor = '';
  rows.forEach(r => {
    if(r.MENTOR !== lastMentor){ height += mentorH; lastMentor = r.MENTOR; }
    height += rowH;
  });
  height = Math.max(430, height);

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
  ctx.fillStyle = '#a8c4e0';
  ctx.fillText('Base Sin Gestion  ·  Generado: ' + generated, margin, 80);

  const cardW = (width - margin * 2 - 28) / 3;
  const cardY = titleH + 24;
  [
    ['Leads distribuidos', totalLeads.toLocaleString()],
    ['Asesores con leads', totalAsesores.toLocaleString()],
    ['Mentores', totalMentores.toLocaleString()]
  ].forEach(([label, value], i) => {
    const x = margin + i * (cardW + 14);
    ctx.fillStyle = '#f4f8fb';
    ctx.fillRect(x, cardY, cardW, 64);
    ctx.strokeStyle = '#d1d9e0';
    ctx.strokeRect(x, cardY, cardW, 64);
    ctx.fillStyle = '#0C2340';
    ctx.font = 'bold 24px Segoe UI, Arial, sans-serif';
    ctx.fillText(value, x + 18, cardY + 31);
    ctx.fillStyle = '#667085';
    ctx.font = '13px Segoe UI, Arial, sans-serif';
    ctx.fillText(label, x + 18, cardY + 52);
  });

  function cutText(text, x, y, maxWidth){
    text = String(text || '');
    if(ctx.measureText(text).width <= maxWidth){ ctx.fillText(text, x, y); return; }
    let out = text;
    while(out.length && ctx.measureText(out + '…').width > maxWidth) out = out.slice(0, -1);
    ctx.fillText(out + '…', x, y);
  }

  let y = titleH + summaryH + 24;
  lastMentor = '';
  rows.forEach((r, idx) => {
    if(r.MENTOR !== lastMentor){
      lastMentor = r.MENTOR;
      const mentorTotal = rows.filter(x => x.MENTOR === lastMentor).reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
      ctx.fillStyle = '#1B365D';
      ctx.fillRect(margin, y, width - margin * 2, 34);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
      cutText('MENTOR: ' + lastMentor, margin + 14, y + 22, 780);
      ctx.font = 'bold 14px Segoe UI, Arial, sans-serif';
      ctx.fillText(mentorTotal.toLocaleString() + ' leads', width - margin - 115, y + 22);
      y += mentorH;
    }

    ctx.fillStyle = idx % 2 ? '#ffffff' : '#f8fafc';
    ctx.fillRect(margin, y - 5, width - margin * 2, rowH);
    ctx.strokeStyle = '#edf1f5';
    ctx.beginPath();
    ctx.moveTo(margin, y + rowH - 5);
    ctx.lineTo(width - margin, y + rowH - 5);
    ctx.stroke();
    ctx.fillStyle = '#2c3e50';
    ctx.font = '14px Segoe UI, Arial, sans-serif';
    cutText(r.Asesor, margin + 16, y + 17, 860);
    ctx.fillStyle = '#0C2340';
    ctx.font = 'bold 16px Segoe UI, Arial, sans-serif';
    ctx.fillText(Number(r['Cantidad leads']).toLocaleString(), width - margin - 80, y + 18);
    y += rowH;
  });

  ctx.fillStyle = '#898D8D';
  ctx.font = '12px Segoe UI, Arial, sans-serif';
  ctx.fillText('Generado desde App Normalizador Contact CUN  ·  ' + generated, margin, height - 18);

  const link = document.createElement('a');
  link.href = canvas.toDataURL('image/jpeg', 0.95);
  link.download = 'Detalle_Asesor_Por_Mentor_Sin_Gestion_' + Date.now() + '.jpg';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('Detalle asesor exportado en JPG');
}

function sgRenderResumenPropietario(){
  const wrap = document.getElementById('sg-owner-summary-wrap');
  if(!wrap) return;
  const data = typeof allData !== 'undefined' ? allData : [];
  if(!data.length){
    wrap.style.display = 'none';
    return;
  }

  const counts = {};
  data.forEach(r => {
    const asesor = String(r['Propietario de Posible Cliente'] || '').trim() || 'Sin asesor';
    counts[asesor] = (counts[asesor] || 0) + 1;
  });

  const rows = Object.entries(counts)
    .map(([asesor, cantidad]) => ({ asesor, cantidad }))
    .sort((a,b) => b.cantidad - a.cantidad || a.asesor.localeCompare(b.asesor));

  const total = rows.reduce((a,b) => a + b.cantidad, 0);
  wrap.style.display = 'block';
  wrap.innerHTML = `
    <div class="owner-summary-card">
      <div class="owner-summary-head">
        <div>
          <h3>📊 Resumen por asesor original</h3>
          <p>Conteo directo de la columna original <strong>Propietario de Posible Cliente</strong>, sin normalizar.</p>
        </div>
        <div class="owner-summary-total"><strong>${total.toLocaleString()}</strong><span>leads</span></div>
      </div>
      <div class="owner-summary-table-wrap">
        <table class="owner-summary-table">
          <thead><tr><th>Asesor</th><th>Cantidad leads</th></tr></thead>
          <tbody>
            ${rows.map(r => `<tr><td title="${sgTextoSeguro(r.asesor)}">${sgTextoSeguro(r.asesor)}</td><td>${r.cantidad.toLocaleString()}</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

// Se engancha al render de la tabla principal para que aparezca justo después de cargar la base.
if(typeof renderTable === 'function' && !window.__sgRenderTableAjustado){
  window.__sgRenderTableAjustado = true;
  window.__renderTableOriginalSG = renderTable;
  renderTable = function(){
    window.__renderTableOriginalSG();
    sgRenderResumenPropietario();
  };
}

// Refuerzo: después de distribuir, siempre se actualizan tabla simple y detalle por mentor.
if(typeof distribuir === 'function' && !window.__sgDistribuirAjustado){
  window.__sgDistribuirAjustado = true;
  window.__distribuirOriginalSG = distribuir;
  distribuir = function(){
    window.__distribuirOriginalSG();
    setTimeout(() => {
      renderDistribucionPreview();
      renderDetalleAsesorSG();
    }, 50);
  };
}


