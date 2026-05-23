/* =========================================================
   UNIFORMIDAD FINAL — SIN GESTIÓN / INTERESADOS / NO CONTACTADO
   - En Distribuir y exportar quedan solo:
     ⚡ Distribuir | ⬇ Exportar distribución | 📋 Copiar IDs distribución
   - El botón de imagen queda solo dentro de la sección 4.
   - Base Sin Gestión muestra el mismo detalle visual por mentor que Interesados y No Contactado.
   - Base Sin Gestión exporta el detalle como JPG.
   ========================================================= */
(function(){
  function safe(v=''){
    return String(v ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#039;');
  }

  function norm(v=''){
    return String(v || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g,'')
      .trim()
      .toLowerCase();
  }

  function email(v=''){
    const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    return m ? m[0].toLowerCase() : '';
  }

  function mentorFinal(asesor, row){
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

    const em = email(asesor);
    const no = norm(asesor);
    const maps = [
      window.SUP_DIRECT_MAP,
      window.SUP_DOM_MAP,
      window.SUP_FINAL_MAP,
      window.SUP_SIMPLE_MAP,
      window.SUPERVISOR_MAP_FINAL,
      window.MAP_SUPERVISOR
    ].filter(Boolean);

    for(const map of maps){
      const info = (em && map[em]) || (no && map[no]) || map[asesor];
      if(!info) continue;
      if(typeof info === 'string') return info;
      if(info.supervisor) return info.supervisor;
      if(info.mentor) return info.mentor;
      if(info.grupo) return info.grupo;
    }
    return 'Sin mentor';
  }

  function getId(row){
    if(typeof getLeadId === 'function') return getLeadId(row);
    return String(row?.['ID de registro'] || row?.ID || '').trim();
  }

  function rowsDetalle(data){
    const grouped = {};
    (data || []).forEach(r => {
      const asesor = String(r?.['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
      const mentor = mentorFinal(asesor, r) || 'Sin mentor';
      const key = mentor + '||' + asesor;
      if(!grouped[key]) grouped[key] = { MENTOR: mentor, Asesor: asesor, 'Cantidad leads': 0 };
      grouped[key]['Cantidad leads'] += 1;
    });

    return Object.values(grouped).sort((a,b) =>
      String(a.MENTOR).localeCompare(String(b.MENTOR)) ||
      Number(b['Cantidad leads']) - Number(a['Cantidad leads']) ||
      String(a.Asesor).localeCompare(String(b.Asesor))
    );
  }

  function renderDetalleVisual(data, boxId, emptyText){
    const box = document.getElementById(boxId);
    if(!box) return;

    const rows = rowsDetalle(data);
    if(!rows.length){
      box.innerHTML = `<div class="detalle-empty">${safe(emptyText || 'Primero distribuye los leads para ver el detalle por mentor y asesor.')}</div>`;
      return;
    }

    const totalLeads = (data || []).length;
    const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
    const totalMentores = new Set(rows.map(r => r.MENTOR)).size;
    const grouped = {};
    rows.forEach(r => {
      if(!grouped[r.MENTOR]) grouped[r.MENTOR] = [];
      grouped[r.MENTOR].push(r);
    });

    box.innerHTML = `
      <div class="detalle-summary">
        <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
        <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
        <div><strong>${totalMentores.toLocaleString()}</strong><span>Mentores</span></div>
      </div>
      ${Object.entries(grouped).map(([mentor, asesores]) => {
        const total = asesores.reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
        return `<div class="detalle-mentor-card">
          <div class="detalle-mentor-head"><span>👤 MENTOR: ${safe(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
          <div class="detalle-asesor-grid">
            ${asesores.map(r => `<div class="detalle-asesor-item"><span><b>${safe(r.Asesor)}</b></span><strong>${Number(r['Cantidad leads']).toLocaleString()}</strong></div>`).join('')}
          </div>
        </div>`;
      }).join('')}
    `;
  }

  function exportDetalleJPG(data, modulo, filePrefix){
    const rows = rowsDetalle(data);
    if(!rows.length){ showToast('Primero distribuye los leads para exportar el detalle.'); return; }

    const totalLeads = (data || []).length;
    const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
    const totalMentores = new Set(rows.map(r => r.MENTOR)).size;

    const scale = 2;
    const width = 1200;
    const margin = 36;
    const titleH = 78;
    const summaryH = 92;
    const mentorH = 44;
    const rowH = 34;
    const footerH = 46;
    let height = titleH + summaryH + footerH + margin;
    let last = '';
    rows.forEach(r => {
      if(r.MENTOR !== last){ height += mentorH; last = r.MENTOR; }
      height += rowH;
    });
    height = Math.max(440, height);

    const canvas = document.createElement('canvas');
    canvas.width = width * scale;
    canvas.height = height * scale;
    const ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0,0,width,height);
    ctx.fillStyle = '#0C2340';
    ctx.fillRect(0,0,width,titleH);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 28px Segoe UI, Arial, sans-serif';
    ctx.fillText('Detalle asesor por MENTOR', margin, 46);
    ctx.font = '14px Segoe UI, Arial, sans-serif';
    ctx.fillText(modulo, width - 230, 46);

    const cardW = (width - margin*2 - 28) / 3;
    const cardY = titleH + 24;
    [['Leads distribuidos', totalLeads], ['Asesores con leads', totalAsesores], ['Mentores', totalMentores]].forEach(([label, value], i) => {
      const x = margin + i * (cardW + 14);
      ctx.fillStyle = '#f4f8fb';
      ctx.fillRect(x, cardY, cardW, 64);
      ctx.strokeStyle = '#d1d9e0';
      ctx.strokeRect(x, cardY, cardW, 64);
      ctx.fillStyle = '#0C2340';
      ctx.font = 'bold 24px Segoe UI, Arial, sans-serif';
      ctx.fillText(Number(value).toLocaleString(), x + 18, cardY + 31);
      ctx.fillStyle = '#667085';
      ctx.font = '13px Segoe UI, Arial, sans-serif';
      ctx.fillText(label, x + 18, cardY + 52);
    });

    function cut(text, x, y, maxW){
      text = String(text || '');
      if(ctx.measureText(text).width <= maxW){ ctx.fillText(text, x, y); return; }
      let out = text;
      while(out.length && ctx.measureText(out + '…').width > maxW) out = out.slice(0, -1);
      ctx.fillText(out + '…', x, y);
    }

    let y = titleH + summaryH + 24;
    last = '';
    rows.forEach((r, i) => {
      if(r.MENTOR !== last){
        last = r.MENTOR;
        const mentorTotal = rows.filter(x => x.MENTOR === last).reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
        ctx.fillStyle = '#1B365D';
        ctx.fillRect(margin, y, width - margin*2, 34);
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
        cut('MENTOR: ' + last, margin + 14, y + 22, 820);
        ctx.font = 'bold 14px Segoe UI, Arial, sans-serif';
        ctx.fillText(mentorTotal.toLocaleString() + ' leads', width - margin - 115, y + 22);
        y += mentorH;
      }
      ctx.fillStyle = i % 2 ? '#ffffff' : '#f8fafc';
      ctx.fillRect(margin, y - 5, width - margin*2, rowH);
      ctx.strokeStyle = '#edf1f5';
      ctx.beginPath();
      ctx.moveTo(margin, y + rowH - 5);
      ctx.lineTo(width - margin, y + rowH - 5);
      ctx.stroke();
      ctx.fillStyle = '#2c3e50';
      ctx.font = '14px Segoe UI, Arial, sans-serif';
      cut(r.Asesor, margin + 16, y + 17, 870);
      ctx.fillStyle = '#0C2340';
      ctx.font = 'bold 16px Segoe UI, Arial, sans-serif';
      ctx.fillText(Number(r['Cantidad leads']).toLocaleString(), width - margin - 78, y + 18);
      y += rowH;
    });

    ctx.fillStyle = '#898D8D';
    ctx.font = '12px Segoe UI, Arial, sans-serif';
    ctx.fillText('Generado desde App Normalizador Contact CUN', margin, height - 20);

    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/jpeg', 0.95);
    link.download = filePrefix + '_' + Date.now() + '.jpg';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('📸 Detalle asesor exportado en JPG');
  }

  function renderPreviewSG(){
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    const prev = document.getElementById('dist-preview');
    if(!prev) return;
    const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
    const rows = (data || []).slice(0,300).map(r => ({
      'ID de registro': getId(r),
      'CORREO ASESOR ASIGNADO': String(r?.['CORREO ASESOR ASIGNADO'] || '').trim()
    }));

    prev.style.display = 'block';
    prev.innerHTML = `
      <table id="dist-tbl" class="dist-simple-table">
        <thead><tr>${cols.map(c => `<th>${safe(c)}</th>`).join('')}</tr></thead>
        <tbody>${rows.map(r => `<tr>${cols.map(c => `<td title="${safe(r[c])}">${safe(r[c])}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>
      <div class="dist-preview-note">Mostrando ${Math.min(300, data.length).toLocaleString()} de ${data.length.toLocaleString()} · exporta para ver todo</div>
    `;
  }

  function copiarIdsSG(){
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    if(!data.length){ showToast('Primero distribuye los leads'); return; }
    navigator.clipboard.writeText(data.map(r => getId(r)).join('\n'))
      .then(() => showToast(`📋 ${data.length.toLocaleString()} IDs de distribución copiados`));
  }

  function exportarDistSGFinal(){
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    if(!data.length){ showToast('Primero distribuye los leads'); return; }
    const rows = data.map(r => ({
      'ID de registro': getId(r),
      'CORREO ASESOR ASIGNADO': String(r?.['CORREO ASESOR ASIGNADO'] || '').trim()
    }));
    const ws = XLSX.utils.json_to_sheet(rows, { header: ['ID de registro', 'CORREO ASESOR ASIGNADO'] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Distribuido');
    XLSX.writeFile(wb, 'Leads_Distribuidos_ID_Asesor_' + Date.now() + '.xlsx');
    showToast(`⬇ ${rows.length.toLocaleString()} leads exportados`);
  }

  function limpiarBotonesDistribucion(){
    const configs = [
      { panel:'#tp-1', copiar:'copiarIdsDistribucionSGFinal()' },
      { panel:'#inter-tp-2', copiar:'copiarInterDistIDs()' },
      { panel:'#nc-tp-2', copiar:'copiarNCDistIDs()' }
    ];

    configs.forEach(cfg => {
      const panel = document.querySelector(cfg.panel);
      if(!panel) return;
      const cards = [...panel.querySelectorAll('.card2')];
      const card = cards.find(c => /Distribuir y exportar/i.test(c.textContent || ''));
      const acts = card ? card.querySelector('.acts') : null;
      if(!acts) return;

      [...acts.querySelectorAll('button')].forEach(btn => {
        const text = btn.textContent || '';
        if(/Exportar detalle/i.test(text)) btn.remove();
      });

      const hasCopy = [...acts.querySelectorAll('button')].some(btn => /Copiar IDs distribución/i.test(btn.textContent || ''));
      if(!hasCopy){
        const btn = document.createElement('button');
        btn.className = 'btn btn-orange';
        btn.setAttribute('onclick', cfg.copiar);
        btn.textContent = '📋 Copiar IDs distribución';
        acts.appendChild(btn);
      }
    });
  }

  function asegurarBotonDetalleSG(){
    const card = document.getElementById('sg-detalle-asesor-card');
    if(!card) return;
    let btn = document.getElementById('btn-export-detalle-asesor-jpg-sg');
    const viejoPng = document.getElementById('btn-export-detalle-asesor-png');
    if(viejoPng) viejoPng.closest('.detalle-png-actions, .acts, div')?.remove();
    if(!btn){
      const p = card.querySelector('p');
      const wrap = document.createElement('div');
      wrap.className = 'acts detalle-png-actions';
      wrap.style.marginTop = '10px';
      wrap.innerHTML = `<button id="btn-export-detalle-asesor-jpg-sg" class="btn btn-green" onclick="exportDetalleAsesorSGJPG()">📸 Exportar detalle asesor JPG</button>`;
      if(p) p.insertAdjacentElement('afterend', wrap);
      else card.insertAdjacentElement('afterbegin', wrap);
    }
  }

  window.copiarIdsDistribucionSGFinal = copiarIdsSG;
  window.exportarDist = exportarDistSGFinal;

  window.renderDetalleAsesorSG = function(){
    const card = document.getElementById('sg-detalle-asesor-card');
    if(card){
      card.style.display = 'block';
      card.classList.add('detalle-visible');
    }
    asegurarBotonDetalleSG();
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    renderDetalleVisual(data, 'sg-detalle-asesor', 'Primero distribuye los leads para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorSGJPG = function(){
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    exportDetalleJPG(data, 'Base Sin Gestión', 'Detalle_Asesor_Por_Mentor_Sin_Gestion');
  };

  const oldDistribuir = window.distribuir;
  window.distribuir = function(){
    if(typeof oldDistribuir === 'function') oldDistribuir.apply(this, arguments);
    [80, 250, 600].forEach(ms => setTimeout(() => {
      renderPreviewSG();
      window.renderDetalleAsesorSG();
      limpiarBotonesDistribucion();
    }, ms));
  };

  const oldInter = window.distribuirInteresados;
  window.distribuirInteresados = function(){
    if(typeof oldInter === 'function') oldInter.apply(this, arguments);
    [80, 250, 600].forEach(ms => setTimeout(limpiarBotonesDistribucion, ms));
  };

  const oldNC = window.distribuirNC;
  window.distribuirNC = function(){
    if(typeof oldNC === 'function') oldNC.apply(this, arguments);
    [80, 250, 600].forEach(ms => setTimeout(limpiarBotonesDistribucion, ms));
  };

  document.addEventListener('DOMContentLoaded', () => {
    [250, 800, 1500].forEach(ms => setTimeout(() => {
      limpiarBotonesDistribucion();
      asegurarBotonDetalleSG();
      if(typeof window.renderDetalleAsesorSG === 'function') window.renderDetalleAsesorSG();
    }, ms));
  });
})();
