/* =========================================================
   AJUSTES FINALES — INTERESADOS Y NO CONTACTADO
   1) Resumen por asesor original en Interesados y No Contactado.
   2) Distribución visual igual a Base Sin Gestión: tabla simple + detalle por MENTOR.
   3) Exportación del detalle asesor en JPG.
   ========================================================= */
(function(){
  function txt(v=''){
    return String(v ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#039;');
  }

  function email(v=''){
    const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    return m ? m[0].toLowerCase() : '';
  }

  function norm(v=''){
    return String(v || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g,'')
      .trim()
      .toLowerCase();
  }

  function mentorPorAsesor(asesor, row){
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

  function distRow(row){
    return {
      'ID de registro': String(row?.['ID de registro'] || '').trim(),
      'CORREO ASESOR ASIGNADO': String(row?.['CORREO ASESOR ASIGNADO'] || '').trim()
    };
  }

  function detalleRowsPorMentor(data){
    const grouped = {};
    (data || []).forEach(r => {
      const asesor = String(r?.['CORREO ASESOR ASIGNADO'] || '').trim() || 'Sin asesor';
      const mentor = mentorPorAsesor(asesor, r);
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

  function renderDetalleGenerico(opts){
    const { data, cardId, boxId, emptyText, exportFnName } = opts;
    const card = document.getElementById(cardId);
    const box = document.getElementById(boxId);
    if(!box) return;
    if(card){
      card.style.display = 'block';
      card.classList.add('detalle-visible');
    }

    const rows = detalleRowsPorMentor(data);
    if(!rows.length){
      box.innerHTML = `<div class="detalle-empty">${txt(emptyText || 'Primero distribuye los leads para ver el detalle por mentor y asesor.')}</div>`;
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
      <div class="detalle-actions">
        <button class="btn btn-green" onclick="${exportFnName}()">📸 Exportar detalle asesor JPG</button>
      </div>
      <div class="detalle-summary">
        <div><strong>${totalLeads.toLocaleString()}</strong><span>Leads distribuidos</span></div>
        <div><strong>${totalAsesores.toLocaleString()}</strong><span>Asesores con leads</span></div>
        <div><strong>${totalMentores.toLocaleString()}</strong><span>Mentores</span></div>
      </div>
      ${Object.entries(grouped).map(([mentor, asesores]) => {
        const total = asesores.reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
        return `<div class="detalle-mentor-card">
          <div class="detalle-mentor-head"><span>👤 MENTOR: ${txt(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
          <div class="detalle-asesor-grid">
            ${asesores.map(r => `<div class="detalle-asesor-item"><span><b>${txt(r.Asesor)}</b></span><strong>${Number(r['Cantidad leads']).toLocaleString()}</strong></div>`).join('')}
          </div>
        </div>`;
      }).join('')}
    `;
  }

  function exportDetalleJPG(data, modulo, filePrefix){
    const rows = detalleRowsPorMentor(data);
    if(!rows.length){ showToast('Primero distribuye los leads para exportar el detalle.'); return; }

    const totalLeads = (data || []).length;
    const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
    const totalMentores = new Set(rows.map(r => r.MENTOR)).size;

    const generated = new Date().toLocaleString('es-CO', {
      year:'numeric', month:'numeric', day:'numeric',
      hour:'numeric', minute:'2-digit', second:'2-digit'
    });

    const scale = 4;
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
    ctx.fillText(modulo + '  ·  Generado: ' + generated, margin, 80);

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
    link.href = canvas.toDataURL('image/jpeg', 0.98);
    link.download = filePrefix + '_' + Date.now() + '.jpg';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('📸 Detalle asesor exportado en JPG');
  }

  function renderDistPreview(data, wrapId, tableId){
    const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
    const rows = (data || []).slice(0,300).map(distRow);
    const wrap = document.getElementById(wrapId);
    const tbl = document.getElementById(tableId);
    if(!wrap || !tbl) return;

    tbl.classList.add('dist-simple-table');
    tbl.innerHTML = '<thead><tr>' + cols.map(c => `<th>${txt(c)}</th>`).join('') + '</tr></thead><tbody>' +
      rows.map(r => '<tr>' + cols.map(c => `<td title="${txt(r[c])}">${txt(r[c])}</td>`).join('') + '</tr>').join('') +
      '</tbody>';
    wrap.style.display = 'block';

    let note = wrap.querySelector('.dist-preview-note');
    if(!note){
      note = document.createElement('div');
      note.className = 'dist-preview-note';
      wrap.appendChild(note);
    }
    note.textContent = `Mostrando ${Math.min(300, (data || []).length).toLocaleString()} de ${(data || []).length.toLocaleString()} · exporta para ver todo`;
  }

  function showStats(data, boxId){
    const box = document.getElementById(boxId);
    if(!box) return;
    const counts = {};
    (data || []).forEach(r => {
      const a = r['CORREO ASESOR ASIGNADO'] || 'Sin asesor';
      counts[a] = (counts[a] || 0) + 1;
    });
    box.style.display = 'block';
    box.innerHTML = `<strong>📊 ${(data || []).length.toLocaleString()} leads distribuidos:</strong><br><br>` +
      Object.entries(counts)
        .sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .map(([a,c]) => `<span class="dist-chip">${txt(a)}: <b>${Number(c).toLocaleString()}</b></span>`)
        .join('');
  }

  function exportDistExcel(data, filename){
    if(!(data || []).length){ showToast('Primero distribuye los leads'); return; }
    const cols = ['ID de registro', 'CORREO ASESOR ASIGNADO'];
    const rows = (data || []).map(distRow);
    const ws = XLSX.utils.json_to_sheet(rows, { header: cols });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Distribucion');
    XLSX.writeFile(wb, filename + '.xlsx');
    showToast(`⬇ ${rows.length.toLocaleString()} asignaciones exportadas`);
  }

  function ensureOwnerSummaryWrap(module){
    const config = {
      inter: { wrapId:'inter-owner-summary-wrap', parentId:'inter-tp-1', beforeId:'inter-filter-panel' },
      nc: { wrapId:'nc-owner-summary-wrap', parentId:'nc-tp-1', beforeSelector:'.filter-card' }
    }[module];
    if(!config) return null;

    let wrap = document.getElementById(config.wrapId);
    if(wrap) return wrap;

    const parent = document.getElementById(config.parentId);
    if(!parent) return null;

    wrap = document.createElement('div');
    wrap.id = config.wrapId;
    wrap.className = 'sg-owner-summary-wrap';
    wrap.style.display = 'none';

    const before = config.beforeId ? document.getElementById(config.beforeId) : parent.querySelector(config.beforeSelector);
    if(before) parent.insertBefore(wrap, before);
    else parent.prepend(wrap);
    return wrap;
  }

  function renderResumenAsesorOriginal(module, data){
    const wrap = ensureOwnerSummaryWrap(module);
    if(!wrap) return;
    if(!(data || []).length){
      wrap.style.display = 'none';
      return;
    }

    const counts = {};
    (data || []).forEach(r => {
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
            <tbody>${rows.map(r => `<tr><td title="${txt(r.asesor)}">${txt(r.asesor)}</td><td>${r.cantidad.toLocaleString()}</td></tr>`).join('')}</tbody>
          </table>
        </div>
      </div>
    `;
  }

  window.renderDetalleAsesorInter = function(){
    renderDetalleGenerico({
      data: (typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []),
      cardId: 'inter-detalle-asesor-card',
      boxId: 'inter-detalle-asesor',
      exportFnName: 'exportDetalleAsesorInterJPG',
      emptyText: 'Primero distribuye los leads para ver el detalle por mentor y asesor.'
    });
  };

  window.renderDetalleAsesorNC = function(){
    renderDetalleGenerico({
      data: (typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []),
      cardId: 'nc-detalle-asesor-card',
      boxId: 'nc-detalle-asesor',
      exportFnName: 'exportDetalleAsesorNCJPG',
      emptyText: 'Primero distribuye los leads para ver el detalle por mentor y asesor.'
    });
  };

  window.exportDetalleAsesorInterJPG = function(){
    exportDetalleJPG((typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []), 'Base Interesados', 'Detalle_Asesor_Por_Mentor_Interesados');
  };

  window.exportDetalleAsesorNCJPG = function(){
    exportDetalleJPG((typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []), 'Base No Contactado', 'Detalle_Asesor_Por_Mentor_No_Contactado');
  };

  window.renderInterDistPreview = function(){
    renderDistPreview((typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []), 'inter-dist-preview', 'inter-dist-tbl');
  };

  window.renderNCDistPreview = function(){
    renderDistPreview((typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []), 'nc-dist-preview', 'nc-dist-tbl');
  };

  window.showInterDistStats = function(){
    showStats((typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []), 'inter-stat-box');
  };

  window.showNCDistStats = function(){
    showStats((typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []), 'nc-stat-box');
  };

  window.exportarDistInteresados = function(){
    const fn = 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' INTER' + (typeof getInterFileLabel==='function'?getInterFileLabel():'');
    exportDistExcel((typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : []), fn);
  };

  window.exportarDistNC = function(){
    const fn = 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' NOCONTAC' + (typeof getNCFileLabel==='function'?getNCFileLabel():'');
    exportDistExcel((typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : []), fn);
  };

  const originalDistribuirInter = window.distribuirInteresados;
  window.distribuirInteresados = function(){
    if(typeof originalDistribuirInter === 'function') originalDistribuirInter();
    setTimeout(() => {
      window.renderInterDistPreview();
      window.showInterDistStats();
      window.renderDetalleAsesorInter();
    }, 80);
  };

  const originalDistribuirNC = window.distribuirNC;
  window.distribuirNC = function(){
    if(typeof originalDistribuirNC === 'function') originalDistribuirNC();
    setTimeout(() => {
      window.renderNCDistPreview();
      window.showNCDistStats();
      window.renderDetalleAsesorNC();
    }, 80);
  };

  if(typeof renderInterTable === 'function'){
    const originalRenderInterTable = renderInterTable;
    window.renderInterTable = renderInterTable = function(){
      originalRenderInterTable();
      renderResumenAsesorOriginal('inter', (typeof interFiltered !== 'undefined' ? interFiltered : (typeof interData !== 'undefined' ? interData : [])));
    };
  }

  if(typeof renderNCTable === 'function'){
    const originalRenderNCTable = renderNCTable;
    window.renderNCTable = renderNCTable = function(){
      originalRenderNCTable();
      renderResumenAsesorOriginal('nc', (typeof ncFiltered !== 'undefined' ? ncFiltered : (typeof ncData !== 'undefined' ? ncData : [])));
    };
  }

  window.renderResumenAsesorOriginalInter = function(){
    renderResumenAsesorOriginal('inter', (typeof interFiltered !== 'undefined' ? interFiltered : (typeof interData !== 'undefined' ? interData : [])));
  };

  window.renderResumenAsesorOriginalNC = function(){
    renderResumenAsesorOriginal('nc', (typeof ncFiltered !== 'undefined' ? ncFiltered : (typeof ncData !== 'undefined' ? ncData : [])));
  };
})();
