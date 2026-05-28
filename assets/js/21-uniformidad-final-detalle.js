/* =========================================================
   UNIFORMIDAD FINAL — SIN GESTIÓN / INTERESADOS / NO CONTACTADO / PE
   - Sección 4 "Detalle asesor por supervisor" funciona en todos los módulos.
   - El JPG se exporta desde el botón "📸 Exportar detalle asesor JPG" ya
     existente en el panel de distribución (limpiarBotonesDistribucion).
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
      .replace(/[̀-ͯ]/g,'')
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
      box.innerHTML = `<div class="detalle-empty">${safe(emptyText || 'Primero distribuye los leads para ver el detalle por supervisor y asesor.')}</div>`;
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
        <div><strong>${totalMentores.toLocaleString()}</strong><span>Supervisores</span></div>
      </div>
      ${Object.entries(grouped).map(([mentor, asesores]) => {
        const total = asesores.reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
        return `<div class="detalle-mentor-card">
          <div class="detalle-mentor-head"><span>👤 ${safe(mentor)}</span><b>${total.toLocaleString()} leads</b></div>
          <div class="detalle-asesor-grid">
            ${asesores.map(r => `<div class="detalle-asesor-item"><span><b>${safe(r.Asesor)}</b></span><strong>${Number(r['Cantidad leads']).toLocaleString()}</strong></div>`).join('')}
          </div>
        </div>`;
      }).join('')}
    `;
  }

  /* exportDetalleJPG: auto-pagina si el alto supera MAX_PAGE_H.
     Cada página se descarga como archivo separado (_P1, _P2, …).
     Si cabe todo en una página sale un solo archivo sin sufijo. */
  function exportDetalleJPG(data, modulo, filePrefix){
    const rows = rowsDetalle(data);
    if(!rows.length){ showToast('Primero distribuye los leads para exportar el detalle.'); return; }

    const totalLeads    = (data || []).length;
    const totalAsesores = new Set(rows.map(r => r.Asesor)).size;
    const totalMentores = new Set(rows.map(r => r.MENTOR)).size;

    const generated = new Date().toLocaleString('es-CO', {
      year:'numeric', month:'numeric', day:'numeric',
      hour:'numeric', minute:'2-digit', second:'2-digit'
    });

    const scale    = 4;
    const width    = 1200;
    const margin   = 36;
    const titleH   = 106;
    const summaryH = 92;
    const mentorH  = 44;
    const rowH     = 34;
    const footerH  = 46;
    const MAX_PAGE_H = 2400; // px lógicos máx por página

    /* ── Agrupar filas por mentor ── */
    const mentorMap = {};
    rows.forEach(r => {
      if(!mentorMap[r.MENTOR]) mentorMap[r.MENTOR] = [];
      mentorMap[r.MENTOR].push(r);
    });
    const mentorEntries = Object.entries(mentorMap); // [[name, rows[]], …]

    /* ── Dividir en páginas ── */
    const baseH = titleH + summaryH + footerH + margin + 24;
    const pages  = [];
    let curPage  = [];
    let curH     = baseH;

    mentorEntries.forEach(([mentor, mRows]) => {
      const groupH = mentorH + mRows.length * rowH;
      if(curPage.length > 0 && curH + groupH > MAX_PAGE_H){
        pages.push(curPage);
        curPage = [[mentor, mRows]];
        curH    = baseH + groupH;
      } else {
        curPage.push([mentor, mRows]);
        curH += groupH;
      }
    });
    if(curPage.length) pages.push(curPage);

    const multiPage = pages.length > 1;

    /* ── Renderizar y descargar cada página ── */
    function renderPage(pageGroups, pageIdx){
      /* calcular alto real de esta página */
      let height = baseH;
      pageGroups.forEach(([, mRows]) => { height += mentorH + mRows.length * rowH; });
      height = Math.max(440, height);

      const canvas = document.createElement('canvas');
      canvas.width  = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.scale(scale, scale);

      /* fondo */
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, width, height);

      /* cabecera */
      ctx.fillStyle = '#0C2340';
      ctx.fillRect(0, 0, width, titleH);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 28px Segoe UI, Arial, sans-serif';
      const pageLabel = multiPage ? `  (${pageIdx + 1} / ${pages.length})` : '';
      ctx.fillText('Detalle asesor por Mentor' + pageLabel, margin, 46);
      ctx.font = '14px Segoe UI, Arial, sans-serif';
      ctx.fillStyle = '#a8c4e0';
      ctx.fillText(modulo + '  ·  Generado: ' + generated, margin, 80);

      /* tarjetas resumen (solo en página 1) */
      if(pageIdx === 0){
        const cardW = (width - margin * 2 - 28) / 3;
        const cardY = titleH + 24;
        [['Leads distribuidos', totalLeads],
         ['Asesores con leads', totalAsesores],
         ['Mentores', totalMentores]
        ].forEach(([label, value], i) => {
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
      }

      function cut(text, x, y, maxW){
        text = String(text || '');
        if(ctx.measureText(text).width <= maxW){ ctx.fillText(text, x, y); return; }
        let out = text;
        while(out.length && ctx.measureText(out + '…').width > maxW) out = out.slice(0,-1);
        ctx.fillText(out + '…', x, y);
      }

      /* filas */
      let y = titleH + summaryH + 24;
      let rowIdx = 0;
      pageGroups.forEach(([mentor, mRows]) => {
        const mentorTotal = mRows.reduce((a,b) => a + Number(b['Cantidad leads'] || 0), 0);
        ctx.fillStyle = '#1B365D';
        ctx.fillRect(margin, y, width - margin * 2, 34);
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
        cut(mentor, margin + 14, y + 22, 820);
        ctx.font = 'bold 14px Segoe UI, Arial, sans-serif';
        ctx.fillText(mentorTotal.toLocaleString() + ' leads', width - margin - 115, y + 22);
        y += mentorH;

        mRows.forEach(r => {
          ctx.fillStyle = rowIdx % 2 ? '#ffffff' : '#f8fafc';
          ctx.fillRect(margin, y - 5, width - margin * 2, rowH);
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
          rowIdx++;
        });
      });

      /* pie */
      ctx.fillStyle = '#898D8D';
      ctx.font = '12px Segoe UI, Arial, sans-serif';
      ctx.fillText('Generado desde App Normalizador Contact CUN  ·  ' + generated, margin, height - 20);

      const suffix   = multiPage ? '_P' + (pageIdx + 1) : '';
      const link     = document.createElement('a');
      link.href      = canvas.toDataURL('image/jpeg', 0.98);
      link.download  = filePrefix + suffix + '.jpg';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }

    /* Descargas con pequeño delay entre páginas para no saturar el browser */
    pages.forEach((pg, idx) => setTimeout(() => renderPage(pg, idx), idx * 400));

    const msg = multiPage
      ? `📸 Exportando ${pages.length} imágenes (P1…P${pages.length})`
      : '📸 Detalle asesor exportado en JPG';
    showToast(msg);
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
    XLSX.writeFile(wb, 'ASIGNACIÓN ' + (typeof getDateStamp==='function'?getDateStamp():'') + ' SG' + (typeof getSGFileLabel==='function'?getSGFileLabel():'') + '.xlsx');
    showToast(`⬇ ${rows.length.toLocaleString()} leads exportados`);
  }

  function limpiarBotonesDistribucion(){
    const configs = [
      { panel:'#tp-1',       copiar:'copiarIdsDistribucionSGFinal()', fotoFn:"exportDetalleDistJPG('sg')" },
      { panel:'#inter-tp-2', copiar:'copiarInterDistIDs()',            fotoFn:"exportDetalleDistJPG('inter')" },
      { panel:'#nc-tp-2',    copiar:'copiarNCDistIDs()',               fotoFn:"exportDetalleDistJPG('nc')" },
      { panel:'#pe-tp-2',    copiar:'copiarPEDistIDs()',               fotoFn:"exportDetalleDistJPG('pe')" }
    ];

    configs.forEach(cfg => {
      const panel = document.querySelector(cfg.panel);
      if(!panel) return;
      const cards = [...panel.querySelectorAll('.card2')];
      const card = cards.find(c => /Distribuir y exportar/i.test(c.textContent || ''));
      const acts = card ? card.querySelector('.acts') : null;
      if(!acts) return;

      // Eliminar botones viejos de detalle duplicados
      [...acts.querySelectorAll('button')].forEach(btn => {
        const text = btn.textContent || '';
        if(/Exportar detalle/i.test(text) && !btn.classList.contains('btn-photo')) btn.remove();
      });

      // Agregar Copiar IDs si no esta
      const hasCopy = [...acts.querySelectorAll('button')].some(btn => /Copiar IDs/i.test(btn.textContent || ''));
      if(!hasCopy){
        const btn = document.createElement('button');
        btn.className = 'btn btn-orange';
        btn.setAttribute('onclick', cfg.copiar);
        btn.textContent = 'Copiar IDs distribucion';
        acts.appendChild(btn);
      }

      // Agregar botón 📸 Exportar detalle asesor JPG si no está
      const hasFoto = [...acts.querySelectorAll('button.btn-photo')].length > 0;
      if(!hasFoto){
        const btn = document.createElement('button');
        btn.className = 'btn btn-photo';
        btn.setAttribute('onclick', cfg.fotoFn);
        btn.textContent = '📸 Exportar detalle asesor JPG';
        acts.appendChild(btn);
      }
    });
  }

  window.copiarIdsDistribucionSGFinal = copiarIdsSG;
  window.exportarDist = exportarDistSGFinal;

  /* ══════════════════════════════════════════════════════
     SIN GESTIÓN — sección 4
  ══════════════════════════════════════════════════════ */
  function addDetalleBtn(cardId, btnId, onclickFn){
    if(document.getElementById(btnId)) return;
    const card = document.getElementById(cardId);
    if(!card) return;
    const wrap = document.createElement('div');
    wrap.className = 'detalle-actions';
    wrap.innerHTML = `<button id="${btnId}" class="btn btn-green" onclick="${onclickFn}">📸 Exportar detalle asesor JPG</button>`;
    card.insertBefore(wrap, card.firstChild);
  }

  window.renderDetalleAsesorSG = function(){
    const card = document.getElementById('sg-detalle-asesor-card');
    if(card){ card.style.display = 'block'; card.classList.add('detalle-visible'); }
    addDetalleBtn('sg-detalle-asesor-card', 'btn-detalle-sg-jpg', 'exportDetalleAsesorSGJPG()');
    const data = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    renderDetalleVisual(data, 'sg-detalle-asesor', 'Primero distribuye los leads para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorSGJPG = function(){
    const data  = typeof distribuidoData !== 'undefined' ? distribuidoData : [];
    const stamp = typeof getDateStamp   === 'function' ? getDateStamp()   : '';
    const area  = typeof getSGFileLabel === 'function' ? getSGFileLabel() : '';
    exportDetalleJPG(data, 'Base Sin Gestión', 'DETALLE_ASESOR_MENTOR_ASIGNACIÓN ' + stamp + ' SG' + area);
  };

  /* ══════════════════════════════════════════════════════
     INTERESADOS — sección 4
  ══════════════════════════════════════════════════════ */
  window.renderDetalleAsesorInter = function(){
    const card = document.getElementById('inter-detalle-asesor-card');
    if(card){ card.style.display = 'block'; card.classList.add('detalle-visible'); }
    addDetalleBtn('inter-detalle-asesor-card', 'btn-detalle-inter-jpg', 'exportDetalleAsesorInterJPG()');
    const data = typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [];
    renderDetalleVisual(data, 'inter-detalle-asesor', 'Primero distribuye los leads para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorInterJPG = function(){
    const data  = typeof interDistribuidoData !== 'undefined' ? interDistribuidoData : [];
    const stamp = typeof getDateStamp     === 'function' ? getDateStamp()     : '';
    const area  = typeof getInterFileLabel === 'function' ? getInterFileLabel() : '';
    exportDetalleJPG(data, 'Base Interesados', 'DETALLE_ASESOR_MENTOR_ASIGNACIÓN ' + stamp + ' INTER' + area);
  };

  /* ══════════════════════════════════════════════════════
     NO CONTACTADO — sección 4
  ══════════════════════════════════════════════════════ */
  window.renderDetalleAsesorNC = function(){
    const card = document.getElementById('nc-detalle-asesor-card');
    if(card){ card.style.display = 'block'; card.classList.add('detalle-visible'); }
    addDetalleBtn('nc-detalle-asesor-card', 'btn-detalle-nc-jpg', 'exportDetalleAsesorNCJPG()');
    const data = typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [];
    renderDetalleVisual(data, 'nc-detalle-asesor', 'Primero distribuye los leads para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorNCJPG = function(){
    const data  = typeof ncDistribuidoData !== 'undefined' ? ncDistribuidoData : [];
    const stamp = typeof getDateStamp   === 'function' ? getDateStamp()   : '';
    const area  = typeof getNCFileLabel === 'function' ? getNCFileLabel() : '';
    exportDetalleJPG(data, 'Base No Contactado', 'DETALLE_ASESOR_MENTOR_ASIGNACIÓN ' + stamp + ' NOCONTAC' + area);
  };

  /* ══════════════════════════════════════════════════════
     PROBLEMAS ECONÓMICOS — sección 4
  ══════════════════════════════════════════════════════ */
  window.renderDetalleAsesorPE = function(){
    const card = document.getElementById('pe-detalle-asesor-card');
    if(card){ card.style.display = 'block'; card.classList.add('detalle-visible'); }
    addDetalleBtn('pe-detalle-asesor-card', 'btn-detalle-pe-jpg', 'exportDetalleAsesorPEJPG()');
    const data = typeof peDistribuidoData !== 'undefined' ? peDistribuidoData : [];
    renderDetalleVisual(data, 'pe-detalle-asesor', 'Primero distribuye los leads para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorPEJPG = function(){
    const data  = typeof peDistribuidoData !== 'undefined' ? peDistribuidoData : [];
    const stamp = typeof getDateStamp   === 'function' ? getDateStamp()   : '';
    const area  = typeof getPEFileLabel === 'function' ? getPEFileLabel() : '';
    exportDetalleJPG(data, 'Problemas Económicos', 'DETALLE_ASESOR_MENTOR_ASIGNACIÓN ' + stamp + ' PROBLEMAS' + area);
  };

  /* ══════════════════════════════════════════════════════
     RECIBOS — sección 2 "Detalle asesor por mentor"
  ══════════════════════════════════════════════════════ */
  window.renderDetalleAsesorRec = function(){
    const card = document.getElementById('rec-detalle-asesor-card');
    if(card){ card.style.display = 'block'; card.classList.add('detalle-visible'); }
    addDetalleBtn('rec-detalle-asesor-card', 'btn-detalle-rec-jpg', 'exportDetalleAsesorRecJPG()');
    const data = typeof recFiltered !== 'undefined' ? recFiltered : [];
    renderDetalleVisual(data, 'rec-detalle-asesor', 'Carga la base para ver el detalle por mentor y asesor.');
  };

  window.exportDetalleAsesorRecJPG = function(){
    const data  = typeof recFiltered !== 'undefined' ? recFiltered : [];
    const stamp = typeof getDateStamp === 'function' ? getDateStamp() : '';
    exportDetalleJPG(data, 'Base Recibos', 'DETALLE_ASESOR_MENTOR_ASIGNACIÓN ' + stamp + ' REC');
  };

  /* ══════════════════════════════════════════════════════
     HOOKS DE DISTRIBUCIÓN — todos los módulos
  ══════════════════════════════════════════════════════ */
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
    [80, 250, 600].forEach(ms => setTimeout(() => {
      window.renderDetalleAsesorInter();
      limpiarBotonesDistribucion();
    }, ms));
  };

  const oldNC = window.distribuirNC;
  window.distribuirNC = function(){
    if(typeof oldNC === 'function') oldNC.apply(this, arguments);
    [80, 250, 600].forEach(ms => setTimeout(() => {
      window.renderDetalleAsesorNC();
      limpiarBotonesDistribucion();
    }, ms));
  };

  const oldPE = window.distribuirPE;
  window.distribuirPE = function(){
    if(typeof oldPE === 'function') oldPE.apply(this, arguments);
    [80, 250, 600].forEach(ms => setTimeout(() => {
      window.renderDetalleAsesorPE();
      limpiarBotonesDistribucion();
    }, ms));
  };

  document.addEventListener('DOMContentLoaded', () => {
    [250, 800, 1500].forEach(ms => setTimeout(() => {
      limpiarBotonesDistribucion();
      if(typeof window.renderDetalleAsesorSG    === 'function') window.renderDetalleAsesorSG();
      if(typeof window.renderDetalleAsesorInter === 'function') window.renderDetalleAsesorInter();
      if(typeof window.renderDetalleAsesorNC    === 'function') window.renderDetalleAsesorNC();
      if(typeof window.renderDetalleAsesorPE    === 'function') window.renderDetalleAsesorPE();
    }, ms));
  });
})();
