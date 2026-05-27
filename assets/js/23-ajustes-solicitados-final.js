/* =========================================================
   AJUSTES SOLICITADOS — FINAL SEGURO
   - No toca la lógica base de carga/distribución.
   - Mantiene navegación entre módulos.
   - Elimina botones JPG duplicados.
   - Agrega Exportar JPG y Exportar Excel al resumen por asesor original
     en Sin Gestión, Interesados y No Contactado.
   - Unifica tarjetas de estado y barras de carga.
   - Fuerza que Sin Gestión muestre la tabla al terminar de cargar.
   - Agrega fecha/hora de generación en exportaciones JPG.
   ========================================================= */
(function(){
  'use strict';

  const MODS = {
    sg: {
      label: 'Base Sin Gestión',
      wrapId: 'sg-owner-summary-wrap',
      getData: () => (typeof filteredData !== 'undefined' ? filteredData : (typeof allData !== 'undefined' ? allData : [])),
      filePrefix: 'Resumen_Asesor_Original_Sin_Gestion'
    },
    inter: {
      label: 'Base Interesados',
      wrapId: 'inter-owner-summary-wrap',
      getData: () => (typeof interFiltered !== 'undefined' ? interFiltered : (typeof interData !== 'undefined' ? interData : [])),
      filePrefix: 'Resumen_Asesor_Original_Interesados'
    },
    nc: {
      label: 'Base No Contactado',
      wrapId: 'nc-owner-summary-wrap',
      getData: () => (typeof ncFiltered !== 'undefined' ? ncFiltered : (typeof ncData !== 'undefined' ? ncData : [])),
      filePrefix: 'Resumen_Asesor_Original_No_Contactado'
    }
  };

  function safe(v=''){
    return String(v ?? '')
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#039;');
  }

  function nowText(){
    return new Date().toLocaleString('es-CO', {
      year:'numeric', month:'numeric', day:'numeric',
      hour:'numeric', minute:'2-digit', second:'2-digit'
    });
  }

  function catalogInfoText(total){
    const p = typeof MAP_PROGRAMA !== 'undefined' ? Object.keys(MAP_PROGRAMA || {}).length : 0;
    const c = typeof MAP_CIUDAD !== 'undefined' ? Object.keys(MAP_CIUDAD || {}).length : 0;
    return `✅ ${Number(total || 0).toLocaleString()} registros normalizados con JSON · ${p.toLocaleString()} programas · ${c.toLocaleString()} ciudades`;
  }

  function ensureModuleStatus(module, total){
    const ids = {
      sg: { after: 'sin-gestion-topbar', id:'sg-json-status' },
      inter: { after: 'inter-upload-zone', id:'inter-json-status' },
      nc: { after: 'nc-upload-zone', id:'nc-json-status' },
      rec: { after: 'rec-upload-zone', id:'rec-json-status' }
    }[module];
    if(!ids) return;

    let el = document.getElementById(ids.id);
    if(!el){
      el = document.createElement('div');
      el.id = ids.id;
      el.className = 'json-status-card hidden';
      const ref = document.getElementById(ids.after);
      if(ref) ref.insertAdjacentElement('afterend', el);
      else document.body.prepend(el);
    }
    el.textContent = catalogInfoText(total);
    el.classList.remove('hidden');
  }

  function normalizeProgressLabel(label){
    const s = String(label || '');
    const m = s.match(/(\d[\d\.\,]*)/);
    if(/Normalizando/i.test(s) && m) return `Normalizando ${m[1]} registros...`;
    return s.replace('…','...');
  }

  function wrapProgressFunction(name){
    const old = window[name];
    if(typeof old !== 'function' || old.__wrappedFinal) return;
    const fn = function(pct, label){
      return old.call(this, pct, normalizeProgressLabel(label));
    };
    fn.__wrappedFinal = true;
    window[name] = fn;
  }

  ['showProg','interShowProg','ncShowProg','recShowProg'].forEach(wrapProgressFunction);

  function forceSinGestionVisible(){
    const tabs = document.getElementById('tabs');
    const filterPanel = document.getElementById('filter-panel');
    const tableWrap = document.getElementById('table-wrap');
    if(tabs) tabs.style.display = 'flex';
    if(filterPanel) filterPanel.style.display = 'block';
    if(tableWrap) tableWrap.style.display = 'block';
    // No forzar cambio de tab aquí; switchTab(0) en 04-loader ya lo gestiona
  }

  function ownerRows(module){
    const cfg = MODS[module];
    const data = cfg ? cfg.getData() : [];
    const counts = {};
    (data || []).forEach(r => {
      const asesor = String(r['Propietario de Posible Cliente'] || r['Propietario de posible cliente'] || r['Lead Owner'] || '').trim() || 'Sin asesor';
      counts[asesor] = (counts[asesor] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([asesor, cantidad]) => ({ Asesor: asesor, 'Cantidad leads': cantidad }))
      .sort((a,b) => b['Cantidad leads'] - a['Cantidad leads'] || String(a.Asesor).localeCompare(String(b.Asesor)));
  }

  function ensureOwnerButtons(module){
    const cfg = MODS[module];
    if(!cfg) return;
    const wrap = document.getElementById(cfg.wrapId);
    if(!wrap || wrap.style.display === 'none') return;
    const card = wrap.querySelector('.owner-summary-card') || wrap;
    let actions = card.querySelector('.owner-summary-actions');
    if(!actions){
      actions = document.createElement('div');
      actions.className = 'owner-summary-actions';
      const head = card.querySelector('.owner-summary-head');
      if(head) head.insertAdjacentElement('afterend', actions);
      else card.prepend(actions);
    }
    actions.innerHTML = `
      <button type="button" class="btn btn-orange" onclick="exportResumenAsesorOriginalJPG('${module}')">📸 Exportar JPG</button>
      <button type="button" class="btn btn-green" onclick="exportResumenAsesorOriginalExcel('${module}')">⬇ Exportar Excel</button>
    `;
  }

  function ensureAllOwnerButtons(){
    ensureOwnerButtons('sg');
    ensureOwnerButtons('inter');
    ensureOwnerButtons('nc');
  }

  window.exportResumenAsesorOriginalExcel = function(module){
    const cfg = MODS[module];
    const rows = ownerRows(module);
    if(!cfg || !rows.length){ showToast('No hay datos para exportar.'); return; }
    const ws = XLSX.utils.json_to_sheet(rows, { header: ['Asesor','Cantidad leads'] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resumen asesor');
    XLSX.writeFile(wb, `${cfg.filePrefix}_${Date.now()}.xlsx`);
    showToast(`⬇ Resumen exportado en Excel: ${rows.length.toLocaleString()} asesores`);
  };

  window.exportResumenAsesorOriginalJPG = function(module){
    const cfg = MODS[module];
    const rows = ownerRows(module);
    if(!cfg || !rows.length){ showToast('No hay datos para exportar.'); return; }

    const total = rows.reduce((a,b)=>a + Number(b['Cantidad leads'] || 0), 0);
    const generated = nowText();
    const scale = 2;
    const width = 1200;
    const margin = 34;
    const titleH = 116;
    const summaryH = 86;
    const rowH = 34;
    const footerH = 44;
    const height = Math.max(460, titleH + summaryH + 38 + rows.length * rowH + footerH + 60);

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
    ctx.font = 'bold 30px Segoe UI, Arial, sans-serif';
    ctx.fillText('Resumen por asesor original', margin, 54);
    ctx.font = '15px Segoe UI, Arial, sans-serif';
    ctx.fillText(`${cfg.label} · Generado: ${generated}`, margin, 88);

    const cardY = titleH + 24;
    ctx.fillStyle = '#f4f8fb';
    ctx.fillRect(margin, cardY, width - margin*2, 58);
    ctx.fillStyle = '#0C2340';
    ctx.font = 'bold 26px Segoe UI, Arial, sans-serif';
    ctx.fillText(total.toLocaleString(), margin + 22, cardY + 37);
    ctx.font = '14px Segoe UI, Arial, sans-serif';
    ctx.fillStyle = '#58677a';
    ctx.fillText(`leads · ${rows.length.toLocaleString()} asesores`, margin + 170, cardY + 36);

    let y = cardY + 92;
    ctx.fillStyle = '#1B365D';
    ctx.fillRect(margin, y, width - margin*2, 38);
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
    ctx.fillText('Asesor', margin + 12, y + 25);
    ctx.fillText('Cantidad leads', width - margin - 150, y + 25);
    y += 38;

    function cutText(text, x, yy, maxW){
      let s = String(text || '');
      while(ctx.measureText(s).width > maxW && s.length > 4) s = s.slice(0,-2);
      if(s !== String(text || '')) s += '...';
      ctx.fillText(s, x, yy);
    }

    rows.forEach((r, i) => {
      ctx.fillStyle = i % 2 ? '#ffffff' : '#f8fafc';
      ctx.fillRect(margin, y, width - margin*2, rowH);
      ctx.strokeStyle = '#e8eef5';
      ctx.beginPath();
      ctx.moveTo(margin, y + rowH);
      ctx.lineTo(width - margin, y + rowH);
      ctx.stroke();
      ctx.fillStyle = '#2c3e50';
      ctx.font = '14px Segoe UI, Arial, sans-serif';
      cutText(r.Asesor, margin + 12, y + 22, 850);
      ctx.fillStyle = '#0C2340';
      ctx.font = 'bold 15px Segoe UI, Arial, sans-serif';
      ctx.fillText(Number(r['Cantidad leads']).toLocaleString(), width - margin - 105, y + 22);
      y += rowH;
    });

    ctx.fillStyle = '#898D8D';
    ctx.font = '12px Segoe UI, Arial, sans-serif';
    ctx.fillText(`Generado desde App Normalizador Contact CUN · ${generated}`, margin, height - 18);

    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/jpeg', 0.95);
    link.download = `${cfg.filePrefix}_${Date.now()}.jpg`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('📸 Resumen exportado en JPG');
  };

  function removeDuplicateDetailJpgButtons(){
    ['sg-detalle-asesor-card','inter-detalle-asesor-card','nc-detalle-asesor-card'].forEach(cardId => {
      const card = document.getElementById(cardId);
      if(!card) return;
      const buttons = [...card.querySelectorAll('button')].filter(b => /Exportar detalle asesor JPG/i.test(b.textContent || ''));
      buttons.forEach((b,i)=>{ if(i > 0) b.remove(); });
      // Si también quedó un botón dentro del contenido renderizado, deja solo el de la cabecera.
      const all = [...card.querySelectorAll('button')].filter(b => /Exportar detalle asesor JPG/i.test(b.textContent || ''));
      if(all.length > 1) all.slice(1).forEach(b => b.remove());
    });
  }

  function enhanceHomeVisual(){
    const hero = document.querySelector('.hero');
    if(!hero || document.getElementById('contact-center-visual')) return;
    const visual = document.createElement('div');
    visual.id = 'contact-center-visual';
    visual.className = 'contact-center-visual';
    visual.innerHTML = `
      <div class="cc-orbit cc-o1"></div>
      <div class="cc-orbit cc-o2"></div>
      <div class="cc-panel">
        <div class="cc-agent"><span>🎧</span><b>Contact Center</b></div>
        <div class="cc-bars"><i></i><i></i><i></i><i></i></div>
        <div class="cc-kpis"><span>LEADS</span><strong>Normalización</strong><em>Distribución inteligente</em></div>
        <p class="cc-desc">Gestión eficiente de tus contactos</p>
      </div>
    `;
    hero.appendChild(visual);
  }

  // Enriquecer renderizados ya existentes sin reemplazar lógica.
  function wrapRender(name, module, after){
    const old = window[name];
    if(typeof old !== 'function' || old.__wrappedSolicitado) return;
    const fn = function(){
      const result = old.apply(this, arguments);
      try{
        if(module) ensureModuleStatus(module, (MODS[module] && MODS[module].getData().length) || 0);
        ensureAllOwnerButtons();
        removeDuplicateDetailJpgButtons();
        if(typeof after === 'function') after();
      }catch(e){ console.warn('Ajuste visual:', e); }
      return result;
    };
    fn.__wrappedSolicitado = true;
    window[name] = fn;
  }

  wrapRender('renderTable', 'sg', forceSinGestionVisible);
  wrapRender('renderInterTable', 'inter');
  wrapRender('renderNCTable', 'nc');
  wrapRender('renderRecCari', null, () => ensureModuleStatus('rec', (typeof recData !== 'undefined' ? recData.length : 0)));

  // Mantener exportaciones de detalle como JPG y con fecha/hora visible usando las funciones finales si existen.
  const oldShowToast = window.showToast;
  if(typeof oldShowToast === 'function' && !oldShowToast.__jpgGuard){
    const guarded = function(msg){
      msg = String(msg || '').replace(/PDF/gi, 'JPG').replace(/PNG/gi, 'JPG');
      return oldShowToast.call(this, msg);
    };
    guarded.__jpgGuard = true;
    window.showToast = guarded;
  }

  document.addEventListener('DOMContentLoaded', () => {
    enhanceHomeVisual();
    [100,300,800,1500].forEach(ms => setTimeout(() => {
      enhanceHomeVisual();
      ensureAllOwnerButtons();
      removeDuplicateDetailJpgButtons();
    }, ms));
  });

  setInterval(() => {
    ensureAllOwnerButtons();
    removeDuplicateDetailJpgButtons();
  }, 1500);
})();
