
// Funciones de interfaz compartidas

function selectModule(type){
  const home = document.getElementById('selector-home');
  const upload = document.getElementById('upload-zone');

  if(type === 'sin-gestion'){
    home.classList.add('hidden');
    upload.classList.remove('hidden');
    const backBtn = document.getElementById('back-home-btn'); if(backBtn) backBtn.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('✅ Base Sin Gestión lista para cargar archivo');
    return;
  }

  if(type === 'interesados'){
    document.getElementById('selector-home').classList.add('hidden');
    const mod = document.getElementById('interesados-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('⭐ Módulo Base de Interesados cargado');
    return;
  }

  if(type === 'no-contactado'){
    document.getElementById('selector-home').classList.add('hidden');
    const mod = document.getElementById('no-contactado-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('📞 Módulo Base No Contactado cargado');
    return;
  }

  if(type === 'recibos'){
    document.getElementById('selector-home').classList.add('hidden');
    const mod = document.getElementById('recibos-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('🧾 Módulo Recibos CARE AI cargado');
    return;
  }
}

function backHome(){
  document.getElementById('selector-home').classList.remove('hidden');
  document.getElementById('upload-zone').classList.add('hidden');
  const backBtn = document.getElementById('back-home-btn'); if(backBtn) backBtn.classList.add('hidden');
  document.getElementById('kpi-row').classList.add('hidden');
  document.getElementById('tabs').style.display = 'none';
  document.getElementById('filter-panel').style.display = 'none';
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  const tp0 = document.getElementById('tp-0');
  if(tp0) tp0.classList.add('active');
  const ncMod = document.getElementById('no-contactado-module');
  if(ncMod) ncMod.classList.add('hidden');
  const ncUpload = document.getElementById('nc-upload-zone');
  if(ncUpload) ncUpload.classList.remove('hidden');
  const ncKpi = document.getElementById('nc-kpi-row');
  if(ncKpi) ncKpi.classList.add('hidden');
  const ncTabs = document.getElementById('nc-tabs');
  if(ncTabs) ncTabs.classList.add('hidden');
  document.querySelectorAll('.nc-panel').forEach(p => p.classList.add('hidden'));

  const interMod = document.getElementById('interesados-module');
  if(interMod) interMod.classList.add('hidden');
  const interUpload = document.getElementById('inter-upload-zone');
  if(interUpload) interUpload.classList.remove('hidden');
  const interKpi = document.getElementById('inter-kpi-row');
  if(interKpi) interKpi.classList.add('hidden');
  const interTabs = document.getElementById('inter-tabs');
  if(interTabs) interTabs.classList.add('hidden');
  document.querySelectorAll('.inter-panel').forEach(p => p.classList.add('hidden'));
  window.scrollTo({ top: 0, behavior: 'smooth' });
}


/* === RESET FINAL DE VISTAS === */
function resetAllModuleViews(){
  // Base Sin Gestión
  const upload = document.getElementById('upload-zone');
  if(upload) upload.classList.add('hidden');
  const backBtn = document.getElementById('back-home-btn');
  if(backBtn) backBtn.classList.add('hidden');
  const kpi = document.getElementById('kpi-row');
  if(kpi) kpi.classList.add('hidden');
  const tabs = document.getElementById('tabs');
  if(tabs) tabs.style.display = 'none';
  const filterPanel = document.getElementById('filter-panel');
  if(filterPanel) filterPanel.style.display = 'none';
  ['tp-0','tp-1','tp-2','tp-3'].forEach(id => {
    const el = document.getElementById(id);
    if(el) el.classList.remove('active');
  });

  // Módulos
  ['interesados-module','no-contactado-module','recibos-module'].forEach(id => {
    const el = document.getElementById(id);
    if(el) el.classList.add('hidden');
  });
}

function selectModule(type){
  resetAllModuleViews();
  const home = document.getElementById('selector-home');
  if(home) home.classList.add('hidden');

  if(type === 'sin-gestion'){
    const upload = document.getElementById('upload-zone');
    if(upload) upload.classList.remove('hidden');
    const backBtn = document.getElementById('back-home-btn');
    if(backBtn) backBtn.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('✅ Base Sin Gestión lista para cargar archivo');
    return;
  }

  if(type === 'interesados'){
    const mod = document.getElementById('interesados-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('⭐ Módulo Base de Interesados cargado');
    return;
  }

  if(type === 'no-contactado'){
    const mod = document.getElementById('no-contactado-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('📞 Módulo Base No Contactado cargado');
    return;
  }

  if(type === 'recibos'){
    const mod = document.getElementById('recibos-module');
    if(mod) mod.classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast('🧾 Módulo Recibos CARE AI cargado');
    return;
  }
}

function backHome(){
  resetAllModuleViews();
  const home = document.getElementById('selector-home');
  if(home) home.classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}


/* === FIX TOPBAR SIN GESTIÓN === */
(function(){
  const oldSelectModule = window.selectModule;
  window.selectModule = function(type){
    if(typeof oldSelectModule === 'function'){
      oldSelectModule(type);
    }

    const sgTop = document.getElementById('sin-gestion-topbar');
    if(sgTop){
      if(type === 'sin-gestion'){
        sgTop.classList.remove('hidden');
      }else{
        sgTop.classList.add('hidden');
      }
    }

    const oldBtn = document.getElementById('back-home-btn');
    if(oldBtn && type === 'sin-gestion'){
      oldBtn.classList.remove('hidden');
    }
  };

  const oldBackHome = window.backHome;
  window.backHome = function(){
    if(typeof oldBackHome === 'function'){
      oldBackHome();
    }

    const sgTop = document.getElementById('sin-gestion-topbar');
    if(sgTop) sgTop.classList.add('hidden');
  };
})();


/* === FIX TOPBAR SIN GESTIÓN DEFINITIVO === */
(function(){
  const oldSelectModule = window.selectModule;
  window.selectModule = function(type){
    if(typeof oldSelectModule === 'function') oldSelectModule(type);
    const sgTop = document.getElementById('sin-gestion-topbar');
    if(sgTop){
      if(type === 'sin-gestion') sgTop.classList.remove('hidden');
      else sgTop.classList.add('hidden');
    }
  };
  const oldBackHome = window.backHome;
  window.backHome = function(){
    if(typeof oldBackHome === 'function') oldBackHome();
    const sgTop = document.getElementById('sin-gestion-topbar');
    if(sgTop) sgTop.classList.add('hidden');
  };
})();

/* === FIX TOPBAR SIN GESTIÓN FINAL VISIBLE === */
(function(){
  const oldSelect=window.selectModule;
  window.selectModule=function(type){
    if(typeof oldSelect==='function') oldSelect(type);
    const sg=document.getElementById('sin-gestion-topbar');
    if(sg){ type==='sin-gestion'?sg.classList.remove('hidden'):sg.classList.add('hidden'); }
  };
  const oldBack=window.backHome;
  window.backHome=function(){
    if(typeof oldBack==='function') oldBack();
    const sg=document.getElementById('sin-gestion-topbar');
    if(sg) sg.classList.add('hidden');
  };
})();

/* === FIX TOPBAR SIMPLE FINAL === */
(function(){
  const oldSelect = window.selectModule;
  window.selectModule = function(type){
    if(typeof oldSelect === 'function') oldSelect(type);
    const sg = document.getElementById('sin-gestion-topbar');
    if(sg) type === 'sin-gestion' ? sg.classList.remove('hidden') : sg.classList.add('hidden');
  };
  const oldBack = window.backHome;
  window.backHome = function(){
    if(typeof oldBack === 'function') oldBack();
    const sg = document.getElementById('sin-gestion-topbar');
    if(sg) sg.classList.add('hidden');
  };
})();
