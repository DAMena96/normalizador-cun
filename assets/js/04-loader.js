/* ════════════════════════════════════════════
   CARGA
   ════════════════════════════════════════════ */
function handleDrop(e){
  e.preventDefault();
  document.getElementById('upload-zone').classList.remove('over');
  loadFile(e.dataTransfer.files[0]);
}

async function loadFile(file){
  if(!file)return;
  showProg(5,'Cargando catálogos JSON…');
  const okCatalogs = await catalogsReady;
  if(!okCatalogs){ hideProg(); return; }
  showProg(10,'Leyendo archivo…');
  const reader=new FileReader();
  reader.onload=ev=>{
    showProg(30,'Parseando Excel…');
    setTimeout(()=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:'array'});

        const sn=wb.SheetNames.find(s=>/SIN\s*GESTI/i.test(s))||wb.SheetNames[0];
        const json=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
        if(json.length<2){showToast('Sin datos en la hoja.');hideProg();return;}
        const headers=json[0].map(h=>String(h).trim());

        showProg(55,`Normalizando ${(json.length-1).toLocaleString()} registros…`);
        setTimeout(()=>{
          allData=[];
          for(let i=1;i<json.length;i++){
            const row=json[i];
            if(row.every(c=>c===''||c===null||c===undefined))continue;
            allData.push(normalizeRow(row,headers));
          }
          // DUPLICADOS: si tel o correo aparece >1 vez acumulado → Duplicado
          const sT={},sC={};
          allData.forEach(r=>{
            const t=r['Número DEF'],c=r['CORREO ELECTRONICO DEF'];
            sT[t]=(sT[t]||0)+1; sC[c]=(sC[c]||0)+1;
            r['DUPLICADO']=(sT[t]+sC[c])>2?'Duplicado':'No';
          });

          showProg(90,'Preparando…');
          setTimeout(()=>{
            filteredData=[...allData];
            buildFilterPanel();
            populateFilters();
            renderKPIs(); renderTable();
            hideProg();
            document.getElementById('upload-zone').classList.add('hidden');
            document.getElementById('kpi-row').classList.remove('hidden');
            document.getElementById('tabs').style.display='flex';
            document.getElementById('filter-panel').style.display='block';
            document.getElementById('table-wrap').style.display='block';
            showToast(`✅ ${allData.length.toLocaleString()} registros normalizados con JSON · `+
              `${Object.keys(MAP_PROGRAMA).length} programas · `+
              `${Object.keys(MAP_CIUDAD).length} ciudades`);
          },80);
        },50);
      }catch(err){showToast('Error: '+err.message);hideProg();console.error(err);}
    },30);
  };
  reader.readAsArrayBuffer(file);
}
function showProg(p,m){
  document.getElementById('progress-wrap').style.display='block';
  document.getElementById('progress-fill').style.width=p+'%';
  document.getElementById('progress-label').textContent=m;
}
function hideProg(){document.getElementById('progress-wrap').style.display='none';}


