/* ════════════════════════════════════════════
   KPIs
   ════════════════════════════════════════════ */
function renderKPIs(){
  const d=filteredData;
  document.getElementById('k-total').textContent=d.length.toLocaleString();
  document.getElementById('k-pres').textContent=d.filter(r=>r['VALIDACION FINAL']==='PRESENCIAL').length.toLocaleString();
  document.getElementById('k-virt').textContent=d.filter(r=>r['VALIDACION FINAL']==='VIRTUAL').length.toLocaleString();
  document.getElementById('k-pos').textContent=d.filter(r=>r['VALIDACION FINAL']==='POSGRADO').length.toLocaleString();
  document.getElementById('k-err').textContent=d.filter(r=>['Número errado','Sin número'].includes(r['NOVEDAD'])).length.toLocaleString();
  document.getElementById('k-dup').textContent=d.filter(r=>r['DUPLICADO']==='Duplicado').length.toLocaleString();
}


/* ════════════════════════════════════════════
   TABLA
   ════════════════════════════════════════════ */
function renderTable(){
  document.getElementById('tbl-head').innerHTML='<tr>'+
    COLS_SHOW.map(c=>`<th onclick="sortBy('${c}')">${c} <span style="opacity:.35">${sortCol===c?(sortAsc?'▲':'▼'):'⇅'}</span></th>`).join('')+'</tr>';

  let data=[...filteredData];
  if(sortCol) data.sort((a,b)=>{
    const va=String(a[sortCol]||''),vb=String(b[sortCol]||'');
    return sortAsc?va.localeCompare(vb):vb.localeCompare(va);
  });

  const total=data.length, pages=Math.max(1,Math.ceil(total/PAGE_SIZE));
  currentPage=Math.min(currentPage,pages);
  const sl=data.slice((currentPage-1)*PAGE_SIZE,currentPage*PAGE_SIZE);

  document.getElementById('tbl-title').textContent=
    `${total.toLocaleString()} registros${filteredData.length<allData.length?' (filtrados)':''}`;

  document.getElementById('tbl-body').innerHTML=sl.map(r=>
    '<tr>'+COLS_SHOW.map(c=>{
      const v=r[c]!==undefined?String(r[c]):'';
      return `<td title="${v}">${getBadge(c,v)||v}</td>`;
    }).join('')+'</tr>'
  ).join('');

  document.getElementById('pag-info').textContent=
    `Página ${currentPage} de ${pages} · ${total.toLocaleString()} registros`;
}

function getBadge(col,val){
  if(['CAMPAÑA','AREA PROGRAMA','VALIDACION FINAL'].includes(col))
    return `<span class="badge b${val.replace(/\s/g,'')}">${val}</span>`;
  if(col==='NOVEDAD'){
    const cls=val==='Normal'?'bNormal':val==='Duplicado'?'bDuplicado':'bErr';
    return `<span class="badge ${cls}">${val}</span>`;
  }
  if(col==='DUPLICADO')
    return val==='Duplicado'?`<span class="badge bDuplicado">${val}</span>`:`<span class="badge bNormal">${val}</span>`;
  return null;
}

function sortBy(col){if(sortCol===col)sortAsc=!sortAsc;else{sortCol=col;sortAsc=true;}renderTable();}
function changePage(dir){currentPage=Math.min(Math.ceil(filteredData.length/PAGE_SIZE),Math.max(1,currentPage+dir));renderTable();}


