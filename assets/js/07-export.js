/* ════════════════════════════════════════════
   EXPORTAR
   ════════════════════════════════════════════ */
function exportData(rows,name){
  const ws=XLSX.utils.json_to_sheet(rows.map(r=>{
    const o={};COLS_ALL.forEach(c=>o[c]=r[c]!==undefined?r[c]:'');return o;
  }));
  const wb2=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2,ws,'Normalizado');
  XLSX.writeFile(wb2,name);
  showToast(`⬇ ${rows.length.toLocaleString()} filas exportadas`);
}
function exportFiltered(){exportData(filteredData,'Base_Filtrada_'+Date.now()+'.xlsx');}
function exportAll(){exportData(allData,'Base_Completa_'+Date.now()+'.xlsx');}
function copiarIDs(){
  navigator.clipboard.writeText(filteredData.map(r=>r['ID de registro']).join('\n'))
    .then(()=>showToast(`📋 ${filteredData.length.toLocaleString()} IDs copiados`));
}


