/* ════════════════════════════════════════════
   ESTADO
   ════════════════════════════════════════════ */
let allData=[], filteredData=[];
let sortCol='', sortAsc=true, currentPage=1;
const PAGE_SIZE=100;
let asesoresList=[], distribuidoData=[];

const COLS_ALL=[
  'ID de registro','Hora de creación','Nombre completo','Teléfono','Correo electrónico',
  'Número de Documento','Modalidad','Periodo','Programa de interes_','Propietario de Posible Cliente',
  'Creado por','Campaña mercadeo','Etapa del Registro','Estado de Posible Cliente',
  'Hora de modificación','Ultimo nivel de escolaridad','Editado por','Periodo Creado',
  'Correo electrónico Continua','Teléfono Continua','Ciudad de residencia','Sub Estado',
  'Etiqueta','Ciudad utm','utm_campaign','utm_content','utm_medium','utm_source','utm_term',
  'conv','Total Notas','Descripción','Descripción actualizacion','Programa',
  'CAMPAÑA','CORREO ELECTRONICO DEF','Número DEF','Asesor','CIUDAD',
  'AREA PROGRAMA','PROGRAMA2','Llave','VALIDACION FINAL','NOVEDAD','MENTOR','DUPLICADO'
];
const COLS_SHOW=[
  'ID de registro','Nombre completo','CORREO ELECTRONICO DEF','Número DEF',
  'Asesor','CIUDAD','PROGRAMA2','AREA PROGRAMA','CAMPAÑA','VALIDACION FINAL',
  'NOVEDAD','Campaña mercadeo','Periodo','conv','Creado por','Total Notas','DUPLICADO'
];
const FILTER_DEFS=[
  {id:'f-camp', col:'CAMPAÑA',          lbl:'Campaña'},
  {id:'f-area', col:'AREA PROGRAMA',    lbl:'Área Programa'},
  {id:'f-val',  col:'VALIDACION FINAL', lbl:'Validación Final'},
  {id:'f-nov',  col:'NOVEDAD',          lbl:'Novedad'},
  {id:'f-ciu',  col:'CIUDAD',           lbl:'Ciudad'},
  {id:'f-ase',  col:'Asesor',           lbl:'Asesor'},
  {id:'f-prog', col:'PROGRAMA2',        lbl:'Programa'},
  {id:'f-crea', col:'Creado por',       lbl:'Creado por'},
  {id:'f-ment', col:'MENTOR',           lbl:'Supervisor'},
  {id:'f-llav', col:'Llave',            lbl:'Llave'},
  {id:'f-nota', col:'Total Notas',      lbl:'Total Notas'},
  {id:'f-per',  col:'Periodo',          lbl:'Periodo'},
  {id:'f-dup',  col:'DUPLICADO',        lbl:'Duplicado'},
];
const msSel={};
FILTER_DEFS.forEach(f=>msSel[f.id]=new Set());


