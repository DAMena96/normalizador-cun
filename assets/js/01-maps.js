/* ════════════════════════════════════════════
   CATÁLOGOS JSON — reemplazan la hoja VAL del Excel
   Edita los archivos en: assets/data/*.json
   ════════════════════════════════════════════ */
let MAP_ASESOR   = {};
let MAP_PROGRAMA = {};
let MAP_CIUDAD   = {};
let MAP_LLAVE    = {};
let MAP_AREA     = {};
let MAP_SUPERVISOR = {};

const CATALOG_FILES = {
  asesores:  'assets/data/asesores.json',
  programas: 'assets/data/programas.json',
  ciudades:  'assets/data/ciudades.json',
  llaves:    'assets/data/llaves.json',
  areas:     'assets/data/areas.json',
  supervisor: 'assets/data/Supervisor_simple.json'
};

async function loadJson(path) {
  const res = await fetch(path, { cache: 'no-store' });
  if (!res.ok) throw new Error(`No se pudo cargar ${path}`);
  return await res.json();
}

async function loadCatalogs() {
  try {
    const [asesores, programas, ciudades, llaves, areas, supervisor] = await Promise.all([
      loadJson(CATALOG_FILES.asesores),
      loadJson(CATALOG_FILES.programas),
      loadJson(CATALOG_FILES.ciudades),
      loadJson(CATALOG_FILES.llaves),
      loadJson(CATALOG_FILES.areas),
      loadJson(CATALOG_FILES.supervisor)
    ]);

    MAP_ASESOR   = asesores;
    MAP_PROGRAMA = programas;
    MAP_CIUDAD   = ciudades;
    MAP_LLAVE    = llaves;
    MAP_AREA     = areas;
    MAP_SUPERVISOR = buildSupervisorMap(supervisor);

    console.log('CATÁLOGOS JSON →',
      'asesores:', Object.keys(MAP_ASESOR).length,
      'programas:', Object.keys(MAP_PROGRAMA).length,
      'ciudades:', Object.keys(MAP_CIUDAD).length,
      'llaves:', Object.keys(MAP_LLAVE).length,
      'áreas:', Object.keys(MAP_AREA).length,
      'supervisor:', Object.keys(MAP_SUPERVISOR).length
    );

    return true;
  } catch (err) {
    console.error(err);
    showToast('Error cargando catálogos JSON. Abre el proyecto con Live Server.');
    return false;
  }
}

const catalogsReady = loadCatalogs();

function cleanText(s = '') {
  return s
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[_-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

/* Lookup flexible: exacto, mayúsculas y versión limpia sin tildes/guiones/underscores */
function lookup(map, key, fallback) {
  if (!key) return fallback;

  if (map[key] !== undefined) return map[key];

  const keyUpper = String(key).toUpperCase();
  if (map[keyUpper] !== undefined) return map[keyUpper];

  const cleanedKey = cleanText(key);
  for (const k in map) {
    if (cleanText(k) === cleanedKey) return map[k];
  }

  console.warn('No encontrado en catálogo:', key);
  return fallback;
}





/* Catálogo Supervisor_simple.json
   En el archivo que enviaste:
   - La primera columna trae el ASESOR.
   - El nombre de la primera columna es el SUPERVISOR.
   - La segunda columna trae el ÁREA.
   - La tercera columna trae GRUPO/TIPO.
*/
function normalizeEmailOrText(value=''){
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function extractEmail(value=''){
  const s = String(value || '').trim();
  const m = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return m ? m[0].toLowerCase() : '';
}

function buildSupervisorMap(data){
  const map = {};
  if(!data) return map;

  /* ── Formato plano: { "asesor@email.com": "NombreSupervisor", … } ──
     Este es el formato de Supervisor_simple.json                        */
  if(!Array.isArray(data) && typeof data === 'object'){
    Object.entries(data).forEach(([asesorRaw, supervisorRaw]) => {
      if(!asesorRaw) return;
      const supervisor = String(supervisorRaw || 'Sin supervisor').trim() || 'Sin supervisor';
      const info = { asesor: asesorRaw, area: 'Sin área', grupo: '', supervisor };

      const email = extractEmail(asesorRaw);
      if(email) map[email] = info;

      map[normalizeEmailOrText(asesorRaw)] = info;
      map[cleanText(asesorRaw)]            = info;
    });
    return map;
  }

  /* ── Formato array (fallback por compatibilidad) ── */
  data.forEach(row => {
    const entries = Object.entries(row || {});
    if(!entries.length) return;

    const supervisor = String(entries[0]?.[0] || 'Sin supervisor').trim() || 'Sin supervisor';
    const asesorRaw  = String(entries[0]?.[1] || '').trim();
    const area       = String(entries[1]?.[1] || 'Sin área').trim() || 'Sin área';
    const grupo      = String(entries[2]?.[1] || '').trim();

    if(!asesorRaw) return;

    const info = { asesor: asesorRaw, area, grupo, supervisor };

    const email = extractEmail(asesorRaw);
    if(email) map[email] = info;

    map[normalizeEmailOrText(asesorRaw)] = info;
    map[cleanText(asesorRaw)]            = info;
  });

  return map;
}

function lookupSupervisorInfo(asesor){
  const raw = String(asesor || '').trim();
  if(!raw) return {asesor:'', area:'Sin área', grupo:'', supervisor:'Sin supervisor'};

  const email = extractEmail(raw);
  if(email && MAP_SUPERVISOR[email]) return MAP_SUPERVISOR[email];

  const norm = normalizeEmailOrText(raw);
  if(MAP_SUPERVISOR[norm]) return MAP_SUPERVISOR[norm];

  const clean = cleanText(raw);
  if(MAP_SUPERVISOR[clean]) return MAP_SUPERVISOR[clean];

  return {asesor: raw, area:'Sin área', grupo:'', supervisor:'Sin supervisor'};
}

function mentorDeAsesor(asesor){
  return lookupSupervisorInfo(asesor).supervisor || 'Sin supervisor';
}

function areaDeAsesor(asesor){
  return lookupSupervisorInfo(asesor).area || 'Sin área';
}

function grupoDeAsesor(asesor){
  return lookupSupervisorInfo(asesor).grupo || '';
}
