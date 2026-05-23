/* ════════════════════════════════════════════
   NORMALIZAR FILA — replica fórmulas del Excel
   ════════════════════════════════════════════ */
function normalizeRow(row, headers) {
  // Lee columnas aunque el encabezado venga con mayúsculas, espacios dobles o tildes.
  const hIndex = new Map(headers.map((h, i) => [cleanHeader(h), i]));
  const g = (...names) => {
    for (const name of names) {
      const i = hIndex.get(cleanHeader(name));
      if (i !== undefined && row[i] != null) return String(row[i]).trim();
    }
    return '';
  };

  // ID: primero intenta por nombre de encabezado; si no lo encuentra, toma la primera columna del Excel.
  // Esto corrige bases donde el encabezado del ID llega con caracteres ocultos o nombre diferente.
  const colA = g(
          'ID de registro',
          'Id de registro',
          'ID Registro',
          'Id Registro',
          'ID',
          'Id',
          'Record ID',
          'Record Id',
          'Lead ID',
          'Lead Id'
        ) || String(row[0] ?? '').trim(),
        colB=g('Hora de creación','Hora de creacion','Created Time'),
        colC=g('Nombre completo','Full Name'),
        colD=g('Teléfono','Telefono','Phone','Mobile'),
        colE=g('Correo electrónico','Correo electronico','Email'),
        colF=g('Número de Documento','Numero de Documento','Documento'),
        colG=g('Modalidad'),
        colH=g('Periodo'),
        colI=g('Programa de interes_','Programa de interés_','Programa de interes','Programa de interés'),
        colJ=g('Propietario de Posible Cliente','Propietario de posible cliente','Lead Owner'),
        colK=g('Creado por','Created By'),
        colL=g('Campaña mercadeo','Campana mercadeo','Campaña de mercadeo'),
        colM=g('Etapa del Registro'),
        colN=g('Estado de Posible Cliente'),
        colO=g('Hora de modificación','Hora de modificacion'),
        colP=g('Ultimo nivel de escolaridad','Último nivel de escolaridad'),
        colQ=g('Editado por'),
        colR=g('Periodo Creado'),
        colS=g('Correo electrónico Continua','Correo electronico Continua','Email Continua'),
        colT=g('Teléfono Continua','Telefono Continua','Phone Continua'),
        colU=g('Ciudad de residencia'),
        colV=g('Sub Estado'),
        colW=g('Etiqueta'),
        colX=g('Ciudad utm','Ciudad UTM'),
        colY=g('utm_campaign','UTM Campaign'),
        colZ=g('utm_content','UTM Content'),
        colAA=g('utm_medium','UTM Medium'),
        colAB=g('utm_source','UTM Source'),
        colAC=g('utm_term','UTM Term'),
        colAD=g('conv','Conv'),
        colAE=g('Total Notas'),
        colAF=g('Descripción','Descripcion'),
        colAG=g('Descripción actualizacion','Descripcion actualizacion','Descripción actualización'),
        colAH=g('Programa');

  /* CORREO DEF: =SI(E="";S;E) */
  const correo = (colE && colE !== '0') ? colE : colS;

  /* NÚMERO DEF: usa Teléfono; si viene vacío, usa Teléfono Continua. Deja solo dígitos y toma los últimos 10. */
  const telSrc = (colD && colD !== '0') ? colD : colT;
  const numTel = normalizarTelefono(telSrc);

  /* ASESOR */
  const asesor = lookup(MAP_ASESOR, colJ, colJ);

  /* PROGRAMA2 */
  const progSrc = (colAH && colAH !== '0') ? colAH : colI;
  const progKey = (!progSrc || progSrc === '0') ? 'SIN PROGRAMA' : progSrc;
  const programa = lookup(MAP_PROGRAMA, progKey, cleanText(progKey));

  /* CIUDAD */
  const ciudadKey = (!colX || colX === '0') ? colU : colX;
  const ciudadFin = (!ciudadKey || ciudadKey === '0') ? 'SIN CIUDAD' : ciudadKey;
  const ciudadRaw = lookup(MAP_CIUDAD, ciudadFin, ciudadFin);
  const ciudad = cleanText(ciudadRaw);

  /* AREA PROGRAMA */
  const area = cleanText(lookup(MAP_AREA, programa, 'OTRO'));

  /* CAMPAÑA */
  let campana;
  if      (area === 'PRESENCIAL')       campana = 'PRESENCIAL';
  else if (/ESPEC/i.test(programa))     campana = 'POSGRADO';
  else if (/v/i.test(colAD))            campana = 'VIRTUAL';
  else if (/S/.test(colAD))             campana = 'POSGRADO';
  else if (/T/.test(colAD))             campana = 'POSGRADO';
  else if (/2026/.test(colAD))          campana = 'PRESENCIAL';
  else if (/2025/.test(colAD))          campana = 'PRESENCIAL';
  else                                  campana = 'OTRO';

  /* LLAVE */
  const llaveKey = `${campana} - ${programa} - ${ciudad}`;
  const llave = normalizarLlave(lookup(MAP_LLAVE, llaveKey, '-'));

  // Debug útil para verificar la fórmula: AP = Llave, AN = Área
  console.log('VALIDACION FINAL DEBUG', { AP_llave: llave, AN_area: area });

  /* VALIDACION FINAL: igual a la fórmula Excel */
  const validacion = calcularValidacionFinal(llave, area);

  /* NOVEDAD */
  const n2 = parseInt(numTel.substring(0, 2), 10);
  let novedad;
  if (!numTel || numTel === '0' || Number(numTel) === 0) novedad = 'Sin número';
  else if (numTel.length < 10)                           novedad = 'Número errado';
  else if (n2 < 30 || n2 > 35)                            novedad = 'Número errado';
  else if (/cun\.edu\.co/i.test(correo))                 novedad = 'Estudiante activo';
  else if (/PROYECCION SOCIAL/i.test(colL))               novedad = 'Campaña erronea';
  else if (/Recunpensa/i.test(colL))                      novedad = 'Recunpensa';
  else if (/MINTIC/i.test(colL))                          novedad = 'Campaña erronea';
  else if (/Espec/i.test(colI))                           novedad = 'POSGRADO';
  else if (/Esp/i.test(colY))                             novedad = 'POSGRADO';
  else                                                    novedad = 'Normal';

  return {
    'ID de registro':colA,'Hora de creación':colB,'Nombre completo':colC,
    'Teléfono':colD,'Correo electrónico':colE,'Número de Documento':colF,
    'Modalidad':colG,'Periodo':colH,'Programa de interes_':colI,
    'Propietario de Posible Cliente':colJ,'Creado por':colK,'Campaña mercadeo':colL,
    'Etapa del Registro':colM,'Estado de Posible Cliente':colN,
    'Hora de modificación':colO,'Ultimo nivel de escolaridad':colP,
    'Editado por':colQ,'Periodo Creado':colR,'Correo electrónico Continua':colS,
    'Teléfono Continua':colT,'Ciudad de residencia':colU,'Sub Estado':colV,
    'Etiqueta':colW,'Ciudad utm':colX,'utm_campaign':colY,'utm_content':colZ,
    'utm_medium':colAA,'utm_source':colAB,'utm_term':colAC,'conv':colAD,
    'Total Notas':colAE,'Descripción':colAF,'Descripción actualizacion':colAG,
    'Programa':colAH,
    'CAMPAÑA':campana,'CORREO ELECTRONICO DEF':correo,'Número DEF':numTel,
    'Asesor':asesor,'CIUDAD':ciudad,'AREA PROGRAMA':area,'PROGRAMA2':programa,
    'Llave':llave,'VALIDACION FINAL':validacion,'NOVEDAD':novedad,
    'MENTOR': (typeof supervisorSimpleLookup==='function' ? supervisorSimpleLookup(asesor) : ''),'DUPLICADO':''
  };
}



function calcularValidacionFinal(llave, area) {
  // Fórmula original de Excel:
  // =SI(Y(AP2="-";AN2="PRESENCIAL");"PRESENCIAL";
  //   SI(Y(AP2="-";O(AN2="MIXTO";AN2="OTRO";AN2="VIRTUAL"));"VIRTUAL";
  //   SI(AP2="PRESENCIAL";"PRESENCIAL";
  //   SI(AN2="PRESENCIAL";"PRESENCIAL";"POSGRADO"))))
  // AP2 = Llave
  // AN2 = AREA PROGRAMA

  const AP = normalizarLlave(llave);
  const AN = cleanText(area || 'OTRO');

  if (AP === '-' && AN === 'PRESENCIAL') {
    return 'PRESENCIAL';
  }

  if (AP === '-' && (AN === 'MIXTO' || AN === 'OTRO' || AN === 'VIRTUAL')) {
    return 'VIRTUAL';
  }

  if (AP === 'PRESENCIAL') {
    return 'PRESENCIAL';
  }

  if (AN === 'PRESENCIAL') {
    return 'PRESENCIAL';
  }

  return 'POSGRADO';
}

function cleanHeader(s = '') {
  return cleanText(s)
    .replace(/[^A-Z0-9]/g, '');
}

function normalizarTelefono(value = '') {
  let s = String(value ?? '').trim();

  // Si Excel lo entrega como texto tipo "3.00123E+09", intentamos convertirlo.
  if (/^\d+(\.\d+)?E\+?\d+$/i.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) s = Math.trunc(n).toString();
  }

  const digits = s.replace(/\D/g, '');
  if (!digits) return '';
  return digits.length >= 10 ? digits.slice(-10) : digits;
}


function normalizarLlave(value = '-') {
  const raw = String(value ?? '').trim();
  // IMPORTANTÍSIMO: en Excel AP2 puede ser exactamente "-".
  // cleanText() convierte guiones en espacios, por eso NO se debe usar directo para AP.
  if (!raw || raw === '-') return '-';
  return raw
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}
