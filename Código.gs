const SPREADSHEET_ID = '1DuyYpJUbYOaVKuhfmUAia1M4pJNI5CErX-DZ7FBXuhU';

const SHEET_PRESUPUESTOS = '05_PRESUPUESTO_SITES';
const SHEET_FTES = '06_FTES_SITES';
const SHEET_FTES_PERSONAS = '07_FTSs_SITES_2';
const SHEET_FTES_KPIS = '08_FTSs_SITES_3';

// KPIs
const SHEET_KPIS = '07_KPIS_SITES';
const SHEET_CATALOGO_KPIS = '11_CATALOGO_KPIs';

// Archivo HTML único donde están Presupuestos y FTEs
const MAIN_INDEX_FILE = 'Index_Presupuestos';

/* =========================================================
   ROUTER PRINCIPAL
   ========================================================= */

function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};

  /* =========================================================
     MODO API PARA GOOGLE SITES / HTML PEGADO DIRECTAMENTE
     Ejemplo:
     .../exec?api=1&action=getPresupuestoData&callback=miCallback
     ========================================================= */
  if (String(e.parameter.api || '') === '1') {
    return handleApiRequest_(e);
  }

  /* =========================================================
     MODO HTML ACTUAL
     Mantiene funcionando tus páginas actuales desde Apps Script:
     - Hub principal
     - KPIs
     - Presupuestos
     - FTEs
     ========================================================= */
  const page = e && e.parameter && e.parameter.page
    ? String(e.parameter.page).trim().toLowerCase()
    : 'ftes';

  const appUrl = getWebAppUrl_();

  if (page === 'kpis') {
    return HtmlService
      .createHtmlOutputFromFile('Index_KPIs')
      .setTitle('KPIs EDC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'catalogo' || page === 'catalogo_kpi' || page === 'catalogo_kpis') {
    return HtmlService
      .createHtmlOutputFromFile('Index_catalogo_KPIs')
      .setTitle('Layout de KPIs EDC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'presupuestos' || page === 'presupuesto' || page === 'ftes') {
    const initialPage = page === 'ftes' ? 'ftes' : 'presupuesto';
    const template = HtmlService.createTemplateFromFile('Index_Presupuestos');
    template.initialPage = initialPage;

    return template
      .evaluate()
      .setTitle(initialPage === 'ftes' ? 'FTEs' : 'Presupuestos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile(MAIN_INDEX_FILE);
  template.initialPage = 'ftes';
  template.appUrl = appUrl;

  return template
    .evaluate()
    .setTitle('Global Transformation Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (error) {
    return '';
  }
}


/* =========================================================
   API JSONP PARA GOOGLE SITES / HTML DIRECTO
   ========================================================= */

function handleApiRequest_(e) {
  const action = String(e.parameter.action || '').trim();

  try {
    let data;

    if (action === 'ping') {
      data = {
        ok: true,
        message: 'API funcionando correctamente',
        timestamp: new Date().toISOString()
      };

    } else if (action === 'getCatalogoKpisData' || action === 'getCatalogoData') {
      data = getCatalogoKpisData();

    } else if (action === 'getPresupuestoData') {
      data = getPresupuestoData();

    } else if (action === 'getFtesData' || action === 'getFTEsData') {
      data = getFtesData();

    } else if (action === 'getFtesPersonasData') {
      data = getFtesPersonasData();

    } else if (action === 'getFtesKpisData') {
      data = getFtesKpisData();

    } else if (action === 'getKpisData') {
      data = getKpisData();

    } else if (
      action === 'getAllFtesPresupuestoData' ||
      action === 'getAllPresupuestoFtesData' ||
      action === 'getAllPresupuestosFtesData'
    ) {
      data = {
        presupuesto: getPresupuestoData(),
        ftes: getFtesData(),
        ftesPersonas: getFtesPersonasData(),
        ftesKpis: getFtesKpisData()
      };

    } else {
      throw new Error('Acción API no reconocida: ' + action);
    }

    return jsonpResponse_(e, {
      status: 'ok',
      action: action,
      data: data
    });

  } catch (error) {
    return jsonpResponse_(e, {
      status: 'error',
      action: action,
      message: error && error.message ? error.message : String(error)
    });
  }
}

function jsonpResponse_(e, payload) {
  const callback = String(e.parameter.callback || '').trim();
  const json = JSON.stringify(payload);

  if (callback) {
    if (!isValidCallbackName_(callback)) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'error',
          message: 'Nombre de callback no válido'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function isValidCallbackName_(callback) {
  return /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback);
}

/* =========================================================
   PRESUPUESTOS
   Pestaña: 05_PRESUPUESTO_SITES
   ========================================================= */

function getPresupuestoData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PRESUPUESTOS);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_PRESUPUESTOS);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      return {
        pais: String(getValue_(obj, ['pais']) || '').trim().toUpperCase(),
        anio: getValue_(obj, ['ano', 'anio']),
        quarter: String(getValue_(obj, ['quarter']) || '').trim().toUpperCase(),
        tipoFila: String(getValue_(obj, ['tipo_fila', 'tipofila']) || '').trim(),
        vistaMoneda: String(getValue_(obj, ['vista_moneda', 'vistamoneda']) || '').trim().toUpperCase(),
        moneda: String(getValue_(obj, ['moneda']) || '').trim().toUpperCase(),

        referencia2026: toNumber_(getValue_(obj, [
          'referencia_2026',
          'referencia2026'
        ])),

        referencia: toNumber_(getValue_(obj, [
          'referencia'
        ])),

        demandado: toNumber_(getValue_(obj, [
          'demandado'
        ])),

        autorizado: toNumber_(getValue_(obj, [
          'autorizado'
        ])),

        ejecutado: toNumber_(getValue_(obj, [
          'ejecutado'
        ]))
      };
    })
    .filter(row => row.pais && row.quarter);

  return rows;
}

/* =========================================================
   FTES PRINCIPAL
   Pestaña: 06_FTES_SITES
   ========================================================= */

function getFtesData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FTES);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_FTES);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      const internosParticipadasRaw =
        getValue_(obj, [
          'internos_participadas',
          'internosparticipadas',
          'internos_participadas_',
          'internos_participadas_pct',
          'pct_internos_participadas'
        ]) || row[11];

      const externosRaw =
        getValue_(obj, [
          'externos',
          'externos_pct',
          'pct_externos'
        ]) || row[12];

      // Nuevas columnas 06_FTES_SITES:
      // Columna N = INGENIERIA / INGENIERÍA
      // Columna O = NEGOCIO
      const ingenieriaRaw =
        getValue_(obj, [
          'ingenieria',
          'ingenieria_',
          'ingenieria_fte',
          'ingenieria_ftes',
          'ing',
          'fte_ingenieria',
          'ftes_ingenieria',
          'asignado_ingenieria',
          'asignados_ingenieria',
          'perfiles_ingenieria',
          'engineering',
          'engineering_assigned',
          'assigned_engineering'
        ]) || row[13];

      const negocioRaw =
        getValue_(obj, [
          'negocio',
          'negocio_fte',
          'negocio_ftes',
          'neg',
          'fte_negocio',
          'ftes_negocio',
          'asignado_negocio',
          'asignados_negocio',
          'perfiles_negocio',
          'business',
          'business_assigned',
          'assigned_business'
        ]) || row[14];

      const internosParticipadas = toPercentDecimal_(internosParticipadasRaw);
      const externos = toPercentDecimal_(externosRaw);
      const ingenieria = toNumber_(ingenieriaRaw);
      const negocio = toNumber_(negocioRaw);

      return {
        pais: String(getValue_(obj, ['pais']) || '').trim().toUpperCase(),
        anio: getValue_(obj, ['ano', 'anio']),
        quarter: String(getValue_(obj, ['quarter']) || '').trim().toUpperCase(),
        tipoFila: String(getValue_(obj, ['tipo_fila', 'tipofila']) || '').trim(),

        referencia2026: toNumber_(getValue_(obj, [
          'referencia_2026',
          'referencia2026'
        ])),

        demand: toNumber_(getValue_(obj, [
          'demand'
        ])),

        asig: toNumber_(getValue_(obj, [
          'asig'
        ])),

        demandQ2Q1: toNumber_(getValue_(obj, [
          'demand_q2q1',
          'demand_q2_q1'
        ])),

        demandQ3Q2: toNumber_(getValue_(obj, [
          'demand_q3q2',
          'demand_q3_q2'
        ])),

        demandQ4Q3: toNumber_(getValue_(obj, [
          'demand_q4q3',
          'demand_q4_q3'
        ])),

        internosParticipadas: internosParticipadas,
        externos: externos,

        pctInternosParticipadas: internosParticipadas,
        pctExternos: externos,

        // Desglose de FTEs asignados por tipo de perfil.
        // Se devuelven varios alias para que el HTML pueda leerlos aunque cambie el nombre.
        ingenieria: ingenieria,
        ingeniería: ingenieria,
        ing: ingenieria,
        ftesIngenieria: ingenieria,
        ftesIngeniería: ingenieria,
        asignadoIngenieria: ingenieria,
        asignadoIngeniería: ingenieria,
        perfilesIngenieria: ingenieria,
        perfilesIngeniería: ingenieria,
        columnaN: ingenieria,

        negocio: negocio,
        neg: negocio,
        ftesNegocio: negocio,
        asignadoNegocio: negocio,
        perfilesNegocio: negocio,
        columnaO: negocio
      };
    })
    .filter(row => row.pais && row.quarter);

  return rows;
}

// Alias para compatibilidad con HTML antiguos que llaman getFTEsData()
function getFTEsData() {
  return getFtesData();
}

/* =========================================================
   FTES PERSONAS
   Pestaña: 07_FTSs_SITES_2
   ========================================================= */

function getFtesPersonasData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FTES_PERSONAS);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_FTES_PERSONAS);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      const paisRaw = getValue_(obj, ['pais']) || row[0];
      const quarterRaw = getValue_(obj, ['quarter']) || row[1];

      const ftesRaw =
        getValue_(obj, [
          'ftes',
          'fte'
        ]) || row[2];

      const personasRaw =
        getValue_(obj, [
          'personas',
          'persona'
        ]) || row[3];

      const personas100Raw =
        getValue_(obj, [
          'personas_al_100',
          'personas_100',
          'pct_personas_al_100',
          'personas_al_100_pct',
          'porcentaje_personas_al_100'
        ]) || row[4];

      const ratioRaw =
        getValue_(obj, [
          'ratio_de_dedicacion',
          'ratio_dedicacion',
          'dedicacion'
        ]) || row[5];

      return {
        pais: String(paisRaw || '').trim().toUpperCase(),
        quarter: String(quarterRaw || '').trim().toUpperCase(),

        ftes: toNumber_(ftesRaw),
        personas: toNumber_(personasRaw),

        personasAl100: toPercentDecimal_(personas100Raw),
        pctPersonasAl100: toPercentDecimal_(personas100Raw),

        ratioDedicacion: toNumber_(ratioRaw),
        ratio: toNumber_(ratioRaw)
      };
    })
    .filter(row => row.pais && row.quarter);

  return rows;
}

/* =========================================================
   FTES KPIS
   Pestaña: 08_FTSs_SITES_3
   ========================================================= */

function getFtesKpisData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FTES_KPIS);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_FTES_KPIS);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      const paisRaw = getValue_(obj, ['pais']) || row[0];
      const quarterRaw = getValue_(obj, ['quarter']) || row[1];

      const productividadRaw =
        getValue_(obj, ['productividad']) || row[2];

      const targetProductividadRaw =
        getValue_(obj, [
          'target_productividad',
          'targetproductividad'
        ]) || row[3];

      const desvProductividadRaw =
        getValue_(obj, [
          'desv_productividad',
          'desviacion_productividad',
          'pct_desv_productividad'
        ]) || row[4];

      const ltRaw =
        getValue_(obj, ['lt']) || row[5];

      const targetLtRaw =
        getValue_(obj, [
          'target_lt',
          'targetlt'
        ]) || row[6];

      const desvLtRaw =
        getValue_(obj, [
          'desv_lt',
          'desviacion_lt'
        ]) || row[7];

      const ctRaw =
        getValue_(obj, ['ct']) || row[8];

      const targetCtRaw =
        getValue_(obj, [
          'target_ct',
          'targetct'
        ]) || row[9];

      const desvCtRaw =
        getValue_(obj, [
          'desv_ct',
          'desviacion_ct',
          'pct_desv_ct'
        ]) || row[10];

      return {
        pais: String(paisRaw || '').trim().toUpperCase(),
        quarter: String(quarterRaw || '').trim().toUpperCase(),

        productividad: toNumber_(productividadRaw),
        targetProductividad: toNumber_(targetProductividadRaw),
        desvProductividad: toPercentDecimal_(desvProductividadRaw),
        pctDesvProductividad: toPercentDecimal_(desvProductividadRaw),

        lt: toNumber_(ltRaw),
        targetLt: toNumber_(targetLtRaw),
        desvLt: toNumber_(desvLtRaw),

        ct: toNumber_(ctRaw),
        targetCt: toNumber_(targetCtRaw),
        desvCt: toPercentDecimal_(desvCtRaw),
        pctDesvCt: toPercentDecimal_(desvCtRaw)
      };
    })
    .filter(row => row.pais && row.quarter);

  return rows;
}

/* =========================================================
   KPIS
   Pestaña: 07_KPIS_SITES
   ========================================================= */

function getKpisData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_KPIS);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_KPIS);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      const kpiRaw =
        getValue_(obj, [
          'kpi',
          'metric',
          'metrica'
        ]) || row[1];

      const paisRaw =
        getValue_(obj, [
          'pais',
          'bu',
          'country'
        ]) || row[2];

      const tipoRaw =
        getValue_(obj, [
          'tipo',
          'status_type',
          'statustype'
        ]) || row[3];

      const unidadRaw =
        getValue_(obj, [
          'unidad',
          'unit'
        ]) || row[4];

      const anioRaw =
        getValue_(obj, [
          'ano',
          'anio',
          'year'
        ]) || row[5];

      const quarterRaw =
        getValue_(obj, [
          'quarter',
          'q'
        ]) || row[6];

      const mesRaw =
        getValue_(obj, [
          'mes',
          'month'
        ]) || row[7];

      const mesNumRaw =
        getValue_(obj, [
          'mes_n',
          'mes_no',
          'mes_num',
          'mes_numero',
          'month_number'
        ]) || row[8];

      const valorRaw =
        getValue_(obj, [
          'valor',
          'value',
          'avance',
          'pct_avance'
        ]) || row[12];

      const targetMensualRaw =
        getValue_(obj, [
          'target_mensual',
          'targetmensual',
          'target_monthly'
        ]) || row[13];

      const targetAnualRaw =
        getValue_(obj, [
          'target_anual',
          'targetanual',
          'target_year'
        ]) || row[14];

      const valor = toKpiNumber_(valorRaw);
      const targetMensual = toKpiNumber_(targetMensualRaw);
      const targetAnual = toKpiNumber_(targetAnualRaw);

      return {
        rawRow: toNumber_(getValue_(obj, ['rawrow']) || row[0]),

        kpi: String(kpiRaw || '').trim(),
        pais: normalizeCountry_(paisRaw),
        paisOriginal: String(paisRaw || '').trim(),

        tipo: String(tipoRaw || '').trim(),
        unidad: String(unidadRaw || '').trim(),

        anio: toNumber_(anioRaw),
        year: toNumber_(anioRaw),

        quarter: String(quarterRaw || '').trim().toUpperCase(),
        mes: String(mesRaw || '').trim(),
        mesNum: toNumber_(mesNumRaw),

        valor: valor,
        targetMensual: targetMensual,
        targetAnual: targetAnual,

        tieneValor: hasKpiValue_(valorRaw),
        tieneTargetMensual: hasKpiValue_(targetMensualRaw),
        tieneTargetAnual: hasKpiValue_(targetAnualRaw)
      };
    })
    .filter(row => row.kpi && row.pais && row.anio && row.quarter && row.mes);

  return rows;
}

/* =========================================================
   TESTS OPCIONALES
   ========================================================= */

function testPresupuestoData() {
  const data = getPresupuestoData();
  Logger.log(JSON.stringify(data.slice(0, 15), null, 2));
}

function testFtesData() {
  const data = getFtesData();
  Logger.log(JSON.stringify(data.slice(0, 20), null, 2));
}

function testFtesPersonasData() {
  const data = getFtesPersonasData();
  Logger.log(JSON.stringify(data.slice(0, 25), null, 2));
}

function testFtesKpisData() {
  const data = getFtesKpisData();
  Logger.log(JSON.stringify(data.slice(0, 25), null, 2));
}

function testKpisData() {
  const data = getKpisData();
  Logger.log(JSON.stringify(data.slice(0, 50), null, 2));
}

/* =========================================================
   HELPERS GENERALES
   ========================================================= */

function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '_')
    .replace(/\+/g, '_')
    .replace(/-/g, '_')
    .replace(/\./g, '')
    .replace(/%/g, 'pct')
    .replace(/[^a-zA-Z0-9_]/g, '')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .toLowerCase();
}

function getValue_(obj, keys) {
  for (let i = 0; i < keys.length; i++) {
    if (
      obj[keys[i]] !== undefined &&
      obj[keys[i]] !== null &&
      obj[keys[i]] !== ''
    ) {
      return obj[keys[i]];
    }
  }

  return '';
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  let text = String(value)
    .trim()
    .replace(/[€%\s]/g, '');

  if (
    text === '#DIV/0!' ||
    text === '#N/A' ||
    text === '#VALUE!' ||
    text === '#REF!' ||
    text === '#ERROR!' ||
    text === '-'
  ) {
    return 0;
  }

  if (text.includes(',') && text.includes('.')) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (text.includes(',')) {
    text = text.replace(',', '.');
  }

  const number = Number(text);

  return isNaN(number) ? 0 : number;
}

function toPercentDecimal_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number') {
    return value > 1 ? value / 100 : value;
  }

  let text = String(value)
    .trim()
    .replace('%', '')
    .replace(/\s/g, '');

  if (
    text === '#DIV/0!' ||
    text === '#N/A' ||
    text === '#VALUE!' ||
    text === '#REF!' ||
    text === '#ERROR!' ||
    text === '-'
  ) {
    return 0;
  }

  if (text.includes(',') && text.includes('.')) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (text.includes(',')) {
    text = text.replace(',', '.');
  }

  const number = Number(text);

  if (isNaN(number)) {
    return 0;
  }

  return number > 1 ? number / 100 : number;
}

function toBoolean_(value) {
  if (value === true || value === 1) {
    return true;
  }

  if (
    value === false ||
    value === 0 ||
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return false;
  }

  const normalized = String(value)
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();

  return [
    'TRUE',
    'VERDADERO',
    'SI',
    'S',
    'YES',
    'Y',
    '1'
  ].includes(normalized);
}


/* =========================================================
   HELPERS KPIS
   ========================================================= */

function normalizeCountry_(value) {
  const text = String(value || '').trim();

  const upper = text
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();

  const map = {
    'TOTAL': 'TOTAL',
    'GLOBAL': 'TOTAL',
    'ALL': 'TOTAL',

    'HOLDING': 'HOLDING',

    'SPA': 'ESPAÑA',
    'ESPANA': 'ESPAÑA',
    'ESPAÑA': 'ESPAÑA',
    'SPAIN': 'ESPAÑA',

    'MEX': 'MÉXICO',
    'MEXICO': 'MÉXICO',
    'MÉXICO': 'MÉXICO',

    'PER': 'PERÚ',
    'PERU': 'PERÚ',
    'PERÚ': 'PERÚ',

    'COL': 'COLOMBIA',
    'COLOMBIA': 'COLOMBIA',

    'ARG': 'ARGENTINA',
    'ARGENTINA': 'ARGENTINA'
  };

  return map[upper] || upper;
}

function hasKpiValue_(value) {
  if (value === null || value === undefined || value === '') {
    return false;
  }

  const text = String(value).trim();

  if (
    text === '-' ||
    text === '#DIV/0!' ||
    text === '#N/A' ||
    text === '#VALUE!' ||
    text === '#REF!' ||
    text === '#ERROR!'
  ) {
    return false;
  }

  return true;
}

function toKpiNumber_(value) {
  if (!hasKpiValue_(value)) {
    return null;
  }

  return toNumber_(value);
}



/* =========================================================
   LAYOUT DE KPIS (CORREGIDO)
   Pestaña: 11_CATALOGO_KPIs
   ========================================================= */

function getCatalogoKpisData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CATALOGO_KPIS);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_CATALOGO_KPIS);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const rows = values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      const id = String(getValue_(obj, ['id', 'id_kpi', 'codigo', 'cod']) || row[0] || '').trim();
      const pais = normalizeCountry_(getValue_(obj, ['pais', 'country', 'geografia']) || row[1] || '');
      const bloque = String(getValue_(obj, ['bloque_funcional', 'bloquefuncional', 'bloque', 'categoria', 'category']) || row[2] || '').trim();
      const solucion = String(getValue_(obj, ['solucion_funcional', 'solucionfuncional', 'solucion', 'nombre', 'kpi', 'indicador']) || row[3] || '').trim();
      const qTarget = getValue_(obj, ['q_target', 'qtarget', 'target', 'target_referencia', 'target_q']) || row[4] || '';
      // Q Real:
      // Columna F de 11_CATALOGO_KPIs.
      // Si la columna F tiene cualquier valor, la funcionalidad cuenta como entregada.
      const qRealRaw = row[5];
      const qReal = qRealRaw === null || qRealRaw === undefined ? '' : qRealRaw;

      // ETPB:
      // Columna G de 11_CATALOGO_KPIs.
      // Se conserva también el valor bruto para diferenciar FALSE explícito de celda vacía.
      const etpbRaw = row[6];
      const etpb = toBoolean_(etpbRaw);
    // Ongoing Akira:
      // Columna H de 11_CATALOGO_KPIs.
      // Si H = TRUE => se cuenta como Ongoing Akira.
      // Si H está vacío o FALSE => no se cuenta.
      const ongoingAkiraRaw = row[7];
      const ongoingAkira = toBoolean_(ongoingAkiraRaw);

      const r5 = toBoolean_(
        getValue_(obj, [
          'r5',
          'robot_5',
          'robot5',
          'robot_5_bei',
          'automatizado',
          'automatizada'
        ]) !== ''
          ? getValue_(obj, [
              'r5',
              'robot_5',
              'robot5',
              'robot_5_bei',
              'automatizado',
              'automatizada'
            ])
          : row[8]
      );

      return {
        id: id,
        pais: pais,
        bloqueFuncional: bloque,
        solucionFuncional: solucion,
        qTarget: qTarget,
        qReal: qReal,
        qRealRaw: qRealRaw,
        etpb: etpb,
        etpbRaw: etpbRaw,
        ongoingAkira: ongoingAkira,
        ongoingAkiraRaw: ongoingAkiraRaw,
        r5: r5,
        robot5: r5,
        automatizado: r5
      };
    })
    .filter(row => row.id || row.solucionFuncional || row.bloqueFuncional);

  return rows;
}
