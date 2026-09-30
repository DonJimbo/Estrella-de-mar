/* =========================================================
   APLICACIÓN WEB UNIFICADA — Global Transformation Hub
   - Páginas embebibles para Google Sites.
   - API GTH: ?api=1&action=...
   - API Dashboard EDC: ?action=...
   ========================================================= */

function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  const params = e.parameter;
if (params.vista) {
  return renderUnifiedSitePage_(params.vista);
}
  // API de presupuestos, FTEs, KPIs, catálogo y adopción.
  if (String(params.api || '') === '1') {
    return handleApiRequest_(e);
  }

  // Conserva tanto las llamadas actuales con action como la llamada antigua
  // que utilizaba únicamente callback para obtener el snapshot EDC.
  if (params.action || params.callback) {
    return handleDashboardEDCApi_(e);
  }

return renderUnifiedSitePage_('inicio');
}

const SPREADSHEET_ID = '1DuyYpJUbYOaVKuhfmUAia1M4pJNI5CErX-DZ7FBXuhU';

const SHEET_PRESUPUESTOS = '05_PRESUPUESTO_SITES';
const SHEET_FTES = '06_FTES_SITES';
const SHEET_FTES_PERSONAS = '07_FTSs_SITES_2';
const SHEET_FTES_KPIS = '08_FTSs_SITES_3';

// KPIs
const SHEET_KPIS = '07_KPIS_SITES';
const SHEET_CATALOGO_KPIS = '11_CATALOGO_KPIs';
const SHEET_OPCIONES_MENU_SITES = '15_OPCIONES_MENU_SITES';

// Archivo HTML único donde están Presupuestos y FTEs
const MAIN_INDEX_FILE = 'Index_Presupuestos';

/* =========================================================
   ROUTER PRINCIPAL
   ========================================================= */

function renderGthPage_(e) {
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

function handleApiRequest_(e) { return gthApiDispatch_(e); }

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

    'ES': 'ESPAÑA',
    'SPA': 'ESPAÑA',
    'ESPANA': 'ESPAÑA',
    'ESPAÑA': 'ESPAÑA',
    'SPAIN': 'ESPAÑA',

    'MX': 'MÉXICO',
    'MEX': 'MÉXICO',
    'MEXICO': 'MÉXICO',
    'MÉXICO': 'MÉXICO',

    'PE': 'PERÚ',
    'PER': 'PERÚ',
    'PERU': 'PERÚ',
    'PERÚ': 'PERÚ',

    'CO': 'COLOMBIA',
    'COL': 'COLOMBIA',
    'COLOMBIA': 'COLOMBIA',

    'AR': 'ARGENTINA',
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
  if (values.length < 2) return [];

  // La primera fila de datos no siempre es la fila 1. Detectamos la cabecera
  // para conservar el funcionamiento aunque se añada un título al Excel.
  const headerRowIndex = findCatalogoHeaderRow_(values);
  const headers = (values[headerRowIndex] || []).map(normalizeHeader_);
  const menuOptionsByKey = getMenuOptionsByKpi_(ss);

  return values.slice(headerRowIndex + 1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => { if (header) obj[header] = row[index]; });

      // Estructura confirmada del catálogo: A país, B KPI ID, C bloque,
      // D solución funcional, E descripción y L comentario de estado.
      const id = String(getCatalogoValue_(obj, row, ['id_kpi', 'kpi_id', 'id', 'codigo', 'cod'], 1) || '').trim();
      const paisRaw = getCatalogoValue_(obj, row, ['pais', 'country', 'geografia'], 0);
      const pais = normalizeCountry_(paisRaw);
      const bloque = String(getCatalogoValue_(obj, row, ['bloque_funcional', 'bloquefuncional', 'bloque', 'categoria', 'category'], 2) || '').trim();
      const solucion = String(getCatalogoValue_(obj, row, ['solucion_funcional', 'solucionfuncional', 'solucion', 'nombre', 'kpi', 'indicador'], 3) || '').trim();
      const descripcion = String(getCatalogoValue_(obj, row, ['descripcion', 'descripcion_solucion', 'detalle', 'description'], 4) || '').trim();
      const qTarget = getCatalogoValue_(obj, row, ['q_target', 'qtarget', 'target', 'target_referencia', 'target_q', 'trimestre_target'], -1);
      const qRealRaw = getCatalogoValue_(obj, row, ['q_real', 'qreal', 'entrega', 'fecha_entrega', 'trimestre_entrega'], 6);
      const etpbRaw = getCatalogoValue_(obj, row, ['etpb'], 7);
      const ongoingAkiraRaw = getCatalogoValue_(obj, row, ['ongoing_akira', 'on_going_akira', 'akira'], 8);
      const r5Raw = getCatalogoValue_(obj, row, ['r5', 'robot_5', 'robot5', 'robot_5_bei', 'automatizado', 'automatizada'], 9);
      const comentario = String(getCatalogoValue_(obj, row, ['comentario', 'comentarios', 'observaciones', 'estado_comentario', 'nota_estado'], 12) || '').trim();
      const key = buildMenuOptionKey_(pais, id);

      return {
        id: id,
        pais: pais,
        paisOriginal: String(paisRaw || '').trim(),
        bloqueFuncional: bloque,
        solucionFuncional: solucion,
        descripcion: descripcion,
        qTarget: qTarget || '',
        qReal: qRealRaw === null || qRealRaw === undefined ? '' : qRealRaw,
        qRealRaw: qRealRaw,
        etpb: toBoolean_(etpbRaw),
        etpbRaw: etpbRaw,
        ongoingAkira: toBoolean_(ongoingAkiraRaw),
        ongoingAkiraRaw: ongoingAkiraRaw,
        r5: toBoolean_(r5Raw),
        r5Raw: r5Raw,
        robot5: toBoolean_(r5Raw),
        automatizado: toBoolean_(r5Raw),
        comentario: comentario,
        // Se conserva highlights para no romper ninguna vista que ya lo use.
        highlights: descripcion,
        menuOptions: menuOptionsByKey[key] || []
      };
    })
    .filter(row => row.id || row.solucionFuncional || row.bloqueFuncional);
}

/** Identifica la fila de cabeceras de 11_CATALOGO_KPIs de forma tolerante. */
function findCatalogoHeaderRow_(values) {
  let bestIndex = 0;
  let bestScore = -1;

  values.slice(0, 12).forEach((row, index) => {
    const headers = row.map(normalizeHeader_);
    let score = 0;
    if (headers.some(header => ['pais', 'country', 'geografia'].indexOf(header) >= 0)) score += 2;
    if (headers.some(header => header === 'id' || header === 'id_kpi' || header === 'kpi_id')) score += 2;
    if (headers.some(header => header.indexOf('solucion') >= 0)) score += 3;
    if (headers.some(header => header.indexOf('bloque') >= 0)) score += 1;
    if (score > bestScore) { bestScore = score; bestIndex = index; }
  });

  return bestScore >= 4 ? bestIndex : 0;
}

/** Devuelve una celda por nombre de cabecera y, si no existe, por posición. */
function getCatalogoValue_(obj, row, names, fallbackIndex) {
  const value = getValue_(obj, names);
  if (value !== '') return value;
  return fallbackIndex >= 0 && row && row[fallbackIndex] !== undefined ? row[fallbackIndex] : '';
}

/** Clave estable para cruzar 15_OPCIONES_MENU_SITES por país y KPI ID. */
function buildMenuOptionKey_(country, id) {
  return normalizeCountry_(country) + '|' + String(id || '').trim().toUpperCase();
}

/**
 * Lee la matriz 15_OPCIONES_MENU_SITES. Cada opción activa de la fila se
 * devuelve como etiqueta; funciona tanto con TRUE/Sí/X como con texto directo.
 */
function getMenuOptionsByKpi_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SHEET_OPCIONES_MENU_SITES);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getValues();
  const headerRowIndex = findMenuOptionsHeaderRow_(values);
  const headers = (values[headerRowIndex] || []).map(value => String(value || '').trim());
  const normalizedHeaders = headers.map(normalizeHeader_);
  const countryIndex = findHeaderIndexInArray_(normalizedHeaders, ['pais', 'country', 'geografia'], 0);
  const idIndex = findHeaderIndexInArray_(normalizedHeaders, ['kpi_id', 'id_kpi', 'id', 'codigo', 'cod'], 1);
  const result = {};

  values.slice(headerRowIndex + 1).forEach(row => {
    const country = normalizeCountry_(row[countryIndex]);
    const id = String(row[idIndex] || '').trim();
    if (!country || !id) return;

    const options = [];
    row.forEach((cell, index) => {
      if (index === countryIndex || index === idIndex || isDisabledMenuOption_(cell)) return;
      const header = headers[index] || '';
      const value = String(cell).trim();

      if (isEnabledMenuOption_(cell)) {
        if (header) options.push(header);
      } else if (value) {
        // Cuando el Excel contiene directamente el nombre de la opción en vez
        // de una marca, el texto de la celda es la etiqueta a mostrar.
        options.push(value);
      }
    });

    result[buildMenuOptionKey_(country, id)] = uniqueTextValues_(options);
  });

  return result;
}

function findMenuOptionsHeaderRow_(values) {
  let bestIndex = 0;
  let bestScore = -1;
  values.slice(0, 12).forEach((row, index) => {
    const headers = row.map(normalizeHeader_);
    const hasCountry = findHeaderIndexInArray_(headers, ['pais', 'country', 'geografia'], -1) >= 0;
    const hasId = findHeaderIndexInArray_(headers, ['kpi_id', 'id_kpi', 'id', 'codigo', 'cod'], -1) >= 0;
    const score = (hasCountry ? 2 : 0) + (hasId ? 2 : 0) + headers.filter(Boolean).length / 100;
    if (score > bestScore) { bestScore = score; bestIndex = index; }
  });
  return bestScore >= 4 ? bestIndex : 0;
}

function findHeaderIndexInArray_(headers, aliases, fallbackIndex) {
  for (let i = 0; i < aliases.length; i++) {
    const wanted = normalizeHeader_(aliases[i]);
    const exactIndex = headers.indexOf(wanted);
    if (exactIndex >= 0) return exactIndex;
  }
  for (let j = 0; j < aliases.length; j++) {
    const wanted = normalizeHeader_(aliases[j]);
    const partialIndex = headers.findIndex(header => header && (header.indexOf(wanted) >= 0 || wanted.indexOf(header) >= 0));
    if (partialIndex >= 0) return partialIndex;
  }
  return fallbackIndex;
}

function isEnabledMenuOption_(value) {
  if (value === true || value === 1) return true;
  const text = String(value || '').trim().toUpperCase();
  return ['TRUE', 'VERDADERO', 'SI', 'SÍ', 'YES', 'Y', 'X', '1', 'OK'].indexOf(text) >= 0;
}

function isDisabledMenuOption_(value) {
  if (value === '' || value === null || value === undefined || value === false || value === 0) return true;
  return ['FALSE', 'FALSO', 'NO', 'N', '0', '-'].indexOf(String(value).trim().toUpperCase()) >= 0;
}

function uniqueTextValues_(values) {
  const seen = {};
  return values.map(value => String(value || '').trim())
    .filter(value => value && !seen[value] && (seen[value] = true));
}
