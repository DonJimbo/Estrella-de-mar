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
const SHEET_RAW_FTE = '03_RAW_FTE';

// KPIs
const SHEET_KPIS = '07_KPIS_SITES';
const SHEET_CATALOGO_KPIS = '11_CATALOGO_KPIs';
const SHEET_OPCIONES_MENU_SITES = '15_OPCIONES_MENU_SITES';
const SHEET_NEXTGEN_SITES = '17_NEXTGEN_SITES';
const SHEET_RAW_KPIS = '04_RAW_KPIS';

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
   REFRESCO DE IMPORTRANGE
   ========================================================= */

/**
 * Un IMPORTRANGE no se recalcula solo en cada lectura: Sheets solo lo
 * refresca cuando alguien tiene abierta la hoja de origen o de destino en
 * el navegador. Una llamada de Apps Script (tanto si viene del editor como
 * de la web pública) puede leer un valor "congelado" desde la última vez
 * que alguien abrió la hoja.
 *
 * Para servir datos realmente en vivo, antes de leer la pestaña de origen
 * borramos y volvemos a escribir cada fórmula IMPORTRANGE que encontremos
 * (solo las celdas ancla, que son las únicas que devuelven el texto de la
 * fórmula; las celdas donde se "derrama" el array no lo hacen). Eso obliga
 * a Sheets a reevaluarlas en el momento.
 *
 * Si algo falla aquí (cuota, permisos, etc.) no debe tirar abajo la carga
 * de datos: se registra el aviso y se sigue con lo que ya hubiera.
 */
function refreshImportRangesInSheet_(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const range = sheet.getDataRange();
    const formulas = range.getFormulas();
    const anchors = [];

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        if (formula && formula.toUpperCase().indexOf('IMPORTRANGE') >= 0) {
          anchors.push({ row: r + 1, col: c + 1, formula: formula });
        }
      }
    }

    if (!anchors.length) return;

    anchors.forEach(function (anchor) {
      sheet.getRange(anchor.row, anchor.col).clearContent();
    });
    SpreadsheetApp.flush();

    anchors.forEach(function (anchor) {
      sheet.getRange(anchor.row, anchor.col).setFormula(anchor.formula);
    });
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log('No se pudo refrescar IMPORTRANGE en "' + sheetName + '": ' + error);
  }
}

/* =========================================================
   FTES PRINCIPAL
   Pestaña: 06_FTES_SITES
   ========================================================= */

function getFtesData() {
  refreshImportRangesInSheet_(SHEET_RAW_FTE);

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

function testRefreshFtesImportRange() {
  refreshImportRangesInSheet_(SHEET_RAW_FTE);
  Logger.log('Refresco de IMPORTRANGE en "' + SHEET_RAW_FTE + '" lanzado. Revisa 06_FTES_SITES para confirmar que ya no hay #DIV/0! ni celdas desactualizadas.');
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
      const descripcion = String(getCatalogoValue_(obj, row, ['descripcion', 'descripcion_solucion', 'detalle', 'description'], 11) || '').trim();
      const qTarget = getCatalogoValue_(obj, row, ['q_target', 'qtarget', 'target', 'target_referencia', 'target_q', 'trimestre_target'], -1);
      const qRealRaw = getCatalogoValue_(obj, row, ['q_real', 'qreal', 'entrega', 'fecha_entrega', 'trimestre_entrega'], 5);
      const etpbRaw = getCatalogoValue_(obj, row, ['etpb'], 6);
      const ongoingAkiraRaw = getCatalogoValue_(obj, row, ['ongoing_akira', 'on_going_akira', 'akira'], 7);
      const r5Raw = getCatalogoValue_(obj, row, ['r5', 'robot_5', 'robot5', 'robot_5_bei', 'automatizado', 'automatizada'], 8);
      const comentario = String(getCatalogoValue_(obj, row, ['comentario', 'comentarios', 'observaciones', 'estado_comentario', 'nota_estado', 'estado'], 9) || '').trim();
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

/* =========================================================
   USO DE NEXTGEN POR BLOQUE FUNCIONAL
   Pestaña: 17_NEXTGEN_SITES
   Columnas: País | Bloque Funcional | Tipo (Legacy/Next gen) | Fecha |
   Avance Total | Uso de NextGen. Todas las fechas están unificadas en
   esa única columna Fecha: cada país+bloque tiene varias filas, una por
   cada corte temporal (hoy jun 26 y sept 26) × Tipo (Legacy/Next gen); se
   conserva solo la fila "Next gen" de cada corte. El resultado se agrupa
   por trimestre ("26Q2", "26Q3"...), deducido de la fecha de cada fila,
   para que el site pueda mostrar el corte que corresponda al "Corte de
   avance" seleccionado. Los países se repiten uno debajo de otro en el
   orden habitual: España, México, Perú, Colombia, Argentina y Global
   (mapeado a TOTAL).
   ========================================================= */
function getNextGenSitesData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NEXTGEN_SITES);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_NEXTGEN_SITES);
  }

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colPais = -1, colBloque = -1, colTipo = -1, colFecha = -1, colAvance = -1, colUso = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iPais = row.findIndex(c => c === 'PAIS' || c.indexOf('PAIS') === 0);
    const iBloque = row.findIndex(c => c.indexOf('BLOQUE') >= 0);
    const iTipo = row.findIndex(c => c === 'TIPO');
    const iFecha = row.indexOf('FECHA');
    const iAvance = row.findIndex(c => c.indexOf('AVANCE') >= 0);
    const iUso = row.findIndex(c => c.indexOf('NEXTGEN') >= 0 || c.indexOf('NEXT GEN') >= 0);
    if (iPais >= 0 && iBloque >= 0 && iTipo >= 0 && iFecha >= 0 && iAvance >= 0 && iUso >= 0) {
      headerRowIndex = r;
      colPais = iPais; colBloque = iBloque; colTipo = iTipo;
      colFecha = iFecha; colAvance = iAvance; colUso = iUso;
      break;
    }
  }

  if (headerRowIndex < 0) {
    throw new Error('No se pudo localizar la cabecera (País / Bloque Funcional / Tipo / Fecha / Avance Total / Uso de NextGen) en ' + SHEET_NEXTGEN_SITES);
  }

  const countryMap = {
    'ESPANA': 'España', 'SPAIN': 'España',
    'MEXICO': 'México',
    'PERU': 'Perú',
    'COLOMBIA': 'Colombia',
    'ARGENTINA': 'Argentina',
    'GLOBAL': 'TOTAL', 'TOTAL': 'TOTAL', 'CONSOLIDADO': 'TOTAL'
  };

  const result = {};
  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const paisRaw = row[colPais];
    const bloqueRaw = row[colBloque];
    if (!paisRaw && !bloqueRaw) continue;

    const country = countryMap[norm(paisRaw)];
    const tipo = norm(row[colTipo]);
    if (!country || tipo.indexOf('NEXT') < 0) continue; // solo la fila "Next gen"

    const bloque = String(bloqueRaw || '').trim();
    if (!bloque) continue;

    const parsed = gthParseMonthYear_(row[colFecha]);
    if (!parsed) continue;
    const quarterLabel = gthMonthToQuarterLabel_(parsed);

    if (!result[quarterLabel]) result[quarterLabel] = {};
    if (!result[quarterLabel][country]) result[quarterLabel][country] = [];
    result[quarterLabel][country].push({
      bloque: bloque,
      avance: gthPercentOrNull_(row[colAvance]),
      nextGen: gthPercentOrNull_(row[colUso])
    });
  }

  return result;
}

function gthNormalizeLabel_(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();
}

function gthPercentOrNull_(value) {
  if (value === null || value === undefined || value === '') return null;
  return Math.round(toPercentDecimal_(value) * 10000) / 100;
}

/** 9 -> 3 (Q3), 6 -> 2 (Q2)... y arma la etiqueta "26Q3" a partir de {month,year}. */
function gthMonthToQuarterLabel_(monthYear) {
  const quarter = Math.ceil(monthYear.month / 3);
  return String(monthYear.year).slice(-2) + 'Q' + quarter;
}

/** "26Q3" -> {month:9, year:2026} (último mes de ese trimestre). */
function gthQuarterLabelToMonthYear_(label) {
  const match = String(label || '').match(/^(\d{2})Q([1-4])$/);
  if (!match) return null;
  return { month: Number(match[2]) * 3, year: 2000 + Number(match[1]) };
}

/* =========================================================
   AVANCE TOTAL POR PAÍS (% EDC Transformed Funcionalities)
   Pestaña: 04_RAW_KPIS
   Son los mismos datos que alimentan "Features Transformadas Web
   Empresas" en KPIs: para cada país (fila con Metric "% EDC Transformed
   Funcionalities" y BU = TOTAL/SPA/MEX/PER/COL/ARG), se toma el último
   valor mensual informado (de derecha a izquierda, ignorando "-" y
   celdas vacías).
   El mes a usar no es "el último informado": es el mismo mes al que está
   fechado 17_NEXTGEN_SITES (columna Fecha, p.ej. "sept 26"), para que el
   Avance Total de país y el de bloque funcional hablen del mismo corte.
   ========================================================= */
const GTH_MONTH_NUMBER_ = {
  ENE: 1, ENERO: 1, JAN: 1, JANUARY: 1,
  FEB: 2, FEBRERO: 2, FEBRUARY: 2,
  MAR: 3, MARZO: 3, MARCH: 3,
  ABR: 4, ABRIL: 4, APR: 4, APRIL: 4,
  MAY: 5, MAYO: 5,
  JUN: 6, JUNIO: 6, JUNE: 6,
  JUL: 7, JULIO: 7, JULY: 7,
  AGO: 8, AGOSTO: 8, AUG: 8, AUGUST: 8,
  SEP: 9, SEPT: 9, SEPTIEMBRE: 9, SEPTEMBER: 9,
  OCT: 10, OCTUBRE: 10, OCTOBER: 10,
  NOV: 11, NOVIEMBRE: 11, NOVEMBER: 11,
  DIC: 12, DICIEMBRE: 12, DEC: 12, DECEMBER: 12
};

/** "sept 26", "Sep 2026", "SEPTIEMBRE DE 2026" o una fecha real de la hoja
 *  (Google Sheets guarda "jun 26" con formato de fecha, no como texto) ->
 *  {month:9, year:2026}. Null si no se reconoce. */
function gthParseMonthYear_(value) {
  if (value instanceof Date) {
    return { month: value.getMonth() + 1, year: value.getFullYear() };
  }

  const text = gthNormalizeLabel_(value).replace(/\./g, '').replace(/\bDE\b/g, ' ').replace(/\s+/g, ' ').trim();
  const match = text.match(/^([A-Z]+)\s*(\d{2,4})$/);
  if (!match) return null;
  const monthNumber = GTH_MONTH_NUMBER_[match[1]];
  if (!monthNumber) return null;
  let year = parseInt(match[2], 10);
  if (year < 100) year += 2000;
  return { month: monthNumber, year: year };
}

/** Trimestres ("26Q2", "26Q3"...) presentes en la columna Fecha de
 *  17_NEXTGEN_SITES (uno por cada fecha de corte distinta que tenga esa
 *  pestaña, ya estén todas en una sola columna o repartidas). */
function getNextGenSitesQuarterLabels_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NEXTGEN_SITES);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colFecha = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iFecha = row.indexOf('FECHA');
    if (iFecha >= 0) { headerRowIndex = r; colFecha = iFecha; break; }
  }
  if (headerRowIndex < 0) return [];

  const labelSet = {};
  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const parsed = gthParseMonthYear_(values[r][colFecha]);
    if (parsed) labelSet[gthMonthToQuarterLabel_(parsed)] = true;
  }
  return Object.keys(labelSet);
}

/** Localiza en 04_RAW_KPIS la cabecera (Metric/BU) y la lista de columnas
 *  mensuales ("MAR 2026", "SEP 2026"...), para no repetir esta búsqueda en
 *  cada función que lee esa pestaña. */
function gthRawKpisHeader_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RAW_KPIS);
  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_RAW_KPIS);
  }

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colMetric = -1, colBU = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iMetric = row.indexOf('METRIC');
    const iBU = row.indexOf('BU');
    if (iMetric >= 0 && iBU >= 0) {
      headerRowIndex = r; colMetric = iMetric; colBU = iBU;
      break;
    }
  }
  if (headerRowIndex < 0) {
    throw new Error('No se pudo localizar la cabecera (Metric / BU) en ' + SHEET_RAW_KPIS);
  }

  const headerRow = values[headerRowIndex];
  const monthCols = []; // [{col, month, year}]
  for (let c = Math.max(colMetric, colBU) + 1; c < headerRow.length; c++) {
    const parsed = gthParseMonthYear_(headerRow[c]);
    if (parsed) monthCols.push({ col: c, month: parsed.month, year: parsed.year });
  }

  return { values: values, headerRowIndex: headerRowIndex, colMetric: colMetric, colBU: colBU, monthCols: monthCols };
}

/** Columnas candidatas para un trimestre dado, de la más adecuada a la
 *  menos: su mes de cierre primero ("26Q2" -> junio) y, si falta, los
 *  meses anteriores hacia atrás (nunca posteriores, que podrían ser
 *  proyecciones/target). */
function gthCandidateMonthCols_(monthCols, quarterLabel) {
  const targetMonth = gthQuarterLabelToMonthYear_(quarterLabel);
  const targetColIndex = targetMonth
    ? monthCols.findIndex(m => m.month === targetMonth.month && m.year === targetMonth.year)
    : -1;
  return targetColIndex >= 0
    ? monthCols.slice(0, targetColIndex + 1).reverse().map(m => m.col)
    : monthCols.slice().reverse().map(m => m.col);
}

/** Primer valor numérico informado de una fila entre las columnas
 *  candidatas (ignora "-" y celdas vacías/no numéricas). */
function gthFirstReportedValue_(row, candidateCols) {
  for (let i = 0; i < candidateCols.length; i++) {
    const raw = row[candidateCols[i]];
    if (raw === '' || raw === null || raw === undefined || raw === '-') continue;
    if (typeof raw !== 'number' && isNaN(Number(String(raw).replace(',', '.').replace('%', '')))) continue;
    return Math.round(toPercentDecimal_(raw) * 10000) / 100;
  }
  return null;
}

const GTH_RAW_KPIS_COUNTRY_MAP_ = { 'TOTAL': 'TOTAL', 'SPA': 'España', 'MEX': 'México', 'PER': 'Perú', 'COL': 'Colombia', 'ARG': 'Argentina' };

function getFeatureTransformedTotalsByCountry_() {
  const header = gthRawKpisHeader_();
  const norm = gthNormalizeLabel_;

  // Un trimestre por cada corte que tenga 17_NEXTGEN_SITES (p.ej. 26Q2 y
  // 26Q3), para que hablen del mismo corte temporal, más cualquier trimestre
  // que ya tenga columna propia en 04_RAW_KPIS aunque 17_NEXTGEN_SITES
  // todavía no lo tenga (para que este dato no se quede esperando a que se
  // actualice esa otra pestaña).
  const quarterLabelSet = {};
  getNextGenSitesQuarterLabels_().forEach(label => { quarterLabelSet[label] = true; });
  header.monthCols.forEach(m => { quarterLabelSet[gthMonthToQuarterLabel_(m)] = true; });
  const quarterLabels = Object.keys(quarterLabelSet);
  const result = {};
  quarterLabels.forEach(quarterLabel => {
    const candidateCols = gthCandidateMonthCols_(header.monthCols, quarterLabel);

    const quarterResult = {};
    for (let r = header.headerRowIndex + 1; r < header.values.length; r++) {
      const row = header.values[r];
      const metric = norm(row[header.colMetric]);
      if (metric.indexOf('TRANSFORMED') < 0) continue;
      if (metric.indexOf('FUNCIONALIT') < 0 && metric.indexOf('FUNCTIONALIT') < 0) continue;

      const country = GTH_RAW_KPIS_COUNTRY_MAP_[norm(row[header.colBU])];
      if (!country) continue;

      quarterResult[country] = gthFirstReportedValue_(row, candidateCols);
    }
    result[quarterLabel] = quarterResult;
  });

  return result;
}

/* =========================================================
   TRAFFIC TO NEXTGEN POR PAÍS (SENDA - Traffic to NextGen)
   Pestaña: 04_RAW_KPIS
   Son los mismos datos que alimentan la gráfica "Distribución (NextGen
   vs Legacy)" en KPIs: el total consolidado está en la fila "SENDA -
   Traffic to NextGen" (BU TOTAL) y el de cada país en su fila "... * -
   Calls to NextGen" (BU SPA/MEX/PER/COL/ARG). Se agrupa por cada
   trimestre que tenga columna mensual en 04_RAW_KPIS (no solo los que
   existan en 17_NEXTGEN_SITES), para cubrir todo el histórico.
   ========================================================= */
function getTrafficToNextGenTotalsByCountry_() {
  const header = gthRawKpisHeader_();
  const norm = gthNormalizeLabel_;

  // A diferencia del Avance Total (ligado a los trimestres que existan en
  // 17_NEXTGEN_SITES, que solo tiene detalle por bloque para 26Q2/26Q3), el
  // Traffic to NextGen se calcula para TODOS los trimestres que tengan
  // columna mensual en 04_RAW_KPIS, para que el histórico completo (25Q3,
  // 25Q4, 26Q1...) tenga también el dato real, no solo los dos últimos.
  const quarterLabelSet = {};
  header.monthCols.forEach(function (m) { quarterLabelSet[gthMonthToQuarterLabel_(m)] = true; });
  const quarterLabels = Object.keys(quarterLabelSet);
  const result = {};
  quarterLabels.forEach(quarterLabel => {
    const candidateCols = gthCandidateMonthCols_(header.monthCols, quarterLabel);

    const quarterResult = {};
    for (let r = header.headerRowIndex + 1; r < header.values.length; r++) {
      const row = header.values[r];
      const metric = norm(row[header.colMetric]);
      if (metric.indexOf('TRAFFIC TO NEXTGEN') < 0) continue;

      const bu = norm(row[header.colBU]);
      let country = null;
      if (metric.indexOf('CALLS TO') < 0 && bu === 'TOTAL') {
        country = 'TOTAL'; // fila consolidada "SENDA - Traffic to NextGen"
      } else if (metric.indexOf('CALLS TO NEXTGEN') >= 0) {
        country = GTH_RAW_KPIS_COUNTRY_MAP_[bu]; // fila "... Calls to NextGen" por país
      }
      if (!country) continue;

      quarterResult[country] = gthFirstReportedValue_(row, candidateCols);
    }
    result[quarterLabel] = quarterResult;
  });

  return result;
}

/* =========================================================
   DIGITAL PENETRATION 3X3
   Pestañas: 19_DIGITALPEN_SITES (cajitas con cifra + variación YoY, a
   cierre del último mes informado) y 20_DIGITALPEN_HISTORICO_RAW
   (histórico mensual, para la variación YoY por trimestre que acompaña
   a la gráfica "Penetración Digital 3x3 por País").
   ========================================================= */
const SHEET_DIGITALPEN_SITES = '19_DIGITALPEN_SITES';
const SHEET_DIGITALPEN_HISTORICO = '20_DIGITALPEN_HISTORICO_RAW';

// Mapa combinado para los dos formatos de "país" que usan estas dos
// pestañas: 19_DIGITALPEN_SITES ya trae una fila "Global" consolidada,
// mientras que 20_DIGITALPEN_HISTORICO_RAW solo trae el agregado sin
// Turquía ni Uruguay bajo "Grupo EXTK y EXUR" (los únicos dos valores que
// no se usan en este site). Ambos mapean a 'TOTAL', como el resto de la app.
const GTH_DIGITALPEN_COUNTRY_MAP_ = {
  'ESPANA': 'España', 'SPAIN': 'España',
  'MEXICO': 'México',
  'PERU': 'Perú',
  'COLOMBIA': 'Colombia',
  'ARGENTINA': 'Argentina',
  'GLOBAL': 'TOTAL',
  'GRUPO EXTK Y EXUR': 'TOTAL'
};

/** Número con coma decimal ("290", "-28", "2,5") sin la conversión a
 *  porcentaje de toPercentDecimal_ (que divide por 100 solo si es > 1 y
 *  por tanto rompe con negativos): aquí el valor ya es el que hay que
 *  mostrar tal cual, points de YoY incluidos. */
function gthParseSignedNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return value;
  let text = String(value).trim();
  if (!text || text === '-' || /^#/.test(text)) return null;
  text = text.replace(/\s/g, '');
  if (text.includes(',') && text.includes('.')) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) text = text.replace(/\./g, '').replace(',', '.');
    else text = text.replace(/,/g, '');
  } else if (text.includes(',')) {
    text = text.replace(',', '.');
  }
  const number = Number(text);
  return isNaN(number) ? null : number;
}

/** Como gthParseSignedNumber_, pero quitando antes un "%" si lo hay: para
 *  columnas de variación YoY que vienen con signo Y con el símbolo de
 *  porcentaje a la vez ("-2,4%", "+8,5%"), donde no hace falta ninguna
 *  conversión de escala (el número ya es el porcentaje a mostrar). */
function gthParsePercentSigned_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return value;
  let text = String(value).trim();
  if (!text || text === '-' || /^#/.test(text)) return null;
  text = text.replace('%', '').replace(/\s/g, '');
  if (text.includes(',') && text.includes('.')) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) text = text.replace(/\./g, '').replace(',', '.');
    else text = text.replace(/,/g, '');
  } else if (text.includes(',')) {
    text = text.replace(',', '.');
  }
  const number = Number(text);
  return isNaN(number) ? null : number;
}

/** Primer valor numérico informado entre las columnas candidatas, usando
 *  gthParseSignedNumber_ en vez de la conversión a porcentaje (para series
 *  como "DP 3x3 YoY (pb)", que son puntos y pueden ser negativas). */
function gthFirstReportedNumber_(row, candidateCols) {
  for (let i = 0; i < candidateCols.length; i++) {
    const value = gthParseSignedNumber_(row[candidateCols[i]]);
    if (value === null) continue;
    return Math.round(value * 100) / 100;
  }
  return null;
}

/** Cajitas de Digital Penetration 3x3: por país, un objeto por segmento
 *  (Total / Commercial (CIB NO incluido) / SMEs...) con la cifra y la
 *  variación YoY de cada fila (Porcentaje, Digital Clients, Total
 *  Clients...). Los "Digital/Total Clients" se devuelven tal cual están
 *  en la hoja (p.ej. "338,8k") porque son solo para mostrar, no para
 *  calcular con ellos. */
function getDigitalPenSitesData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_DIGITALPEN_SITES);
  if (!sheet) throw new Error('No existe la pestaña: ' + SHEET_DIGITALPEN_SITES);

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colPais = -1, colMes = -1, colSegmento = -1, colAgrupacion = -1, colCifra = -1, colVariacion = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iPais = row.indexOf('PAIS');
    const iMes = row.indexOf('MES');
    const iSeg = row.indexOf('CLIENTES');
    const iAgr = row.indexOf('AGRUPACION');
    const iCifra = row.indexOf('CIFRA');
    const iVar = row.findIndex(c => c.indexOf('VARIACION') >= 0);
    if (iPais >= 0 && iMes >= 0 && iSeg >= 0 && iAgr >= 0 && iCifra >= 0 && iVar >= 0) {
      headerRowIndex = r; colPais = iPais; colMes = iMes; colSegmento = iSeg; colAgrupacion = iAgr; colCifra = iCifra; colVariacion = iVar;
      break;
    }
  }
  if (headerRowIndex < 0) {
    throw new Error('No se pudo localizar la cabecera (País / Mes / Clientes / Agrupación / Cifra / Variación) en ' + SHEET_DIGITALPEN_SITES);
  }

  const result = {};
  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const country = GTH_DIGITALPEN_COUNTRY_MAP_[norm(row[colPais])];
    if (!country) continue;
    const segmento = String(row[colSegmento] || '').trim();
    if (!segmento) continue;
    const agrupacion = norm(row[colAgrupacion]);

    if (!result[country]) result[country] = { mes: String(row[colMes] || '').trim(), segments: {} };
    if (!result[country].segments[segmento]) result[country].segments[segmento] = {};
    const target = result[country].segments[segmento];

    if (agrupacion === 'PORCENTAJE') {
      target.porcentaje = gthPercentOrNull_(row[colCifra]);
      target.porcentajeYoy = gthParseSignedNumber_(row[colVariacion]);
    } else if (agrupacion.indexOf('TARGET DIGITAL CLIENTS') >= 0) {
      target.targetDigitalClients = String(row[colCifra] || '').trim();
      target.targetDigitalClientsYoy = gthParsePercentSigned_(row[colVariacion]);
    } else if (agrupacion.indexOf('TARGET TOTAL CLIENTS') >= 0) {
      target.targetTotalClients = String(row[colCifra] || '').trim();
      target.targetTotalClientsYoy = gthParsePercentSigned_(row[colVariacion]);
    } else if (agrupacion.indexOf('DIGITAL CLIENTS') >= 0) {
      target.digitalClients = String(row[colCifra] || '').trim();
      target.digitalClientsYoy = gthParsePercentSigned_(row[colVariacion]);
    } else if (agrupacion.indexOf('TOTAL CLIENTS') >= 0) {
      target.totalClients = String(row[colCifra] || '').trim();
      target.totalClientsYoy = gthParsePercentSigned_(row[colVariacion]);
    }
  }
  return result;
}

/** Localiza en 20_DIGITALPEN_HISTORICO_RAW la cabecera (Region/Segments/
 *  Metric) y la lista de columnas mensuales, igual que gthRawKpisHeader_
 *  para 04_RAW_KPIS. */
function gthDigitalPenHeader_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_DIGITALPEN_HISTORICO);
  if (!sheet) throw new Error('No existe la pestaña: ' + SHEET_DIGITALPEN_HISTORICO);

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colRegion = -1, colSegments = -1, colMetric = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iRegion = row.indexOf('REGION');
    const iSegments = row.indexOf('SEGMENTS');
    const iMetric = row.indexOf('METRIC');
    if (iRegion >= 0 && iSegments >= 0 && iMetric >= 0) {
      headerRowIndex = r; colRegion = iRegion; colSegments = iSegments; colMetric = iMetric;
      break;
    }
  }
  if (headerRowIndex < 0) {
    throw new Error('No se pudo localizar la cabecera (Region / Segments / Metric) en ' + SHEET_DIGITALPEN_HISTORICO);
  }

  const headerRow = values[headerRowIndex];
  const monthCols = [];
  for (let c = Math.max(colRegion, colSegments, colMetric) + 1; c < headerRow.length; c++) {
    const parsed = gthParseMonthYear_(headerRow[c]);
    if (parsed) monthCols.push({ col: c, month: parsed.month, year: parsed.year });
  }

  return { values: values, headerRowIndex: headerRowIndex, colRegion: colRegion, colSegments: colSegments, colMetric: colMetric, monthCols: monthCols };
}

/** Variación YoY (en puntos) del Digital Penetration 3x3 total ("TOTAL
 *  PJs" · "DP 3x3 YoY (pb)"), un valor por trimestre y país, para
 *  acompañar la gráfica de barras de Penetración Digital 3x3 por País.
 *  Se calcula para todos los trimestres con columna mensual en la hoja,
 *  igual que getTrafficToNextGenTotalsByCountry_. */
function getDigitalPenYoyByCountry_() {
  const header = gthDigitalPenHeader_();
  const norm = gthNormalizeLabel_;

  const quarterLabelSet = {};
  header.monthCols.forEach(function (m) { quarterLabelSet[gthMonthToQuarterLabel_(m)] = true; });
  const quarterLabels = Object.keys(quarterLabelSet);

  const result = {};
  quarterLabels.forEach(function (quarterLabel) {
    const candidateCols = gthCandidateMonthCols_(header.monthCols, quarterLabel);
    const quarterResult = {};
    for (let r = header.headerRowIndex + 1; r < header.values.length; r++) {
      const row = header.values[r];
      const segment = norm(row[header.colSegments]);
      if (segment.indexOf('TOTAL') < 0) continue; // solo el agregado, no Commercial/SMEs sueltos
      const metric = norm(row[header.colMetric]);
      if (metric.indexOf('DP 3X3 YOY') < 0) continue;

      const country = GTH_DIGITALPEN_COUNTRY_MAP_[norm(row[header.colRegion])];
      if (!country) continue;
      quarterResult[country] = gthFirstReportedNumber_(row, candidateCols);
    }
    result[quarterLabel] = quarterResult;
  });

  return result;
}

/** % de Digital Penetration 3x3 por trimestre, país y segmento (Total /
 *  Commercial (CIB NO incluido) / SMEs), para comparar año actual vs año
 *  anterior en la gráfica de histórico. A diferencia de
 *  getDigitalPenYoyByCountry_ (solo el agregado TOTAL PJs y la fila de
 *  variación), aquí se guarda la fila "Digital Penetration" (el % en sí,
 *  nunca negativo) de los tres segmentos. */
function getDigitalPenHistoricoByCountry_() {
  const header = gthDigitalPenHeader_();
  const norm = gthNormalizeLabel_;

  const quarterLabelSet = {};
  header.monthCols.forEach(function (m) { quarterLabelSet[gthMonthToQuarterLabel_(m)] = true; });
  const quarterLabels = Object.keys(quarterLabelSet);

  const result = {}; // result[country][segmento][quarterLabel] = valor
  quarterLabels.forEach(function (quarterLabel) {
    const candidateCols = gthCandidateMonthCols_(header.monthCols, quarterLabel);
    for (let r = header.headerRowIndex + 1; r < header.values.length; r++) {
      const row = header.values[r];
      const metric = norm(row[header.colMetric]);
      if (metric.indexOf('DIGITAL PENETRATION') < 0 || metric.indexOf('YOY') >= 0) continue;

      const segmentRaw = norm(row[header.colSegments]);
      let segmento = null;
      if (segmentRaw.indexOf('TOTAL') >= 0) segmento = 'Total';
      else if (segmentRaw.indexOf('COMMERCIAL') >= 0) segmento = 'Commercial (CIB NO incluido)';
      else if (segmentRaw.indexOf('SMES') >= 0) segmento = 'SMEs';
      if (!segmento) continue;

      const country = GTH_DIGITALPEN_COUNTRY_MAP_[norm(row[header.colRegion])];
      if (!country) continue;

      const value = gthFirstReportedValue_(row, candidateCols);
      if (value === null) continue;

      if (!result[country]) result[country] = {};
      if (!result[country][segmento]) result[country][segmento] = {};
      result[country][segmento][quarterLabel] = value;
    }
  });

  return result;
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
