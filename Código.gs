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
             Avance Total | Uso de NextGen
   Cada país+bloque tiene dos filas (Legacy y Next gen); se conserva solo
   la fila "Next gen", cuya columna "Uso de NextGen" es el % que se pinta
   por bloque y cuya "Avance Total" alimenta el dato secundario "Av.".
   Los países se repiten uno debajo de otro en el orden habitual:
   España, México, Perú, Colombia, Argentina y Global (mapeado a TOTAL).
   ========================================================= */
function getNextGenSitesData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NEXTGEN_SITES);

  if (!sheet) {
    throw new Error('No existe la pestaña: ' + SHEET_NEXTGEN_SITES);
  }

  const values = sheet.getDataRange().getValues();
  const norm = gthNormalizeLabel_;

  let headerRowIndex = -1, colPais = -1, colBloque = -1, colTipo = -1, colAvance = -1, colUso = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(norm);
    const iPais = row.findIndex(c => c === 'PAIS' || c.indexOf('PAIS') === 0);
    const iBloque = row.findIndex(c => c.indexOf('BLOQUE') >= 0);
    const iTipo = row.findIndex(c => c === 'TIPO');
    const iAvance = row.findIndex(c => c.indexOf('AVANCE') >= 0);
    const iUso = row.findIndex(c => c.indexOf('NEXTGEN') >= 0 || c.indexOf('NEXT GEN') >= 0);
    if (iPais >= 0 && iBloque >= 0 && iTipo >= 0 && iAvance >= 0 && iUso >= 0) {
      headerRowIndex = r;
      colPais = iPais; colBloque = iBloque; colTipo = iTipo; colAvance = iAvance; colUso = iUso;
      break;
    }
  }

  if (headerRowIndex < 0) {
    throw new Error('No se pudo localizar la cabecera (País / Bloque Funcional / Tipo / Avance Total / Uso de NextGen) en ' + SHEET_NEXTGEN_SITES);
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

    if (!result[country]) result[country] = [];
    result[country].push({
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

/* =========================================================
   AVANCE TOTAL POR PAÍS (% EDC Transformed Funcionalities)
   Pestaña: 04_RAW_KPIS
   Son los mismos datos que alimentan "Features Transformadas Web
   Empresas" en KPIs: para cada país (fila con Metric "% EDC Transformed
   Funcionalities" y BU = TOTAL/SPA/MEX/PER/COL/ARG), se toma el último
   valor mensual informado (de derecha a izquierda, ignorando "-" y
   celdas vacías).
   ========================================================= */
function getFeatureTransformedTotalsByCountry_() {
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

  // Columnas de datos mensuales: todas las que, tras Metric/BU, tienen una
  // cabecera con pinta de mes ("MAR 2026", "SEP 2026"...). Se toman solo
  // estas (no "hasta el final de la fila") para no pisar el último valor
  // real con alguna columna de Target/objetivo que pueda venir después.
  const headerRow = values[headerRowIndex].map(norm);
  const monthPattern = /^[A-Z]{3,9}\.?\s*\d{4}$/;
  const monthCols = [];
  for (let c = Math.max(colMetric, colBU) + 1; c < headerRow.length; c++) {
    if (monthPattern.test(headerRow[c])) monthCols.push(c);
  }
  if (!monthCols.length) {
    for (let c = Math.max(colMetric, colBU) + 1; c < headerRow.length; c++) monthCols.push(c);
  }

  const countryMap = { 'TOTAL': 'TOTAL', 'SPA': 'España', 'MEX': 'México', 'PER': 'Perú', 'COL': 'Colombia', 'ARG': 'Argentina' };

  const result = {};
  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const metric = norm(row[colMetric]);
    if (metric.indexOf('TRANSFORMED') < 0) continue;
    if (metric.indexOf('FUNCIONALIT') < 0 && metric.indexOf('FUNCTIONALIT') < 0) continue;

    const country = countryMap[norm(row[colBU])];
    if (!country) continue;

    let lastValue = null;
    for (let i = monthCols.length - 1; i >= 0; i--) {
      const raw = row[monthCols[i]];
      if (raw === '' || raw === null || raw === undefined || raw === '-') continue;
      if (typeof raw !== 'number' && isNaN(Number(String(raw).replace(',', '.').replace('%', '')))) continue;
      lastValue = Math.round(toPercentDecimal_(raw) * 10000) / 100;
      break;
    }
    result[country] = lastValue;
  }

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
