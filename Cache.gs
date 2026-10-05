/**
 * GLOBAL TRANSFORMATION HUB · API en vivo y matriz de adopción
 * ---------------------------------------------------------------
 * Este fichero se AÑADE al proyecto de Apps Script «Presupuestos» (el que
 * contiene getPresupuestoData, getFtesData, getKpisData, getCatalogoKpisData…).
 * No sustituye a nada: reutiliza esas funciones.
 *
 * Sin caché: cada petición recalcula los datos directamente desde las
 * pestañas. Es más lento que servir una copia guardada, pero evita que el
 * site muestre información desactualizada (por ejemplo, cuando una pestaña
 * se alimenta de IMPORTRANGE, cuyos cambios no siempre disparan los
 * triggers de edición que invalidarían una caché).
 *
 * Todos los identificadores llevan el prefijo GTH_/gth para no chocar con los
 * que ya existen en el proyecto.
 */

/* =========================================================
   CONFIGURACIÓN
   ========================================================= */

const GTH_CACHE = {
  TZ: 'Europe/Madrid'
};

/** Conjuntos de datos que sirve la API. La clave es la que viaja en `action`. */
const GTH_DATASETS = {
  presupuesto:   function () { return getPresupuestoData(); },
  ftes:          function () { return getFtesData(); },
  ftesPersonas:  function () { return getFtesPersonasData(); },
  ftesKpis:      function () { return getFtesKpisData(); },
  kpis:          function () { return getKpisData(); },
  catalogo:      function () { return getCatalogoKpisData(); },
  nextgenSites:  function () { return getNextGenSitesData(); },
  featureTransformedTotals: function () { return getFeatureTransformedTotalsByCountry_(); },
  adopcion:      function () { return gthBuildAdopcion_(); }
};

/** Países del programa, con bandera y etiqueta corta para el front. */
const GTH_COUNTRIES = [
  { code: 'ESPANA',    label: 'España',    flag: 'es', short: 'ES' },
  { code: 'MEXICO',    label: 'México',    flag: 'mx', short: 'MX' },
  { code: 'PERU',      label: 'Perú',      flag: 'pe', short: 'PE' },
  { code: 'COLOMBIA',  label: 'Colombia',  flag: 'co', short: 'CO' },
  { code: 'ARGENTINA', label: 'Argentina', flag: 'ar', short: 'AR' }
];

/* =========================================================
   ROUTER DE LA API
   =========================================================
   Sustituye el cuerpo de tu `handleApiRequest_(e)` por:
       function handleApiRequest_(e) { return gthApiDispatch_(e); }
   ========================================================= */

function gthApiDispatch_(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || '').trim();
  const started = Date.now();

  try {
    let data = null;
    let meta = {};

    if (action === 'ping') {
      data = { ok: true, timestamp: new Date().toISOString() };

    } else if (action === 'getCacheMeta') {
      // Se mantiene por compatibilidad con llamadas antiguas del front-end:
      // ya no hay caché, así que solo informa de la hora actual del servidor.
      data = gthAllMeta_();

    } else if (action === 'getBootstrap' || action === 'refreshCache' || action === 'refreshAll') {
      // Una sola conexión para todo el site. Sin distinción "normal"/"forzado":
      // siempre se calcula en el momento.
      const parts = {};
      Object.keys(GTH_DATASETS).forEach(function (name) {
        const result = gthCompute_(name);
        parts[name] = result.data;
        meta[name] = result.meta;
      });
      data = parts;

    } else {
      const name = gthResolveDataset_(action);
      if (!name) throw new Error('Acción API no reconocida: ' + action);
      const result = gthCompute_(name);
      data = result.data;
      meta = result.meta;
    }

    return gthJsonp_(e, {
      status: 'ok',
      action: action,
      data: data,
      meta: meta,
      elapsedMs: Date.now() - started
    });

  } catch (error) {
    return gthJsonp_(e, {
      status: 'error',
      action: action,
      message: error && error.message ? error.message : String(error)
    });
  }
}

/** Traduce los nombres de acción antiguos y nuevos al conjunto de datos. */
function gthResolveDataset_(action) {
  const map = {
    getPresupuestoData: 'presupuesto',
    getFtesData: 'ftes',
    getFTEsData: 'ftes',
    getFtesPersonasData: 'ftesPersonas',
    getFtesKpisData: 'ftesKpis',
    getKpisData: 'kpis',
    getCatalogoKpisData: 'catalogo',
    getCatalogoData: 'catalogo',
    getNextGenSitesData: 'nextgenSites',
    getFeatureTransformedTotalsByCountry: 'featureTransformedTotals',
    getAdopcionSnapshot: 'adopcion',
    refreshAdopcion: 'adopcion',
    refreshCatalogo: 'catalogo',
    refreshPresupuesto: 'presupuesto',
    refreshFtes: 'ftes',
    refreshNextgenSites: 'nextgenSites',
    refreshFeatureTransformedTotals: 'featureTransformedTotals'
  };
  return map[action] || (GTH_DATASETS[action] ? action : null);
}

function gthJsonp_(e, payload) {
  const callback = String(((e && e.parameter) || {}).callback || '').trim();
  const json = JSON.stringify(payload);
  if (!callback) {
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error', message: 'Nombre de callback no válido'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* =========================================================
   CÁLCULO EN VIVO
   ========================================================= */

/** Calcula un conjunto de datos directamente, sin pasar por ninguna caché. */
function gthCompute_(name) {
  if (!GTH_DATASETS[name]) throw new Error('Conjunto de datos desconocido: ' + name);
  const data = GTH_DATASETS[name]();
  const generatedAtMs = Date.now();
  return {
    data: data,
    meta: {
      dataset: name,
      version: String(generatedAtMs),
      generatedAtMs: generatedAtMs,
      generatedAt: gthFormatDate_(generatedAtMs),
      source: 'en vivo'
    }
  };
}

function gthAllMeta_() {
  const generatedAtMs = Date.now();
  const meta = {
    dataset: 'todos',
    version: String(generatedAtMs),
    generatedAtMs: generatedAtMs,
    generatedAt: gthFormatDate_(generatedAtMs),
    source: 'en vivo'
  };
  const result = {};
  Object.keys(GTH_DATASETS).forEach(function (name) { result[name] = meta; });
  return result;
}

function gthFormatDate_(ms) {
  return Utilities.formatDate(new Date(Number(ms)), GTH_CACHE.TZ, "dd/MM/yyyy 'a las' HH:mm");
}

/* =========================================================
   SNAPSHOT DE ADOPCIÓN
   =========================================================
   Convierte el catálogo en una matriz país × funcionalidad lista para
   pintar. El navegador no vuelve a normalizar ni a agrupar nada: solo
   compara dos números para saber si una entrega cae dentro del corte.
   ========================================================= */

function gthBuildAdopcion_() {
  const rows = getCatalogoKpisData() || [];

  const countryIndex = {};
  GTH_COUNTRIES.forEach(function (country, index) { countryIndex[country.code] = index; });

  const blocks = [];
  const blockIndex = {};
  const features = [];
  const featureIndex = {};
  const cells = [];
  const quarters = {};

  rows.forEach(function (row) {
    const countryCode = gthCountryCode_(row.pais || row.paisOriginal);
    const ci = countryIndex[countryCode];
    if (ci === undefined) return;   // filas GLOBAL/TOTAL no entran en la matriz

    const blockLabel = String(row.bloqueFuncional || row.bloque || 'Sin bloque').trim();
    const blockKey = gthKey_(blockLabel);
    if (blockIndex[blockKey] === undefined) {
      blockIndex[blockKey] = blocks.length;
      blocks.push(blockLabel);
    }
    const bi = blockIndex[blockKey];

    const featureLabel = String(row.solucionFuncional || row.solucion || row.id || 'Sin nombre').trim();
    const featureKey = blockKey + '|' + gthKey_(featureLabel);
    if (featureIndex[featureKey] === undefined) {
      featureIndex[featureKey] = features.length;
      features.push([featureLabel, bi, String(row.descripcion || '').trim()]);
    }
    const fi = featureIndex[featureKey];

    const targetLabel = String(row.qTarget || '').trim();
    const realLabel = String(
      row.qRealRaw !== undefined && row.qRealRaw !== null ? row.qRealRaw : (row.qReal || '')
    ).trim();

    const targetRank = gthQuarterRank_(targetLabel);
    const realRank = gthQuarterRank_(realLabel);
    if (targetLabel) quarters[gthQuarterLabel_(targetLabel)] = gthQuarterRank_(targetLabel);
    if (realLabel) quarters[gthQuarterLabel_(realLabel)] = realRank;

    let flags = 0;
    if (gthTruthy_(row.etpbRaw !== undefined ? row.etpbRaw : row.etpb)) flags |= 1;
    if (gthTruthy_(row.ongoingAkiraRaw !== undefined ? row.ongoingAkiraRaw : row.ongoingAkira)) flags |= 2;
    if (gthTruthy_(row.r5Raw !== undefined ? row.r5Raw : row.r5)) flags |= 4;

    cells.push([
      ci,                                   // 0 país
      fi,                                   // 1 funcionalidad
      targetRank || 0,                      // 2 rango del target
      realRank || 0,                        // 3 rango de la entrega
      flags,                                // 4 ETPB / Akira / Robot 5
      String(row.id || ''),                 // 5 id del catálogo
      targetLabel,                          // 6 target legible
      realLabel,                            // 7 entrega legible
      String(row.comentario || '').trim(),  // 8 comentario
      Array.isArray(row.menuOptions) ? row.menuOptions : []   // 9 opciones de menú
    ]);
  });

  const quarterList = Object.keys(quarters)
    .filter(function (label) { return quarters[label]; })
    .sort(function (a, b) { return quarters[a] - quarters[b]; });

  return {
    countries: GTH_COUNTRIES,
    blocks: blocks,
    features: features,
    cells: cells,
    quarters: quarterList,
    totals: { countries: GTH_COUNTRIES.length, blocks: blocks.length, features: features.length, cells: cells.length }
  };
}

/** «26Q2», «2026 Q2», «Q2 26» → 2026*4+2. Devuelve 0 si no hay trimestre. */
function gthQuarterRank_(value) {
  const text = String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '');
  let yearText = '';
  let quarterText = '';

  // Formato habitual: 26Q2 o 2026Q2.
  let match = text.match(/(?:^|[^0-9])(\d{4}|\d{2})Q([1-4])(?![0-9])/);
  if (match) {
    yearText = match[1];
    quarterText = match[2];
  } else {
    // Formato alternativo: Q2-26, Q2DE2026.
    match = text.match(/Q([1-4])[^0-9]*(\d{4}|\d{2})(?![0-9])/);
    if (!match) return 0;
    quarterText = match[1];
    yearText = match[2];
  }

  const year = yearText.length === 2 ? 2000 + Number(yearText) : Number(yearText);
  const quarter = Number(quarterText);
  return (year && quarter) ? year * 4 + quarter : 0;
}

function gthQuarterLabel_(value) {
  const rank = gthQuarterRank_(value);
  if (!rank) return String(value || '').trim();
  const year = Math.floor((rank - 1) / 4);
  const quarter = rank - year * 4;
  return String(year).slice(2) + 'Q' + quarter;
}

function gthCountryCode_(value) {
  const text = String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();
  if (['ES', 'SPA', 'ESPANA', 'SPAIN'].indexOf(text) >= 0) return 'ESPANA';
  if (['MX', 'MEX', 'MEXICO'].indexOf(text) >= 0) return 'MEXICO';
  if (['PE', 'PER', 'PERU'].indexOf(text) >= 0) return 'PERU';
  if (['CO', 'COL', 'COLOMBIA'].indexOf(text) >= 0) return 'COLOMBIA';
  if (['AR', 'ARG', 'ARGENTINA'].indexOf(text) >= 0) return 'ARGENTINA';
  return text;
}

function gthKey_(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function gthTruthy_(value) {
  if (value === true || value === 1) return true;
  const text = String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();
  return ['TRUE', 'VERDADERO', 'SI', 'S', 'YES', 'Y', '1', 'X', 'OK'].indexOf(text) >= 0;
}

/* =========================================================
   COMPROBACIONES
   ========================================================= */

function gthProbarAdopcion() {
  const snapshot = gthBuildAdopcion_();
  Logger.log('Países: %s · Bloques: %s · Funcionalidades: %s · Celdas: %s',
    snapshot.totals.countries, snapshot.totals.blocks,
    snapshot.totals.features, snapshot.totals.cells);
  Logger.log('Trimestres detectados: ' + snapshot.quarters.join(', '));
  Logger.log(JSON.stringify(snapshot.cells.slice(0, 5), null, 2));
}

function gthProbarVelocidad() {
  Object.keys(GTH_DATASETS).forEach(function (name) {
    const t0 = Date.now();
    gthCompute_(name);
    Logger.log(name + ': ' + (Date.now() - t0) + ' ms');
  });
}
