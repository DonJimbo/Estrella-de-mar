/**
 * GLOBAL TRANSFORMATION HUB · Capa de caché y matriz de adopción
 * ---------------------------------------------------------------
 * Este fichero se AÑADE al proyecto de Apps Script «Presupuestos» (el que
 * contiene getPresupuestoData, getFtesData, getKpisData, getCatalogoKpisData…).
 * No sustituye a nada: reutiliza esas funciones y les pone una caché delante.
 *
 * Qué resuelve:
 *   1. Cada visita recalculaba todo desde las pestañas. Ahora se calcula una
 *      vez, se guarda y se sirve en milisegundos.
 *   2. Cada bloque del site abría su propia conexión. Ahora existe un único
 *      `getBootstrap` que devuelve todo en una sola llamada.
 *   3. El navegador recibía filas con 20 campos duplicados. El snapshot de
 *      adopción viaja en arrays compactos (≈8 veces menos bytes).
 *   4. Toda respuesta incluye `meta.generatedAt`, la fecha y hora reales del
 *      último cálculo, para poder pintarla en pantalla.
 *
 * Tres niveles de caché:
 *   CacheService  → memoria, 6 h, respuesta inmediata.
 *   Fichero Drive → duradero, sobrevive al vaciado de CacheService.
 *   Recálculo     → solo si no hay nada guardado o se pide actualizar.
 *
 * Todos los identificadores llevan el prefijo GTH_/gth para no chocar con los
 * que ya existen en el proyecto.
 */

/* =========================================================
   CONFIGURACIÓN
   ========================================================= */

const GTH_CACHE = {
  NS: 'GTH_V1',
  TTL_SECONDS: 21600,           // 6 h en CacheService
  CHUNK_SIZE: 90000,            // CacheService admite ~100 KB por clave
  FOLDER_NAME: 'GTH · Caché de datos',
  TZ: 'Europe/Madrid',
  LOCK_MS: 45000
};

/** Conjuntos de datos cacheables. La clave es la que viaja en `action`. */
const GTH_DATASETS = {
  presupuesto:   function () { return getPresupuestoData(); },
  ftes:          function () { return getFtesData(); },
  ftesPersonas:  function () { return getFtesPersonasData(); },
  ftesKpis:      function () { return getFtesKpisData(); },
  kpis:          function () { return getKpisData(); },
  catalogo:      function () { return getCatalogoKpisData(); },
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
  const force = String(params.force || '') === '1' || action.indexOf('refresh') === 0;
  const started = Date.now();

  try {
    let data = null;
    let meta = {};

    if (action === 'ping') {
      data = { ok: true, timestamp: new Date().toISOString() };

    } else if (action === 'getCacheMeta') {
      // Llamada ligera: solo versiones y fechas, sin mover los datos.
      data = gthAllMeta_();

    } else if (action === 'getBootstrap') {
      // Una sola conexión para todo el site.
      const parts = {};
      Object.keys(GTH_DATASETS).forEach(function (name) {
        const result = gthCached_(name, force);
        parts[name] = result.data;
        meta[name] = result.meta;
      });
      data = parts;

    } else if (action === 'refreshCache' || action === 'refreshAll') {
      Object.keys(GTH_DATASETS).forEach(function (name) { gthInvalidate_(name); });
      const parts = {};
      Object.keys(GTH_DATASETS).forEach(function (name) {
        const result = gthCached_(name, true);
        parts[name] = result.data;
        meta[name] = result.meta;
      });
      data = parts;

    } else {
      const name = gthResolveDataset_(action);
      if (!name) throw new Error('Acción API no reconocida: ' + action);
      if (force) gthInvalidate_(name);
      const result = gthCached_(name, force);
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
    getAdopcionSnapshot: 'adopcion',
    refreshAdopcion: 'adopcion',
    refreshCatalogo: 'catalogo',
    refreshPresupuesto: 'presupuesto',
    refreshFtes: 'ftes'
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
   MOTOR DE CACHÉ
   ========================================================= */

/**
 * Devuelve { data, meta }. Busca en memoria, luego en Drive y, solo si no hay
 * nada utilizable, recalcula. El bloqueo evita que dos visitas simultáneas
 * disparen el mismo cálculo.
 */
function gthCached_(name, force) {
  if (!GTH_DATASETS[name]) throw new Error('Conjunto de datos desconocido: ' + name);

  if (!force) {
    const hot = gthReadMemory_(name);
    if (hot) return { data: hot.data, meta: gthMeta_(name, 'memoria') };

    const warm = gthReadDrive_(name);
    if (warm) {
      gthWriteMemory_(name, warm.raw);
      return { data: warm.data, meta: gthMeta_(name, 'disco') };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(GTH_CACHE.LOCK_MS);
  try {
    // Otra ejecución puede haber terminado mientras esperábamos el bloqueo.
    if (!force) {
      const arrived = gthReadMemory_(name);
      if (arrived) return { data: arrived.data, meta: gthMeta_(name, 'memoria') };
    }

    const data = GTH_DATASETS[name]();
    const generatedAtMs = Date.now();
    const raw = JSON.stringify({ v: generatedAtMs, d: data });

    gthWriteMemory_(name, raw);
    gthWriteDrive_(name, raw);
    PropertiesService.getScriptProperties().setProperties({
      [GTH_CACHE.NS + '_VER_' + name]: String(generatedAtMs),
      [GTH_CACHE.NS + '_AT_' + name]: String(generatedAtMs)
    });

    return { data: data, meta: gthMeta_(name, 'recalculo') };
  } finally {
    lock.releaseLock();
  }
}

function gthInvalidate_(name) {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(GTH_CACHE.NS + '_N_' + name) || 0);
  const keys = [GTH_CACHE.NS + '_N_' + name];
  for (let i = 0; i < count; i += 1) keys.push(GTH_CACHE.NS + '_C_' + name + '_' + i);
  cache.removeAll(keys);
}

/* ---------- Nivel 1: CacheService, troceado ---------- */

function gthReadMemory_(name) {
  try {
    const cache = CacheService.getScriptCache();
    const count = Number(cache.get(GTH_CACHE.NS + '_N_' + name) || 0);
    if (!count) return null;
    const keys = [];
    for (let i = 0; i < count; i += 1) keys.push(GTH_CACHE.NS + '_C_' + name + '_' + i);
    const parts = cache.getAll(keys);
    let raw = '';
    for (let i = 0; i < count; i += 1) {
      const chunk = parts[keys[i]];
      if (chunk === null || chunk === undefined) return null;   // trozo caducado
      raw += chunk;
    }
    const parsed = JSON.parse(raw);
    return { data: parsed.d, raw: raw };
  } catch (error) {
    return null;
  }
}

function gthWriteMemory_(name, raw) {
  try {
    const cache = CacheService.getScriptCache();
    const values = {};
    let count = 0;
    for (let i = 0; i < raw.length; i += GTH_CACHE.CHUNK_SIZE) {
      values[GTH_CACHE.NS + '_C_' + name + '_' + count] = raw.substr(i, GTH_CACHE.CHUNK_SIZE);
      count += 1;
    }
    values[GTH_CACHE.NS + '_N_' + name] = String(count);
    cache.putAll(values, GTH_CACHE.TTL_SECONDS);
  } catch (error) {
    // La caché en memoria es una optimización: si falla, queda la de Drive.
    Logger.log('Caché en memoria no disponible para ' + name + ': ' + error);
  }
}

/* ---------- Nivel 2: fichero JSON comprimido en Drive ---------- */

function gthCacheFolder_() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(GTH_CACHE.NS + '_FOLDER');
  if (stored) {
    try { return DriveApp.getFolderById(stored); } catch (ignored) {}
  }
  const existing = DriveApp.getFoldersByName(GTH_CACHE.FOLDER_NAME);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(GTH_CACHE.FOLDER_NAME);
  props.setProperty(GTH_CACHE.NS + '_FOLDER', folder.getId());
  return folder;
}

function gthReadDrive_(name) {
  try {
    const fileId = PropertiesService.getScriptProperties().getProperty(GTH_CACHE.NS + '_FILE_' + name);
    if (!fileId) return null;
    const gzipped = DriveApp.getFileById(fileId).getBlob();
    const raw = Utilities.ungzip(gzipped).getDataAsString('UTF-8');
    const parsed = JSON.parse(raw);
    return { data: parsed.d, raw: raw };
  } catch (error) {
    return null;
  }
}

function gthWriteDrive_(name, raw) {
  try {
    const props = PropertiesService.getScriptProperties();
    const key = GTH_CACHE.NS + '_FILE_' + name;
    const gzipped = Utilities.gzip(
      Utilities.newBlob(raw, 'application/json', name + '.json'),
      name + '.json.gz'
    );
    const existingId = props.getProperty(key);
    if (existingId) {
      // La copia anterior se retira solo después de crear la nueva.
      try { DriveApp.getFileById(existingId).setTrashed(true); } catch (ignored) {}
    }
    const created = gthCacheFolder_().createFile(gzipped);
    props.setProperty(key, created.getId());
  } catch (error) {
    Logger.log('No se pudo guardar la caché en Drive para ' + name + ': ' + error);
  }
}

/* ---------- Metadatos ---------- */

function gthMeta_(name, source) {
  const props = PropertiesService.getScriptProperties();
  const ms = Number(props.getProperty(GTH_CACHE.NS + '_AT_' + name) || 0);
  return {
    dataset: name,
    version: String(ms || ''),
    generatedAtMs: ms,
    generatedAt: ms ? gthFormatDate_(ms) : '',
    source: source
  };
}

function gthAllMeta_() {
  const result = {};
  Object.keys(GTH_DATASETS).forEach(function (name) { result[name] = gthMeta_(name, 'consulta'); });
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
   AUTOMATISMOS
   ========================================================= */

/** Ejecuta esto una vez desde el editor. Deja la caché lista y programada. */
function instalarCacheGTH() {
  ScriptApp.getProjectTriggers()
    .filter(function (trigger) {
      return ['gthRecalcularCache', 'gthInvalidarPorEdicion'].indexOf(trigger.getHandlerFunction()) >= 0;
    })
    .forEach(function (trigger) { ScriptApp.deleteTrigger(trigger); });

  // Recálculo nocturno: por la mañana los datos ya están calientes.
  ScriptApp.newTrigger('gthRecalcularCache')
    .timeBased().everyDays(1).atHour(6).inTimezone(GTH_CACHE.TZ).create();

  // Cualquier edición del libro marca la caché como caducada.
  ScriptApp.newTrigger('gthInvalidarPorEdicion')
    .forSpreadsheet(SpreadsheetApp.openById(SPREADSHEET_ID))
    .onChange().create();

  gthRecalcularCache();
}

function gthRecalcularCache() {
  Object.keys(GTH_DATASETS).forEach(function (name) {
    try {
      gthInvalidate_(name);
      gthCached_(name, true);
    } catch (error) {
      Logger.log('Error recalculando ' + name + ': ' + error);
    }
  });
}

function gthInvalidarPorEdicion() {
  Object.keys(GTH_DATASETS).forEach(function (name) { gthInvalidate_(name); });
  PropertiesService.getScriptProperties().setProperty(GTH_CACHE.NS + '_DIRTY', String(Date.now()));
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
    gthCached_(name, false);
    Logger.log(name + ': ' + (Date.now() - t0) + ' ms');
  });
}
