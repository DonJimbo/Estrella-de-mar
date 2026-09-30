/**
 * Dashboard EDC App
 *
 * Este proyecto sirve la base que alimenta el site, el directorio Equipos EDC
 * y el seguimiento de las solicitudes de cambio. Está comentado en castellano
 * para que las responsabilidades de cada bloque sean fáciles de mantener.
 */

/* =========================================================
   CONFIGURACIÓN
   ========================================================= */

const CONFIG = {
  TARGET_SHEET_NAME: 'Backlog FT',
  MONITORED_COLUMNS: [9, 11, 13, 15, 17],
  EMAIL_TO: 'maria.cabanas@bbva.com',
  LOG_SHEET_NAME: '_CHANGE_LOG',
  SNAPSHOT_SHEET_NAME: '_DAILY_SNAPSHOT',
  TIMEZONE: Session.getScriptTimeZone() || 'Europe/Madrid',
  ID_COLUMN: 4,
  NAME_COLUMN: 3,
  HEADER_ROW: 3,
  DATA_START_ROW: 4
};

const EDC_DASHBOARD = {
  // Libro Layout KPIs EDC / fuente de los datos que consume el dashboard.
  DASHBOARD_SPREADSHEET_ID: '1mlRMrS5AnkjrPBU6XHGIk1yU-qMKHYzIfk_OwYk8wAM',

  // Libro de Equipos EDC. La pestaña operativa se llama «Modelo de Gobierno».
  EQUIPOS_SPREADSHEET_ID: '1Id49v5fUc8CJbjIEgWByl2xQzLKYSOw0Fjc9aChxU7M',
  EQUIPOS_SHEET_NAME: 'Modelo de Gobierno',

  // Formulario de Solicitud de Cambios compartido por el equipo.
  FORM_ID: '1FAIpQLSeUakNmL159DcWndrEhStWV4UsCJoOPPSbiBDlwORM9qMDdPA',
  FORMULARIO_SHEET_NAME: 'Formulario',

  TIMEZONE: 'Europe/Madrid',
  CACHE_FILE_NAME: 'EDC Site · copia de datos para caché.xlsx',
  PROP_CACHE_FILE_ID: 'EDC_SITE_CACHE_FILE_ID',
  PROP_CACHE_SOURCE_UPDATED: 'EDC_SITE_CACHE_SOURCE_UPDATED',
  PROP_CACHE_REFRESHED_AT: 'EDC_SITE_CACHE_REFRESHED_AT',
  PROP_RELEVANT_DATA_VERSION: 'EDC_RELEVANT_DATA_VERSION'
};

/* =========================================================
   PUNTO DE ENTRADA WEB (JSONP)
   ========================================================= */

/**
 * Mantiene la compatibilidad con la llamada antigua del HTML:
 *   /exec?callback=miFuncion
 * y añade las acciones nuevas que utiliza el site.
 */
function handleDashboardEDCApi_(e) {
  e = e || {};
  const params = e.parameter || {};
  const action = String(params.action || 'getSnapshot').trim();

  try {
    let payload;

    if (action === 'getSnapshot' || action === 'getData') {
      payload = getSnapshotPayload_();
    } else if (action === 'getCacheStatus') {
      payload = {
        status: 'success',
        data: getCacheStatus_()
      };
    } else if (action === 'refreshCache' || action === 'refreshData') {
      payload = refreshDashboardCache_();
    } else if (action === 'getEquiposEDC') {
      payload = {
        status: 'success',
        data: getEquiposEDC_()
      };
    } else if (action === 'getMisSolicitudes') {
      payload = {
        status: 'success',
        data: getMisSolicitudes_()
      };
    } else {
      throw new Error('Acción no reconocida: ' + action);
    }

    return edcJsonpResponse_(params, payload);
  } catch (error) {
    return edcJsonpResponse_(params, {
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  }
}

/** Devuelve JSON o JSONP, validando el nombre de la función de callback. */
function edcJsonpResponse_(params, payload) {
  const callback = String((params && params.callback) || '').trim();
  const json = JSON.stringify(payload);

  if (!callback) {
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Nombre de callback no válido.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* =========================================================
   CACHÉ DEL DASHBOARD
   ========================================================= */

/**
 * CacheService no admite un XLSX completo. Por ello la caché persistente se
 * guarda como una copia XLSX en Drive y el navegador conserva además su propia
 * copia en IndexedDB. Así se evita exportar el libro en cada visita.
 */
function getSnapshotPayload_() {
  const properties = PropertiesService.getScriptProperties();
  let fileId = properties.getProperty(EDC_DASHBOARD.PROP_CACHE_FILE_ID);

  if (!fileId) {
    refreshDashboardCache_();
    fileId = properties.getProperty(EDC_DASHBOARD.PROP_CACHE_FILE_ID);
  }

  if (!fileId) {
    throw new Error('No se ha podido crear la copia de datos del dashboard.');
  }

  const cacheFile = DriveApp.getFileById(fileId);
  const base64 = Utilities.base64Encode(cacheFile.getBlob().getBytes());

  return {
    status: 'success',
    data: base64,
    meta: getCacheStatus_()
  };
}

/**
 * Reexporta la fuente una sola vez y sustituye de forma segura la copia usada
 * por el site. Se usa un bloqueo para que dos visitantes no la recreen a la vez.
 */
function refreshDashboardCache_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sourceUpdatedAt = getRelevantDataVersion_();
    const xlsxBlob = exportDashboardSpreadsheet_();
    const newCacheFile = DriveApp.createFile(xlsxBlob)
      .setName(EDC_DASHBOARD.CACHE_FILE_NAME);

    const properties = PropertiesService.getScriptProperties();
    const previousFileId = properties.getProperty(EDC_DASHBOARD.PROP_CACHE_FILE_ID);
    const refreshedAt = Date.now();

    properties.setProperties({
      [EDC_DASHBOARD.PROP_CACHE_FILE_ID]: newCacheFile.getId(),
      [EDC_DASHBOARD.PROP_CACHE_SOURCE_UPDATED]: String(sourceUpdatedAt),
      [EDC_DASHBOARD.PROP_CACHE_REFRESHED_AT]: String(refreshedAt)
    });

    // Solo se envía a la papelera después de haber guardado la nueva copia.
    if (previousFileId && previousFileId !== newCacheFile.getId()) {
      try {
        DriveApp.getFileById(previousFileId).setTrashed(true);
      } catch (ignored) {
        // No se interrumpe la actualización si una copia antigua no se puede eliminar.
      }
    }

    return {
      status: 'success',
      data: Utilities.base64Encode(newCacheFile.getBlob().getBytes()),
      meta: getCacheStatus_()
    };
  } finally {
    lock.releaseLock();
  }
}

/** Exporta el spreadsheet mediante Drive API, igual que hacía el código previo. */
function exportDashboardSpreadsheet_() {
  const url = 'https://www.googleapis.com/drive/v3/files/' +
    EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID +
    '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('No se pudo exportar el libro de datos (' + response.getResponseCode() + ').');
  }

  return response.getBlob().setName(EDC_DASHBOARD.CACHE_FILE_NAME);
}

/**
 * Esta consulta es ligera: solo lee la fecha del archivo fuente. El HTML la usa
 * para saber si debe mostrar el aviso de actualización sin descargar el XLSX.
 */
function getCacheStatus_() {
  const properties = PropertiesService.getScriptProperties();
  const sourceUpdatedAt = getRelevantDataVersion_();
  const cachedSourceUpdatedAt = Number(properties.getProperty(EDC_DASHBOARD.PROP_CACHE_SOURCE_UPDATED) || 0);
  const refreshedAt = Number(properties.getProperty(EDC_DASHBOARD.PROP_CACHE_REFRESHED_AT) || 0);

  return {
    cacheVersion: cachedSourceUpdatedAt ? String(cachedSourceUpdatedAt) : '',
    sourceVersion: String(sourceUpdatedAt),
    sourceUpdatedAt: sourceUpdatedAt ? formatMadrid_(sourceUpdatedAt) : '',
    refreshedAt: refreshedAt ? formatMadrid_(refreshedAt) : '',
    hasCache: !!properties.getProperty(EDC_DASHBOARD.PROP_CACHE_FILE_ID),
    hasNewData: sourceUpdatedAt > cachedSourceUpdatedAt
  };
}

/**
 * Ejecutar una vez desde el editor para dejar instalados los automatismos.
 * Apps Script programa la ejecución dentro de la franja de las 08:00 de Madrid.
 */
function instalarActualizacionDiariaEDC() {
  deleteTriggersByHandler_('actualizarCacheProgramada');
  ScriptApp.newTrigger('actualizarCacheProgramada')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone(EDC_DASHBOARD.TIMEZONE)
    .create();
}

/** Función llamada cada mañana por el activador horario. */
function actualizarCacheProgramada() {
  refreshDashboardCache_();
}

/**
 * Preparación única: crea la copia inicial, instala la actualización diaria y
 * registra el trigger que copia las solicitudes del formulario a «Formulario».
 */
function instalarDashboardEDC() {
  instalarVigilanciaCambiosBaseEDC();
  refreshDashboardCache_();
  instalarActualizacionDiariaEDC();
  configurarFormularioIdentificadoEDC();
  instalarTriggerFormularioEDC();
}

function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handlerName)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Registra los cambios de las pestañas que alimentan el dashboard. Al ser un
 * trigger de edición de spreadsheet, las escrituras del script en «Formulario»
 * no provocan avisos falsos de actualización en el site.
 */
function instalarVigilanciaCambiosBaseEDC() {
  deleteTriggersByHandler_('registrarCambioBaseEDC');
  const spreadsheet = SpreadsheetApp.openById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID);
  ScriptApp.newTrigger('registrarCambioBaseEDC')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(EDC_DASHBOARD.PROP_RELEVANT_DATA_VERSION)) {
    const updatedAt = DriveApp.getFileById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID).getLastUpdated().getTime();
    properties.setProperty(EDC_DASHBOARD.PROP_RELEVANT_DATA_VERSION, String(updatedAt));
  }
}

function registrarCambioBaseEDC(e) {
  const sheetName = e && e.range && e.range.getSheet ? e.range.getSheet().getName() : '';
  if (edcNormalizeHeader_(sheetName) === edcNormalizeHeader_(EDC_DASHBOARD.FORMULARIO_SHEET_NAME)) return;
  PropertiesService.getScriptProperties().setProperty(EDC_DASHBOARD.PROP_RELEVANT_DATA_VERSION, String(Date.now()));
}

function getRelevantDataVersion_() {
  const properties = PropertiesService.getScriptProperties();
  const storedVersion = Number(properties.getProperty(EDC_DASHBOARD.PROP_RELEVANT_DATA_VERSION) || 0);
  if (storedVersion) return storedVersion;
  return DriveApp.getFileById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID).getLastUpdated().getTime();
}

function formatMadrid_(milliseconds) {
  return Utilities.formatDate(new Date(Number(milliseconds)), EDC_DASHBOARD.TIMEZONE, 'dd/MM/yyyy HH:mm');
}

/* =========================================================
   EQUIPOS EDC
   ========================================================= */

/** Lee el directorio desde «Modelo de Gobierno» y lo normaliza para el HTML. */
function getEquiposEDC_() {
  const spreadsheet = SpreadsheetApp.openById(EDC_DASHBOARD.EQUIPOS_SPREADSHEET_ID);
  const sheet = findEquiposSheet_(spreadsheet);

  if (!sheet) {
    throw new Error('No se encuentra la pestaña «' + EDC_DASHBOARD.EQUIPOS_SHEET_NAME + '» en el fichero de Equipos EDC.');
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const richValues = range.getRichTextValues();
  const headerRow = findEquiposHeaderRow_(displayValues);

  if (headerRow < 0) {
    throw new Error('No se han identificado las cabeceras del directorio Equipos EDC.');
  }

  const headers = displayValues[headerRow].map(edcNormalizeHeader_);
  const indexes = {
    pais: findHeaderIndex_(headers, ['pais', 'country', 'geografia']),
    nivel: findHeaderIndex_(headers, ['nivel']),
    subNivel: findHeaderIndex_(headers, ['subnivel', 'sub_nivel']),
    rol: findHeaderIndex_(headers, ['rol', 'role', 'cargo', 'funcion']),
    nombre: findHeaderIndex_(headers, ['nombre', 'contacto', 'persona', 'name']),
    descripcion: findHeaderIndex_(headers, ['descripcion', 'description']),
    linea: findHeaderIndex_(headers, ['linea', 'area', 'equipo', 'team', 'unidad']),
    foto: findHeaderIndex_(headers, ['foto', 'photo', 'imagen', 'image']),
    contacto: findHeaderIndex_(headers, ['contacto', 'correo', 'email', 'e_mail', 'mail'])
  };

  return values.slice(headerRow + 1).map((row, relativeIndex) => {
    const rowNumber = headerRow + 1 + relativeIndex;
    const displayedRow = displayValues[rowNumber] || [];
    const richRow = richValues[rowNumber] || [];
    const nombre = cleanText_(getRowValue_(displayedRow, indexes.nombre));
    const pais = formatCountry_(getRowValue_(displayedRow, indexes.pais));

    if (!nombre && !pais) return null;

    const email = findTeamEmail_(displayedRow, indexes.contacto);
    const foto = indexes.foto >= 0
      ? extractPhotoUrl_(row[indexes.foto], richRow[indexes.foto])
      : '';

    return {
      pais: pais || 'Sin país',
      codigoPais: cleanText_(getRowValue_(displayedRow, indexes.pais)),
      nivel: cleanText_(getRowValue_(displayedRow, indexes.nivel)),
      subNivel: cleanText_(getRowValue_(displayedRow, indexes.subNivel)),
      rol: cleanText_(getRowValue_(displayedRow, indexes.rol)),
      nombre: titleCase_(nombre),
      descripcion: cleanText_(getRowValue_(displayedRow, indexes.descripcion)),
      area: cleanText_(getRowValue_(displayedRow, indexes.linea)),
      linea: cleanText_(getRowValue_(displayedRow, indexes.linea)),
      email: email,
      contacto: email,
      foto: foto
    };
  }).filter(person => person && (person.nombre || person.pais));
}

function findEquiposSheet_(spreadsheet) {
  const direct = spreadsheet.getSheetByName(EDC_DASHBOARD.EQUIPOS_SHEET_NAME);
  if (direct) return direct;

  const wanted = edcNormalizeHeader_(EDC_DASHBOARD.EQUIPOS_SHEET_NAME);
  return spreadsheet.getSheets().find(sheet => {
    const name = edcNormalizeHeader_(sheet.getName());
    return name === wanted || name.indexOf('modelo_de_gobierno') >= 0 || name.indexOf('equipos_edc') >= 0;
  }) || null;
}

/** Localiza cabeceras aunque estén en la tercera fila del fichero. */
function findEquiposHeaderRow_(values) {
  let bestIndex = -1;
  let bestScore = 0;

  values.slice(0, 20).forEach((row, index) => {
    const headers = row.map(edcNormalizeHeader_);
    let score = 0;
    if (findHeaderIndex_(headers, ['nombre', 'contacto', 'persona', 'name']) >= 0) score += 4;
    if (findHeaderIndex_(headers, ['pais', 'country', 'geografia']) >= 0) score += 2;
    if (findHeaderIndex_(headers, ['rol', 'role', 'cargo']) >= 0) score += 1;
    if (findHeaderIndex_(headers, ['linea', 'area', 'equipo', 'team']) >= 0) score += 1;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 5 ? bestIndex : -1;
}

function findHeaderIndex_(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const wanted = edcNormalizeHeader_(names[i]);
    const exact = headers.indexOf(wanted);
    if (exact >= 0) return exact;
  }

  for (let j = 0; j < names.length; j++) {
    const wanted = edcNormalizeHeader_(names[j]);
    const partial = headers.findIndex(header => header && (header.indexOf(wanted) >= 0 || wanted.indexOf(header) >= 0));
    if (partial >= 0) return partial;
  }
  return -1;
}

function getRowValue_(row, index) {
  return index >= 0 && row && row[index] !== undefined ? row[index] : '';
}

function findTeamEmail_(row, preferredIndex) {
  const preferred = cleanText_(getRowValue_(row, preferredIndex));
  if (isEmail_(preferred)) return preferred;
  return row.map(cleanText_).find(isEmail_) || '';
}

function extractPhotoUrl_(cell, richTextValue) {
  try {
    if (cell && typeof cell === 'object' && typeof cell.getContentUrl === 'function') {
      const contentUrl = cell.getContentUrl();
      if (contentUrl) return contentUrl;
    }
  } catch (ignored) {}

  try {
    if (richTextValue && richTextValue.getLinkUrl) {
      const richUrl = richTextValue.getLinkUrl();
      if (richUrl) return toDirectDriveUrl_(richUrl);
    }
  } catch (ignored) {}

  const value = cleanText_(cell);
  return /^https?:\/\//i.test(value) ? toDirectDriveUrl_(value) : '';
}

function toDirectDriveUrl_(url) {
  const match = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return match ? 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400' : String(url || '');
}

function formatCountry_(value) {
  const key = edcNormalizeHeader_(value).toUpperCase();
  const countries = {
    ES: 'España', ESPANA: 'España', SPAIN: 'España', SPA: 'España',
    MX: 'México', MEXICO: 'México', MEX: 'México',
    PE: 'Perú', PERU: 'Perú', PER: 'Perú',
    CO: 'Colombia', COL: 'Colombia', COLOMBIA: 'Colombia',
    AR: 'Argentina', ARG: 'Argentina', ARGENTINA: 'Argentina',
    HOLDING: 'Holding', GLOBAL: 'Holding', TOTAL: 'Holding'
  };
  return countries[key] || cleanText_(value);
}

function titleCase_(value) {
  return cleanText_(value).toLowerCase().replace(/\S+/g, word => word.charAt(0).toUpperCase() + word.slice(1));
}

function cleanText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function isEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText_(value));
}

function edcNormalizeHeader_(value) {
  return cleanText_(value)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_|_$/g, '')
    .toLowerCase();
}

/* =========================================================
   SOLICITUDES DE CAMBIO
   ========================================================= */

/** Activa la recogida de correo en el formulario cuando la política lo permite. */
function configurarFormularioIdentificadoEDC() {
  try {
    FormApp.openById(EDC_DASHBOARD.FORM_ID).setCollectEmail(true);
  } catch (error) {
    // Algunas políticas corporativas solo permiten activar esta opción desde
    // la interfaz de Forms. El trigger y el site seguirán funcionando.
    Logger.log('No se pudo activar la recogida de correo automáticamente: ' + error);
  }
}

/** Instala el trigger del formulario una única vez. */
function instalarTriggerFormularioEDC() {
  deleteTriggersByHandler_('registrarSolicitudDeCambio');
  const form = FormApp.openById(EDC_DASHBOARD.FORM_ID);
  ScriptApp.newTrigger('registrarSolicitudDeCambio')
    .forForm(form)
    .onFormSubmit()
    .create();
}

/**
 * Copia cada respuesta a la pestaña «Formulario». La columna Estado nace como
 * «Nueva» y el equipo puede actualizar Estado/Comentario de seguimiento allí.
 */
function registrarSolicitudDeCambio(e) {
  if (!e || !e.response) {
    throw new Error('La función debe ejecutarse desde el trigger de envío del formulario.');
  }

  const response = e.response;
  const responseValues = response.getItemResponses().map(itemResponse => ({
    title: itemResponse.getItem().getTitle(),
    value: Array.isArray(itemResponse.getResponse())
      ? itemResponse.getResponse().join(', ')
      : String(itemResponse.getResponse() || '')
  }));
  const respondentEmail = getResponseEmail_(response, responseValues);
  const sheet = getFormularioSheet_();
  const columns = ensureFormularioColumns_(sheet);
  const columnCount = Math.max(sheet.getLastColumn(), Object.keys(columns).length);
  const row = new Array(columnCount).fill('');

  setFormularioValue_(row, columns, 'id_solicitud', response.getId());
  setFormularioValue_(row, columns, 'fecha_solicitud', response.getTimestamp());
  setFormularioValue_(row, columns, 'solicitante', respondentEmail);
  setFormularioValue_(row, columns, 'estado', 'Nueva');
  setFormularioValue_(row, columns, 'ultima_actualizacion', new Date());
  setFormularioValue_(row, columns, 'comentario_seguimiento', '');
  setFormularioValue_(row, columns, 'respuestas_formulario', responseValues.map(item => item.title + ': ' + item.value).join(' | '));

  sheet.appendRow(row);
}

/** Recupera solamente las solicitudes del usuario autenticado en el site. */
function getMisSolicitudes_() {
  const email = getActiveUserEmail_();
  if (!email) {
    return {
      identityAvailable: false,
      user: '',
      requests: []
    };
  }

  const sheet = getFormularioSheet_();
  if (sheet.getLastRow() < 2) {
    return { identityAvailable: true, user: email, requests: [] };
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(edcNormalizeHeader_);
  const idx = aliases => findHeaderIndex_(headers, aliases);
  const idIndex = idx(['id_solicitud', 'id', 'solicitud_id']);
  const dateIndex = idx(['fecha_solicitud', 'timestamp', 'fecha']);
  const requesterIndex = idx(['solicitante', 'email', 'correo', 'respondent_email']);
  const statusIndex = idx(['estado', 'status']);
  const updatedIndex = idx(['ultima_actualizacion', 'actualizado', 'fecha_actualizacion']);
  const commentIndex = idx(['comentario_seguimiento', 'comentario', 'observaciones']);
  const detailIndex = idx(['respuestas_formulario', 'respuestas', 'detalle']);

  const requests = values.slice(1).map((row, index) => ({
    id: cleanText_(getRowValue_(row, idIndex)) || 'SOL-' + (index + 1),
    date: cleanText_(getRowValue_(row, dateIndex)),
    requester: cleanText_(getRowValue_(row, requesterIndex)),
    status: cleanText_(getRowValue_(row, statusIndex)) || 'Nueva',
    updatedAt: cleanText_(getRowValue_(row, updatedIndex)),
    comment: cleanText_(getRowValue_(row, commentIndex)),
    detail: cleanText_(getRowValue_(row, detailIndex))
  })).filter(request => request.requester.toLowerCase() === email.toLowerCase());

  return {
    identityAvailable: true,
    user: email,
    requests: requests.reverse()
  };
}

function getFormularioSheet_() {
  const spreadsheet = SpreadsheetApp.openById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID);
  return spreadsheet.getSheetByName(EDC_DASHBOARD.FORMULARIO_SHEET_NAME) ||
    spreadsheet.insertSheet(EDC_DASHBOARD.FORMULARIO_SHEET_NAME);
}

function ensureFormularioColumns_(sheet) {
  const required = [
    'ID solicitud',
    'Fecha solicitud',
    'Solicitante',
    'Estado',
    'Última actualización',
    'Comentario de seguimiento',
    'Respuestas formulario'
  ];

  const currentColumns = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, currentColumns).getDisplayValues()[0];
  const normalHeaders = currentHeaders.map(edcNormalizeHeader_);
  let nextColumn = currentHeaders.length;

  required.forEach(label => {
    const key = edcNormalizeHeader_(label);
    if (normalHeaders.indexOf(key) < 0) {
      nextColumn += 1;
      sheet.getRange(1, nextColumn).setValue(label);
      normalHeaders[nextColumn - 1] = key;
    }
  });

  return normalHeaders.reduce((result, header, index) => {
    if (header) result[header] = index;
    return result;
  }, {});
}

function setFormularioValue_(row, columns, key, value) {
  const index = columns[key];
  if (index !== undefined) row[index] = value;
}

function getResponseEmail_(response, responses) {
  try {
    const email = cleanText_(response.getRespondentEmail());
    if (isEmail_(email)) return email;
  } catch (ignored) {}

  const emailQuestion = responses
    .map(item => item.value)
    .map(cleanText_)
    .find(isEmail_);
  return emailQuestion || '';
}

function getActiveUserEmail_() {
  try {
    return cleanText_(Session.getActiveUser().getEmail());
  } catch (ignored) {
    return '';
  }
}

/* =========================================================
   UTILIDADES DE PRUEBA Y AUTORIZACIÓN
   ========================================================= */

/** Ejecuta esta función después de pegar el código para conceder permisos. */
function autorizarDashboardEDC() {
  DriveApp.getFileById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID).getName();
  SpreadsheetApp.openById(EDC_DASHBOARD.EQUIPOS_SPREADSHEET_ID).getName();
  SpreadsheetApp.openById(EDC_DASHBOARD.DASHBOARD_SPREADSHEET_ID).getName();
  FormApp.openById(EDC_DASHBOARD.FORM_ID).getTitle();
}

function probarEquiposEDC() {
  const people = getEquiposEDC_();
  Logger.log('Contactos leídos: ' + people.length);
  Logger.log(JSON.stringify(people.slice(0, 5), null, 2));
}

function probarEstadoCacheEDC() {
  Logger.log(JSON.stringify(getCacheStatus_(), null, 2));
}
