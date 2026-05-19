const CONFIG = {
  MS_API_BASE: 'https://api.moysklad.ru/api/remap/1.2',
  get MS_TOKEN() { return getScriptPropOrThrow('MS_TOKEN'); },
  get MS_ORGANIZATION_ID() { return getScriptPropOrThrow('MS_ORGANIZATION_ID'); },
  get MS_STORE_ID() { return getScriptProp('MS_STORE_ID', ''); },
  get EXTERNAL_SPREADSHEET_ID() { return getScriptPropOrThrow('EXTERNAL_SPREADSHEET_ID'); },
  get EXTERNAL_SHEET_NAME() { return getScriptProp('EXTERNAL_SHEET_NAME', 'Закуплено'); }
};

function getScriptProp(key, defaultValue) {
  const fallback = defaultValue === undefined ? '' : defaultValue;
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value == null ? fallback : String(value).trim();
}

function getScriptPropOrThrow(key) {
  const value = getScriptProp(key, '');
  if (!value) throw new Error(`Не заполнен Script Property: ${key}`);
  return value;
}

function logInfo(message, context) {
  Logger.log(`[INFO] ${message}${context ? ' | ' + JSON.stringify(context) : ''}`);
}

function logWarn(message, context) {
  Logger.log(`[WARN] ${message}${context ? ' | ' + JSON.stringify(context) : ''}`);
}

function logError(message, error, context) {
  const errText = error ? (error.stack || error.message || String(error)) : '';
  Logger.log(`[ERROR] ${message}${errText ? ' | ' + errText : ''}${context ? ' | ' + JSON.stringify(context) : ''}`);
}

function getMsHeaders() {
  return {
    Authorization: 'Bearer ' + CONFIG.MS_TOKEN,
    'Content-Type': 'application/json'
  };
}

function parseMsError(responseCode, text, endpoint, method) {
  let json = null;
  try { json = JSON.parse(text); } catch (e) {}
  const apiMessage = json && json.errors
    ? json.errors.map(err => `${err.error || 'Ошибка'}${err.moreInfo ? ' (' + err.moreInfo + ')' : ''}`).join('; ')
    : text;
  return {
    success: false,
    error: apiMessage || `HTTP ${responseCode}`,
    statusCode: responseCode,
    endpoint: endpoint,
    method: method
  };
}

function msApiRequest(endpoint, method, payload) {
  const requestMethod = (method || 'get').toLowerCase();
  const options = {
    method: requestMethod,
    headers: getMsHeaders(),
    muteHttpExceptions: true
  };
  if (payload !== null && payload !== undefined) {
    options.payload = JSON.stringify(payload);
  }

  const url = CONFIG.MS_API_BASE + endpoint;
  try {
    const response = UrlFetchApp.fetch(url, options);
    Utilities.sleep(50);
    const statusCode = response.getResponseCode();
    const text = response.getContentText();
    let json = null;
    try { json = JSON.parse(text); } catch (e) {}

    if (statusCode >= 200 && statusCode < 300) {
      return { success: true, data: json, statusCode: statusCode };
    }

    const err = parseMsError(statusCode, text, endpoint, requestMethod);
    logWarn('Ошибка API МойСклад', err);
    return err;
  } catch (error) {
    logError('Сбой запроса к API МойСклад', error, { endpoint: endpoint, method: requestMethod });
    return {
      success: false,
      error: `Сбой запроса: ${error.message}`,
      endpoint: endpoint,
      method: requestMethod
    };
  }
}

function msGet(endpoint) { return msApiRequest(endpoint, 'get', null); }
function msPost(endpoint, payload) { return msApiRequest(endpoint, 'post', payload); }
function msPut(endpoint, payload) { return msApiRequest(endpoint, 'put', payload); }
function msPatch(endpoint, payload) { return msApiRequest(endpoint, 'patch', payload); }
function msDelete(endpoint) { return msApiRequest(endpoint, 'delete', null); }

function msFetch(endpoint, method, payload) {
  return msApiRequest(endpoint, method || 'get', payload);
}

// normalizeMsEntityId_, validateMsOrganization_, fetchMsOrganizationsList_ — в main.gs
