/**
 * Остатки Wildberries для листа «Планирование закупок».
 *
 * Script Properties:
 *   WB_API_TOKEN — токен WB (в ЛК: настройки доступа к API). Для отчёта warehouse_remains
 *     нужна категория «Аналитика»; для supplier/stocks — «Статистика» (можно один токен с обеими галочками).
 *   WB_STOCK_SOURCE — analytics | statistics | auto (по умолчанию auto: сначала warehouse_remains, при ошибке — supplier/stocks).
 *   WB_ANALYTICS_API_BASE — по умолчанию https://seller-analytics-api.wildberries.ru
 *   WB_STATISTICS_API_BASE — по умолчанию https://statistics-api.wildberries.ru
 *   WB_STOCK_USE_QUANTITY_FULL — если 1, в режиме statistics суммировать quantityFull вместо quantity.
 */

const WB_DEFAULT_ANALYTICS_BASE = 'https://seller-analytics-api.wildberries.ru';
const WB_DEFAULT_STATISTICS_BASE = 'https://statistics-api.wildberries.ru';

function wbGetProp_(key, defaultValue) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, defaultValue);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? defaultValue : String(v).trim();
}

function wbGetTokenOrThrow_() {
  const t = wbGetProp_('WB_API_TOKEN', '');
  if (!t) throw new Error('Не заполнен WB_API_TOKEN в свойствах скрипта.');
  return t.trim();
}

/** Количество из массива warehouses отчёта warehouse_remains: приоритет строки «Всего на складах…». */
function wbWarehouseLineQtyFromWarehouses_(warehouses) {
  if (!warehouses || !warehouses.length) return 0;
  const norm = function (s) {
    return String(s || '')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  };
  for (let i = 0; i < warehouses.length; i++) {
    const n = norm(warehouses[i].warehouseName);
    if (n.indexOf('всего') >= 0 && n.indexOf('склад') >= 0 && n.indexOf('пути') < 0) {
      return Math.round(parseFloat(warehouses[i].quantity) || 0);
    }
  }
  let sum = 0;
  for (let i = 0; i < warehouses.length; i++) {
    const n = norm(warehouses[i].warehouseName);
    if (n.indexOf('пути') >= 0) continue;
    sum += parseFloat(warehouses[i].quantity) || 0;
  }
  return Math.round(sum);
}

function wbFetchGet_(baseUrl, pathAndQuery, token) {
  const url = String(baseUrl).replace(/\/$/, '') + pathAndQuery;
  const options = {
    method: 'get',
    headers: { Authorization: token },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText() || '';
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {}
  return { code: code, text: text, json: json };
}

function wbFormatApiError_(res) {
  const j = res.json;
  if (j && j.title) return String(j.title) + (j.detail ? ': ' + j.detail : '');
  if (j && j.error) return String(j.error);
  if (j && j.errors && j.errors.length) return j.errors.map(function (e) { return e.message || e.error || String(e); }).join('; ');
  return res.text ? res.text.slice(0, 500) : 'HTTP ' + res.code;
}

/**
 * Собирает остатки WB только для ключей из плана (needWb / needBc — как в МС).
 * @return {{ stockByWb: Object.<string, number>, stockByBarcode: Object.<string, number>, source: string }}
 */
function wbBuildStockLookupForProcurementPlanning_(needWb, needBc) {
  const mode = String(wbGetProp_('WB_STOCK_SOURCE', 'auto') || 'auto').toLowerCase();
  const token = wbGetTokenOrThrow_();
  if (mode === 'statistics') {
    return wbBuildFromSupplierStocks_(token, needWb, needBc);
  }
  if (mode === 'analytics') {
    return wbBuildFromWarehouseRemains_(token, needWb, needBc);
  }
  try {
    return wbBuildFromWarehouseRemains_(token, needWb, needBc);
  } catch (e) {
    if (typeof logWarn === 'function') logWarn('WB warehouse_remains failed, fallback supplier/stocks', { err: String(e.message || e) });
    return wbBuildFromSupplierStocks_(token, needWb, needBc);
  }
}

function wbBuildFromWarehouseRemains_(token, needWb, needBc) {
  const base = wbGetProp_('WB_ANALYTICS_API_BASE', WB_DEFAULT_ANALYTICS_BASE);
  const path =
    '/api/v1/warehouse_remains?locale=ru&groupByNm=true&groupByBarcode=false&groupBySa=false&groupByBrand=false&groupBySubject=false&groupBySize=false&filterPics=0&filterVolume=0';
  const res = wbFetchGet_(base, path, token);
  if (res.code < 200 || res.code >= 300) throw new Error('WB warehouse_remains (создание): ' + wbFormatApiError_(res));
  const taskId =
    res.json && res.json.data && res.json.data.taskId
      ? String(res.json.data.taskId)
      : res.json && res.json.data && res.json.data.task_id
      ? String(res.json.data.task_id)
      : '';
  if (!taskId) throw new Error('WB warehouse_remains: в ответе нет taskId.');
  const deadline = Date.now() + 5 * 60 * 1000;
  let status = '';
  while (Date.now() < deadline) {
    Utilities.sleep(5000);
    const st = wbFetchGet_(base, '/api/v1/warehouse_remains/tasks/' + encodeURIComponent(taskId) + '/status', token);
    if (st.code < 200 || st.code >= 300) throw new Error('WB warehouse_remains (статус): ' + wbFormatApiError_(st));
    status = st.json && st.json.data && st.json.data.status ? String(st.json.data.status) : '';
    if (status === 'done') break;
    if (status === 'canceled' || status === 'purged' || status === 'error') {
      throw new Error('WB warehouse_remains: задание завершилось со статусом «' + status + '».');
    }
  }
  if (status !== 'done') throw new Error('WB warehouse_remains: таймаут ожидания отчёта (5 мин).');
  Utilities.sleep(500);
  const dl = wbFetchGet_(base, '/api/v1/warehouse_remains/tasks/' + encodeURIComponent(taskId) + '/download', token);
  if (dl.code === 204) {
    return { stockByWb: {}, stockByBarcode: {}, source: 'warehouse_remains(empty)' };
  }
  if (dl.code < 200 || dl.code >= 300) throw new Error('WB warehouse_remains (скачивание): ' + wbFormatApiError_(dl));
  let rows = dl.json;
  if (rows && rows.data && Array.isArray(rows.data)) rows = rows.data;
  if (!Array.isArray(rows)) throw new Error('WB warehouse_remains: неожиданный формат ответа (ожидался массив строк).');
  const stockByWb = {};
  const stockByBarcode = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const qty = wbWarehouseLineQtyFromWarehouses_(row.warehouses);
    const nmKey = row.nmId != null ? ppCanonArticle_(String(row.nmId)) : '';
    const bcKey = row.barcode != null ? ppCanonBarcodeForStock_(String(row.barcode)) : '';
    const vcKey = row.vendorCode != null ? ppCanonArticle_(String(row.vendorCode)) : '';
    if (nmKey && needWb[nmKey]) stockByWb[nmKey] = (stockByWb[nmKey] || 0) + qty;
    if (bcKey && needBc[bcKey]) stockByBarcode[bcKey] = (stockByBarcode[bcKey] || 0) + qty;
    if (vcKey && needWb[vcKey]) stockByWb[vcKey] = (stockByWb[vcKey] || 0) + qty;
  }
  return { stockByWb: stockByWb, stockByBarcode: stockByBarcode, source: 'warehouse_remains' };
}

function wbBuildFromSupplierStocks_(token, needWb, needBc) {
  const base = wbGetProp_('WB_STATISTICS_API_BASE', WB_DEFAULT_STATISTICS_BASE);
  const useFull = wbGetProp_('WB_STOCK_USE_QUANTITY_FULL', '0') === '1';
  const all = [];
  let dateFrom = '2019-06-20T00:00:00';
  for (let page = 0; page < 40; page++) {
    const path = '/api/v1/supplier/stocks?dateFrom=' + encodeURIComponent(dateFrom);
    const res = wbFetchGet_(base, path, token);
    if (res.code < 200 || res.code >= 300) throw new Error('WB supplier/stocks: ' + wbFormatApiError_(res));
    const rows = Array.isArray(res.json) ? res.json : [];
    if (!rows.length) break;
    for (let ri = 0; ri < rows.length; ri++) all.push(rows[ri]);
    if (rows.length < 60000) break;
    const last = rows[rows.length - 1];
    const nextFrom = last && last.lastChangeDate != null ? String(last.lastChangeDate) : '';
    if (!nextFrom || nextFrom === dateFrom) break;
    dateFrom = nextFrom;
    Utilities.sleep(61000);
  }
  const stockByWb = {};
  const stockByBarcode = {};
  for (let i = 0; i < all.length; i++) {
    const r = all[i];
    const qRaw = useFull ? r.quantityFull : r.quantity;
    const qty = Math.round(parseFloat(qRaw) || 0);
    const nmKey = r.nmId != null ? ppCanonArticle_(String(r.nmId)) : '';
    const bcKey = r.barcode != null ? ppCanonBarcodeForStock_(String(r.barcode)) : '';
    const saKey = r.supplierArticle != null ? ppCanonArticle_(String(r.supplierArticle)) : '';
    if (nmKey && needWb[nmKey]) stockByWb[nmKey] = (stockByWb[nmKey] || 0) + qty;
    if (bcKey && needBc[bcKey]) stockByBarcode[bcKey] = (stockByBarcode[bcKey] || 0) + qty;
    if (saKey && needWb[saKey]) stockByWb[saKey] = (stockByWb[saKey] || 0) + qty;
  }
  return { stockByWb: stockByWb, stockByBarcode: stockByBarcode, source: 'supplier/stocks' };
}
