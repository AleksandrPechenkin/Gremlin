/**
 * Остатки Ozon для листа «Планирование закупок» (книга 03).
 *
 * Сопоставление с планом (согласовано):
 *   offer_id (Ozon)  ↔ «Артикул поставщика»  (canon через ppCanonArticle_)
 *   barcode  (Ozon)  ↔ «ШК»                  (canon через ppCanonBarcodeForStock_)
 * Это симметрично логике WB (там nmId/supplierArticle ↔ «Артикул ВБ», barcode ↔ «ШК»),
 * но для Ozon отдельный набор `needSupplier` строится из колонки «Артикул поставщика»
 * листа плана — артикул ВБ для Ozon не используется (его в Ozon API нет).
 *
 * Источники остатков:
 *   POST /v4/product/info/stocks      — список товаров + stocks[] с разбивкой по типу склада
 *                                       (`fbo`, `fbs`); пагинация через cursor.
 *                                       Ozon в 2026 перевёл этот метод с v3 на v4
 *                                       (v3 теперь отдаёт 404 page not found).
 *                                       Схема запроса/ответа совместима со старой v3.
 *   POST /v3/product/info/list        — для тех offer_id, что вернулись из stocks, забираем
 *                                       поле `barcodes` пакетами по 500, чтобы построить
 *                                       barcode→qty (FBO/FBS). Можно отключить свойством
 *                                       OZON_FETCH_BARCODES=0 (тогда матчинг — только по offer_id).
 *
 * Script Properties:
 *   OZON_API_TOKEN     — строка «<client_id>:<api_key>» (Client-Id и Api-Key Ozon Seller API).
 *   OZON_API_BASE      — опционально, по умолчанию https://api-seller.ozon.ru
 *   OZON_FETCH_BARCODES — «0» отключает дозагрузку barcodes через /v3/product/info/list
 *                         (тогда `stockByBarcodeFbo/Fbs` останутся пустыми).
 *   OZON_STOCKS_PAGE_LIMIT — размер страницы /v4/product/info/stocks (по умолчанию 1000).
 *   OZON_INFO_LIST_BATCH — размер пачки offer_id для /v3/product/info/list (по умолчанию 500).
 *
 * Контракт результата `ozBuildStockLookupForProcurementPlanning_`:
 *   {
 *     stockBySupplierFbo: { canon(offer_id): qty },
 *     stockBySupplierFbs: { canon(offer_id): qty },
 *     stockByBarcodeFbo:  { canon(barcode):  qty },
 *     stockByBarcodeFbs:  { canon(barcode):  qty },
 *     source: 'v4 product/info/stocks(+v3 info/list)',
 *     stats: { itemsFromStocks, offerHits, barcodeHits, infoListCalls }
 *   }
 */

const OZON_DEFAULT_API_BASE = 'https://api-seller.ozon.ru';
const OZON_STOCKS_DEFAULT_PAGE_LIMIT = 1000;
const OZON_INFO_LIST_DEFAULT_BATCH = 500;

function ozGetProp_(key, defaultValue) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, defaultValue);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? defaultValue : String(v).trim();
}

/**
 * Парсит OZON_API_TOKEN формата «client_id:api_key». Возвращает { clientId, apiKey }.
 * Бросает понятную ошибку, если формат не соблюдён.
 */
function ozParseToken_() {
  const raw = ozGetProp_('OZON_API_TOKEN', '');
  if (!raw) {
    throw new Error(
      'Не заполнен OZON_API_TOKEN в свойствах скрипта.\n' +
        'Формат: «<Client-Id>:<Api-Key>» — оба значения из ЛК Ozon (Настройки → Seller API).'
    );
  }
  const idx = raw.indexOf(':');
  if (idx <= 0 || idx >= raw.length - 1) {
    throw new Error(
      'OZON_API_TOKEN должен иметь формат «<Client-Id>:<Api-Key>» (значения через двоеточие).'
    );
  }
  return { clientId: raw.slice(0, idx).trim(), apiKey: raw.slice(idx + 1).trim() };
}

function ozApiBase_() {
  return ozGetProp_('OZON_API_BASE', OZON_DEFAULT_API_BASE).replace(/\/$/, '');
}

function ozPost_(path, payload, token) {
  const url = ozApiBase_() + path;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Client-Id': token.clientId,
      'Api-Key': token.apiKey
    },
    payload: JSON.stringify(payload || {}),
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

function ozFormatApiError_(res) {
  const j = res.json;
  if (j && j.message) return String(j.message) + (j.code != null ? ' (code ' + j.code + ')' : '');
  if (j && j.error && j.error.message) return String(j.error.message);
  if (j && j.details && j.details.length) {
    return j.details.map(function (d) { return d.typ ? d.typ + ': ' + (d.message || '') : (d.message || JSON.stringify(d)); }).join('; ');
  }
  return res.text ? res.text.slice(0, 500) : 'HTTP ' + res.code;
}

/**
 * Извлекает qty по типу склада из stocks[] ответа /v4/product/info/stocks.
 * Формула: max(present - reserved, 0). Не учитываем «sku_in_transit».
 * @param {Array} stocks
 * @param {'fbo'|'fbs'} type
 */
function ozStockAvailableByType_(stocks, type) {
  if (!stocks || !stocks.length) return 0;
  let sum = 0;
  for (let i = 0; i < stocks.length; i++) {
    const s = stocks[i];
    if (!s) continue;
    const t = String(s.type || '').toLowerCase();
    if (t !== type) continue;
    const present = parseFloat(s.present) || 0;
    const reserved = parseFloat(s.reserved) || 0;
    sum += Math.max(present - reserved, 0);
  }
  return Math.round(sum);
}

/**
 * Основная точка вызова из procurement_planning.gs.
 * @param {Object.<string, boolean>} needSupplier canon(«Артикул поставщика» из плана) → true
 * @param {Object.<string, boolean>} needBc canon(«ШК» из плана) → true
 */
function ozBuildStockLookupForProcurementPlanning_(needSupplier, needBc) {
  const token = ozParseToken_();
  const pageLimit = Math.max(
    100,
    Math.min(1000, parseInt(ozGetProp_('OZON_STOCKS_PAGE_LIMIT', String(OZON_STOCKS_DEFAULT_PAGE_LIMIT)), 10) || OZON_STOCKS_DEFAULT_PAGE_LIMIT)
  );
  const fetchBarcodes = String(ozGetProp_('OZON_FETCH_BARCODES', '1')).trim() !== '0';

  const stockBySupplierFbo = {};
  const stockBySupplierFbs = {};
  const stockByBarcodeFbo = {};
  const stockByBarcodeFbs = {};

  // Запоминаем offer_id → { fbo, fbs }, чтобы при дозагрузке barcodes из info/list
  // прицепить количество к barcode-ключу (один запрос — два индекса).
  const offerQtyMap = {};
  let itemsFromStocks = 0;
  let offerHits = 0;
  let barcodeHits = 0;
  let infoListCalls = 0;

  let cursor = '';
  for (let page = 0; page < 200; page++) {
    const res = ozPost_('/v4/product/info/stocks', {
      cursor: cursor,
      filter: { visibility: 'ALL' },
      limit: pageLimit
    }, token);
    if (res.code < 200 || res.code >= 300) {
      throw new Error('Ozon /v4/product/info/stocks: ' + ozFormatApiError_(res));
    }
    const items = (res.json && res.json.items) || [];
    if (!items.length) break;
    for (let i = 0; i < items.length; i++) {
      itemsFromStocks++;
      const it = items[i];
      const offerCanon = it.offer_id != null ? ppCanonArticle_(String(it.offer_id)) : '';
      const fboQty = ozStockAvailableByType_(it.stocks, 'fbo');
      const fbsQty = ozStockAvailableByType_(it.stocks, 'fbs');
      if (fboQty || fbsQty) {
        offerQtyMap[String(it.offer_id || '')] = { fbo: fboQty, fbs: fbsQty };
      }
      if (offerCanon && needSupplier[offerCanon]) {
        if (fboQty) stockBySupplierFbo[offerCanon] = (stockBySupplierFbo[offerCanon] || 0) + fboQty;
        if (fbsQty) stockBySupplierFbs[offerCanon] = (stockBySupplierFbs[offerCanon] || 0) + fbsQty;
        if (fboQty || fbsQty) offerHits++;
      }
    }
    cursor = (res.json && (res.json.cursor || res.json.next_cursor)) || '';
    if (!cursor) break;
  }

  if (fetchBarcodes && Object.keys(needBc).length) {
    const batchSize = Math.max(
      50,
      Math.min(1000, parseInt(ozGetProp_('OZON_INFO_LIST_BATCH', String(OZON_INFO_LIST_DEFAULT_BATCH)), 10) || OZON_INFO_LIST_DEFAULT_BATCH)
    );
    const offerIds = Object.keys(offerQtyMap);
    for (let start = 0; start < offerIds.length; start += batchSize) {
      const batch = offerIds.slice(start, start + batchSize);
      const res = ozPost_('/v3/product/info/list', { offer_id: batch, product_id: [], sku: [] }, token);
      infoListCalls++;
      if (res.code < 200 || res.code >= 300) {
        throw new Error('Ozon /v3/product/info/list (barcodes): ' + ozFormatApiError_(res));
      }
      const items = (res.json && (res.json.items || (res.json.result && res.json.result.items))) || [];
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const oid = String(it.offer_id || '');
        const qtyRec = offerQtyMap[oid];
        if (!qtyRec) continue;
        const barcodes = ozCollectBarcodes_(it);
        for (let b = 0; b < barcodes.length; b++) {
          const bcCanon = ppCanonBarcodeForStock_(barcodes[b]);
          if (!bcCanon || !needBc[bcCanon]) continue;
          if (qtyRec.fbo) stockByBarcodeFbo[bcCanon] = (stockByBarcodeFbo[bcCanon] || 0) + qtyRec.fbo;
          if (qtyRec.fbs) stockByBarcodeFbs[bcCanon] = (stockByBarcodeFbs[bcCanon] || 0) + qtyRec.fbs;
          if (qtyRec.fbo || qtyRec.fbs) barcodeHits++;
        }
      }
    }
  }

  return {
    stockBySupplierFbo: stockBySupplierFbo,
    stockBySupplierFbs: stockBySupplierFbs,
    stockByBarcodeFbo: stockByBarcodeFbo,
    stockByBarcodeFbs: stockByBarcodeFbs,
    source: fetchBarcodes ? 'v4 product/info/stocks + v3 info/list' : 'v4 product/info/stocks',
    stats: {
      itemsFromStocks: itemsFromStocks,
      offerHits: offerHits,
      barcodeHits: barcodeHits,
      infoListCalls: infoListCalls
    }
  };
}

/** Собирает все возможные штрихкоды из item ответа /v3/product/info/list: `barcodes`, `barcode`. */
function ozCollectBarcodes_(item) {
  const out = [];
  if (item && Array.isArray(item.barcodes)) {
    for (let i = 0; i < item.barcodes.length; i++) {
      const b = item.barcodes[i];
      if (b != null && String(b).trim()) out.push(String(b));
    }
  }
  if (item && item.barcode != null && String(item.barcode).trim()) {
    out.push(String(item.barcode));
  }
  return out;
}
