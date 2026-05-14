/**
 * Факт продаж WB за текущий месяц для расчёта закупки (книга 03).
 *
 * Источники (выбор через WB_SALES_SOURCE):
 *   analytics  — POST https://seller-analytics-api.wildberries.ru
 *                /api/analytics/v3/sales-funnel/products
 *                Лимит 3 запроса/мин, burst 3, интервал 20 с — мягче и стабильнее.
 *                Возвращает по карточкам метрики за период (orderCount/buyoutCount/cancelCount).
 *                Barcode не возвращает — матчинг с планом идёт по nmId/vendorCode.
 *   statistics — GET https://statistics-api.wildberries.ru/api/v1/supplier/sales
 *                (наследие; жёсткий лимит ~1 запрос/мин, после серии 429 банит токен надолго).
 *
 * Метрика (выбор через WB_SALES_ANALYTICS_METRIC, действует для analytics):
 *   orderCount     — заказы за период (по умолчанию; согласуется с планом «Заказали, шт»).
 *   buyoutCount    — выкупы за период.
 *   netOrderCount  — max(0, orderCount - cancelCount), «чистые заказы».
 *
 * Script Properties:
 *   WB_API_TOKEN                — токен WB (для analytics нужны права на «Аналитика»).
 *   WB_SALES_SOURCE             — analytics | statistics | auto; по умолчанию analytics.
 *                                 «auto» = сначала analytics, при ошибке — statistics.
 *   WB_SALES_ANALYTICS_API_BASE — опционально, по умолчанию https://seller-analytics-api.wildberries.ru
 *   WB_STATISTICS_API_BASE      — опционально, по умолчанию https://statistics-api.wildberries.ru
 *   WB_SALES_ANALYTICS_METRIC   — orderCount | buyoutCount | netOrderCount; по умолчанию orderCount.
 *
 *   WB_SALES_RETRY_MS           — базовая пауза при 429/too many requests; по умолчанию 65000.
 *   WB_SALES_RETRY_ATTEMPTS     — число повторов при 429; по умолчанию 2 (паузы растут линейно).
 *   WB_SALES_TIME_BUDGET_MS     — максимум времени на сбор продаж; по умолчанию 240000 (4 мин).
 *                                 Если до следующего retry мы выходим за бюджет — пропускаем
 *                                 retry, чтобы не убить расчёт 6-минутным лимитом Apps Script.
 *   WB_SALES_SKIP               — «1»/«true»/«yes» — пропустить WB sales вовсе.
 *   WB_SALES_CACHE_MIN          — сколько минут кешировать удачную выборку; по умолчанию 30.
 *                                 Кеш в DocumentCache, ключ версионируется по source+metric.
 *
 * Контракт `wbFetchCurrentMonthSalesForPlanning_`:
 *   {
 *     byWb:       { canon(nmId|supplierArticle): qty },
 *     byBarcode:  { canon(barcode): qty },          // у analytics всегда пусто
 *     rows:       <число строк/карточек, попавших в выборку>,
 *     source?:    'analytics' | 'statistics',
 *     metric?:    'orderCount' | 'buyoutCount' | 'netOrderCount' | 'sales',
 *     fromCache?: true,
 *     skipped?:   true
 *   }
 */

const WB_SALES_STATISTICS_BASE_DEFAULT = 'https://statistics-api.wildberries.ru';
const WB_SALES_ANALYTICS_BASE_DEFAULT = 'https://seller-analytics-api.wildberries.ru';
const WB_SALES_ANALYTICS_PATH = '/api/analytics/v3/sales-funnel/products';
const WB_SALES_ANALYTICS_PAGE_SIZE = 1000;
const WB_SALES_ANALYTICS_MAX_PAGES = 50;

function wbSalesProp_(key, fallback) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, fallback);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? fallback : String(v).trim();
}

function wbSalesBoolProp_(key, fallback) {
  const v = String(wbSalesProp_(key, fallback) || '').trim().toLowerCase();
  return v === '1' || v === 'true' || v === 'yes' || v === 'on';
}

function wbSalesTokenOrThrow_() {
  const t = wbSalesProp_('WB_API_TOKEN', '');
  if (!t) throw new Error('Не заполнен WB_API_TOKEN');
  return t;
}

function wbSalesApiErr_(res) {
  if (res.json && res.json.title) return String(res.json.title) + (res.json.detail ? ': ' + res.json.detail : '');
  if (res.json && res.json.error) return String(res.json.error);
  if (res.json && res.json.message) return String(res.json.message);
  return res.text ? res.text.slice(0, 500) : 'HTTP ' + res.code;
}

function wbSalesIsRateLimited_(res) {
  if (res.code === 429) return true;
  const err = wbSalesApiErr_(res).toLowerCase();
  return err.indexOf('too many requests') >= 0;
}

function wbSalesApiFetch_(method, base, path, token, body) {
  const url = String(base).replace(/\/$/, '') + path;
  const opts = {
    method: method,
    headers: { Authorization: token },
    muteHttpExceptions: true
  };
  if (body !== undefined && body !== null) {
    opts.contentType = 'application/json';
    opts.payload = JSON.stringify(body);
  }
  const res = UrlFetchApp.fetch(url, opts);
  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let json = null;
  try { json = JSON.parse(text); } catch (e) {}
  return { code: code, text: text, json: json };
}

function wbSalesApiGet_(base, path, token) {
  return wbSalesApiFetch_('get', base, path, token, null);
}

function wbSalesApiCallWithRetry_(method, base, path, token, body, deadlineMs) {
  let res = wbSalesApiFetch_(method, base, path, token, body);
  if (!wbSalesIsRateLimited_(res)) return res;

  const retryMs = Math.max(1000, parseInt(wbSalesProp_('WB_SALES_RETRY_MS', '65000'), 10) || 65000);
  const attempts = Math.max(1, parseInt(wbSalesProp_('WB_SALES_RETRY_ATTEMPTS', '2'), 10) || 2);
  for (let i = 1; i <= attempts; i++) {
    const wait = retryMs * i;
    // Apps Script-таймаут — 6 минут. Если за это время мы выходим за общий бюджет,
    // лучше вернуть текущий 429 и не убить расчёт.
    if (deadlineMs && Date.now() + wait > deadlineMs) break;
    Utilities.sleep(wait);
    res = wbSalesApiFetch_(method, base, path, token, body);
    if (!wbSalesIsRateLimited_(res)) return res;
  }
  return res;
}

function wbSalesApiGetWithRetry_(base, path, token, deadlineMs) {
  return wbSalesApiCallWithRetry_('get', base, path, token, null, deadlineMs);
}

function wbSalesApiPostWithRetry_(base, path, token, body, deadlineMs) {
  return wbSalesApiCallWithRetry_('post', base, path, token, body, deadlineMs);
}

function wbSalesCurrentMonthFromIso_() {
  const d = new Date();
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') + 'T00:00:00';
}

function wbSalesCurrentMonthRangeYmd_() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const first = new Date(now.getFullYear(), now.getMonth(), 1);
  // pastPeriod должен идти строго до selectedPeriod (валидатор у WB требует «past < current»).
  // Берём предыдущий месяц с теми же датами; getDate() автоматически корректирует «несуществующие»
  // даты типа 31-е (когда в прошлом месяце 30) — JS Date перенесёт на ближайшее валидное.
  const pastStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const pastEnd = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
  return {
    start: Utilities.formatDate(first, tz, 'yyyy-MM-dd'),
    end: Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
    pastStart: Utilities.formatDate(pastStart, tz, 'yyyy-MM-dd'),
    pastEnd: Utilities.formatDate(pastEnd, tz, 'yyyy-MM-dd')
  };
}

function wbSalesQty_(row) {
  const cands = [row.quantity, row.saleQty, row.qty, row.realizedQuantity];
  for (let i = 0; i < cands.length; i++) {
    const v = cands[i];
    if (v == null || v === '') continue;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
    if (isFinite(n)) return n;
  }
  return 1;
}

function wbSalesNormSource_() {
  const s = String(wbSalesProp_('WB_SALES_SOURCE', 'analytics') || 'analytics').trim().toLowerCase();
  if (s === 'statistics' || s === 'stats' || s === 'legacy') return 'statistics';
  if (s === 'auto') return 'auto';
  return 'analytics';
}

function wbSalesNormMetric_() {
  const m = String(wbSalesProp_('WB_SALES_ANALYTICS_METRIC', 'orderCount') || 'orderCount').trim();
  const lm = m.toLowerCase();
  if (lm === 'buyoutcount' || lm === 'buyout' || lm === 'buyouts') return 'buyoutCount';
  if (lm === 'netordercount' || lm === 'netorders' || lm === 'net') return 'netOrderCount';
  return 'orderCount';
}

function wbSalesCacheKey_(source, metric, dateFromIso) {
  return 'wb_sales_month_v2:' + source + ':' + metric + ':' + dateFromIso;
}

function wbSalesNegCacheKey_(source) {
  return 'wb_sales_neg:' + source;
}

function wbSalesNegCacheHas_(source) {
  try {
    const cache = CacheService.getDocumentCache();
    if (!cache) return false;
    return !!cache.get(wbSalesNegCacheKey_(source));
  } catch (e) { return false; }
}

function wbSalesNegCachePut_(source, seconds) {
  try {
    const cache = CacheService.getDocumentCache();
    if (!cache) return;
    cache.put(wbSalesNegCacheKey_(source), '1', Math.max(30, seconds | 0));
  } catch (e) {}
}

function wbSalesLoadFromCache_(source, metric, dateFromIso, needWb, needBc) {
  try {
    const cache = CacheService.getDocumentCache();
    if (!cache) return null;
    const raw = cache.get(wbSalesCacheKey_(source, metric, dateFromIso));
    if (!raw) return null;
    const all = JSON.parse(raw);
    if (!all || typeof all !== 'object') return null;
    const byWb = {};
    const byBarcode = {};
    const wb = all.byWb || {};
    for (const k in wb) if (needWb[k]) byWb[k] = wb[k];
    const bc = all.byBarcode || {};
    for (const k in bc) if (needBc[k]) byBarcode[k] = bc[k];
    return {
      byWb: byWb,
      byBarcode: byBarcode,
      rows: Number(all.rows) || 0,
      source: source,
      metric: metric,
      fromCache: true
    };
  } catch (e) {
    return null;
  }
}

function wbSalesSaveToCache_(source, metric, dateFromIso, byWb, byBarcode, rows) {
  try {
    const cache = CacheService.getDocumentCache();
    if (!cache) return;
    const minutes = Math.max(1, parseInt(wbSalesProp_('WB_SALES_CACHE_MIN', '30'), 10) || 30);
    cache.put(
      wbSalesCacheKey_(source, metric, dateFromIso),
      JSON.stringify({ byWb: byWb, byBarcode: byBarcode, rows: rows }),
      minutes * 60
    );
  } catch (e) {}
}

/**
 * Извлекает массив числовых nmId из ключей needWb (артикулы плана, уже канонизированные).
 * Артикулы поставщика (vendorCode) на этом эндпоинте как фильтр не работают, нужны именно nmId.
 */
function wbSalesExtractNmIdsFromNeed_(needWb) {
  const out = [];
  const seen = {};
  for (const k in needWb) {
    if (!k) continue;
    if (/^\d+$/.test(k)) {
      const n = parseInt(k, 10);
      if (n > 0 && !seen[n]) { seen[n] = true; out.push(n); }
    }
  }
  return out;
}

/**
 * Сборщик через seller-analytics-api /api/analytics/v3/sales-funnel/products.
 * Стратегия — один (или несколько по 1000) запросов с явным фильтром nmIds из плана:
 * это даёт максимум одну страницу при ~188 SKU и резко снижает шанс упереться в
 * лимит 3 запроса/мин у токена. Если у плана нет ни одного числового nmId — запрос
 * вовсе не отправляется. Barcode из эндпоинта не возвращается — byBarcode пустой.
 */
function wbFetchCurrentMonthSalesViaAnalytics_(needWb, needBc) {
  const token = wbSalesTokenOrThrow_();
  const base = wbSalesProp_('WB_SALES_ANALYTICS_API_BASE', WB_SALES_ANALYTICS_BASE_DEFAULT);
  const metric = wbSalesNormMetric_();
  const range = wbSalesCurrentMonthRangeYmd_();
  const timeBudget = Math.max(30000, parseInt(wbSalesProp_('WB_SALES_TIME_BUDGET_MS', '240000'), 10) || 240000);
  const deadlineMs = Date.now() + timeBudget;

  const byWb = {};
  let totalCards = 0;

  const nmIds = wbSalesExtractNmIdsFromNeed_(needWb);
  if (!nmIds.length) {
    return {
      byWb: {}, byBarcode: {}, rows: 0,
      source: 'analytics', metric: metric
    };
  }

  // Бьём nmIds на батчи <= 1000 (максимум limit одной страницы). Между батчами выдерживаем
  // 20-секундный интервал, как требует документация. Обычно у плана < 1000 SKU — будет 1 запрос.
  const BATCH = WB_SALES_ANALYTICS_PAGE_SIZE;
  for (let start = 0; start < nmIds.length; start += BATCH) {
    const batch = nmIds.slice(start, start + BATCH);
    const body = {
      selectedPeriod: { start: range.start, end: range.end },
      pastPeriod: { start: range.pastStart, end: range.pastEnd },
      nmIds: batch,
      limit: BATCH,
      offset: 0
    };
    const res = wbSalesApiPostWithRetry_(base, WB_SALES_ANALYTICS_PATH, token, body, deadlineMs);
    if (res.code < 200 || res.code >= 300) {
      throw new Error('WB analytics: ' + wbSalesApiErr_(res));
    }
    const items = (res.json && res.json.data && Array.isArray(res.json.data.products))
      ? res.json.data.products
      : [];
    totalCards += items.length;

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const product = it && it.product ? it.product : {};
      const sel = it && it.statistic && it.statistic.selected ? it.statistic.selected : {};
      const orderCount = Number(sel.orderCount) || 0;
      const buyoutCount = Number(sel.buyoutCount) || 0;
      const cancelCount = Number(sel.cancelCount) || 0;
      let qty;
      if (metric === 'buyoutCount') qty = buyoutCount;
      else if (metric === 'netOrderCount') qty = Math.max(0, orderCount - cancelCount);
      else qty = orderCount;
      if (!(qty > 0)) continue;

      const nm = product.nmId != null ? ppCanonArticle_(String(product.nmId)) : '';
      const sa = product.vendorCode != null ? ppCanonArticle_(String(product.vendorCode)) : '';
      if (nm && needWb[nm]) byWb[nm] = (byWb[nm] || 0) + qty;
      if (sa && needWb[sa]) byWb[sa] = (byWb[sa] || 0) + qty;
    }

    if (start + BATCH < nmIds.length && Date.now() + 21000 < deadlineMs) {
      Utilities.sleep(21000);
    }
  }

  return {
    byWb: byWb,
    byBarcode: {},
    rows: totalCards,
    source: 'analytics',
    metric: metric
  };
}

/**
 * Старый путь через statistics-api /api/v1/supplier/sales. Оставлен как fallback
 * для совместимости. У него жёсткие лимиты (см. шапку файла).
 */
function wbFetchCurrentMonthSalesViaStatistics_(needWb, needBc) {
  const token = wbSalesTokenOrThrow_();
  const base = wbSalesProp_('WB_STATISTICS_API_BASE', WB_SALES_STATISTICS_BASE_DEFAULT);
  let cursor = wbSalesCurrentMonthFromIso_();
  const byWb = {};
  const byBarcode = {};
  const allByWb = {};
  const allByBarcode = {};
  let totalRows = 0;
  const timeBudget = Math.max(30000, parseInt(wbSalesProp_('WB_SALES_TIME_BUDGET_MS', '240000'), 10) || 240000);
  const deadlineMs = Date.now() + timeBudget;

  for (let page = 0; page < 40; page++) {
    const res = wbSalesApiGetWithRetry_(
      base,
      '/api/v1/supplier/sales?dateFrom=' + encodeURIComponent(cursor),
      token,
      deadlineMs
    );
    if (res.code < 200 || res.code >= 300) throw new Error('WB sales: ' + wbSalesApiErr_(res));
    const rows = Array.isArray(res.json) ? res.json : [];
    if (!rows.length) break;
    totalRows += rows.length;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const qty = wbSalesQty_(r);
      const nm = r.nmId != null ? ppCanonArticle_(String(r.nmId)) : '';
      const sa = r.supplierArticle != null ? ppCanonArticle_(String(r.supplierArticle)) : '';
      const bc = r.barcode != null ? ppCanonBarcodeForStock_(String(r.barcode)) : '';
      if (nm) allByWb[nm] = (allByWb[nm] || 0) + qty;
      if (sa) allByWb[sa] = (allByWb[sa] || 0) + qty;
      if (bc) allByBarcode[bc] = (allByBarcode[bc] || 0) + qty;
      if (nm && needWb[nm]) byWb[nm] = (byWb[nm] || 0) + qty;
      if (sa && needWb[sa]) byWb[sa] = (byWb[sa] || 0) + qty;
      if (bc && needBc[bc]) byBarcode[bc] = (byBarcode[bc] || 0) + qty;
    }
    if (rows.length < 60000) break;
    const last = rows[rows.length - 1];
    const nextFrom = last && last.lastChangeDate ? String(last.lastChangeDate) : '';
    if (!nextFrom || nextFrom === cursor) break;
    cursor = nextFrom;
    Utilities.sleep(61000);
  }
  return {
    byWb: byWb,
    byBarcode: byBarcode,
    rows: totalRows,
    source: 'statistics',
    metric: 'sales',
    _allByWb: allByWb,
    _allByBarcode: allByBarcode
  };
}

/**
 * Продажи WB за текущий месяц — публичная точка входа для procurement_planning.gs.
 * Маршрутизация:
 *   - WB_SALES_SOURCE=analytics  → только analytics
 *   - WB_SALES_SOURCE=statistics → только statistics
 *   - WB_SALES_SOURCE=auto       → analytics, при ошибке — statistics (фолбэк)
 * Поверх — общий слой: skip / кеш / возвращение контракта планировщику.
 */
function wbFetchCurrentMonthSalesForPlanning_(needWb, needBc) {
  if (wbSalesBoolProp_('WB_SALES_SKIP', '')) {
    return { byWb: {}, byBarcode: {}, rows: 0, skipped: true };
  }
  const source = wbSalesNormSource_();
  const metric = source === 'statistics' ? 'sales' : wbSalesNormMetric_();
  const dateFromIso = wbSalesCurrentMonthFromIso_();

  // Сначала — позитивный кеш текущего источника/метрики.
  const cached = wbSalesLoadFromCache_(source === 'auto' ? 'analytics' : source, metric, dateFromIso, needWb, needBc);
  if (cached) return cached;

  const isRateLimitedMsg = function (msg) {
    return msg && /too many requests|429/i.test(String(msg));
  };

  const tryAnalytics = function () {
    // Если только что упирались в 429 — не долбим WB снова, пока бан не отпустит.
    if (wbSalesNegCacheHas_('analytics')) {
      const err = new Error('WB analytics: too many requests (negative cache, повторите через минуту)');
      err.rateLimited = true;
      throw err;
    }
    try {
      const res = wbFetchCurrentMonthSalesViaAnalytics_(needWb, needBc);
      wbSalesSaveToCache_('analytics', res.metric, dateFromIso, res.byWb || {}, {}, res.rows);
      return res;
    } catch (e) {
      if (isRateLimitedMsg(e && e.message)) {
        wbSalesNegCachePut_('analytics', 90);
      }
      throw e;
    }
  };
  const tryStatistics = function () {
    if (wbSalesNegCacheHas_('statistics')) {
      const err = new Error('WB statistics: too many requests (negative cache, повторите через минуту)');
      err.rateLimited = true;
      throw err;
    }
    try {
      const res = wbFetchCurrentMonthSalesViaStatistics_(needWb, needBc);
      wbSalesSaveToCache_('statistics', 'sales', dateFromIso, res._allByWb || {}, res._allByBarcode || {}, res.rows);
      delete res._allByWb;
      delete res._allByBarcode;
      return res;
    } catch (e) {
      if (isRateLimitedMsg(e && e.message)) {
        wbSalesNegCachePut_('statistics', 120);
      }
      throw e;
    }
  };

  if (source === 'statistics') return tryStatistics();
  if (source === 'analytics') return tryAnalytics();

  // auto: analytics, при ошибке — statistics.
  try {
    return tryAnalytics();
  } catch (e) {
    try {
      return tryStatistics();
    } catch (e2) {
      throw new Error(
        'WB sales (auto): analytics → ' + (e && e.message ? e.message : String(e)) +
          '; statistics → ' + (e2 && e2.message ? e2.message : String(e2))
      );
    }
  }
}
