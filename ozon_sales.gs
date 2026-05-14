/**
 * Факт продаж Ozon (FBO + FBS) за текущий месяц для расчёта закупки в книге 03.
 *
 * Источники:
 *   POST /v2/posting/fbo/list  — отправления FBO; в каждом products[] с offer_id и quantity.
 *                                У Ozon для FBO актуальна именно v2 (v3 для FBO нет —
 *                                на него API отвечает синтетической ошибкой валидации
 *                                limit (0, 100], а не 404, что путает).
 *   POST /v3/posting/fbs/list  — отправления FBS; в каждом products[] аналогично.
 *
 * Период — с 00:00 первого числа текущего месяца по «сейчас», по Script Timezone.
 * Из суммы исключаются отмены: posting со status == 'cancelled'.
 *
 * Сопоставление с планом — только по offer_id ↔ canon(«Артикул поставщика»).
 * Barcode-fallback для продаж не делаем, чтобы не вызывать ещё один пакетный info/list:
 * для продаж это обычно избыточно (продажу всегда формирует наш товар с offer_id),
 * а нагрузка на квоты ниже.
 *
 * Script Properties:
 *   OZON_API_TOKEN          — как в ozon_stocks.gs (формат «client_id:api_key»).
 *   OZON_API_BASE           — опционально.
 *   OZON_SALES_PAGE_LIMIT   — размер страницы posting/fbo|fbs/list (по умолчанию 1000, максимум 1000).
 *   OZON_SALES_MAX_PAGES    — предохранитель от бесконечной пагинации (по умолчанию 200).
 *
 * Контракт `ozFetchCurrentMonthSalesForPlanning_`:
 *   {
 *     bySupplier: { canon(offer_id): qty },
 *     byBarcode:  {},                          // пусто (см. выше)
 *     rows: <число просмотренных posting>,
 *     stats: { fboPostings, fbsPostings, fboItems, fbsItems, cancelledSkipped }
 *   }
 */

const OZON_SALES_DEFAULT_PAGE_LIMIT = 1000;
const OZON_SALES_DEFAULT_MAX_PAGES = 200;

function ozSalesPostingsRange_() {
  const now = new Date();
  const since = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const tz = Session.getScriptTimeZone() || 'UTC';
  return {
    since: Utilities.formatDate(since, tz, "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    to: Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ss'Z'")
  };
}

function ozSalesPagedList_(path, token, sinceTo, pageLimit, maxPages) {
  const postings = [];
  let offset = 0;
  for (let page = 0; page < maxPages; page++) {
    const res = ozPost_(path, {
      dir: 'ASC',
      filter: { since: sinceTo.since, to: sinceTo.to },
      limit: pageLimit,
      offset: offset,
      with: { analytics_data: false, financial_data: false }
    }, token);
    if (res.code < 200 || res.code >= 300) {
      throw new Error('Ozon ' + path + ': ' + ozFormatApiError_(res));
    }
    const result = (res.json && (res.json.result || res.json)) || {};
    const batch = Array.isArray(result.postings) ? result.postings : (Array.isArray(result) ? result : []);
    if (!batch.length) break;
    for (let i = 0; i < batch.length; i++) postings.push(batch[i]);
    if (batch.length < pageLimit) break;
    offset += batch.length;
  }
  return postings;
}

function ozFetchCurrentMonthSalesForPlanning_(needSupplier /* unused needBc placeholder kept for symmetry */, needBc) {
  const token = ozParseToken_();
  const pageLimit = Math.max(
    1,
    Math.min(1000, parseInt(ozGetProp_('OZON_SALES_PAGE_LIMIT', String(OZON_SALES_DEFAULT_PAGE_LIMIT)), 10) || OZON_SALES_DEFAULT_PAGE_LIMIT)
  );
  const maxPages = Math.max(1, parseInt(ozGetProp_('OZON_SALES_MAX_PAGES', String(OZON_SALES_DEFAULT_MAX_PAGES)), 10) || OZON_SALES_DEFAULT_MAX_PAGES);
  const sinceTo = ozSalesPostingsRange_();

  const fbo = ozSalesPagedList_('/v2/posting/fbo/list', token, sinceTo, pageLimit, maxPages);
  const fbs = ozSalesPagedList_('/v3/posting/fbs/list', token, sinceTo, pageLimit, maxPages);

  const bySupplier = {};
  const byBarcode = {};
  let fboItems = 0;
  let fbsItems = 0;
  let cancelledSkipped = 0;
  let totalRows = 0;

  function consumePosting_(p, sourceTag) {
    totalRows++;
    const status = String(p.status || '').toLowerCase();
    if (status === 'cancelled') { cancelledSkipped++; return; }
    const products = Array.isArray(p.products) ? p.products : [];
    for (let i = 0; i < products.length; i++) {
      const prod = products[i];
      const qty = parseFloat(prod.quantity) || 0;
      if (!qty) continue;
      const sa = prod.offer_id != null ? ppCanonArticle_(String(prod.offer_id)) : '';
      if (sa && needSupplier && needSupplier[sa]) {
        bySupplier[sa] = (bySupplier[sa] || 0) + qty;
      }
      if (sourceTag === 'fbo') fboItems++;
      else fbsItems++;
    }
  }

  for (let i = 0; i < fbo.length; i++) consumePosting_(fbo[i], 'fbo');
  for (let i = 0; i < fbs.length; i++) consumePosting_(fbs[i], 'fbs');

  return {
    bySupplier: bySupplier,
    byBarcode: byBarcode,
    rows: totalRows,
    stats: {
      fboPostings: fbo.length,
      fbsPostings: fbs.length,
      fboItems: fboItems,
      fbsItems: fbsItems,
      cancelledSkipped: cancelledSkipped
    }
  };
}
