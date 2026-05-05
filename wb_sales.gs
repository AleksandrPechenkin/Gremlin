/**
 * Факт продаж WB за текущий месяц для расчёта закупки.
 *
 * Script Properties:
 *   WB_API_TOKEN
 *   WB_STATISTICS_API_BASE (опционально, по умолчанию https://statistics-api.wildberries.ru)
 */

const WB_SALES_STATISTICS_BASE_DEFAULT = 'https://statistics-api.wildberries.ru';

function wbSalesProp_(key, fallback) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, fallback);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? fallback : String(v).trim();
}

function wbSalesTokenOrThrow_() {
  const t = wbSalesProp_('WB_API_TOKEN', '');
  if (!t) throw new Error('Не заполнен WB_API_TOKEN');
  return t;
}

function wbSalesApiGet_(base, path, token) {
  const url = String(base).replace(/\/$/, '') + path;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {}
  return { code: code, text: text, json: json };
}

function wbSalesApiErr_(res) {
  if (res.json && res.json.title) return String(res.json.title) + (res.json.detail ? ': ' + res.json.detail : '');
  if (res.json && res.json.error) return String(res.json.error);
  return res.text ? res.text.slice(0, 500) : 'HTTP ' + res.code;
}

function wbSalesCurrentMonthFromIso_() {
  const d = new Date();
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') + 'T00:00:00';
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

/**
 * Продажи WB за текущий месяц, агрегированные по ключам из плана.
 * @param {Object.<string, boolean>} needWb
 * @param {Object.<string, boolean>} needBc
 * @return {{byWb:Object.<string,number>, byBarcode:Object.<string,number>, rows:number}}
 */
function wbFetchCurrentMonthSalesForPlanning_(needWb, needBc) {
  const token = wbSalesTokenOrThrow_();
  const base = wbSalesProp_('WB_STATISTICS_API_BASE', WB_SALES_STATISTICS_BASE_DEFAULT);
  let dateFrom = wbSalesCurrentMonthFromIso_();
  const byWb = {};
  const byBarcode = {};
  let totalRows = 0;

  for (let page = 0; page < 40; page++) {
    const res = wbSalesApiGet_(base, '/api/v1/supplier/sales?dateFrom=' + encodeURIComponent(dateFrom), token);
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
      if (nm && needWb[nm]) byWb[nm] = (byWb[nm] || 0) + qty;
      if (sa && needWb[sa]) byWb[sa] = (byWb[sa] || 0) + qty;
      if (bc && needBc[bc]) byBarcode[bc] = (byBarcode[bc] || 0) + qty;
    }
    if (rows.length < 60000) break;
    const last = rows[rows.length - 1];
    const nextFrom = last && last.lastChangeDate ? String(last.lastChangeDate) : '';
    if (!nextFrom || nextFrom === dateFrom) break;
    dateFrom = nextFrom;
    Utilities.sleep(61000);
  }
  return { byWb: byWb, byBarcode: byBarcode, rows: totalRows };
}
