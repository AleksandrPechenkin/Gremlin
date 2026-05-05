/**
 * Планирование закупок: сбор планов продаж («Заказали, шт») по «Артикул ВБ»
 * с вкладок отделов в отдельной книге, запись свода на лист «Планирование закупок»
 * в активной книге (файл заказов).
 *
 * Важно: листы «Проставление планов …» в книге отделов переименовывать нельзя — на них
 * завязаны другие скрипты. Константы ниже совпадают с текущими именами; свойства
 * SALES_PLANS_SHEET_* заполняйте только если имя листа осознанно меняется вместе
 * со всеми зависимыми скриптами.
 *
 * Script Properties (опционально, есть значения по умолчанию):
 *   SALES_PLANS_SPREADSHEET_ID — ID книги с «Проставление планов …» (по умолчанию задан в коде).
 *   SALES_PLANS_SHEET_GREMLIN — имя листа (по умолчанию: Проставление планов Gremlin).
 *   SALES_PLANS_SHEET_OBSHCHIY — имя листа (по умолчанию: Проставление планов Общий).
 *   SALES_PLANS_SHEET_DEPT3 — третий отдел: имя листа в первой книге; если пусто — не читаем.
 *   SALES_PLANS_EXTRA_SPREADSHEET_ID — вторая книга с планом (по умолчанию задана в коде); пусто или OFF — не читать.
 *   SALES_PLANS_EXTRA_SHEET_NAME — лист во второй книге (по умолчанию: Проставление планов).
 *   Доп. книга: в разметке может не быть «Артикул ВБ» — тогда строки сопоставляются со справочником по «Артикул поставщика», количество — «Заказали, шт» (как в остальных планах, не «Продажи»).
 *   PROCUREMENT_PLANNING_SHEET — куда писать свод (по умолчанию: Планирование закупок).
 *   PROCUREMENT_PLANNING_MONTHS — месяцы года, через запятую «5,6,7,8» (номера месяцев, по умолчанию май–август 2026).
 *   PROCUREMENT_PLANNING_YEAR — год (по умолчанию 2026).
 *   PRODUCT_REFERENCE_SPREADSHEET_ID — книга «Справочник товаров» (по умолчанию как в supplier_invoices.gs).
 *   PRODUCT_REFERENCE_SHEET_NAME — лист справочника (по умолчанию: Справочник с названием товаров).
 *   PRODUCT_REFERENCE_PCS_PER_BOX_COL — номер колонки «шт в коробке» (по умолчанию 20 = T).
 *   MS_STOCK_STORES_SHEET — лист перечня складов для остатков (по умолчанию: Склады МС (остатки)).
 *   MS_STORE_SYNC_INCLUDE_ARCHIVED — если 1, в выгрузку попадают и архивные склады МС (по умолчанию: только неархивные).
 *   Учётный остаток: меню «Записать учётный остаток МС…» — /report/stock/all (stock или quantity), склады с «Использовать»;
 *   ключ номенклатуры: meta.href или assortment.meta.href; сопоставление с планом по article, code, externalCode, штрихкодам; в плане ВБ/ШК читаются как на экране (display).
 *   Остаток WB: wildberries_stocks.gs — WB_API_TOKEN, меню «Записать остаток WB…»; см. комментарий в начале wildberries_stocks.gs.
 */

const PP_DEFAULT_SOURCE_SPREADSHEET_ID = '1lMD9ilxKgrFCmFtPCPAFCUMuQh25jaN6lnI7L2WEEnM';
/** Дополнительная книга: вкладка «Проставление планов». */
const PP_DEFAULT_EXTRA_SOURCE_SPREADSHEET_ID = '1SAlQlrYSKYNL72qsYEJU0Syy40w7ZHa1opUZrIXrVzo';
const PP_DEFAULT_EXTRA_SOURCE_SHEET_NAME = 'Проставление планов';
const PP_DEFAULT_PRODUCT_REF_SPREADSHEET_ID = '1PXWd05ENcZGvPYYbAVvf-1EPevdwkxr4IvjRbbojOlg';
const PP_DEFAULT_PRODUCT_REF_SHEET_NAME = 'Справочник с названием товаров';
/** Колонка T в Google Sheets = 20-я колонка (1-based). */
const PP_DEFAULT_PCS_PER_BOX_COL_1BASED = 20;
/** Имена совпадают с листами в книге планов; не переименовывать без согласования со всеми скриптами. */
const PP_DEFAULT_SHEET_GREMLIN = 'Проставление планов Gremlin';
const PP_DEFAULT_SHEET_OBSHCHIY = 'Проставление планов Общий';
const PP_DEFAULT_OUT_SHEET = 'Планирование закупок';
const PP_DEFAULT_MS_STOCK_STORES_SHEET = 'Склады МС (остатки)';
const PP_PROC_PLAN_HEADER_ROW = 5;
const PP_PROC_PLAN_DATA_START_ROW = 6;
const PP_MS_STOCK_COL_HEADER = 'Остаток МС (учётный), шт';
const PP_WB_STOCK_COL_HEADER = 'Остаток ВБ (склады), шт';
const PP_STOCK_CHECK_COL_HEADER = 'Проверка сопоставления (МС/WB)';
const PP_PURCHASE_INBOUND_COL_HEADER = 'В пути до конца горизонта, шт';
const PP_PURCHASE_ADJ_CURR_COL_HEADER = 'План тек. месяца (скорр.), шт';
const PP_PURCHASE_DEFICIT_COL_HEADER = 'Макс дефицит горизонта, шт';
const PP_PURCHASE_RECO_COL_HEADER = 'Рекомендовано к заказу, шт';
const PP_PURCHASE_PRIORITY_COL_HEADER = 'Приоритет отгрузки из Китая';
const PP_PURCHASE_REPORT_SHEET_DEFAULT = 'Планирование закупок (расчёт)';
const PP_SHIPPED_STATUS_CODE_DEFAULT = 'S13_IN_TRANSIT_MOSCOW';
const PP_SCAN_TOP_ROWS = 10;
const PP_HEADER_SCAN_ROWS = 8;
const PP_MAX_COLS_FROM_ANCHOR = 45;

function ppGetProp_(key, defaultValue) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, defaultValue);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? defaultValue : String(v).trim();
}

function ppNormArticle_(s) {
  if (typeof invNorm_ === 'function') return invNorm_(s);
  return String(s == null ? '' : s).trim();
}

function ppCanonArticle_(s) {
  if (typeof invCanon_ === 'function') return invCanon_(ppNormArticle_(s));
  return ppNormArticle_(s)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/ё/g, 'е');
}

function ppParseMonthCell_(val) {
  if (val == null || val === '') return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    const d = val;
    if (isNaN(d.getTime())) return null;
    return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  }
  const str = String(val).trim();
  const m = str.match(/^(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})/);
  if (!m) return null;
  let year = parseInt(m[3], 10);
  if (year < 100) year += 2000;
  return { y: year, m: parseInt(m[2], 10), d: parseInt(m[1], 10) };
}

function ppIsFirstOfMonth_(parsed) {
  return parsed && parsed.d === 1 && parsed.m >= 1 && parsed.m <= 12;
}

function ppIsWbArticleHeader_(h) {
  const c = ppCanonArticle_(h);
  return c.includes('артикул') && (c.includes('вб') || c.includes('wb'));
}

function ppIsOrderedPlanHeader_(h) {
  const c = ppCanonArticle_(h);
  if (!c.includes('заказали') || !c.includes('шт')) return false;
  if (c.includes('вдень')) return false;
  return true;
}

/** Колонка артикула поставщика / Supplier SKU (без ВБ в заголовке). */
function ppIsSupplierSkuColumnHeader_(h) {
  const c = ppCanonArticle_(String(h || ''));
  if (c.includes('артикул') && c.includes('поставщ') && !c.includes('вб') && !c.includes('wb')) return true;
  if (ppCanonHeaderSimple_(h).indexOf('suppliersku') >= 0) return true;
  return false;
}

/**
 * Для листов, где нет «Артикул ВБ»: «Артикул поставщика» и «Заказали, шт» в одной строке шапки, в полосе около якоря месяца.
 */
function ppFindSupplierSkuAndOrderedCols_(headerMatrix, anchorCol) {
  const nRows = headerMatrix.length;
  const nCols = headerMatrix[0] ? headerMatrix[0].length : 0;
  const c0 = Math.max(0, anchorCol - 20);
  const c1 = Math.min(nCols - 1, anchorCol + PP_MAX_COLS_FROM_ANCHOR);
  let best = { skuCol: -1, planCol: -1, headerRow: -1 };
  for (let r = 0; r < nRows; r++) {
    let skuCol = -1;
    let planCol = -1;
    for (let c = c0; c <= c1; c++) {
      const h = headerMatrix[r][c];
      if (ppIsSupplierSkuColumnHeader_(h)) skuCol = c;
      if (ppIsOrderedPlanHeader_(h)) planCol = c;
    }
    if (skuCol >= 0 && planCol >= 0) {
      return { skuCol: skuCol, planCol: planCol, headerRow: r };
    }
    if (skuCol >= 0 && best.skuCol < 0) {
      best = { skuCol: skuCol, planCol: -1, headerRow: r };
    }
  }
  return best;
}

/**
 * План с листа «только артикул поставщика»: суммирует в byArticle по ключу артикула ВБ (через supplierToWb).
 */
function ppExtractPlansForMonthFromSupplierSkuLayout_(snap, year, month, supplierToWb, wbDisplayByKey) {
  const values = snap.values;
  const out = { error: '', byArticle: {}, displayByCanon: {}, unmappableSamples: [] };
  if (!values.length) {
    out.error = 'Пустой лист';
    return out;
  }
  const topRows = Math.min(PP_SCAN_TOP_ROWS, values.length);
  const topMatrix = [];
  for (let r = 0; r < topRows; r++) topMatrix.push(values[r]);
  const anchorCol = ppFindMonthAnchorCol_(topMatrix, year, month);
  if (anchorCol < 0) {
    out.error = 'Не найдена дата 1-го числа для ' + year + '-' + String(month).padStart(2, '0');
    return out;
  }
  const headerEnd = Math.min(values.length, PP_HEADER_SCAN_ROWS);
  const headerMatrix = [];
  for (let r = 0; r < headerEnd; r++) headerMatrix.push(values[r]);
  const found = ppFindSupplierSkuAndOrderedCols_(headerMatrix, anchorCol);
  if (found.skuCol < 0 || found.planCol < 0) {
    out.error =
      'Не найдены колонки «Артикул поставщика» и «Заказали, шт» у месяца ' + year + '-' + String(month).padStart(2, '0');
    return out;
  }
  const skuCol = found.skuCol;
  const planCol = found.planCol;
  const dataStart = found.headerRow + 2;
  const seenUnmapped = {};
  for (let r = dataStart - 1; r < values.length; r++) {
    const row = values[r];
    const rawSku = skuCol < row.length ? row[skuCol] : '';
    const sku = ppNormArticle_(rawSku);
    if (!sku) continue;
    const skuKey = ppCanonArticle_(sku);
    const wbKey = supplierToWb[skuKey];
    if (!wbKey) {
      if (!seenUnmapped[skuKey] && out.unmappableSamples.length < 8) {
        seenUnmapped[skuKey] = true;
        out.unmappableSamples.push(sku);
      }
      continue;
    }
    const rawPlan = planCol < row.length ? row[planCol] : '';
    const qty = typeof parseNumber === 'function' ? parseNumber(rawPlan) : null;
    const n = qty == null || isNaN(qty) ? 0 : qty;
    out.byArticle[wbKey] = (out.byArticle[wbKey] || 0) + n;
    if (!out.displayByCanon[wbKey]) {
      out.displayByCanon[wbKey] = wbDisplayByKey[wbKey] || sku;
    }
  }
  return out;
}

/**
 * Ищет колонку с датой 1-го числа целевого месяца (в первых PP_SCAN_TOP_ROWS строках).
 * Берётся самая левая подходящая колонка для данного месяца.
 */
function ppFindMonthAnchorCol_(topMatrix, year, month) {
  const maxCol = topMatrix[0] ? topMatrix[0].length : 0;
  for (let c = 0; c < maxCol; c++) {
    for (let r = 0; r < topMatrix.length; r++) {
      const parsed = ppParseMonthCell_(topMatrix[r][c]);
      if (!ppIsFirstOfMonth_(parsed)) continue;
      if (parsed.y === year && parsed.m === month) return c;
    }
  }
  return -1;
}

/**
 * В строках под шапкой ищем заголовок «Артикул ВБ» и «Заказали, шт» около якоря месяца.
 */
function ppFindWbAndPlanCols_(headerMatrix, anchorCol) {
  const nRows = headerMatrix.length;
  const nCols = headerMatrix[0] ? headerMatrix[0].length : 0;
  const c0 = Math.max(0, anchorCol - 3);
  const c1 = Math.min(nCols - 1, anchorCol + PP_MAX_COLS_FROM_ANCHOR);
  let wbCol = -1;
  let headerRow = -1;
  for (let r = 0; r < nRows; r++) {
    for (let c = c0; c <= c1; c++) {
      const h = headerMatrix[r][c];
      if (ppIsWbArticleHeader_(h)) {
        wbCol = c;
        headerRow = r;
        break;
      }
    }
    if (wbCol >= 0) break;
  }
  if (wbCol < 0) {
    wbCol = anchorCol;
    headerRow = Math.min(headerMatrix.length - 1, 3);
  }
  let planCol = -1;
  const scanRow = headerRow >= 0 ? headerRow : Math.min(headerMatrix.length - 1, 3);
  for (let c = wbCol + 1; c <= Math.min(nCols - 1, wbCol + PP_MAX_COLS_FROM_ANCHOR); c++) {
    if (ppIsOrderedPlanHeader_(headerMatrix[scanRow][c])) {
      planCol = c;
      break;
    }
  }
  return { wbCol: wbCol, planCol: planCol, headerRow: headerRow >= 0 ? headerRow : scanRow };
}

function ppReadSheetSnapshot_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { values: [], lastRow: 0, lastCol: 0 };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  return { values: values, lastRow: lastRow, lastCol: lastCol };
}

function ppExtractPlansForMonth_(snap, year, month) {
  const values = snap.values;
  if (!values.length) return { error: 'Пустой лист', byArticle: {}, displayByCanon: {} };
  const topRows = Math.min(PP_SCAN_TOP_ROWS, values.length);
  const topMatrix = [];
  for (let r = 0; r < topRows; r++) topMatrix.push(values[r]);
  const anchorCol = ppFindMonthAnchorCol_(topMatrix, year, month);
  if (anchorCol < 0) {
    return {
      error: 'Не найдена дата 1-го числа для ' + year + '-' + String(month).padStart(2, '0'),
      byArticle: {},
      displayByCanon: {}
    };
  }
  const headerEnd = Math.min(values.length, PP_HEADER_SCAN_ROWS);
  const headerMatrix = [];
  for (let r = 0; r < headerEnd; r++) headerMatrix.push(values[r]);
  const found = ppFindWbAndPlanCols_(headerMatrix, anchorCol);
  if (found.planCol < 0) {
    return {
      error: 'Не найдена колонка «Заказали, шт» у месяца ' + year + '-' + String(month).padStart(2, '0'),
      byArticle: {},
      displayByCanon: {}
    };
  }
  const wbCol = found.wbCol;
  const planCol = found.planCol;
  const dataStart = found.headerRow + 2;
  const byArticle = {};
  const displayByCanon = {};
  for (let r = dataStart - 1; r < values.length; r++) {
    const row = values[r];
    const rawWb = wbCol < row.length ? row[wbCol] : '';
    const wb = ppNormArticle_(rawWb);
    if (!wb) continue;
    const key = ppCanonArticle_(wb);
    if (!displayByCanon[key]) displayByCanon[key] = wb;
    const rawPlan = planCol < row.length ? row[planCol] : '';
    const qty = typeof parseNumber === 'function' ? parseNumber(rawPlan) : null;
    const n = qty == null || isNaN(qty) ? 0 : qty;
    byArticle[key] = (byArticle[key] || 0) + n;
  }
  return { error: '', byArticle: byArticle, displayByCanon: displayByCanon, wbCol: wbCol, planCol: planCol };
}

function ppParseMonthList_() {
  const year = parseInt(ppGetProp_('PROCUREMENT_PLANNING_YEAR', '2026'), 10) || 2026;
  const raw = ppGetProp_('PROCUREMENT_PLANNING_MONTHS', '5,6,7,8');
  const parts = String(raw)
    .split(/[,;\s]+/)
    .map(function (s) {
      return parseInt(String(s).trim(), 10);
    })
    .filter(function (m) {
      return m >= 1 && m <= 12;
    });
  const months = parts.length ? parts : [5, 6, 7, 8];
  return { year: year, months: months };
}

function ppMergeInto_(target, addMap) {
  for (const k in addMap) {
    target[k] = (target[k] || 0) + addMap[k];
  }
}

function ppMergeDisplay_(target, addDisp) {
  if (!addDisp) return;
  for (const k in addDisp) {
    if (!target[k]) target[k] = addDisp[k];
  }
}

/**
 * Суммирует планы с указанных листов одной книги в perMonth / displayByCanon.
 * @param {string} sourceTag подпись в предупреждениях (например имя книги).
 */
function ppMergeSourceSheets_(ss, sheetNames, year, months, monthKeys, perMonth, displayByCanon, warnings, sourceTag) {
  const tag = sourceTag ? '[' + sourceTag + '] ' : '';
  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) {
      warnings.push(tag + 'Лист не найден: «' + sheetNames[si] + '»');
      continue;
    }
    const snap = ppReadSheetSnapshot_(sh);
    for (let mi = 0; mi < months.length; mi++) {
      const m = months[mi];
      const mk = monthKeys[mi];
      const res = ppExtractPlansForMonth_(snap, year, m);
      if (res.error) warnings.push(tag + '«' + sheetNames[si] + '» / ' + mk + ': ' + res.error);
      ppMergeInto_(perMonth[mk], res.byArticle);
      ppMergeDisplay_(displayByCanon, res.displayByCanon);
    }
  }
}

/**
 * Доп. книга: только «Артикул поставщика» + «Заказали, шт»; ключи схлопываются в артикул ВБ через справочник.
 */
function ppMergeExtraPlanSheets_(
  ss,
  sheetNames,
  year,
  months,
  monthKeys,
  perMonth,
  displayByCanon,
  warnings,
  sourceTag,
  supplierToWb,
  wbDisplayByKey
) {
  const tag = sourceTag ? '[' + sourceTag + '] ' : '';
  const stw = supplierToWb || {};
  const disp = wbDisplayByKey || {};
  for (let si = 0; si < sheetNames.length; si++) {
    const sh = ss.getSheetByName(sheetNames[si]);
    if (!sh) {
      warnings.push(tag + 'Лист не найден: «' + sheetNames[si] + '»');
      continue;
    }
    const snap = ppReadSheetSnapshot_(sh);
    for (let mi = 0; mi < months.length; mi++) {
      const m = months[mi];
      const mk = monthKeys[mi];
      const res = ppExtractPlansForMonthFromSupplierSkuLayout_(snap, year, m, stw, disp);
      if (res.error) warnings.push(tag + '«' + sheetNames[si] + '» / ' + mk + ': ' + res.error);
      else {
        ppMergeInto_(perMonth[mk], res.byArticle);
        ppMergeDisplay_(displayByCanon, res.displayByCanon);
        if (res.unmappableSamples && res.unmappableSamples.length) {
          warnings.push(
            tag +
              '«' +
              sheetNames[si] +
              '» / ' +
              mk +
              ': не сопоставлены со справочником (примеры): ' +
              res.unmappableSamples.join(', ')
          );
        }
      }
    }
  }
}

function ppGetExtraSalesPlansSpreadsheetId_() {
  const raw = PropertiesService.getScriptProperties().getProperty('SALES_PLANS_EXTRA_SPREADSHEET_ID');
  if (raw === null) return PP_DEFAULT_EXTRA_SOURCE_SPREADSHEET_ID;
  const t = String(raw).trim();
  if (t === '' || t.toUpperCase() === 'OFF' || t === '-') return '';
  return t;
}

function ppCanonHeaderSimple_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/ё/g, 'е')
    .replace(/\s+/g, '');
}

function ppFindWbColInRefHeaders_(headers) {
  for (let i = 0; i < headers.length; i++) {
    if (ppIsWbArticleHeader_(headers[i])) return i;
  }
  return 4;
}

function ppFindBarcodeColInRefHeaders_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(headers[i]);
    if (c === 'шк' || c === 'barcode' || c.indexOf('штрих') >= 0 || c.indexOf('ean') >= 0) return i;
  }
  return -1;
}

function ppFindSupplierArticleColInRefHeaders_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(headers[i]);
    if (c.indexOf('артикул') < 0) continue;
    if (c.indexOf('поставщ') < 0) continue;
    if (c.indexOf('вб') >= 0 || c.indexOf('wb') >= 0) continue;
    return i;
  }
  return -1;
}

function ppParsePositiveNumber_(v) {
  if (typeof parseNumber === 'function') {
    const n = parseNumber(v);
    return n != null && !isNaN(n) && n > 0 ? n : 0;
  }
  const s = String(v == null ? '' : v)
    .replace(/\s/g, '')
    .replace(',', '.');
  const n = parseFloat(s);
  return isFinite(n) && n > 0 ? n : 0;
}

/**
 * Округление плана закупки вверх до целого числа коробок.
 * @param {number} qty суммарный план, шт
 * @param {number} pcsPerBox штук в коробке (>0)
 */
function ppRoundUpToCarton_(qty, pcsPerBox) {
  if (qty == null || qty <= 0) return 0;
  if (!(pcsPerBox > 0)) return qty;
  return Math.ceil(qty / pcsPerBox) * pcsPerBox;
}

/**
 * Справочник: refMap по артикулу ВБ, supplierToWb (канон артикула поставщика → канон ВБ), wbDisplayByKey.
 * @return {{ refMap: Object, supplierToWb: Object, wbDisplayByKey: Object }}
 */
function ppLoadProductReferenceBundle_(warnings) {
  const empty = { refMap: {}, supplierToWb: {}, wbDisplayByKey: {} };
  const refId = ppGetProp_('PRODUCT_REFERENCE_SPREADSHEET_ID', PP_DEFAULT_PRODUCT_REF_SPREADSHEET_ID);
  const sheetName = ppGetProp_('PRODUCT_REFERENCE_SHEET_NAME', PP_DEFAULT_PRODUCT_REF_SHEET_NAME);
  const pcsCol1 = parseInt(ppGetProp_('PRODUCT_REFERENCE_PCS_PER_BOX_COL', String(PP_DEFAULT_PCS_PER_BOX_COL_1BASED)), 10);
  const pcsCol0 = (isFinite(pcsCol1) && pcsCol1 > 0 ? pcsCol1 : PP_DEFAULT_PCS_PER_BOX_COL_1BASED) - 1;

  const map = {};
  const supplierToWb = {};
  const wbDisplayByKey = {};
  if (!refId) {
    warnings.push('Справочник товаров: не задан PRODUCT_REFERENCE_SPREADSHEET_ID.');
    return empty;
  }
  try {
    const ss = SpreadsheetApp.openById(refId);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      warnings.push('Справочник товаров: лист «' + sheetName + '» не найден.');
      return empty;
    }
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), pcsCol0 + 1);
    if (lastRow < 2) {
      warnings.push('Справочник товаров: нет строк данных.');
      return empty;
    }

    let headerRow1Based = 1;
    let headers = [];
    for (let hr = 1; hr <= Math.min(5, lastRow); hr++) {
      const rowH = sh.getRange(hr, 1, 1, lastCol).getValues()[0];
      for (let i = 0; i < rowH.length; i++) {
        if (ppIsWbArticleHeader_(rowH[i])) {
          headers = rowH;
          headerRow1Based = hr;
          break;
        }
      }
      if (headers.length) break;
    }
    if (!headers.length) {
      headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      headerRow1Based = 1;
    }

    const wbCol = ppFindWbColInRefHeaders_(headers);
    const bcCol = ppFindBarcodeColInRefHeaders_(headers);
    const saCol = ppFindSupplierArticleColInRefHeaders_(headers);
    const firstDataRow = headerRow1Based + 1;
    if (firstDataRow > lastRow) {
      warnings.push('Справочник товаров: нет строк под шапкой.');
      return empty;
    }

    const numDataRows = lastRow - firstDataRow + 1;
    const data = sh.getRange(firstDataRow, 1, numDataRows, lastCol).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const wb = ppNormArticle_(wbCol < row.length ? row[wbCol] : '');
      if (!wb) continue;
      const key = ppCanonArticle_(wb);
      const barcode = bcCol >= 0 && bcCol < row.length && row[bcCol] != null ? String(row[bcCol]).trim() : '';
      const supplierArticle = saCol >= 0 && saCol < row.length && row[saCol] != null ? String(row[saCol]).trim() : '';
      const pcs = ppParsePositiveNumber_(pcsCol0 < row.length ? row[pcsCol0] : '');

      if (!wbDisplayByKey[key]) wbDisplayByKey[key] = wb;

      if (!map[key]) {
        map[key] = { barcode: '', supplierArticle: '', pcsPerBox: 0 };
      }
      const slot = map[key];
      if (barcode && !slot.barcode) slot.barcode = barcode;
      if (supplierArticle && !slot.supplierArticle) slot.supplierArticle = supplierArticle;
      if (pcs > 0 && !(slot.pcsPerBox > 0)) slot.pcsPerBox = pcs;

      if (supplierArticle) {
        const sk = ppCanonArticle_(supplierArticle);
        if (sk && supplierToWb[sk] == null) supplierToWb[sk] = key;
      }
    }
  } catch (e) {
    warnings.push('Справочник товаров: ' + (e.message || String(e)));
  }
  return { refMap: map, supplierToWb: supplierToWb, wbDisplayByKey: wbDisplayByKey };
}

/**
 * Собирает планы с листов отделов и записывает свод на «Планирование закупок».
 */
function refreshProcurementPlanningFromSalesSheets() {
  const sourceId = ppGetProp_('SALES_PLANS_SPREADSHEET_ID', PP_DEFAULT_SOURCE_SPREADSHEET_ID);
  if (!sourceId) {
    SpreadsheetApp.getUi().alert('Задайте SALES_PLANS_SPREADSHEET_ID в свойствах скрипта.');
    return;
  }
  const sheetNames = [
    ppGetProp_('SALES_PLANS_SHEET_GREMLIN', PP_DEFAULT_SHEET_GREMLIN),
    ppGetProp_('SALES_PLANS_SHEET_OBSHCHIY', PP_DEFAULT_SHEET_OBSHCHIY)
  ];
  const third = ppGetProp_('SALES_PLANS_SHEET_DEPT3', '');
  if (third) sheetNames.push(third);

  const { year, months } = ppParseMonthList_();
  const monthKeys = months.map(function (m) {
    return year + '-' + String(m).padStart(2, '0');
  });

  const perMonth = {};
  const displayByCanon = {};
  const warnings = [];
  for (let mi = 0; mi < months.length; mi++) {
    perMonth[monthKeys[mi]] = {};
  }

  const refBundle = ppLoadProductReferenceBundle_(warnings);

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceId);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Не удалось открыть книгу планов по ID. Проверьте доступ и SALES_PLANS_SPREADSHEET_ID.\n' + e.message);
    return;
  }

  ppMergeSourceSheets_(sourceSs, sheetNames, year, months, monthKeys, perMonth, displayByCanon, warnings, sourceSs.getName());

  let extraSs = null;
  const extraId = ppGetExtraSalesPlansSpreadsheetId_();
  const extraSheetName = ppGetProp_('SALES_PLANS_EXTRA_SHEET_NAME', PP_DEFAULT_EXTRA_SOURCE_SHEET_NAME);
  if (extraId) {
    try {
      extraSs = SpreadsheetApp.openById(extraId);
      ppMergeExtraPlanSheets_(
        extraSs,
        [extraSheetName],
        year,
        months,
        monthKeys,
        perMonth,
        displayByCanon,
        warnings,
        extraSs.getName(),
        refBundle.supplierToWb,
        refBundle.wbDisplayByKey
      );
    } catch (e) {
      warnings.push('Доп. книга планов: не удалось открыть — ' + (e.message || String(e)));
    }
  }

  const allKeys = {};
  for (let ki = 0; ki < monthKeys.length; ki++) {
    const mk = monthKeys[ki];
    for (const k in perMonth[mk]) allKeys[k] = true;
  }
  const sorted = Object.keys(allKeys).sort();

  const refMap = refBundle.refMap;
  let missingRefCount = 0;
  let missingPcsCount = 0;

  const destSs = SpreadsheetApp.getActiveSpreadsheet();
  const outName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  let out = destSs.getSheetByName(outName);
  if (!out) out = destSs.insertSheet(outName);

  const header = ['Артикул ВБ', 'ШК', 'Наименование'].concat(monthKeys);
  const rows = [header];
  for (let i = 0; i < sorted.length; i++) {
    const key = sorted[i];
    const ref = refMap[key];
    if (!ref) {
      missingRefCount++;
    }
    const slot = ref || { barcode: '', supplierArticle: '', pcsPerBox: 0 };
    const line = [displayByCanon[key] || key, slot.barcode, slot.supplierArticle];
    let rowNeedsPcs = false;
    for (let j = 0; j < monthKeys.length; j++) {
      const raw = perMonth[monthKeys[j]][key] || 0;
      let outQty = ppRoundUpToCarton_(raw, slot.pcsPerBox);
      if (raw > 0 && !(slot.pcsPerBox > 0)) {
        outQty = raw;
        rowNeedsPcs = true;
      }
      line.push(outQty);
    }
    if (rowNeedsPcs) missingPcsCount++;
    rows.push(line);
  }

  if (missingRefCount > 0) {
    warnings.push(
      'Артикулов без совпадения в справочнике товаров: ' +
        missingRefCount +
        ' (ШК и наименование пустые, по месяцам — без округления до коробки).'
    );
  }
  if (missingPcsCount > 0) {
    warnings.push(
      'Строк с планом >0, но без положительного «шт в коробке» (колонка T): ' +
        missingPcsCount +
        ' (количества по месяцам оставлены как в плане).'
    );
  }

  out.clearContents();
  out.getRange(1, 1).setValue('Обновлено (планы продаж)');
  out.getRange(1, 2).setValue(new Date());
  out.getRange(2, 1).setValue('Источник (книги)');
  const row2Urls = extraSs ? sourceSs.getUrl() + '\n' + extraSs.getUrl() : sourceSs.getUrl();
  out.getRange(2, 2).setValue(row2Urls).setWrap(true);
  out.getRange(3, 1).setValue('Листы');
  let row3Sheets = sheetNames.join(', ');
  if (extraSs) row3Sheets += ' | доп. книга: «' + extraSheetName + '»';
  out.getRange(3, 2).setValue(row3Sheets);
  out.getRange(4, 1).setValue(
    'Наименование = артикул поставщика из справочника. Месяцы: план округлён вверх до целых коробок (шт/кор из справочника, по умолчанию колонка T). Колонку остатка МС после полного обновления планов заполните снова: меню «Записать учётный остаток МС…».'
  );
  const startRow = 5;
  if (rows.length) {
    out.getRange(startRow, 1, rows.length, header.length).setValues(rows);
    out.getRange(startRow, 1, rows.length, 1).setNumberFormat('@');
    if (rows.length > 1) {
      out.getRange(startRow + 1, 2, rows.length - 1, 1).setNumberFormat('@');
    }
    const numMonthCols = monthKeys.length;
    if (numMonthCols > 0) {
      out.getRange(startRow + 1, 4, rows.length - 1, numMonthCols).setNumberFormat('0');
    }
  }

  const msg =
    'Готово. Артикулов в своде: ' +
    sorted.length +
    (warnings.length ? '\n\nПредупреждения:\n' + warnings.slice(0, 12).join('\n') + (warnings.length > 12 ? '\n…' : '') : '');
  SpreadsheetApp.getUi().alert(msg);
  if (typeof logInfo === 'function') logInfo('procurement_planning', { warnings: warnings, articles: sorted.length });
}

function ppMsStoreHref_(storeId) {
  return 'https://api.moysklad.ru/api/remap/1.2/entity/store/' + String(storeId).trim();
}

function ppReadEnabledStoreIdsForStock_(ss) {
  const name = ppGetProp_('MS_STOCK_STORES_SHEET', PP_DEFAULT_MS_STOCK_STORES_SHEET);
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 1, n, 2).getValues();
  const ids = [];
  for (let i = 0; i < vals.length; i++) {
    const use = vals[i][0];
    const id = String(vals[i][1] || '').trim();
    if (!id) continue;
    if (use === true || String(use).toLowerCase() === 'да') ids.push(id);
  }
  return ids;
}

function ppExtractBarcodeStringFromMs_(b) {
  if (b == null) return '';
  if (typeof b === 'string') return ppNormArticle_(b);
  if (typeof b === 'object') {
    const o = b;
    const keys = ['ean13', 'ean8', 'code128', 'gtin', 'code'];
    for (let k = 0; k < keys.length; k++) {
      if (o[keys[k]] != null && String(o[keys[k]]).trim() !== '') return String(o[keys[k]]).trim();
    }
  }
  return '';
}

function ppCanonBarcodeForStock_(s) {
  return String(ppNormArticle_(s).replace(/\s/g, '')).toLowerCase().replace(/ё/g, 'е');
}

/** Ссылка на товар/модификацию в строке отчёта stock/all: вложенный assortment или корневой meta (Remap 1.2). */
function ppMsStockRowAssortmentHref_(row) {
  const asm = row.assortment;
  if (asm && asm.meta && asm.meta.href) return String(asm.meta.href).trim();
  if (row.meta && row.meta.href) return String(row.meta.href).trim();
  return '';
}

/** Учётный остаток: stock; если пусто — quantity (разные версии отчёта МС). */
function ppParseMsStockRowQty_(row) {
  const raw = row.stock != null && row.stock !== '' ? row.stock : row.quantity;
  const stk = typeof raw === 'number' ? raw : parseFloat(String(raw == null ? '' : raw).replace(/\s/g, '').replace(',', '.'));
  return isFinite(stk) ? stk : 0;
}

/** Собирает артикулы/коды/штрихкоды из строки отчёта и из вложенного assortment (если есть). */
function ppMsStockRowPushMeta_(m, row) {
  const ra = row.article != null ? String(row.article).trim() : '';
  if (ra) m.articles.push(ra);
  const rc = row.code != null ? String(row.code).trim() : '';
  if (rc) m.codes.push(rc);
  const ext = row.externalCode != null ? String(row.externalCode).trim() : '';
  if (ext) m.codes.push(ext);
  const bars = row.barcodes;
  if (bars && bars.length) {
    for (let bi = 0; bi < bars.length; bi++) m.barcodes.push(bars[bi]);
  }
  const asm = row.assortment;
  if (!asm || typeof asm !== 'object') return;
  const ra2 = asm.article != null ? String(asm.article).trim() : '';
  if (ra2) m.articles.push(ra2);
  const rc2 = asm.code != null ? String(asm.code).trim() : '';
  if (rc2) m.codes.push(rc2);
  const ext2 = asm.externalCode != null ? String(asm.externalCode).trim() : '';
  if (ext2) m.codes.push(ext2);
  const bars2 = asm.barcodes;
  if (bars2 && bars2.length) {
    for (let bi = 0; bi < bars2.length; bi++) m.barcodes.push(bars2[bi]);
  }
}

function ppFetchStockAllRowsForStore_(storeId) {
  const filterVal = 'store=' + ppMsStoreHref_(storeId);
  const out = [];
  let offset = 0;
  const limit = 1000;
  while (true) {
    const path =
      '/report/stock/all?limit=' +
      limit +
      '&offset=' +
      offset +
      '&include=' +
      encodeURIComponent('zeroLines') +
      '&filter=' +
      encodeURIComponent(filterVal);
    const res = msFetch(path, 'get');
    if (!res.success) {
      if (offset === 0) throw new Error(res.error || 'Ошибка отчёта /report/stock/all');
      break;
    }
    const rows = res.data && res.data.rows ? res.data.rows : [];
    if (!rows.length) break;
    for (let i = 0; i < rows.length; i++) out.push(rows[i]);
    if (rows.length < limit) break;
    offset += limit;
    Utilities.sleep(60);
  }
  return out;
}

/** Ключи артикулов ВБ и ШК из колонок планирования (только непустые ячейки). */
function ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals) {
  const needWb = {};
  const needBc = {};
  for (let i = 0; i < wbVals.length; i++) {
    const w = ppCanonArticle_(wbVals[i][0]);
    if (w) needWb[w] = true;
    const b = ppCanonBarcodeForStock_(bcVals[i][0]);
    if (b) needBc[b] = true;
  }
  return { needWb: needWb, needBc: needBc };
}

/**
 * Суммирует учётный остаток по выбранным складам только для ключей из планирования (needWb / needBc).
 * Сопоставление с МС: article, code, штрихкоды из отчёта (мета по assortment.meta.href сливается со всех строк).
 * @return {{ stockByWb: Object.<string, number>, stockByBarcode: Object.<string, number> }}
 */
function ppBuildMsAccountingStockLookup_(storeIds, needWb, needBc) {
  const stockByWb = {};
  const stockByBarcode = {};
  const byHref = {};
  const metaByHref = {};

  for (let si = 0; si < storeIds.length; si++) {
    const rows = ppFetchStockAllRowsForStore_(storeIds[si]);
    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const href = ppMsStockRowAssortmentHref_(row);
      if (!href) continue;
      const n = ppParseMsStockRowQty_(row);
      byHref[href] = (byHref[href] || 0) + n;
      if (!metaByHref[href]) {
        metaByHref[href] = { articles: [], codes: [], barcodes: [] };
      }
      ppMsStockRowPushMeta_(metaByHref[href], row);
    }
  }

  for (const href in byHref) {
    const qty = byHref[href];
    const meta = metaByHref[href] || { articles: [], codes: [], barcodes: [] };
    const wbHit = {};
    const bcHit = {};

    for (let ai = 0; ai < meta.articles.length; ai++) {
      const k = ppCanonArticle_(meta.articles[ai]);
      if (k && needWb[k]) wbHit[k] = true;
    }
    for (let ci = 0; ci < meta.codes.length; ci++) {
      const raw = meta.codes[ci];
      const kArt = ppCanonArticle_(raw);
      if (kArt && needWb[kArt]) wbHit[kArt] = true;
      const kBc = ppCanonBarcodeForStock_(raw);
      if (kBc && needBc[kBc]) bcHit[kBc] = true;
    }
    for (let bi = 0; bi < meta.barcodes.length; bi++) {
      const s = ppExtractBarcodeStringFromMs_(meta.barcodes[bi]);
      if (!s) continue;
      const bk = ppCanonBarcodeForStock_(s);
      if (bk && needBc[bk]) bcHit[bk] = true;
      const kAsWb = ppCanonArticle_(s);
      if (kAsWb && needWb[kAsWb]) wbHit[kAsWb] = true;
    }

    for (const k in wbHit) stockByWb[k] = (stockByWb[k] || 0) + qty;
    for (const k in bcHit) stockByBarcode[k] = (stockByBarcode[k] || 0) + qty;
  }

  return { stockByWb: stockByWb, stockByBarcode: stockByBarcode };
}

function ppPlanningFindWbBarcodeCol0_(headers) {
  let wb = -1;
  let bc = -1;
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (wb < 0 && c.indexOf('артикул') >= 0 && (c.indexOf('вб') >= 0 || c.indexOf('wb') >= 0)) wb = i;
    if (bc < 0 && (c === 'шк' || c.indexOf('штрих') >= 0 || c === 'barcode')) bc = i;
  }
  if (wb < 0) wb = 0;
  if (bc < 0) bc = 1;
  return { wb: wb, bc: bc };
}

function ppPlanningFindMsStockCol0_(headers) {
  const target = ppCanonHeaderSimple_(PP_MS_STOCK_COL_HEADER);
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (c === target) return i;
    if (c.indexOf('остатокмс') >= 0 && (c.indexOf('учет') >= 0 || c.indexOf('учёт') >= 0)) return i;
  }
  return -1;
}

function ppPlanningFindWbStockCol0_(headers) {
  const target = ppCanonHeaderSimple_(PP_WB_STOCK_COL_HEADER);
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (c === target) return i;
    if (c.indexOf('остатоквб') >= 0 || (c.indexOf('остаток') >= 0 && c.indexOf('wb') >= 0)) return i;
  }
  return -1;
}

function ppPlanningFindStockCheckCol0_(headers) {
  const target = ppCanonHeaderSimple_(PP_STOCK_CHECK_COL_HEADER);
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (c === target) return i;
    if (c.indexOf('проверкасопоставления') >= 0 || c.indexOf('проверкаостатков') >= 0) return i;
  }
  return -1;
}

function ppPlanningFindColByHeaderVariants_(headers, variants) {
  const canonVariants = variants.map(function (v) {
    return ppCanonHeaderSimple_(v);
  });
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    for (let j = 0; j < canonVariants.length; j++) {
      if (c === canonVariants[j]) return i;
    }
  }
  return -1;
}

function ppPlanningFindMonthCols_(headers) {
  const cols = [];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (/^\d{4}-\d{2}$/.test(h)) cols.push({ key: h, col0: i });
  }
  return cols;
}

function ppParseAnyDate_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = String(v).trim();
  if (!s) return null;
  const m1 = s.match(/^(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})$/);
  if (m1) {
    let y = parseInt(m1[3], 10);
    if (y < 100) y += 2000;
    return new Date(y, parseInt(m1[2], 10) - 1, parseInt(m1[1], 10));
  }
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function ppMonthStartFromPeriodToken_(token) {
  const s = String(token || '').trim();
  const m = s.match(/^(\d{1,2})[\/\.\-](\d{2,4})$/);
  if (!m) return null;
  let y = parseInt(m[2], 10);
  if (y < 100) y += 2000;
  const mm = parseInt(m[1], 10);
  if (!(mm >= 1 && mm <= 12)) return null;
  return new Date(y, mm - 1, 1);
}

function ppMonthKeyFromDate_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function ppMonthDiff_(a, b) {
  return (a.getFullYear() - b.getFullYear()) * 12 + (a.getMonth() - b.getMonth());
}

function ppMonthFirstDateFromKey_(k) {
  const m = String(k || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, 1);
}

function ppToNumOr0_(v) {
  const n = typeof parseNumber === 'function' ? parseNumber(v) : null;
  return n == null || !isFinite(n) ? 0 : n;
}

function ppReadInboundRowsForPlanning_(ss) {
  const sh = ss.getSheetByName('Сводная');
  if (!sh || sh.getLastRow() < 2) return { rows: [], headerRow: 1 };
  const lastCol = sh.getLastColumn();
  const maxScan = Math.min(sh.getLastRow(), 5);
  let headerRow = 1;
  for (let r = 1; r <= maxScan; r++) {
    const hdr = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const idx = ppPlanningFindColByHeaderVariants_(hdr, ['Артикул ВБ', 'Артикул WB']);
    if (idx >= 0) {
      headerRow = r;
      break;
    }
  }
  const rowsCount = sh.getLastRow() - headerRow;
  if (rowsCount <= 0) return { rows: [], headerRow: headerRow };
  const hdr = sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const vals = sh.getRange(headerRow + 1, 1, rowsCount, lastCol).getValues();
  const disp = sh.getRange(headerRow + 1, 1, rowsCount, lastCol).getDisplayValues();
  return { rows: vals, rowsDisplay: disp, headers: hdr, headerRow: headerRow };
}

function ppBuildInboundByMonthForPlanning_(ss, monthKeys, needWb, needBc) {
  const dump = ppReadInboundRowsForPlanning_(ss);
  if (!dump.rows.length) return { byWbMonth: {}, byBcMonth: {}, errors: ['Лист «Сводная» пуст или не найден.'], usedRows: 0 };
  const hdr = dump.headers;
  const idxArticle = ppPlanningFindColByHeaderVariants_(hdr, ['Артикул ВБ', 'Артикул WB']);
  const idxBarcode = ppPlanningFindColByHeaderVariants_(hdr, ['ШК', 'Barcode']);
  const idxQty = ppPlanningFindColByHeaderVariants_(hdr, ['Итоговое количество', 'Количество', 'Кол-во']);
  const idxReady = ppPlanningFindColByHeaderVariants_(hdr, ['Дата готовности']);
  const idxStatus = ppPlanningFindColByHeaderVariants_(hdr, ['Статус заказа', 'status_name', 'Статус']);
  const idxEta = ppPlanningFindColByHeaderVariants_(hdr, ['Плановая дата поступления', 'ETA', 'Дата поступления', 'Дата прибытия']);
  const idxPeriod = ppPlanningFindColByHeaderVariants_(hdr, ['Период (MM/YY)', 'Период']);
  if (idxArticle < 0 || idxQty < 0 || idxReady < 0 || idxStatus < 0) {
    return { byWbMonth: {}, byBcMonth: {}, errors: ['В «Сводная» не найдены обязательные колонки для in-transit.'], usedRows: 0 };
  }

  const monthSet = {};
  for (let i = 0; i < monthKeys.length; i++) monthSet[monthKeys[i]] = true;
  const byWbMonth = {};
  const byBcMonth = {};
  const shippedCode = ppCanonArticle_(ppGetProp_('PROCUREMENT_SHIPPED_STATUS_CODE', PP_SHIPPED_STATUS_CODE_DEFAULT));
  let usedRows = 0;

  for (let i = 0; i < dump.rows.length; i++) {
    const r = dump.rows[i];
    const d = dump.rowsDisplay[i];
    const artRaw = idxArticle < r.length ? d[idxArticle] : '';
    const bcRaw = idxBarcode >= 0 && idxBarcode < r.length ? d[idxBarcode] : '';
    const qty = ppToNumOr0_(idxQty < r.length ? r[idxQty] : '');
    if (!(qty > 0)) continue;
    const art = ppCanonArticle_(artRaw);
    const bc = ppCanonBarcodeForStock_(bcRaw);
    if (!(art && needWb[art]) && !(bc && needBc[bc])) continue;
    const st = idxStatus < r.length ? ppCanonArticle_(d[idxStatus]) : '';
    const isShipped = st === shippedCode;
    let arrival = null;
    if (isShipped && idxEta >= 0 && idxEta < r.length) {
      arrival = ppParseAnyDate_(r[idxEta]);
      if (!arrival) arrival = ppParseAnyDate_(d[idxEta]);
    }
    if (!arrival) {
      if (idxPeriod >= 0 && idxPeriod < r.length) {
        arrival = ppMonthStartFromPeriodToken_(d[idxPeriod]);
      }
      if (!arrival) arrival = ppParseAnyDate_(r[idxReady]);
      if (!arrival) arrival = ppParseAnyDate_(d[idxReady]);
      if (arrival) arrival = new Date(arrival.getFullYear(), arrival.getMonth(), 1);
    }
    if (!arrival) continue;
    const mk = ppMonthKeyFromDate_(arrival);
    if (!monthSet[mk]) continue;
    if (art && needWb[art]) {
      const k = art + '|' + mk;
      byWbMonth[k] = (byWbMonth[k] || 0) + qty;
    }
    if (bc && needBc[bc]) {
      const k2 = bc + '|' + mk;
      byBcMonth[k2] = (byBcMonth[k2] || 0) + qty;
    }
    usedRows++;
  }
  return { byWbMonth: byWbMonth, byBcMonth: byBcMonth, errors: [], usedRows: usedRows };
}

function ppLookupByRowKey_(wbCell, bcCell, byWb, byBc) {
  const art = ppCanonArticle_(wbCell);
  if (art && byWb[art] !== undefined) return byWb[art];
  const bc = ppCanonBarcodeForStock_(bcCell);
  if (bc && byBc[bc] !== undefined) return byBc[bc];
  return 0;
}

function ppLookupInboundByRowMonth_(wbCell, bcCell, monthKey, byWbMonth, byBcMonth) {
  const art = ppCanonArticle_(wbCell);
  if (art) {
    const k = art + '|' + monthKey;
    if (byWbMonth[k] !== undefined) return byWbMonth[k];
  }
  const bc = ppCanonBarcodeForStock_(bcCell);
  if (bc) {
    const k2 = bc + '|' + monthKey;
    if (byBcMonth[k2] !== undefined) return byBcMonth[k2];
  }
  return 0;
}

function ppLookupMsAccountingStockForRow_(wbCell, bcCell, stockByWb, stockByBarcode) {
  const wa = ppCanonArticle_(wbCell);
  if (wa && stockByWb[wa] !== undefined) return stockByWb[wa];
  const cb = ppCanonBarcodeForStock_(bcCell);
  if (cb && stockByBarcode[cb] !== undefined) return stockByBarcode[cb];
  return '';
}

/**
 * Заполняет на листе «Планирование закупок» колонку учётного остатка МС (сумма по складам с «Использовать»).
 * Сопоставление: сначала артикул ВБ, затем ШК.
 */
function updateProcurementPlanningMsAccountingStock() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (typeof getScriptPropOrThrow === 'function') getScriptPropOrThrow('MS_TOKEN');
    else if (!PropertiesService.getScriptProperties().getProperty('MS_TOKEN')) throw new Error('Не заполнен MS_TOKEN');
  } catch (e) {
    ui.alert('Настройка', 'Нужен MS_TOKEN в свойствах скрипта.', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storeIds = ppReadEnabledStoreIdsForStock_(ss);
  if (!storeIds.length) {
    ui.alert(
      'Склады',
      'Нет складов с включённым «Использовать» на листе «' +
        ppGetProp_('MS_STOCK_STORES_SHEET', PP_DEFAULT_MS_STOCK_STORES_SHEET) +
        '».',
      ui.ButtonSet.OK
    );
    return;
  }

  const planName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  const plan = ss.getSheetByName(planName);
  if (!plan) {
    ui.alert('Лист', 'Не найден лист «' + planName + '».', ui.ButtonSet.OK);
    return;
  }
  if (plan.getLastRow() < PP_PROC_PLAN_DATA_START_ROW) {
    ui.alert('Данные', 'На листе «' + planName + '» нет строк таблицы (ожидается шапка в строке ' + PP_PROC_PLAN_HEADER_ROW + ').', ui.ButtonSet.OK);
    return;
  }

  const lastCol = plan.getLastColumn();
  const hdr = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const cols = ppPlanningFindWbBarcodeCol0_(hdr);
  let msCol0 = ppPlanningFindMsStockCol0_(hdr);
  if (msCol0 < 0) {
    msCol0 = hdr.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, msCol0 + 1).setValue(PP_MS_STOCK_COL_HEADER).setFontWeight('bold');
  }

  const lastRow = plan.getLastRow();
  const numRows = lastRow - PP_PROC_PLAN_DATA_START_ROW + 1;
  const wbCol = cols.wb + 1;
  const bcCol = cols.bc + 1;
  const msCol = msCol0 + 1;
  const wbVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, wbCol, numRows, 1).getDisplayValues();
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, bcCol, numRows, 1).getDisplayValues();
  const need = ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals);
  const needWbCount = Object.keys(need.needWb).length;
  const needBcCount = Object.keys(need.needBc).length;
  if (needWbCount === 0 && needBcCount === 0) {
    ui.alert('Планирование', 'В колонках «Артикул ВБ» и «ШК» нет ни одного заполненного значения — нечего сопоставлять с МС.', ui.ButtonSet.OK);
    return;
  }

  let lookup;
  try {
    lookup = ppBuildMsAccountingStockLookup_(storeIds, need.needWb, need.needBc);
  } catch (e) {
    ui.alert('МойСклад', 'Ошибка при получении остатков:\n' + (e.message || String(e)), ui.ButtonSet.OK);
    return;
  }
  const out = [];
  let filled = 0;
  for (let i = 0; i < numRows; i++) {
    const q = ppLookupMsAccountingStockForRow_(wbVals[i][0], bcVals[i][0], lookup.stockByWb, lookup.stockByBarcode);
    if (q !== '') filled++;
    out.push([q === '' ? '' : q]);
  }
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, msCol, numRows, 1).setValues(out);
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, msCol, numRows, 1).setNumberFormat('0');

  ui.alert(
    'Готово',
    'Складов в расчёте: ' +
      storeIds.length +
      '\nУникальных артикулов ВБ в плане: ' +
      needWbCount +
      ', уникальных ШК в плане: ' +
      needBcCount +
      '\nСтрок таблицы: ' +
      numRows +
      '\nСтрок, где остаток найден и записан (в т.ч. 0): ' +
      filled +
      '\n\nУчитываются только позиции из планирования. Колонка: «' +
      PP_MS_STOCK_COL_HEADER +
      '».',
    ui.ButtonSet.OK
  );
  if (typeof logInfo === 'function') {
    logInfo('updateProcurementPlanningMsAccountingStock', {
      stores: storeIds.length,
      rows: numRows,
      filled: filled,
      needWb: needWbCount,
      needBc: needBcCount
    });
  }
}

/**
 * Заполняет колонку остатка Wildberries (отчёт warehouse_remains или supplier/stocks — см. wildberries_stocks.gs).
 * Сопоставление: сначала артикул ВБ (nmId / артикул продавца), затем ШК.
 */
function updateProcurementPlanningWbStock() {
  const ui = SpreadsheetApp.getUi();
  if (typeof wbBuildStockLookupForProcurementPlanning_ !== 'function') {
    ui.alert('Код', 'Не найден модуль wildberries_stocks.gs (функция wbBuildStockLookupForProcurementPlanning_).', ui.ButtonSet.OK);
    return;
  }
  if (!wbGetProp_('WB_API_TOKEN', '')) {
    ui.alert('Настройка', 'Нужен WB_API_TOKEN в свойствах скрипта (токен WB с доступом «Аналитика» и/или «Статистика»).', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  const plan = ss.getSheetByName(planName);
  if (!plan) {
    ui.alert('Лист', 'Не найден лист «' + planName + '».', ui.ButtonSet.OK);
    return;
  }
  if (plan.getLastRow() < PP_PROC_PLAN_DATA_START_ROW) {
    ui.alert('Данные', 'На листе «' + planName + '» нет строк таблицы (ожидается шапка в строке ' + PP_PROC_PLAN_HEADER_ROW + ').', ui.ButtonSet.OK);
    return;
  }

  const lastCol = plan.getLastColumn();
  const hdr = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const cols = ppPlanningFindWbBarcodeCol0_(hdr);
  let wbStockCol0 = ppPlanningFindWbStockCol0_(hdr);
  if (wbStockCol0 < 0) {
    wbStockCol0 = hdr.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, wbStockCol0 + 1).setValue(PP_WB_STOCK_COL_HEADER).setFontWeight('bold');
  }

  const lastRow = plan.getLastRow();
  const numRows = lastRow - PP_PROC_PLAN_DATA_START_ROW + 1;
  const wbCol = cols.wb + 1;
  const bcCol = cols.bc + 1;
  const outCol = wbStockCol0 + 1;
  const wbVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, wbCol, numRows, 1).getDisplayValues();
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, bcCol, numRows, 1).getDisplayValues();
  const need = ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals);
  const needWbCount = Object.keys(need.needWb).length;
  const needBcCount = Object.keys(need.needBc).length;
  if (needWbCount === 0 && needBcCount === 0) {
    ui.alert('Планирование', 'В колонках «Артикул ВБ» и «ШК» нет ни одного заполненного значения — нечего сопоставлять с WB.', ui.ButtonSet.OK);
    return;
  }

  let lookup;
  try {
    lookup = wbBuildStockLookupForProcurementPlanning_(need.needWb, need.needBc);
  } catch (e) {
    ui.alert('Wildberries', 'Ошибка при получении остатков:\n' + (e.message || String(e)), ui.ButtonSet.OK);
    return;
  }
  const out = [];
  let filled = 0;
  for (let i = 0; i < numRows; i++) {
    const q = ppLookupMsAccountingStockForRow_(wbVals[i][0], bcVals[i][0], lookup.stockByWb, lookup.stockByBarcode);
    if (q !== '') filled++;
    out.push([q === '' ? '' : q]);
  }
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, outCol, numRows, 1).setValues(out);
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, outCol, numRows, 1).setNumberFormat('0');

  ui.alert(
    'Готово',
    'Источник: ' +
      (lookup.source || 'WB') +
      '\nУникальных артикулов ВБ в плане: ' +
      needWbCount +
      ', уникальных ШК в плане: ' +
      needBcCount +
      '\nСтрок таблицы: ' +
      numRows +
      '\nСтрок, где остаток найден и записан (в т.ч. 0): ' +
      filled +
      '\n\nКолонка: «' +
      PP_WB_STOCK_COL_HEADER +
      '».',
    ui.ButtonSet.OK
  );
  if (typeof logInfo === 'function') {
    logInfo('updateProcurementPlanningWbStock', {
      source: lookup.source,
      rows: numRows,
      filled: filled,
      needWb: needWbCount,
      needBc: needBcCount
    });
  }
}

/**
 * Динамический расчёт потребности закупки на горизонте месячных колонок листа планирования.
 * Формула: stock_t = stock_{t-1} + inbound_t - salesAdjusted_t.
 */
function computeProcurementPurchasePlan() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  const plan = ss.getSheetByName(planName);
  if (!plan) return ui.alert('Лист', 'Не найден лист «' + planName + '».', ui.ButtonSet.OK);
  if (plan.getLastRow() < PP_PROC_PLAN_DATA_START_ROW) return ui.alert('Данные', 'На листе нет строк для расчёта.', ui.ButtonSet.OK);

  const lastCol = plan.getLastColumn();
  const headers = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const cols = ppPlanningFindWbBarcodeCol0_(headers);
  const msCol0 = ppPlanningFindMsStockCol0_(headers);
  const wbCol0 = ppPlanningFindWbStockCol0_(headers);
  if (msCol0 < 0 || wbCol0 < 0) {
    return ui.alert('Колонки', 'Сначала заполните «' + PP_MS_STOCK_COL_HEADER + '» и «' + PP_WB_STOCK_COL_HEADER + '».', ui.ButtonSet.OK);
  }
  const monthCols = ppPlanningFindMonthCols_(headers);
  if (!monthCols.length) return ui.alert('Горизонт', 'Не найдены колонки месяцев вида YYYY-MM.', ui.ButtonSet.OK);

  let inboundCol0 = ppPlanningFindColByHeaderVariants_(headers, [PP_PURCHASE_INBOUND_COL_HEADER]);
  if (inboundCol0 < 0) {
    inboundCol0 = headers.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, inboundCol0 + 1).setValue(PP_PURCHASE_INBOUND_COL_HEADER).setFontWeight('bold');
  }
  const headers2 = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, plan.getLastColumn()).getDisplayValues()[0];
  let adjCol0 = ppPlanningFindColByHeaderVariants_(headers2, [PP_PURCHASE_ADJ_CURR_COL_HEADER]);
  if (adjCol0 < 0) {
    adjCol0 = headers2.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, adjCol0 + 1).setValue(PP_PURCHASE_ADJ_CURR_COL_HEADER).setFontWeight('bold');
  }
  const headers3 = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, plan.getLastColumn()).getDisplayValues()[0];
  let deficitCol0 = ppPlanningFindColByHeaderVariants_(headers3, [PP_PURCHASE_DEFICIT_COL_HEADER]);
  if (deficitCol0 < 0) {
    deficitCol0 = headers3.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, deficitCol0 + 1).setValue(PP_PURCHASE_DEFICIT_COL_HEADER).setFontWeight('bold');
  }
  const headers4 = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, plan.getLastColumn()).getDisplayValues()[0];
  let recoCol0 = ppPlanningFindColByHeaderVariants_(headers4, [PP_PURCHASE_RECO_COL_HEADER]);
  if (recoCol0 < 0) {
    recoCol0 = headers4.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, recoCol0 + 1).setValue(PP_PURCHASE_RECO_COL_HEADER).setFontWeight('bold');
  }
  const headers5 = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, plan.getLastColumn()).getDisplayValues()[0];
  let prioCol0 = ppPlanningFindColByHeaderVariants_(headers5, [PP_PURCHASE_PRIORITY_COL_HEADER]);
  if (prioCol0 < 0) {
    prioCol0 = headers5.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, prioCol0 + 1).setValue(PP_PURCHASE_PRIORITY_COL_HEADER).setFontWeight('bold');
  }

  const lastRow = plan.getLastRow();
  const numRows = lastRow - PP_PROC_PLAN_DATA_START_ROW + 1;
  const wbVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, cols.wb + 1, numRows, 1).getDisplayValues();
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, cols.bc + 1, numRows, 1).getDisplayValues();
  const nameCol0 = ppPlanningFindColByHeaderVariants_(headers, ['Наименование', 'Артикул поставщика']);
  const nameVals =
    nameCol0 >= 0
      ? plan.getRange(PP_PROC_PLAN_DATA_START_ROW, nameCol0 + 1, numRows, 1).getDisplayValues()
      : [];
  const msVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, msCol0 + 1, numRows, 1).getValues();
  const wbStockVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, wbCol0 + 1, numRows, 1).getValues();
  const monthVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, monthCols[0].col0 + 1, numRows, monthCols.length).getValues();

  const need = ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals);
  const inbound = ppBuildInboundByMonthForPlanning_(ss, monthCols.map(function (m) { return m.key; }), need.needWb, need.needBc);
  let salesFact = { byWb: {}, byBarcode: {}, rows: 0 };
  let salesErr = '';
  try {
    if (typeof wbFetchCurrentMonthSalesForPlanning_ !== 'function') throw new Error('Не найден модуль wb_sales.gs');
    salesFact = wbFetchCurrentMonthSalesForPlanning_(need.needWb, need.needBc);
  } catch (e) {
    salesErr = e && e.message ? e.message : String(e);
  }

  const currentMonthDate = new Date();
  const currentMonthKey = currentMonthDate.getFullYear() + '-' + String(currentMonthDate.getMonth() + 1).padStart(2, '0');
  const currentMonthIndex = monthCols.findIndex(function (m) { return m.key === currentMonthKey; });
  const targetMonthKey = ppGetProp_('PROCUREMENT_TARGET_CONTROL_MONTH', monthCols[monthCols.length - 1].key);
  const targetMonthIndex = monthCols.findIndex(function (m) {
    return m.key === targetMonthKey;
  });
  const controlMonthIndex = targetMonthIndex >= 0 ? targetMonthIndex : monthCols.length - 1;
  const controlMonthKey = monthCols[controlMonthIndex].key;

  const outInbound = [];
  const outAdj = [];
  const outDeficit = [];
  const outReco = [];
  const outPrio = [];
  const detailHeaders = [
    'Артикул ВБ',
    'ШК',
    'Артикул поставщика (название)',
    'Остаток МС, шт',
    'Остаток ВБ, шт',
    'Стартовый остаток, шт',
    'Продажи WB тек. месяца, шт',
    PP_PURCHASE_ADJ_CURR_COL_HEADER,
    PP_PURCHASE_INBOUND_COL_HEADER,
    PP_PURCHASE_DEFICIT_COL_HEADER,
    PP_PURCHASE_RECO_COL_HEADER,
    PP_PURCHASE_PRIORITY_COL_HEADER,
    'Потребность на старт ' + controlMonthKey + ', шт',
    'Потребность на конец ' + controlMonthKey + ', шт',
    'Дельта срезов, шт'
  ];
  for (let m = 0; m < monthCols.length; m++) {
    detailHeaders.push('Старт ' + monthCols[m].key + ', шт');
    detailHeaders.push('План ' + monthCols[m].key + ', шт');
    detailHeaders.push('Приход ' + monthCols[m].key + ', шт');
    detailHeaders.push('Конец ' + monthCols[m].key + ', шт');
  }
  const detailRows = [detailHeaders];
  let withDeficit = 0;
  let positiveNeedStart = 0;
  let positiveNeedEnd = 0;
  let splitDiffRows = 0;
  for (let r = 0; r < numRows; r++) {
    const wbCell = wbVals[r][0];
    const bcCell = bcVals[r][0];
    const supplierName = nameVals.length ? nameVals[r][0] : '';
    const stockStart = ppToNumOr0_(msVals[r][0]) + ppToNumOr0_(wbStockVals[r][0]);
    const soldCurrent = ppLookupByRowKey_(wbCell, bcCell, salesFact.byWb, salesFact.byBarcode);
    let adjustedCurrent = '';
    let runStock = stockStart;
    let minStock = stockStart;
    let sumInbound = 0;
    let firstDeficitMonth = '';
    const monthlyTrace = [];
    let startTarget = null;
    let endTarget = null;
    let planTargetAdjusted = 0;
    for (let m = 0; m < monthCols.length; m++) {
      const mk = monthCols[m].key;
      const inboundQty = ppLookupInboundByRowMonth_(wbCell, bcCell, mk, inbound.byWbMonth, inbound.byBcMonth);
      sumInbound += inboundQty;
      // Отрицательный остаток не переносим: старт следующего месяца не ниже нуля.
      const startMonth = Math.max(runStock, 0);
      runStock = startMonth;
      let planQty = ppToNumOr0_(monthVals[r][m]);
      if (m === currentMonthIndex) {
        planQty = Math.max(0, planQty - soldCurrent);
        adjustedCurrent = planQty;
      }
      runStock = runStock + inboundQty - planQty;
      monthlyTrace.push({ start: startMonth, plan: planQty, inbound: inboundQty, end: runStock, key: mk });
      if (m === controlMonthIndex) {
        startTarget = startMonth;
        endTarget = runStock;
        planTargetAdjusted = planQty;
      }
      if (runStock < minStock) minStock = runStock;
      if (!firstDeficitMonth && runStock < 0) firstDeficitMonth = mk;
    }
    const deficit = minStock < 0 ? Math.ceil(Math.abs(minStock)) : 0;
    const needAtStartTarget =
      startTarget == null ? 0 : Math.max(0, Math.ceil(planTargetAdjusted - startTarget)); // сравнение старта месяца с его планом
    const needAtEndTarget = endTarget == null ? 0 : Math.max(0, Math.ceil(-endTarget)); // контрольный срез по концу месяца
    const reco = needAtStartTarget; // финальный единый срез выберем после тестов
    let priority = 'Нет дефицита';
    if (firstDeficitMonth) {
      const dFirst = ppMonthFirstDateFromKey_(firstDeficitMonth);
      const dNow = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth(), 1);
      const diff = dFirst ? ppMonthDiff_(dFirst, dNow) : 99;
      if (diff <= 0) priority = 'Критично';
      else if (diff === 1) priority = 'Высокий';
      else if (diff <= 2) priority = 'Средний';
      else priority = 'Низкий';
      withDeficit++;
    }
    if (needAtStartTarget > 0) positiveNeedStart++;
    if (needAtEndTarget > 0) positiveNeedEnd++;
    if (needAtStartTarget !== needAtEndTarget) splitDiffRows++;
    outInbound.push([sumInbound || '']);
    outAdj.push([adjustedCurrent === '' ? '' : adjustedCurrent]);
    outDeficit.push([deficit || '']);
    outReco.push([reco || '']);
    outPrio.push([priority]);
    const traceCols = [];
    for (let t = 0; t < monthlyTrace.length; t++) {
      traceCols.push(monthlyTrace[t].start, monthlyTrace[t].plan, monthlyTrace[t].inbound, monthlyTrace[t].end);
    }
    detailRows.push(
      [
        wbCell,
        bcCell,
        supplierName,
        ppToNumOr0_(msVals[r][0]),
        ppToNumOr0_(wbStockVals[r][0]),
        stockStart,
        soldCurrent,
        adjustedCurrent === '' ? 0 : adjustedCurrent,
        sumInbound,
        deficit,
        reco,
        priority,
        needAtStartTarget,
        needAtEndTarget,
        needAtStartTarget - needAtEndTarget
      ].concat(traceCols)
    );
  }

  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, inboundCol0 + 1, numRows, 1).setValues(outInbound).setNumberFormat('0');
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, adjCol0 + 1, numRows, 1).setValues(outAdj).setNumberFormat('0');
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, deficitCol0 + 1, numRows, 1).setValues(outDeficit).setNumberFormat('0');
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, recoCol0 + 1, numRows, 1).setValues(outReco).setNumberFormat('0');
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, prioCol0 + 1, numRows, 1).setValues(outPrio);

  const reportSheetName = ppGetProp_('PROCUREMENT_PURCHASE_REPORT_SHEET', PP_PURCHASE_REPORT_SHEET_DEFAULT);
  let report = ss.getSheetByName(reportSheetName);
  if (!report) report = ss.insertSheet(reportSheetName);
  report.clearContents();
  report.getRange(1, 1).setValue('Обновлено');
  report.getRange(1, 2).setValue(new Date());
  report.getRange(2, 1).setValue('Источник');
  report.getRange(2, 2).setValue('Расчёт потребности закупки из листа «' + planName + '»');
  report.getRange(1, 4).setValue('Контрольный месяц');
  report.getRange(1, 5).setValue(controlMonthKey);
  report.getRange(2, 4).setValue('SKU с потребностью на старт');
  report.getRange(2, 5).setValue(positiveNeedStart);
  report.getRange(3, 4).setValue('SKU с потребностью на конец');
  report.getRange(3, 5).setValue(positiveNeedEnd);
  report.getRange(4, 4).setValue('SKU с расхождением срезов');
  report.getRange(4, 5).setValue(splitDiffRows);
  const reportStartRow = 6;
  report.getRange(reportStartRow, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
  report.getRange(reportStartRow, 1, 1, detailRows[0].length).setFontWeight('bold');
  if (detailRows.length > 1) {
    report.getRange(reportStartRow + 1, 3, detailRows.length - 1, detailRows[0].length - 2).setNumberFormat('0');
  }
  report.setFrozenRows(reportStartRow);

  const warn = [];
  if (inbound.errors.length) warn.push('В пути: ' + inbound.errors.join('; '));
  if (salesErr) warn.push('WB продажи: ' + salesErr);
  ui.alert(
    'Потребность закупки',
    'Строк: ' +
      numRows +
      '\nГоризонт месяцев: ' +
      monthCols.length +
      '\nСтрок с дефицитом: ' +
      withDeficit +
      '\nПотребность на старт ' +
      controlMonthKey +
      ': ' +
      positiveNeedStart +
      ' SKU' +
      '\nПотребность на конец ' +
      controlMonthKey +
      ': ' +
      positiveNeedEnd +
      ' SKU' +
      '\nРасхождение двух срезов: ' +
      splitDiffRows +
      ' SKU' +
      '\nВ пути учтено строк: ' +
      inbound.usedRows +
      '\nWB продаж обработано строк: ' +
      salesFact.rows +
      '\nЛист детализации: «' +
      reportSheetName +
      '»' +
      (warn.length ? '\n\nПредупреждения:\n' + warn.join('\n') : ''),
    ui.ButtonSet.OK
  );
  if (typeof logInfo === 'function') {
    logInfo('computeProcurementPurchasePlan', {
      rows: numRows,
      months: monthCols.length,
      withDeficit: withDeficit,
      controlMonth: controlMonthKey,
      positiveNeedStart: positiveNeedStart,
      positiveNeedEnd: positiveNeedEnd,
      splitDiffRows: splitDiffRows,
      inboundRows: inbound.usedRows,
      wbSalesRows: salesFact.rows,
      warnings: warn
    });
  }
}

/**
 * Проверяет сопоставление позиций плана с остатками МС/WB по артикулу ВБ и ШК.
 */
function checkProcurementPlanningStocksCoverage() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  const plan = ss.getSheetByName(planName);
  if (!plan) {
    ui.alert('Лист', 'Не найден лист «' + planName + '».', ui.ButtonSet.OK);
    return;
  }
  if (plan.getLastRow() < PP_PROC_PLAN_DATA_START_ROW) {
    ui.alert('Данные', 'На листе «' + planName + '» нет строк таблицы.', ui.ButtonSet.OK);
    return;
  }

  const lastCol = plan.getLastColumn();
  const hdr = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const cols = ppPlanningFindWbBarcodeCol0_(hdr);
  let checkCol0 = ppPlanningFindStockCheckCol0_(hdr);
  if (checkCol0 < 0) {
    checkCol0 = hdr.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, checkCol0 + 1).setValue(PP_STOCK_CHECK_COL_HEADER).setFontWeight('bold');
  }

  const lastRow = plan.getLastRow();
  const numRows = lastRow - PP_PROC_PLAN_DATA_START_ROW + 1;
  const wbVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, cols.wb + 1, numRows, 1).getDisplayValues();
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, cols.bc + 1, numRows, 1).getDisplayValues();
  const need = ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals);
  const needWbCount = Object.keys(need.needWb).length;
  const needBcCount = Object.keys(need.needBc).length;
  if (needWbCount === 0 && needBcCount === 0) {
    ui.alert('Планирование', 'В колонках «Артикул ВБ» и «ШК» нет заполненных значений — нечего проверять.', ui.ButtonSet.OK);
    return;
  }

  let msLookup = { stockByWb: {}, stockByBarcode: {} };
  let wbLookup = { stockByWb: {}, stockByBarcode: {} };
  let inboundLookup = { byWbMonth: {}, byBcMonth: {}, errors: [] };
  let salesLookup = { byWb: {}, byBarcode: {}, rows: 0 };
  let msReady = false;
  let wbReady = false;
  let inboundReady = false;
  let salesReady = false;
  let msErr = '';
  let wbErr = '';
  let inboundErr = '';
  let salesErr = '';

  try {
    if (typeof getScriptPropOrThrow === 'function') getScriptPropOrThrow('MS_TOKEN');
    else if (!PropertiesService.getScriptProperties().getProperty('MS_TOKEN')) throw new Error('Не заполнен MS_TOKEN');
    const storeIds = ppReadEnabledStoreIdsForStock_(ss);
    if (!storeIds.length) throw new Error('Нет включённых складов «Использовать» на листе складов МС');
    msLookup = ppBuildMsAccountingStockLookup_(storeIds, need.needWb, need.needBc);
    msReady = true;
  } catch (e) {
    msErr = e && e.message ? e.message : String(e);
  }

  try {
    if (typeof wbBuildStockLookupForProcurementPlanning_ !== 'function') throw new Error('Не найден модуль wildberries_stocks.gs');
    if (!wbGetProp_('WB_API_TOKEN', '')) throw new Error('Не заполнен WB_API_TOKEN');
    wbLookup = wbBuildStockLookupForProcurementPlanning_(need.needWb, need.needBc);
    wbReady = true;
  } catch (e) {
    wbErr = e && e.message ? e.message : String(e);
  }

  try {
    const monthCols = ppPlanningFindMonthCols_(hdr);
    inboundLookup = ppBuildInboundByMonthForPlanning_(
      ss,
      monthCols.map(function (m) {
        return m.key;
      }),
      need.needWb,
      need.needBc
    );
    inboundReady = inboundLookup.errors.length === 0;
    if (!inboundReady) inboundErr = inboundLookup.errors.join('; ');
  } catch (e) {
    inboundErr = e && e.message ? e.message : String(e);
  }

  try {
    if (typeof wbFetchCurrentMonthSalesForPlanning_ !== 'function') throw new Error('Не найден модуль wb_sales.gs');
    salesLookup = wbFetchCurrentMonthSalesForPlanning_(need.needWb, need.needBc);
    salesReady = true;
  } catch (e) {
    salesErr = e && e.message ? e.message : String(e);
  }

  if (!msReady && !wbReady) {
    ui.alert('Проверка', 'Не удалось получить lookup ни из МС, ни из WB.\nМС: ' + msErr + '\nWB: ' + wbErr, ui.ButtonSet.OK);
    return;
  }

  const outCheck = [];
  let okBoth = 0;
  let onlyMs = 0;
  let onlyWb = 0;
  let none = 0;
  let noneByWb = 0;
  let noneByBc = 0;
  let noneByBothKeys = 0;
  let emptyKeys = 0;
  let noInboundMatch = 0;
  let noSalesMatch = 0;

  for (let i = 0; i < numRows; i++) {
    const art = ppCanonArticle_(wbVals[i][0]);
    const bc = ppCanonBarcodeForStock_(bcVals[i][0]);
    const msHit = (art && msLookup.stockByWb[art] !== undefined) || (bc && msLookup.stockByBarcode[bc] !== undefined);
    const wbHit = (art && wbLookup.stockByWb[art] !== undefined) || (bc && wbLookup.stockByBarcode[bc] !== undefined);
    const currentMonthKey = new Date().getFullYear() + '-' + String(new Date().getMonth() + 1).padStart(2, '0');
    const inboundHit = ppLookupInboundByRowMonth_(wbVals[i][0], bcVals[i][0], currentMonthKey, inboundLookup.byWbMonth, inboundLookup.byBcMonth) > 0;
    const salesHit = ppLookupByRowKey_(wbVals[i][0], bcVals[i][0], salesLookup.byWb, salesLookup.byBarcode) > 0;
    if (!inboundHit && (art || bc)) noInboundMatch++;
    if (!salesHit && (art || bc)) noSalesMatch++;

    if (msHit && wbHit) {
      okBoth++;
      outCheck.push(['OK: МС+WB']);
    } else if (msHit) {
      onlyMs++;
      outCheck.push(['Только МС']);
    } else if (wbHit) {
      onlyWb++;
      outCheck.push(['Только ВБ']);
    } else {
      none++;
      if (!art && !bc) {
        emptyKeys++;
        outCheck.push(['Ошибка: пустые Артикул ВБ и ШК']);
      } else if (art && bc) {
        noneByBothKeys++;
        outCheck.push(['Не найдено по артикулу и ШК']);
      } else if (art) {
        noneByWb++;
        outCheck.push(['Не найдено по артикулу ВБ']);
      } else {
        noneByBc++;
        outCheck.push(['Не найдено по ШК']);
      }
    }
  }

  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, checkCol0 + 1, numRows, 1).setValues(outCheck);

  const sourceLine =
    'Проверка по источникам: ' +
    (msReady ? 'МС OK' : 'МС недоступен') +
    ', ' +
    (wbReady ? 'WB OK' : 'WB недоступен') +
    ', ' +
    (inboundReady ? 'В пути OK' : 'В пути недоступен') +
    ', ' +
    (salesReady ? 'WB Sales OK' : 'WB Sales недоступен');
  ui.alert(
    'Проверка остатков',
    sourceLine +
      '\nСтрок таблицы: ' +
      numRows +
      '\nOK (МС+WB): ' +
      okBoth +
      '\nТолько МС: ' +
      onlyMs +
      '\nТолько ВБ: ' +
      onlyWb +
      '\nНе найдено в обоих источниках: ' +
      none +
      '\n  └ пустые ключи: ' +
      emptyKeys +
      '\n  └ нет по артикулу ВБ: ' +
      noneByWb +
      '\n  └ нет по ШК: ' +
      noneByBc +
      '\n  └ нет по артикулу и ШК: ' +
      noneByBothKeys +
      '\nНет совпадения в «в пути» (текущий месяц): ' +
      noInboundMatch +
      '\nНет факта продаж WB текущего месяца: ' +
      noSalesMatch +
      (inboundErr ? '\nВ пути ошибка: ' + inboundErr : '') +
      (salesErr ? '\nWB Sales ошибка: ' + salesErr : '') +
      '\n\nЗаписаны колонки: «' +
      PP_STOCK_CHECK_COL_HEADER +
      '».',
    ui.ButtonSet.OK
  );
  if (typeof logInfo === 'function') {
    logInfo('checkProcurementPlanningStocksCoverage', {
      rows: numRows,
      msReady: msReady,
      wbReady: wbReady,
      okBoth: okBoth,
      onlyMs: onlyMs,
      onlyWb: onlyWb,
      none: none,
      emptyKeys: emptyKeys,
      noneByWb: noneByWb,
      noneByBc: noneByBc,
      noneByBothKeys: noneByBothKeys,
      inboundReady: inboundReady,
      salesReady: salesReady,
      noInboundMatch: noInboundMatch,
      noSalesMatch: noSalesMatch
    });
  }
}

/** Лист перечня складов: A=использовать, B=id, C=название, D=код, E=есть в МС, F=обновлено. */
const PP_MS_STORE_COL = { USE: 0, ID: 1, NAME: 2, CODE: 3, IN_MS: 4, SYNCED: 5 };
const PP_MS_STORE_HEADER = ['Использовать', 'ID склада', 'Название', 'Код', 'Есть в МС', 'Обновлено из МС'];

function ensureMsStockStoresSheet_(ss) {
  const name = ppGetProp_('MS_STOCK_STORES_SHEET', PP_DEFAULT_MS_STOCK_STORES_SHEET);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (String(sh.getRange(1, 2).getValue() || '').trim() !== 'ID склада') {
    sh.getRange(1, 1, 1, PP_MS_STORE_HEADER.length).setValues([PP_MS_STORE_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const lrFmt = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 2, lrFmt - 1, 1).setNumberFormat('@');
  return sh;
}

function fetchAllStoresFromMs_() {
  const includeArchived = ppGetProp_('MS_STORE_SYNC_INCLUDE_ARCHIVED', '0') === '1';
  const list = [];
  let offset = 0;
  const limit = 1000;
  while (true) {
    let path = '/entity/store?limit=' + limit + '&offset=' + offset;
    if (!includeArchived) {
      path += '&filter=' + encodeURIComponent('archived=false');
    }
    const res = msFetch(path, 'get');
    if (!res.success || !res.data || !res.data.rows) {
      if (offset === 0 && res.error) {
        throw new Error(res.error);
      }
      break;
    }
    const chunk = res.data.rows;
    for (let i = 0; i < chunk.length; i++) list.push(chunk[i]);
    if (chunk.length < limit) break;
    offset += limit;
    Utilities.sleep(80);
  }
  return list;
}

/**
 * Подтягивает склады из МС на лист «Склады МС (остатки)»: новые с «Использовать» = нет;
 * у существующих строк колонка «Использовать» не меняется.
 */
function syncMsStockStoresSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (typeof getScriptPropOrThrow === 'function') getScriptPropOrThrow('MS_TOKEN');
    else if (!PropertiesService.getScriptProperties().getProperty('MS_TOKEN')) {
      throw new Error('Не заполнен MS_TOKEN');
    }
  } catch (e) {
    ui.alert('Настройка', 'Нужен MS_TOKEN в свойствах скрипта.', ui.ButtonSet.OK);
    return;
  }

  let stores;
  try {
    stores = fetchAllStoresFromMs_();
  } catch (e) {
    ui.alert('МойСклад', 'Не удалось получить склады:\n' + (e.message || String(e)), ui.ButtonSet.OK);
    return;
  }

  if (!stores.length) {
    ui.alert('МойСклад', 'Список складов пуст (проверьте фильтр архивных и права токена).', ui.ButtonSet.OK);
    return;
  }

  const apiMap = {};
  for (let i = 0; i < stores.length; i++) {
    const st = stores[i];
    if (!st || !st.id) continue;
    apiMap[st.id] = {
      name: st.name != null ? String(st.name) : '',
      code: st.code != null ? String(st.code) : '',
      archived: !!st.archived
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureMsStockStoresSheet_(ss);
  const now = new Date();
  const lastRow = sh.getLastRow();
  const existingById = {};

  if (lastRow >= 2) {
    const n = lastRow - 1;
    const block = sh.getRange(2, 1, n, PP_MS_STORE_HEADER.length).getValues();
    for (let i = 0; i < block.length; i++) {
      const id = String(block[i][PP_MS_STORE_COL.ID] || '').trim();
      if (!id) continue;
      existingById[id] = { row: i + 2 };
    }
  }

  let updated = 0;
  let markedMissing = 0;
  for (const id in existingById) {
    const row = existingById[id].row;
    const inf = apiMap[id];
    if (inf) {
      const dispName = inf.archived ? '(архив) ' + inf.name : inf.name;
      sh.getRange(row, PP_MS_STORE_COL.NAME + 1, 1, 2).setValues([[dispName, inf.code]]);
      sh.getRange(row, PP_MS_STORE_COL.IN_MS + 1).setValue('Да');
      sh.getRange(row, PP_MS_STORE_COL.SYNCED + 1).setValue(now);
      updated++;
    } else {
      sh.getRange(row, PP_MS_STORE_COL.IN_MS + 1).setValue('Нет');
      sh.getRange(row, PP_MS_STORE_COL.SYNCED + 1).setValue(now);
      markedMissing++;
    }
  }

  const newRows = [];
  for (const id in apiMap) {
    if (existingById[id]) continue;
    const inf = apiMap[id];
    const dispName = inf.archived ? '(архив) ' + inf.name : inf.name;
    newRows.push([false, id, dispName, inf.code, 'Да', now]);
  }

  if (newRows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, newRows.length, PP_MS_STORE_HEADER.length).setValues(newRows);
    sh.getRange(start, 2, newRows.length, 1).setNumberFormat('@');
  }

  const lr = sh.getLastRow();
  if (lr >= 2) {
    const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sh.getRange(2, PP_MS_STORE_COL.USE + 1, lr - 2 + 1, 1).setDataValidation(cb);
  }

  const msg =
    'Складов в МС: ' +
    Object.keys(apiMap).length +
    '\nОбновлено строк: ' +
    updated +
    '\nНе найдено в последней выгрузке МС: ' +
    markedMissing +
    '\nДобавлено новых строк: ' +
    newRows.length +
    '\n\nОтметьте «Использовать» у складов, которые участвуют в сумме учётного остатка (остатки на планировании закупок).';
  ui.alert('Склады МС (остатки)', msg, ui.ButtonSet.OK);
  if (typeof logInfo === 'function') {
    logInfo('syncMsStockStoresSheet', { api: Object.keys(apiMap).length, newRows: newRows.length });
  }
}
