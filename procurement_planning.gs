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
const PP_OZON_FBO_STOCK_COL_HEADER = 'Остаток Ozon FBO, шт';
const PP_OZON_FBS_STOCK_COL_HEADER = 'Остаток Ozon FBS, шт';
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
    const n = ppParseQtyForPlan_(rawPlan);
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
    const n = ppParseQtyForPlan_(rawPlan);
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
  // Согласовано с syncManagerCanonHeader_ (main.gs): обрабатывает №, #, точки, скобки,
  // слэши, кавычки, дефисы и подчёркивания. Без этого «Артикул_ВБ», «№ ВБ», «Шт./коробка»
  // не сматчатся с «Артикул ВБ», «WB», «Шт в коробке».
  // Дальше схлопываем пробелы И УБИРАЕМ их совсем — оставшиеся includes/indexOf-проверки
  // в этом модуле ожидают «слитный» канон (например, `c.indexOf('заказали')`).
  return String(s == null ? '' : s)
    .toLowerCase()
    .replace(/\u00A0/g, ' ')
    .replace(/ё/g, 'е')
    .replace(/[№#]/g, ' ')
    .replace(/[.,:;()/\\\[\]{}'"“”«»]/g, ' ')
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, '');
}

function ppFindWbColInRefHeaders_(headers) {
  for (let i = 0; i < headers.length; i++) {
    if (ppIsWbArticleHeader_(headers[i])) return i;
  }
  // Раньше тут был фолбэк `4` (колонка E) — он опасен: если у листа другая структура,
  // данные читались по неверному столбцу. Безопаснее вернуть -1 и пусть вызывающий код
  // решает, как реагировать (обычно — сообщение пользователю про переименование шапки).
  return -1;
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
  const n = ppParseQtyForPlan_(v);
  return n > 0 ? n : 0;
}

/**
 * Локальный парсер количеств для планов: не зависит от глобального parseNumber.
 * Принимает числа, строки с пробелами/неразрывными пробелами и запятой как
 * десятичным разделителем. Возвращает 0 при пустом/нечисловом значении.
 *
 * Зачем: в книге 03 может отсутствовать helpers.gs — тогда parseNumber undefined,
 * и старая логика молча возвращала 0 для всех планов (см. диагностику от 2026-05-13).
 */
function ppParseQtyForPlan_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v).replace(/\u00a0/g, '').replace(/\s/g, '').replace(',', '.');
  if (!s) return 0;
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
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
    if (wbCol < 0) {
      // Раньше тут был тихий фолбэк на колонку E, что приводило к «магическим»
      // ошибкам — справочник читался по неверной колонке. Лучше явно сказать пользователю.
      warnings.push(
        'Справочник товаров: не нашёл колонку «Артикул ВБ». ' +
        'Шапка строки ' + headerRow1Based + ': ' +
        headers.map(function (h) { return String(h == null ? '' : h).trim(); }).filter(Boolean).slice(0, 20).join(' | ')
      );
      return empty;
    }
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

/**
 * Конвертирует 0-based индекс колонки в A1-нотацию (0→A, 25→Z, 26→AA, …).
 */
function ppColToA1_(col0) {
  let n = col0 + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Диагностический пункт меню: открывает книги-источники планов и для каждого месяца
 * каждой вкладки показывает, какие колонки реально нашёл скрипт (anchor / WB / plan)
 * и сколько строк дают по этой паре количество > 0.
 *
 * Результат пишется на лист «Планирование закупок (диагностика)» в активной книге.
 * Сам свод плана не пересобирается — это чисто read-only отчёт.
 */
function diagnoseProcurementPlanningSourceSheets() {
  const ui = SpreadsheetApp.getUi();
  const sourceId = ppGetProp_('SALES_PLANS_SPREADSHEET_ID', PP_DEFAULT_SOURCE_SPREADSHEET_ID);
  if (!sourceId) {
    ui.alert('Задайте SALES_PLANS_SPREADSHEET_ID в свойствах скрипта.');
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

  const out = [
    [
      'Книга',
      'Вкладка',
      'Месяц',
      'Anchor (дата 1-го)',
      'WB-заголовок',
      'WB-ячейка',
      'WB-текст',
      'Plan-заголовок',
      'Plan-ячейка',
      'Plan-текст',
      'Строк под шапкой',
      'Из них с qty>0',
      'Пример строки',
      'Ошибка'
    ]
  ];

  /**
   * Накатывает на одну вкладку анализ ровно так же, как продукционный код,
   * и собирает строку отчёта для каждого месяца.
   */
  function inspectStandardSheet_(bookLabel, sh) {
    const snap = ppReadSheetSnapshot_(sh);
    if (!snap.values.length) {
      out.push([bookLabel, sh.getName(), '—', '—', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Пустой лист']);
      return;
    }
    const topRows = Math.min(PP_SCAN_TOP_ROWS, snap.values.length);
    const topMatrix = [];
    for (let r = 0; r < topRows; r++) topMatrix.push(snap.values[r]);
    const headerEnd = Math.min(snap.values.length, PP_HEADER_SCAN_ROWS);
    const headerMatrix = [];
    for (let r = 0; r < headerEnd; r++) headerMatrix.push(snap.values[r]);
    for (let mi = 0; mi < months.length; mi++) {
      const m = months[mi];
      const mk = monthKeys[mi];
      const anchorCol = ppFindMonthAnchorCol_(topMatrix, year, m);
      if (anchorCol < 0) {
        out.push([
          bookLabel, sh.getName(), mk, 'не найдена',
          '—', '—', '—', '—', '—', '—', 0, 0, '',
          'Дата 1-го числа не найдена в первых ' + PP_SCAN_TOP_ROWS + ' строках'
        ]);
        continue;
      }
      // Найдём строку, в которой стоит anchor — это нужно, чтобы показать A1-нотацию.
      let anchorRow0 = -1;
      for (let r = 0; r < topMatrix.length; r++) {
        const parsed = ppParseMonthCell_(topMatrix[r][anchorCol]);
        if (ppIsFirstOfMonth_(parsed) && parsed.y === year && parsed.m === m) {
          anchorRow0 = r;
          break;
        }
      }
      const anchorA1 = ppColToA1_(anchorCol) + (anchorRow0 + 1);
      const found = ppFindWbAndPlanCols_(headerMatrix, anchorCol);
      const wbText = found.wbCol >= 0 && found.headerRow >= 0 && found.wbCol < (headerMatrix[found.headerRow] || []).length
        ? String(headerMatrix[found.headerRow][found.wbCol] || '').trim()
        : '';
      const planText = found.planCol >= 0 && found.headerRow >= 0 && found.planCol < (headerMatrix[found.headerRow] || []).length
        ? String(headerMatrix[found.headerRow][found.planCol] || '').trim()
        : '';
      const wbA1 = found.wbCol >= 0 ? ppColToA1_(found.wbCol) + (found.headerRow + 1) : '';
      const planA1 = found.planCol >= 0 ? ppColToA1_(found.planCol) + (found.headerRow + 1) : '';
      let underHeader = 0;
      let withQty = 0;
      let sample = '';
      let firstAnyRow = '';
      if (found.wbCol >= 0 && found.planCol >= 0) {
        const dataStart = found.headerRow + 2;
        for (let r = dataStart - 1; r < snap.values.length; r++) {
          const row = snap.values[r];
          const rawWb = found.wbCol < row.length ? row[found.wbCol] : '';
          if (!ppNormArticle_(rawWb)) continue;
          underHeader++;
          const rawPlan = found.planCol < row.length ? row[found.planCol] : '';
          const n = ppParseQtyForPlan_(rawPlan);
          if (!firstAnyRow) {
            firstAnyRow = String(rawWb).trim() + ' | raw="' + String(rawPlan) + '" | parsed=' + n;
          }
          if (n > 0) {
            withQty++;
            if (!sample) sample = String(rawWb).trim() + ' → ' + String(rawPlan).trim() + ' (=' + n + ')';
          }
        }
      }
      out.push([
        bookLabel, sh.getName(), mk, anchorA1,
        wbText || (found.wbCol >= 0 ? '(пусто)' : 'не найдено'),
        wbA1, wbText,
        planText || (found.planCol >= 0 ? '(пусто)' : 'не найдено'),
        planA1, planText,
        underHeader, withQty, sample || firstAnyRow, ''
      ]);
    }
  }

  function inspectExtraSheet_(bookLabel, sh, supplierToWb) {
    const snap = ppReadSheetSnapshot_(sh);
    if (!snap.values.length) {
      out.push([bookLabel, sh.getName(), '—', '—', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Пустой лист']);
      return;
    }
    const topRows = Math.min(PP_SCAN_TOP_ROWS, snap.values.length);
    const topMatrix = [];
    for (let r = 0; r < topRows; r++) topMatrix.push(snap.values[r]);
    const headerEnd = Math.min(snap.values.length, PP_HEADER_SCAN_ROWS);
    const headerMatrix = [];
    for (let r = 0; r < headerEnd; r++) headerMatrix.push(snap.values[r]);
    for (let mi = 0; mi < months.length; mi++) {
      const m = months[mi];
      const mk = monthKeys[mi];
      const anchorCol = ppFindMonthAnchorCol_(topMatrix, year, m);
      if (anchorCol < 0) {
        out.push([bookLabel, sh.getName(), mk, 'не найдена', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Дата 1-го числа не найдена']);
        continue;
      }
      let anchorRow0 = -1;
      for (let r = 0; r < topMatrix.length; r++) {
        const parsed = ppParseMonthCell_(topMatrix[r][anchorCol]);
        if (ppIsFirstOfMonth_(parsed) && parsed.y === year && parsed.m === m) { anchorRow0 = r; break; }
      }
      const anchorA1 = ppColToA1_(anchorCol) + (anchorRow0 + 1);
      const found = ppFindSupplierSkuAndOrderedCols_(headerMatrix, anchorCol);
      const skuText = found.skuCol >= 0 && found.headerRow >= 0 && found.skuCol < (headerMatrix[found.headerRow] || []).length
        ? String(headerMatrix[found.headerRow][found.skuCol] || '').trim()
        : '';
      const planText = found.planCol >= 0 && found.headerRow >= 0 && found.planCol < (headerMatrix[found.headerRow] || []).length
        ? String(headerMatrix[found.headerRow][found.planCol] || '').trim()
        : '';
      const skuA1 = found.skuCol >= 0 ? ppColToA1_(found.skuCol) + (found.headerRow + 1) : '';
      const planA1 = found.planCol >= 0 ? ppColToA1_(found.planCol) + (found.headerRow + 1) : '';
      let underHeader = 0;
      let withQty = 0;
      let mapped = 0;
      let sample = '';
      let firstAnyRow = '';
      if (found.skuCol >= 0 && found.planCol >= 0) {
        const dataStart = found.headerRow + 2;
        for (let r = dataStart - 1; r < snap.values.length; r++) {
          const row = snap.values[r];
          const rawSku = found.skuCol < row.length ? row[found.skuCol] : '';
          if (!ppNormArticle_(rawSku)) continue;
          underHeader++;
          const rawPlan = found.planCol < row.length ? row[found.planCol] : '';
          const n = ppParseQtyForPlan_(rawPlan);
          if (!firstAnyRow) {
            firstAnyRow = String(rawSku).trim() + ' | raw="' + String(rawPlan) + '" | parsed=' + n;
          }
          if (n > 0) {
            withQty++;
            if (!sample) sample = String(rawSku).trim() + ' → ' + String(rawPlan).trim() + ' (=' + n + ')';
          }
          const sk = ppCanonArticle_(String(rawSku));
          if (sk && supplierToWb[sk]) mapped++;
        }
      }
      out.push([
        bookLabel, sh.getName(), mk, anchorA1,
        skuText || (found.skuCol >= 0 ? '(пусто)' : 'не найдено'),
        skuA1, skuText,
        planText || (found.planCol >= 0 ? '(пусто)' : 'не найдено'),
        planA1, planText,
        underHeader, withQty, sample || firstAnyRow,
        underHeader > 0 ? ('сопоставлено со справочником: ' + mapped + ' из ' + underHeader) : ''
      ]);
    }
  }

  // Основная книга планов
  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceId);
  } catch (e) {
    ui.alert('Не открылась основная книга планов: ' + (e.message || String(e)));
    return;
  }
  const bookLabelMain = sourceSs.getName();
  for (let si = 0; si < sheetNames.length; si++) {
    const sh = sourceSs.getSheetByName(sheetNames[si]);
    if (!sh) {
      out.push([bookLabelMain, sheetNames[si], '—', '—', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Вкладка не найдена']);
      continue;
    }
    inspectStandardSheet_(bookLabelMain, sh);
  }

  // Доп. книга планов (Ozon-style: только «Артикул поставщика»)
  const warningsForRef = [];
  const refBundle = ppLoadProductReferenceBundle_(warningsForRef);
  const extraId = ppGetExtraSalesPlansSpreadsheetId_();
  if (extraId) {
    try {
      const extraSs = SpreadsheetApp.openById(extraId);
      const extraSheetName = ppGetProp_('SALES_PLANS_EXTRA_SHEET_NAME', PP_DEFAULT_EXTRA_SOURCE_SHEET_NAME);
      const sh = extraSs.getSheetByName(extraSheetName);
      if (!sh) {
        out.push([extraSs.getName(), extraSheetName, '—', '—', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Вкладка не найдена']);
      } else {
        inspectExtraSheet_(extraSs.getName(), sh, refBundle.supplierToWb);
      }
    } catch (e) {
      out.push(['(доп. книга)', '—', '—', '—', '—', '—', '—', '—', '—', '—', 0, 0, '', 'Не открылась: ' + (e.message || String(e))]);
    }
  }

  // Записываем результат на лист в активной книге.
  const dest = SpreadsheetApp.getActiveSpreadsheet();
  const diagName = 'Планирование закупок (диагностика)';
  let diag = dest.getSheetByName(diagName);
  if (!diag) diag = dest.insertSheet(diagName);
  diag.clearContents();
  diag.getRange(1, 1, out.length, out[0].length).setValues(out);
  diag.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  try { diag.setFrozenRows(1); } catch (e) {}

  let dataRows = 0;
  let qtyRows = 0;
  for (let i = 1; i < out.length; i++) {
    dataRows += Number(out[i][10]) || 0;
    qtyRows += Number(out[i][11]) || 0;
  }
  ui.alert(
    'Диагностика книги планов',
    'Записано строк отчёта: ' + (out.length - 1) +
      '\nИтого строк под шапками: ' + dataRows +
      '\nИз них с qty>0 в найденных plan-колонках: ' + qtyRows +
      '\n\nЛист: «' + diagName + '».',
    ui.ButtonSet.OK
  );
}

/**
 * Диагностика «Сводной» в книге 03 относительно колонки «В пути до конца горизонта, шт».
 * Проходит по тем же правилам, что и `ppBuildInboundByMonthForPlanning_`, но не молча:
 * собирает причины отбраковки каждой строки и кладёт отчёт на лист
 * «Планирование закупок (диагностика Сводной)».
 *
 * Read-only: ничего на «Сводной» и «Планирование закупок» не правит.
 */
function diagnoseProcurementInboundFromSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planName = ppGetProp_('PROCUREMENT_PLANNING_SHEET', PP_DEFAULT_OUT_SHEET);
  const plan = ss.getSheetByName(planName);
  if (!plan) {
    ui.alert('Лист', 'Не найден лист «' + planName + '».', ui.ButtonSet.OK);
    return;
  }
  if (plan.getLastRow() < PP_PROC_PLAN_DATA_START_ROW) {
    ui.alert('Данные', 'На листе «' + planName + '» нет строк для диагностики.', ui.ButtonSet.OK);
    return;
  }
  const planLastCol = plan.getLastColumn();
  const planHeaders = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, planLastCol).getDisplayValues()[0];
  const monthCols = ppPlanningFindMonthCols_(planHeaders);
  const planCols = ppPlanningFindWbBarcodeCol0_(planHeaders);

  const sv = ss.getSheetByName('Сводная');
  const out = [];
  const push = function () { out.push(Array.prototype.slice.call(arguments)); };

  push('Параметр', 'Значение', 'Пояснение');
  if (!sv) {
    push('Лист «Сводная»', 'НЕ НАЙДЕН', 'Запустите «Операционные потоки → Сводная 01→03 (боевой)» в книге 04');
    ppWriteInboundDiagSheet_(ss, out);
    ui.alert(
      'Диагностика «Сводной»',
      'Лист «Сводная» в этой книге отсутствует.\n' +
        'Запустите в книге 04: «Операционные потоки → Сводная 01→03 (боевой)».',
      ui.ButtonSet.OK
    );
    return;
  }
  push('Лист «Сводная»', 'найден', 'строк всего: ' + sv.getLastRow() + ', колонок: ' + sv.getLastColumn());
  push('Лист «Планирование закупок»', planName, 'строк под шапкой: ' + (plan.getLastRow() - PP_PROC_PLAN_DATA_START_ROW + 1));
  push('Горизонт месяцев в плане', monthCols.map(function (m) { return m.key; }).join(', ') || '(не найдены)', 'колонки YYYY-MM на «' + planName + '»');

  const dump = ppReadInboundRowsForPlanning_(ss);
  push('Шапка «Сводной»', 'строка ' + dump.headerRow, 'найдена сканом первых 10 строк');

  if (!dump.rows.length) {
    push('Состояние', 'нет строк под шапкой', 'либо лист пуст, либо шапка не распознана');
    ppWriteInboundDiagSheet_(ss, out);
    ui.alert('Диагностика «Сводной»', 'На листе «Сводная» нет строк под шапкой.', ui.ButtonSet.OK);
    return;
  }

  const hdr = dump.headers;
  const idxArticle = ppPlanningFindColByHeaderVariants_(hdr, ['Артикул ВБ', 'Артикул WB']);
  const idxBarcode = ppPlanningFindColByHeaderVariants_(hdr, ['ШК', 'Barcode']);
  const idxQty = ppPlanningFindColByHeaderVariants_(hdr, ['Итоговое количество', 'Количество', 'Кол-во']);
  const idxReady = ppPlanningFindColByHeaderVariants_(hdr, ['Дата готовности']);
  const idxStatus = ppPlanningFindColByHeaderVariants_(hdr, ['Статус заказа', 'status_name', 'Статус']);
  const idxEta = ppPlanningFindColByHeaderVariants_(hdr, ['Плановая дата поступления', 'ETA', 'Дата поступления', 'Дата прибытия']);
  const idxPeriod = ppPlanningFindColByHeaderVariants_(hdr, ['Период (MM/YY)', 'Период']);

  const colDescr = function (label, idx) {
    if (idx < 0) return [label, 'НЕ НАЙДЕНО', '(нет такого заголовка в шапке)'];
    const a1 = ppColToA1_(idx) + (dump.headerRow);
    return [label, a1, hdr[idx]];
  };
  push.apply(null, colDescr('Артикул ВБ', idxArticle));
  push.apply(null, colDescr('ШК', idxBarcode));
  push.apply(null, colDescr('Итоговое количество', idxQty));
  push.apply(null, colDescr('Статус заказа', idxStatus));
  push.apply(null, colDescr('ETA (плановая дата поступления)', idxEta));
  push.apply(null, colDescr('Период (MM/YY)', idxPeriod));
  push.apply(null, colDescr('Дата готовности', idxReady));

  const shippedRaw = ppGetProp_('PROCUREMENT_SHIPPED_STATUS_CODE', PP_SHIPPED_STATUS_CODE_DEFAULT);
  const shippedCode = ppCanonArticle_(shippedRaw);
  push('Код «отгружено»', shippedRaw, 'canon: «' + shippedCode + '» (свойство PROCUREMENT_SHIPPED_STATUS_CODE)');

  if (idxArticle < 0 || idxQty < 0 || idxReady < 0 || idxStatus < 0) {
    push('Состояние', 'отсутствуют обязательные колонки', 'нужны: Артикул ВБ, Итоговое количество, Дата готовности, Статус заказа');
    ppWriteInboundDiagSheet_(ss, out);
    ui.alert(
      'Диагностика «Сводной»',
      'В «Сводной» отсутствуют обязательные колонки.\nПодробности на листе «Планирование закупок (диагностика Сводной)».',
      ui.ButtonSet.OK
    );
    return;
  }

  const wbVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, planCols.wb + 1, plan.getLastRow() - PP_PROC_PLAN_DATA_START_ROW + 1, 1).getDisplayValues();
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, planCols.bc + 1, plan.getLastRow() - PP_PROC_PLAN_DATA_START_ROW + 1, 1).getDisplayValues();
  const need = ppPlanningNeedSetsFromWbBcColumns_(wbVals, bcVals);
  const monthSet = {};
  for (let i = 0; i < monthCols.length; i++) monthSet[monthCols[i].key] = true;

  let total = dump.rows.length;
  let dropQty = 0;
  let dropMatch = 0;
  let acceptedTotal = 0;
  let sumAcceptedTotal = 0;
  let acceptedTotalByArt = 0;
  let acceptedTotalByBc = 0;
  let noDateRows = 0;
  let sumNoDateQty = 0;
  let outOfHorizonRows = 0;
  let sumOutOfHorizonQty = 0;
  let inHorizonRows = 0;
  let sumInHorizonQty = 0;
  const exQty = [];
  const exMatch = [];
  const exDate = [];
  const exHorizon = [];
  const byMonth = {};
  const byWbMonth = {};
  const byBcMonth = {};
  const byWbTotal = {};
  const byBcTotal = {};

  const exHeader = ['Артикул ВБ', 'ШК', 'qty (сырое)', 'Статус', 'ETA (сырое)', 'Период (сырое)', 'Дата готовности (сырое)', 'Распознанный месяц'];
  const pickExample = function (bucket, i, monthKey) {
    if (bucket.length >= 5) return;
    const r = dump.rows[i];
    const d = dump.rowsDisplay[i];
    bucket.push([
      String(idxArticle < r.length ? d[idxArticle] : '').trim(),
      String(idxBarcode >= 0 && idxBarcode < r.length ? d[idxBarcode] : '').trim(),
      String(idxQty < r.length ? d[idxQty] : '').trim(),
      String(idxStatus < r.length ? d[idxStatus] : '').trim(),
      String(idxEta >= 0 && idxEta < r.length ? d[idxEta] : '').trim(),
      String(idxPeriod >= 0 && idxPeriod < r.length ? d[idxPeriod] : '').trim(),
      String(idxReady < r.length ? d[idxReady] : '').trim(),
      monthKey || ''
    ]);
  };

  for (let i = 0; i < dump.rows.length; i++) {
    const r = dump.rows[i];
    const d = dump.rowsDisplay[i];
    const qty = ppParseQtyForPlan_(idxQty < r.length ? r[idxQty] : '');
    if (!(qty > 0)) { dropQty++; pickExample(exQty, i, ''); continue; }
    const artRaw = idxArticle < r.length ? d[idxArticle] : '';
    const bcRaw = idxBarcode >= 0 && idxBarcode < r.length ? d[idxBarcode] : '';
    const art = ppCanonArticle_(artRaw);
    const bc = ppCanonBarcodeForStock_(bcRaw);
    const hitArt = !!(art && need.needWb[art]);
    const hitBc = !!(bc && need.needBc[bc]);
    if (!hitArt && !hitBc) { dropMatch++; pickExample(exMatch, i, ''); continue; }
    acceptedTotal++;
    sumAcceptedTotal += qty;
    if (hitArt) { acceptedTotalByArt++; byWbTotal[art] = (byWbTotal[art] || 0) + qty; }
    if (hitBc) { acceptedTotalByBc++; byBcTotal[bc] = (byBcTotal[bc] || 0) + qty; }

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
    if (!arrival) {
      noDateRows++;
      sumNoDateQty += qty;
      pickExample(exDate, i, '');
      continue;
    }
    const mk = ppMonthKeyFromDate_(arrival);
    if (!monthSet[mk]) {
      outOfHorizonRows++;
      sumOutOfHorizonQty += qty;
      pickExample(exHorizon, i, mk);
      continue;
    }
    inHorizonRows++;
    sumInHorizonQty += qty;
    if (!byMonth[mk]) byMonth[mk] = { rows: 0, qty: 0, art: 0, bc: 0 };
    byMonth[mk].rows++;
    byMonth[mk].qty += qty;
    if (hitArt) { byMonth[mk].art++; const k = art + '|' + mk; byWbMonth[k] = (byWbMonth[k] || 0) + qty; }
    if (hitBc) { byMonth[mk].bc++; const k2 = bc + '|' + mk; byBcMonth[k2] = (byBcMonth[k2] || 0) + qty; }
  }

  push('', '', '');
  push('--- СВОДКА ПО ОБРАБОТКЕ ---', '', '');
  push('Всего строк под шапкой', String(total), '');
  push('Отброшено: qty<=0 или пусто', String(dropQty), 'служебные/итоговые строки');
  push('Отброшено: не найдено в плане', String(dropMatch), 'нет совпадения по «Артикул ВБ» и по «ШК»');
  push('Учтено в «В пути …» (всего, по плану)', String(acceptedTotal), 'сумма qty: ' + sumAcceptedTotal + ' — это значение идёт в колонку «' + PP_PURCHASE_INBOUND_COL_HEADER + '»');
  push('  из них по «Артикул ВБ»', String(acceptedTotalByArt), '');
  push('  из них по «ШК»', String(acceptedTotalByBc), '');
  push('  в т.ч. без распознанной даты', String(noDateRows), 'сумма qty: ' + sumNoDateQty + ' — учтены в «итого в пути», но пока без месяца поступления');
  push('  в т.ч. с месяцем вне горизонта плана', String(outOfHorizonRows), 'сумма qty: ' + sumOutOfHorizonQty + ' — например, 2026-04 уже отгружено, но в горизонт не входит');
  push('  в т.ч. с месяцем в горизонте плана', String(inHorizonRows), 'сумма qty: ' + sumInHorizonQty + ' — раскладка по месяцам ниже');

  push('', '', '');
  push('--- РАЗЛОЖКА УЧТЁННЫХ ПО МЕСЯЦАМ ---', '', '');
  push('Месяц (YYYY-MM)', 'Строк', 'Сумма qty / по «Артикул ВБ» / по «ШК»');
  const monthKeysSorted = Object.keys(byMonth).sort();
  for (let i = 0; i < monthKeysSorted.length; i++) {
    const mk = monthKeysSorted[i];
    const x = byMonth[mk];
    push(mk, String(x.rows), x.qty + ' / ' + x.art + ' / ' + x.bc);
  }

  const writeExamples = function (title, bucket) {
    push('', '', '');
    push('--- ' + title + ' ---', '', '');
    if (!bucket.length) {
      push('(нет таких строк)', '', '');
      return;
    }
    push.apply(null, exHeader);
    for (let i = 0; i < bucket.length; i++) push.apply(null, bucket[i]);
  };
  writeExamples('ПРИМЕРЫ: отброшено по qty<=0', exQty);
  writeExamples('ПРИМЕРЫ: отброшено по «нет даты»', exDate);
  writeExamples('ПРИМЕРЫ: отброшено по «месяц вне горизонта»', exHorizon);
  writeExamples('ПРИМЕРЫ: отброшено по «не найдено в плане»', exMatch);

  let planRows = wbVals.length;
  let coveredAny = 0;
  let coveredByArtOnly = 0;
  let coveredByBcOnly = 0;
  let coveredBoth = 0;
  let coveredNone = 0;
  let sumCoverageQty = 0;
  for (let i = 0; i < planRows; i++) {
    const art = ppCanonArticle_(wbVals[i][0]);
    const bc = ppCanonBarcodeForStock_(bcVals[i][0]);
    const hitArt = !!(art && byWbTotal[art] > 0);
    const hitBc = !!(bc && byBcTotal[bc] > 0);
    let rowQty = 0;
    if (art && byWbTotal[art] !== undefined) rowQty = byWbTotal[art];
    else if (bc && byBcTotal[bc] !== undefined) rowQty = byBcTotal[bc];
    sumCoverageQty += rowQty;
    if (hitArt && hitBc) coveredBoth++;
    else if (hitArt) coveredByArtOnly++;
    else if (hitBc) coveredByBcOnly++;
    else coveredNone++;
    if (hitArt || hitBc) coveredAny++;
  }
  push('', '', '');
  push('--- ПОКРЫТИЕ ПЛАНА ---', '', '');
  push('Строк в плане', String(planRows), '');
  push('Нашли «в пути» (всего, по плану)', String(coveredAny), 'сумма qty: ' + sumCoverageQty + ' — это попадёт в колонку «' + PP_PURCHASE_INBOUND_COL_HEADER + '»');
  push('  только по «Артикул ВБ»', String(coveredByArtOnly), '');
  push('  только по «ШК»', String(coveredByBcOnly), '');
  push('  и по «Артикул ВБ», и по «ШК»', String(coveredBoth), '');
  push('Ничего не найдено', String(coveredNone), 'эти строки получат 0 в «В пути …»');

  ppWriteInboundDiagSheet_(ss, out);
  ui.alert(
    'Диагностика «Сводной»',
    'Готово. Лист «Планирование закупок (диагностика Сводной)».\n' +
      'Всего строк: ' + total + ', учтено по плану: ' + acceptedTotal + ' (qty ' + sumAcceptedTotal + ').\n' +
      'Отброшено: qty=' + dropQty + ', не в плане=' + dropMatch + '.\n' +
      'В т.ч. без даты: ' + noDateRows + ' (qty ' + sumNoDateQty + '), вне горизонта: ' + outOfHorizonRows + ' (qty ' + sumOutOfHorizonQty + '), в горизонте: ' + inHorizonRows + ' (qty ' + sumInHorizonQty + ').\n' +
      'Покрыто строк плана: ' + coveredAny + ' из ' + planRows + ' (не покрыто: ' + coveredNone + ').',
    ui.ButtonSet.OK
  );
}

function ppWriteInboundDiagSheet_(ss, rows) {
  const name = 'Планирование закупок (диагностика Сводной)';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const norm = rows.map(function (r) {
    const out = r.slice();
    while (out.length < width) out.push('');
    return out;
  });
  const range = sh.getRange(1, 1, norm.length, width);
  // Принудительно делаем формат «обычный текст», иначе Google интерпретирует значения,
  // начинающиеся с «=», как формулу и показывает #ERROR! в заголовках секций.
  range.setNumberFormat('@');
  range.setValues(norm);
  sh.getRange(1, 1, 1, width).setFontWeight('bold');
  try { sh.setFrozenRows(1); } catch (e) {}
  try { sh.autoResizeColumns(1, width); } catch (e) {}
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

function ppPlanningFindOzonFboStockCol0_(headers) {
  const target = ppCanonHeaderSimple_(PP_OZON_FBO_STOCK_COL_HEADER);
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (c === target) return i;
    if (c.indexOf('остатокozon') >= 0 && c.indexOf('fbo') >= 0) return i;
    if (c.indexOf('остатокозон') >= 0 && c.indexOf('fbo') >= 0) return i;
  }
  return -1;
}

function ppPlanningFindOzonFbsStockCol0_(headers) {
  const target = ppCanonHeaderSimple_(PP_OZON_FBS_STOCK_COL_HEADER);
  for (let i = 0; i < headers.length; i++) {
    const c = ppCanonHeaderSimple_(String(headers[i] || ''));
    if (c === target) return i;
    if (c.indexOf('остатокozon') >= 0 && c.indexOf('fbs') >= 0) return i;
    if (c.indexOf('остатокозон') >= 0 && c.indexOf('fbs') >= 0) return i;
  }
  return -1;
}

function ppPlanningFindSupplierArticleCol0_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const raw = String(headers[i] || '');
    if (ppIsSupplierSkuColumnHeader_(raw)) return i;
  }
  return ppPlanningFindColByHeaderVariants_(headers, ['Артикул поставщика', 'Наименование']);
}

/** Набор canon(«Артикул поставщика») для строк плана. Используется в матчинге Ozon. */
function ppPlanningNeedSupplierSet_(supplierVals) {
  const need = {};
  for (let i = 0; i < supplierVals.length; i++) {
    const s = ppCanonArticle_(supplierVals[i][0]);
    if (s) need[s] = true;
  }
  return need;
}

/** Универсальный lookup по (артикул поставщика, ШК) — для Ozon-карт остатков и продаж. */
function ppLookupBySupplierBarcode_(supplierCell, bcCell, bySupplier, byBarcode) {
  const sa = ppCanonArticle_(supplierCell);
  if (sa && bySupplier && bySupplier[sa] !== undefined) return bySupplier[sa];
  const bc = ppCanonBarcodeForStock_(bcCell);
  if (bc && byBarcode && byBarcode[bc] !== undefined) return byBarcode[bc];
  return 0;
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
  // ВАЖНО: ищем в порядке вариантов. Сначала первый вариант — по всей шапке;
  // только если не нашли — пробуем второй и т.д. Это нужно, например, чтобы
  // для idxQty взять «Итоговое количество» (G в «Сводной»), а не «Количество» (D),
  // — обе колонки есть в шапке, и при цикле «по колонкам сначала» брался ошибочный.
  const canonHeaders = [];
  for (let i = 0; i < headers.length; i++) {
    canonHeaders.push(ppCanonHeaderSimple_(String(headers[i] || '')));
  }
  for (let j = 0; j < variants.length; j++) {
    const v = ppCanonHeaderSimple_(variants[j]);
    for (let i = 0; i < canonHeaders.length; i++) {
      if (canonHeaders[i] === v) return i;
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
  return ppParseQtyForPlan_(v);
}

function ppReadInboundRowsForPlanning_(ss) {
  const sh = ss.getSheetByName('Сводная');
  if (!sh || sh.getLastRow() < 2) return { rows: [], headerRow: 1 };
  const lastCol = sh.getLastColumn();
  // Сводная в книге 03 — это снимок из 01. Иногда сверху появляются служебные
  // строки (заголовок «Сводная», описание, разделители). Сканируем первые 10
  // строк в поисках реальной шапки с «Артикул ВБ».
  const maxScan = Math.min(sh.getLastRow(), 10);
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
  const empty = { byWbMonth: {}, byBcMonth: {}, byWbTotal: {}, byBcTotal: {}, errors: [], usedRows: 0, totalRows: 0 };
  const dump = ppReadInboundRowsForPlanning_(ss);
  if (!dump.rows.length) {
    return Object.assign({}, empty, { errors: ['Лист «Сводная» пуст или не найден.'] });
  }
  const hdr = dump.headers;
  const idxArticle = ppPlanningFindColByHeaderVariants_(hdr, ['Артикул ВБ', 'Артикул WB']);
  const idxBarcode = ppPlanningFindColByHeaderVariants_(hdr, ['ШК', 'Barcode']);
  const idxQty = ppPlanningFindColByHeaderVariants_(hdr, ['Итоговое количество', 'Количество', 'Кол-во']);
  const idxReady = ppPlanningFindColByHeaderVariants_(hdr, ['Дата готовности']);
  const idxStatus = ppPlanningFindColByHeaderVariants_(hdr, ['Статус заказа', 'status_name', 'Статус']);
  const idxEta = ppPlanningFindColByHeaderVariants_(hdr, ['Плановая дата поступления', 'ETA', 'Дата поступления', 'Дата прибытия']);
  const idxPeriod = ppPlanningFindColByHeaderVariants_(hdr, ['Период (MM/YY)', 'Период']);
  if (idxArticle < 0 || idxQty < 0 || idxReady < 0 || idxStatus < 0) {
    return Object.assign({}, empty, { errors: ['В «Сводная» не найдены обязательные колонки для in-transit.'] });
  }

  const monthSet = {};
  for (let i = 0; i < monthKeys.length; i++) monthSet[monthKeys[i]] = true;
  // По месяцам (для будущего пересчёта остатков с учётом даты прихода).
  const byWbMonth = {};
  const byBcMonth = {};
  // Итого без оглядки на дату/горизонт: всё, что qty>0 и есть в плане.
  // Это значение идёт в колонку «В пути до конца горизонта, шт» на листе планирования —
  // пользователь хочет видеть «полное в пути» сейчас, а планирование по датам поступления
  // (раскладка по месяцам) появится на следующем этапе.
  const byWbTotal = {};
  const byBcTotal = {};
  const shippedCode = ppCanonArticle_(ppGetProp_('PROCUREMENT_SHIPPED_STATUS_CODE', PP_SHIPPED_STATUS_CODE_DEFAULT));
  let usedRows = 0;
  let totalRows = 0;

  for (let i = 0; i < dump.rows.length; i++) {
    const r = dump.rows[i];
    const d = dump.rowsDisplay[i];
    const artRaw = idxArticle < r.length ? d[idxArticle] : '';
    const bcRaw = idxBarcode >= 0 && idxBarcode < r.length ? d[idxBarcode] : '';
    const qty = ppToNumOr0_(idxQty < r.length ? r[idxQty] : '');
    if (!(qty > 0)) continue;
    const art = ppCanonArticle_(artRaw);
    const bc = ppCanonBarcodeForStock_(bcRaw);
    const hitArt = !!(art && needWb[art]);
    const hitBc = !!(bc && needBc[bc]);
    if (!hitArt && !hitBc) continue;
    // Итого «в пути» по плану — независимо от даты.
    if (hitArt) byWbTotal[art] = (byWbTotal[art] || 0) + qty;
    if (hitBc) byBcTotal[bc] = (byBcTotal[bc] || 0) + qty;
    totalRows++;

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
    if (hitArt) {
      const k = art + '|' + mk;
      byWbMonth[k] = (byWbMonth[k] || 0) + qty;
    }
    if (hitBc) {
      const k2 = bc + '|' + mk;
      byBcMonth[k2] = (byBcMonth[k2] || 0) + qty;
    }
    usedRows++;
  }
  return {
    byWbMonth: byWbMonth,
    byBcMonth: byBcMonth,
    byWbTotal: byWbTotal,
    byBcTotal: byBcTotal,
    errors: [],
    usedRows: usedRows,
    totalRows: totalRows
  };
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
 * Заполняет на листе «Планирование закупок» колонки «Остаток Ozon FBO, шт» и «Остаток Ozon FBS, шт».
 *
 * Сопоставление: offer_id ↔ canon(«Артикул поставщика»), barcode ↔ canon(«ШК»).
 * Если колонок Ozon ещё нет — создаются в конце шапки. Если в плане нет колонки
 * «Артикул поставщика», матчинг идёт только по ШК (что соответствует ограниченному
 * покрытию — пользователь увидит это в финальном сообщении: «без supplier-колонки»).
 */
function updateProcurementPlanningOzonStock() {
  const ui = SpreadsheetApp.getUi();
  if (!ozGetProp_('OZON_API_TOKEN', '')) {
    ui.alert(
      'Настройка',
      'Нужен OZON_API_TOKEN в свойствах скрипта.\n' +
        'Формат: «<Client-Id>:<Api-Key>» (значения из ЛК Ozon → Настройки → Seller API).',
      ui.ButtonSet.OK
    );
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
  const supplierCol0 = ppPlanningFindSupplierArticleCol0_(hdr);

  let fboCol0 = ppPlanningFindOzonFboStockCol0_(hdr);
  if (fboCol0 < 0) {
    fboCol0 = hdr.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, fboCol0 + 1).setValue(PP_OZON_FBO_STOCK_COL_HEADER).setFontWeight('bold');
  }
  const hdr2 = plan.getRange(PP_PROC_PLAN_HEADER_ROW, 1, 1, plan.getLastColumn()).getValues()[0];
  let fbsCol0 = ppPlanningFindOzonFbsStockCol0_(hdr2);
  if (fbsCol0 < 0) {
    fbsCol0 = hdr2.length;
    plan.getRange(PP_PROC_PLAN_HEADER_ROW, fbsCol0 + 1).setValue(PP_OZON_FBS_STOCK_COL_HEADER).setFontWeight('bold');
  }

  const lastRow = plan.getLastRow();
  const numRows = lastRow - PP_PROC_PLAN_DATA_START_ROW + 1;
  const bcVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, cols.bc + 1, numRows, 1).getDisplayValues();
  const supplierVals = supplierCol0 >= 0
    ? plan.getRange(PP_PROC_PLAN_DATA_START_ROW, supplierCol0 + 1, numRows, 1).getDisplayValues()
    : [];
  const needSupplier = supplierVals.length ? ppPlanningNeedSupplierSet_(supplierVals) : {};
  const needBc = {};
  for (let i = 0; i < bcVals.length; i++) {
    const b = ppCanonBarcodeForStock_(bcVals[i][0]);
    if (b) needBc[b] = true;
  }
  const needSupplierCount = Object.keys(needSupplier).length;
  const needBcCount = Object.keys(needBc).length;
  if (!needSupplierCount && !needBcCount) {
    ui.alert(
      'Планирование',
      'В колонках «Артикул поставщика» и «ШК» нет ни одного заполненного значения — нечего сопоставлять с Ozon.',
      ui.ButtonSet.OK
    );
    return;
  }

  let lookup;
  try {
    lookup = ozBuildStockLookupForProcurementPlanning_(needSupplier, needBc);
  } catch (e) {
    ui.alert('Ozon', 'Ошибка при получении остатков:\n' + (e.message || String(e)), ui.ButtonSet.OK);
    return;
  }

  const outFbo = [];
  const outFbs = [];
  let filledFbo = 0;
  let filledFbs = 0;
  for (let i = 0; i < numRows; i++) {
    const supCell = supplierVals.length ? supplierVals[i][0] : '';
    const bcCell = bcVals[i][0];
    const qFbo = ppLookupBySupplierBarcode_(supCell, bcCell, lookup.stockBySupplierFbo, lookup.stockByBarcodeFbo);
    const qFbs = ppLookupBySupplierBarcode_(supCell, bcCell, lookup.stockBySupplierFbs, lookup.stockByBarcodeFbs);
    if (qFbo) filledFbo++;
    if (qFbs) filledFbs++;
    outFbo.push([qFbo || '']);
    outFbs.push([qFbs || '']);
  }
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, fboCol0 + 1, numRows, 1).setValues(outFbo).setNumberFormat('0');
  plan.getRange(PP_PROC_PLAN_DATA_START_ROW, fbsCol0 + 1, numRows, 1).setValues(outFbs).setNumberFormat('0');

  ui.alert(
    'Готово',
    'Источник: ' +
      (lookup.source || 'Ozon Seller API') +
      '\nТоваров в /v3/product/info/stocks: ' +
      lookup.stats.itemsFromStocks +
      '\nСовпадений по «Артикул поставщика»: ' +
      lookup.stats.offerHits +
      (supplierCol0 < 0 ? ' (колонка «Артикул поставщика» не найдена — матчинг только по ШК)' : '') +
      '\nСовпадений по ШК: ' +
      lookup.stats.barcodeHits +
      '\nСтрок таблицы: ' +
      numRows +
      ', с FBO≠0: ' +
      filledFbo +
      ', с FBS≠0: ' +
      filledFbs +
      '\n\nКолонки: «' +
      PP_OZON_FBO_STOCK_COL_HEADER +
      '», «' +
      PP_OZON_FBS_STOCK_COL_HEADER +
      '».',
    ui.ButtonSet.OK
  );
  if (typeof logInfo === 'function') {
    logInfo('updateProcurementPlanningOzonStock', {
      source: lookup.source,
      rows: numRows,
      filledFbo: filledFbo,
      filledFbs: filledFbs,
      needSupplier: needSupplierCount,
      needBc: needBcCount,
      stats: lookup.stats
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
  const supplierCol0Compute = ppPlanningFindSupplierArticleCol0_(headers);
  const supplierValsForOzon = supplierCol0Compute >= 0
    ? plan.getRange(PP_PROC_PLAN_DATA_START_ROW, supplierCol0Compute + 1, numRows, 1).getDisplayValues()
    : nameVals;
  const msVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, msCol0 + 1, numRows, 1).getValues();
  const wbStockVals = plan.getRange(PP_PROC_PLAN_DATA_START_ROW, wbCol0 + 1, numRows, 1).getValues();
  // Колонки Ozon — необязательные. Если их нет, считаем что вклад в stockStart нулевой.
  const ozFboCol0 = ppPlanningFindOzonFboStockCol0_(headers);
  const ozFbsCol0 = ppPlanningFindOzonFbsStockCol0_(headers);
  const ozFboVals = ozFboCol0 >= 0
    ? plan.getRange(PP_PROC_PLAN_DATA_START_ROW, ozFboCol0 + 1, numRows, 1).getValues()
    : [];
  const ozFbsVals = ozFbsCol0 >= 0
    ? plan.getRange(PP_PROC_PLAN_DATA_START_ROW, ozFbsCol0 + 1, numRows, 1).getValues()
    : [];
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

  // Продажи Ozon (опционально): тянем только если есть токен и хотя бы одна supplier-строка,
  // иначе сразу пустой результат — расчёт работает как раньше, без Ozon.
  const needSupplier = supplierValsForOzon.length ? ppPlanningNeedSupplierSet_(supplierValsForOzon) : {};
  const needBcForOzon = need.needBc;
  let ozSalesFact = { bySupplier: {}, byBarcode: {}, rows: 0 };
  let ozSalesErr = '';
  const ozTokenPresent = !!(typeof ozGetProp_ === 'function' && ozGetProp_('OZON_API_TOKEN', ''));
  if (ozTokenPresent && typeof ozFetchCurrentMonthSalesForPlanning_ === 'function' && Object.keys(needSupplier).length) {
    try {
      ozSalesFact = ozFetchCurrentMonthSalesForPlanning_(needSupplier, needBcForOzon);
    } catch (e) {
      ozSalesErr = e && e.message ? e.message : String(e);
    }
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
    PP_OZON_FBO_STOCK_COL_HEADER,
    PP_OZON_FBS_STOCK_COL_HEADER,
    'Стартовый остаток, шт',
    'Продажи WB тек. месяца, шт',
    'Продажи Ozon тек. месяца, шт',
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
    const supplierForOzon = supplierValsForOzon.length ? supplierValsForOzon[r][0] : '';
    const ozFboQty = ozFboVals.length ? ppToNumOr0_(ozFboVals[r][0]) : 0;
    const ozFbsQty = ozFbsVals.length ? ppToNumOr0_(ozFbsVals[r][0]) : 0;
    const stockStart = ppToNumOr0_(msVals[r][0]) + ppToNumOr0_(wbStockVals[r][0]) + ozFboQty + ozFbsQty;
    const soldCurrentWb = ppLookupByRowKey_(wbCell, bcCell, salesFact.byWb, salesFact.byBarcode);
    const soldCurrentOzon = ppLookupBySupplierBarcode_(supplierForOzon, bcCell, ozSalesFact.bySupplier, ozSalesFact.byBarcode);
    const soldCurrent = soldCurrentWb + soldCurrentOzon;
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
        // Скорректированный план текущего месяца = план − факт WB − факт Ozon.
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
    // В колонку «В пути до конца горизонта, шт» пишем полное «итого в пути»
    // (qty>0 + есть в плане), независимо от того, удалось ли распознать дату.
    // Это видимость для пользователя «сколько всего едет». Помесячная разнёска
    // по датам поступления — следующий этап (планирование товаров в пути по датам).
    const totalInbound = ppLookupByRowKey_(wbCell, bcCell, inbound.byWbTotal, inbound.byBcTotal);
    outInbound.push([totalInbound || '']);
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
        ozFboQty,
        ozFbsQty,
        stockStart,
        soldCurrentWb,
        soldCurrentOzon,
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
  if (ozSalesErr) warn.push('Ozon продажи: ' + ozSalesErr);
  const ozonStockColsPresent = ozFboCol0 >= 0 || ozFbsCol0 >= 0;
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
      (inbound.totalRows || inbound.usedRows) +
      ' (по плану всего; в т.ч. с распознанной датой в горизонте: ' + inbound.usedRows + ')' +
      '\nWB продаж обработано строк: ' +
      salesFact.rows +
      (salesFact.source || salesFact.metric
        ? ' (' +
          [salesFact.source, salesFact.metric].filter(function (x) { return x; }).join('/') +
          (salesFact.fromCache ? ', из кеша' : '') +
          ')'
        : salesFact.fromCache ? ' (из кеша)' : '') +
      (salesFact.skipped ? ' (пропущено WB_SALES_SKIP)' : '') +
      '\nOzon: остатки в stockStart — ' + (ozonStockColsPresent ? 'учтены (FBO+FBS из листа)' : 'нет колонок на листе') +
      '; продажи тек. месяца обработано posting: ' + (ozSalesFact.rows || 0) +
      (ozTokenPresent ? '' : ' (OZON_API_TOKEN не задан)') +
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
      wbSalesSource: salesFact.source || null,
      wbSalesMetric: salesFact.metric || null,
      wbSalesFromCache: !!salesFact.fromCache,
      wbSalesSkipped: !!salesFact.skipped,
      ozonStockColsPresent: ozonStockColsPresent,
      ozonSalesRows: ozSalesFact.rows || 0,
      ozonTokenPresent: ozTokenPresent,
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
