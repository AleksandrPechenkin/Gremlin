const STOCK_CFG = {
  SUMMARY_SHEET: 'Сводная',
  STATUSES_SHEET: 'Статусы',
  MOVEMENTS_SHEET: 'Stock_Movements',
  BALANCE_SHEET: 'Транзитный склад',
  FORECAST_SHEET: 'Планирование отгрузок',
  HEADER_ROW: 2,
  DATA_START_ROW: 3
};

function addSenderStockMenu_(ui) {
  ui.createMenu('🚚 Транзитный склад')
    .addItem('Исправить формат "Номер спецификации"', 'fixSpecFormatEverywhere')
    .addItem('Пересчитать движения и остатки', 'rebuildSenderStockData')
    .addItem('Только пересчитать остатки', 'rebuildTransitBalanceOnly')
    .addSeparator()
    .addItem('Обновить "Планирование отгрузок"', 'buildShipmentForecastFromReadyDate')
    .addItem('Позиции по строке планирования (выделение)', 'showShipmentForecastDetailForSelection')
    .addToUi();
}

function rebuildSenderStockData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movementsSheet = ensureStockMovementsSheet_(ss);
  const balanceSheet = ensureTransitBalanceSheet_(ss);
  const statusEventMap = loadStatusEventsMap_(ss);
  const summary = ss.getSheetByName(STOCK_CFG.SUMMARY_SHEET);
  if (!summary) throw new Error('Не найден лист "Сводная".');
  const sourceSheets = [summary];
  let events = [];
  const totalStats = { total: 0, withEvent: 0, noSender: 0, noQty: 0, noArticle: 0 };

  sourceSheets.forEach(function (sh) {
    const result = buildStockEventsFromSheet_(sh, statusEventMap);
    events = events.concat(result.events);
    totalStats.total += result.stats.total;
    totalStats.withEvent += result.stats.withEvent;
    totalStats.noSender += result.stats.noSender;
    totalStats.noQty += result.stats.noQty;
    totalStats.noArticle += result.stats.noArticle;
  });

  // Защита от повторного учета: оставляем только уникальные движения по movement_key.
  events = dedupeStockEventsByKey_(events);

  writeStockMovements_(movementsSheet, events);
  rebuildTransitBalance_(balanceSheet, events);

  SpreadsheetApp.getUi().alert(
    'Готово',
    'Движений записано: ' + events.length +
      '\nИсточников (листов): ' + sourceSheets.length +
      '\nСтрок проверено: ' + totalStats.total +
      '\nСо статусом движения: ' + totalStats.withEvent +
      '\nОтброшено (пустой отправитель): ' + totalStats.noSender +
      '\nОтброшено (qty <= 0): ' + totalStats.noQty +
      '\nОтброшено (пустой артикул): ' + totalStats.noArticle,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function rebuildTransitBalanceOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movements = ss.getSheetByName(STOCK_CFG.MOVEMENTS_SHEET);
  const balance = ensureTransitBalanceSheet_(ss);
  if (!movements) throw new Error('Не найден лист Stock_Movements.');
  const events = readStockMovements_(movements);
  rebuildTransitBalance_(balance, events);
  SpreadsheetApp.getUi().alert('Остатки пересчитаны.');
}

function ensureStockMovementsSheet_(ss) {
  let sh = ss.getSheetByName(STOCK_CFG.MOVEMENTS_SHEET);
  if (!sh) sh = ss.insertSheet(STOCK_CFG.MOVEMENTS_SHEET);
  const headers = [[
    'created_at',
    'source_sheet',
    'source_row',
    'status_name',
    'stock_event',
    'sender_warehouse',
    'wb_article',
    'barcode',
    'supplier',
    'spec_number',
    'plan_period',
    'qty_in',
    'qty_out',
    'qty_signed',
    'volume_in',
    'volume_out',
    'volume_signed',
    'weight_in',
    'weight_out',
    'weight_signed',
    'movement_key'
  ]];
  sh.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function ensureTransitBalanceSheet_(ss) {
  let sh = ss.getSheetByName(STOCK_CFG.BALANCE_SHEET);
  if (!sh) sh = ss.insertSheet(STOCK_CFG.BALANCE_SHEET);
  const headers = [[
    'Источник лист',
    'Источник строка',
    'Отгрузка через',
    'Артикул ВБ',
    'ШК',
    'Поставщик',
    'Номер спецификации',
    'Период плана (MM/YY)',
    'Приход',
    'Расход',
    'Остаток',
    'Объем приход',
    'Объем расход',
    'Объем остаток',
    'Вес приход',
    'Вес расход',
    'Вес остаток'
  ]];
  sh.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Номер спецификации должен храниться как текст (например "16/1"), не как дата.
  sh.getRange('G:G').setNumberFormat('@');
  return sh;
}

function loadStatusEventsMap_(ss) {
  const sh = ss.getSheetByName(STOCK_CFG.STATUSES_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;

  // Поддержка гибкой схемы: B=status_name, D=stock_event
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(4, sh.getLastColumn())).getValues();
  values.forEach(function (r) {
    const statusName = String(r[1] || '').trim();
    const stockEvent = String(r[3] || '').trim().toUpperCase();
    if (!statusName) return;
    if (stockEvent === 'IN' || stockEvent === 'OUT') {
      map[canonStock_(statusName)] = stockEvent;
    }
  });

  // Фолбэк, если в D не заполнено.
  if (Object.keys(map).length === 0) {
    map[canonStock_('5. На складе логистов')] = 'IN';
    map[canonStock_('06. Зарезервирован в отправку')] = 'OUT';
    map[canonStock_('4. В пути в Москву')] = 'OUT';
  }
  return map;
}

function buildStockEventsFromSheet_(sheet, statusEventMap) {
  const headerRow = detectHeaderRow_(sheet);
  if (!headerRow) return { events: [], stats: { total: 0, withEvent: 0, noSender: 0, noQty: 0, noArticle: 0 } };
  const dataStart = headerRow + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStart) {
    return { events: [], stats: { total: 0, withEvent: 0, noSender: 0, noQty: 0, noArticle: 0 } };
  }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const hmap = indexMapStock_(headers);

  const idxStatus = findHeaderStock_(hmap, ['Статус заказа'], 11);
  const idxSpec = findHeaderStock_(hmap, ['Номер спецификации', 'Номер Спецификации'], 12);
  const idxArticle = findHeaderStock_(hmap, ['Артикул ВБ', 'Артикул WB'], 1);
  const idxBarcode = findHeaderStock_(hmap, ['ШК', 'Barcode'], 3);
  const idxSupplier = findHeaderStock_(hmap, ['Поставщик'], 9);
  // В "Сводная" поле "Отгрузка через" обычно в колонке Q (17-я, 1-based).
  const idxSender = findHeaderStock_(hmap, ['Отгрузка через'], 17);
  const idxQty = findHeaderStock_(hmap, ['Итоговое количество', 'Количество'], 7);
  const idxPeriod = findHeaderStock_(hmap, ['Период (MM/YY)'], 32);
  const idxVolume = findHeaderStock_(hmap, ['Объем', 'Volume'], 16);
  const idxWeight = findHeaderStock_(hmap, ['Вес', 'Weight'], 17);

  const range = sheet.getRange(dataStart, 1, lastRow - dataStart + 1, lastCol);
  const data = range.getValues();
  const dataDisplay = range.getDisplayValues();

  const now = new Date();
  const events = [];
  const stats = { total: 0, withEvent: 0, noSender: 0, noQty: 0, noArticle: 0 };
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const status = String(r[idxStatus] || '').trim();
    const event = resolveStockEvent_(status, statusEventMap);
    stats.total++;
    if (!event) continue;
    stats.withEvent++;

    const qty = toNumStock_(r[idxQty]);
    const article = String(r[idxArticle] || '').trim();
    const sender = normalizeShipViaForStock_(String(r[idxSender] || '').trim());
    if (!sender) { stats.noSender++; continue; }
    if (!article) { stats.noArticle++; continue; }
    if (!(qty > 0)) { stats.noQty++; continue; }

    const qtyIn = event === 'IN' ? qty : 0;
    const qtyOut = event === 'OUT' ? qty : 0;
    const absRow = dataStart + i;
    const spec = normalizeSpecStock_(r[idxSpec], dataDisplay[i][idxSpec]);
    const barcode = String(r[idxBarcode] || '').trim();
    const supplier = String(r[idxSupplier] || '').trim();
    const period = String(r[idxPeriod] || '').trim();
    const volume = toNumStock_(r[idxVolume]);
    const weight = toNumStock_(r[idxWeight]);
    const volumeIn = event === 'IN' ? volume : 0;
    const volumeOut = event === 'OUT' ? volume : 0;
    const weightIn = event === 'IN' ? weight : 0;
    const weightOut = event === 'OUT' ? weight : 0;
    const signed = qtyIn - qtyOut;
    const key = [sender, article, spec, absRow, event].join('|');

    events.push([
      now,
      sheet.getName(),
      absRow,
      status,
      event,
      sender,
      article,
      barcode,
      supplier,
      spec,
      period,
      qtyIn,
      qtyOut,
      signed,
      volumeIn,
      volumeOut,
      volumeIn - volumeOut,
      weightIn,
      weightOut,
      weightIn - weightOut,
      key
    ]);
  }
  return { events: events, stats: stats };
}

function detectHeaderRow_(sheet) {
  const maxCol = Math.max(sheet.getLastColumn(), 20);
  const rowsToCheck = Math.min(3, sheet.getMaxRows());
  const probe = sheet.getRange(1, 1, rowsToCheck, maxCol).getValues();
  for (let r = 0; r < probe.length; r++) {
    const map = indexMapStock_(probe[r]);
    if (map[canonStock_('Статус заказа')] != null && map[canonStock_('Отгрузка через')] != null) {
      return r + 1;
    }
  }
  return null;
}

function writeStockMovements_(sheet, events) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 21).clearContent();
  if (!events.length) return;
  sheet.getRange(2, 1, events.length, 21).setValues(events);
}

function readStockMovements_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 21).getValues();
}

function rebuildTransitBalance_(balanceSheet, events) {
  const rows = events.map(function (e) {
    const inQty = toNumStock_(e[11]);
    const outQty = toNumStock_(e[12]);
    const inVol = toNumStock_(e[14]);
    const outVol = toNumStock_(e[15]);
    const inW = toNumStock_(e[17]);
    const outW = toNumStock_(e[18]);
    return [
      String(e[1] || '').trim(),  // source sheet
      String(e[2] || '').trim(),  // source row
      String(e[5] || '').trim(),  // sender
      String(e[6] || '').trim(),  // article
      String(e[7] || '').trim(),  // barcode
      String(e[8] || '').trim(),  // supplier
      "'" + String(e[9] || '').trim(),  // spec (жестко текст)
      String(e[10] || '').trim(), // period
      inQty,
      outQty,
      inQty - outQty,
      inVol,
      outVol,
      inVol - outVol,
      inW,
      outW,
      inW - outW
    ];
  }).sort(function (a, b) {
    return String(a[0]).localeCompare(String(b[0])) ||
      Number(a[1]) - Number(b[1]) ||
      String(a[2]).localeCompare(String(b[2])) ||
      String(a[7]).localeCompare(String(b[7])) ||
      String(a[3]).localeCompare(String(b[3]));
  });

  const last = balanceSheet.getLastRow();
  if (last > 1) balanceSheet.getRange(2, 1, last - 1, 17).clearContent();
  if (rows.length) {
    balanceSheet.getRange(2, 1, rows.length, 17).setValues(rows);
    balanceSheet.getRange(2, 7, rows.length, 1).setNumberFormat('@');
  }
}

function indexMapStock_(headers) {
  const map = {};
  headers.forEach(function (h, i) {
    const c = canonStock_(h);
    if (c && map[c] == null) map[c] = i;
  });
  return map;
}

function findHeaderStock_(map, names, fallback1Based) {
  for (let i = 0; i < names.length; i++) {
    const key = canonStock_(names[i]);
    if (map[key] != null) return map[key];
  }
  return fallback1Based - 1;
}

/** Индекс колонки по заголовку или null, если ни один вариант не найден. */
function findHeaderIdxOptional_(map, names) {
  for (let i = 0; i < names.length; i++) {
    const key = canonStock_(names[i]);
    if (map[key] != null) return map[key];
  }
  return null;
}

function canonStock_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/ё/g, 'е')
    .replace(/[.,;:!?()[\]{}"']/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Подпись для «Отгрузка через», когда груз без вашего транзитного узла (прямая отгрузка с фабрики). */
const SHIP_VIA_NO_TRANSIT_CANON_ = 'Не транзит (прямая с фабрики)';

/** Приводит варианты вроде «условно не транзит» к одной строке для остатков и прогноза. */
function normalizeShipViaForStock_(raw) {
  const t = String(raw || '').trim();
  if (!t) return '';
  const k = canonStock_(t);
  if (!k) return '';
  if (shipViaNoTransitKeys_()[k]) return SHIP_VIA_NO_TRANSIT_CANON_;
  return t;
}

function shipViaNoTransitKeys_() {
  if (shipViaNoTransitKeys_.cache) return shipViaNoTransitKeys_.cache;
  const labels = [
    SHIP_VIA_NO_TRANSIT_CANON_,
    'условно не транзит',
    'не транзит',
    'прямая с фабрики',
    'без транзита',
    'direct from factory'
  ];
  const map = {};
  for (let i = 0; i < labels.length; i++) {
    const key = canonStock_(labels[i]);
    if (key) map[key] = true;
  }
  shipViaNoTransitKeys_.cache = map;
  return map;
}

function toNumStock_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v || '').replace(/\s/g, '').replace(',', '.');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function normalizeSpecStock_(rawValue, displayValue) {
  const shown = String(displayValue || '').trim();
  if (rawValue instanceof Date) {
    // Если в источнике спецификация была распознана как дата (например 16/1),
    // возвращаем компактный формат d/M, чтобы не терять исходный смысл.
    return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'd/M');
  }
  const txt = String(rawValue || '').trim();
  if (txt) return txt;
  return shown;
}

function resolveStockEvent_(statusName, statusEventMap) {
  const key = canonStock_(statusName);
  if (!key) return '';
  if (statusEventMap[key]) return statusEventMap[key];

  // Гибкий фолбэк по смыслу, если названия статусов поменяли.
  if (key.indexOf('на складе логистов') >= 0) return 'IN';
  if (key.indexOf('зарезервирован в отправку') >= 0) return 'OUT';
  if (key.indexOf('в пути в москву') >= 0) return 'OUT';
  return '';
}

function dedupeStockEventsByKey_(events) {
  const seen = {};
  const out = [];
  for (let i = 0; i < events.length; i++) {
    const row = events[i];
    const key = String(row[20] || row[14] || '').trim();
    if (!key) {
      out.push(row);
      continue;
    }
    if (seen[key]) continue;
    seen[key] = true;
    out.push(row);
  }
  return out;
}

function fixSpecFormatEverywhere() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const managerPattern = /\s\d{2}\/\d{2}$/;
  let fixedSheets = 0;
  let fixedCells = 0;

  sheets.forEach(function (sh) {
    const name = sh.getName();
    const isTarget =
      name === STOCK_CFG.SUMMARY_SHEET ||
      name === STOCK_CFG.BALANCE_SHEET ||
      managerPattern.test(name);
    if (!isTarget) return;

    const res = fixSpecFormatInSheet_(sh);
    if (res.fixed > 0 || res.checked > 0) fixedSheets++;
    fixedCells += res.fixed;
  });

  SpreadsheetApp.getUi().alert(
    'Исправление завершено',
    'Листов обработано: ' + fixedSheets + '\nЯчеек исправлено: ' + fixedCells,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function fixSpecFormatInSheet_(sheet) {
  const headerRow = detectHeaderRow_(sheet) || 2;
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return { checked: 0, fixed: 0 };

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const hmap = indexMapStock_(headers);
  const specCol = findHeaderStock_(hmap, ['Номер спецификации', 'Номер Спецификации'], 12) + 1;

  const rows = lastRow - headerRow;
  const rng = sheet.getRange(headerRow + 1, specCol, rows, 1);
  const values = rng.getValues();
  const display = rng.getDisplayValues();

  let fixed = 0;
  for (let i = 0; i < values.length; i++) {
    const raw = values[i][0];
    if (raw instanceof Date) {
      values[i][0] = Utilities.formatDate(raw, Session.getScriptTimeZone(), 'd/M');
      fixed++;
    } else {
      const txt = String(raw || '').trim();
      if (txt) values[i][0] = txt;
    }
    if (!values[i][0] && display[i][0]) values[i][0] = String(display[i][0]).trim();
  }

  rng.setNumberFormat('@');
  rng.setValues(values);
  rng.setNumberFormat('@');
  return { checked: rows, fixed: fixed };
}

const FORECAST_SENDER_EMPTY_ = '— Отгрузка через не заполнено';
const FORECAST_READY_EMPTY_WEEK_ = '—';
const FORECAST_READY_EMPTY_MONTH_ = '—';
const FORECAST_READY_EMPTY_BUCKET_ = 'Дата готовности не заполнена';

function buildShipmentForecastFromReadyDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(STOCK_CFG.SUMMARY_SHEET);
  if (!summary) throw new Error('Не найден лист "Сводная".');

  let forecast = ss.getSheetByName(STOCK_CFG.FORECAST_SHEET);
  if (!forecast) forecast = ss.insertSheet(STOCK_CFG.FORECAST_SHEET);

  const lastRow = summary.getLastRow();
  const lastCol = summary.getLastColumn();
  if (lastRow < STOCK_CFG.DATA_START_ROW) {
    forecast.clearContents();
    forecast.getRange(1, 1).setValue('Нет данных в Сводной');
    return;
  }

  const headers = summary.getRange(STOCK_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const hmap = indexMapStock_(headers);
  const idxReady = findHeaderStock_(hmap, ['Дата готовности'], 13);
  const idxSender = findHeaderStock_(hmap, ['Отгрузка через'], 17);
  const idxArticle = findHeaderStock_(hmap, ['Артикул ВБ', 'Артикул WB'], 1);
  const idxSpec = findHeaderStock_(hmap, ['Номер спецификации', 'Номер Спецификации'], 12);
  const idxQty = findHeaderStock_(hmap, ['Итоговое количество', 'Количество'], 7);
  const idxVolume = findHeaderStock_(hmap, ['Объем', 'Volume'], 16);
  const idxWeight = findHeaderStock_(hmap, ['Вес', 'Weight'], 17);
  const idxStatus = findHeaderStock_(hmap, ['Статус заказа'], 11);

  const srcRange = summary.getRange(STOCK_CFG.DATA_START_ROW, 1, lastRow - STOCK_CFG.DATA_START_ROW + 1, lastCol);
  const data = srcRange.getValues();
  const dataDisplay = srcRange.getDisplayValues();
  const summaryGid = summary.getSheetId();
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  const agg = {};
  data.forEach(function (r, i) {
    const senderRaw = normalizeShipViaForStock_(String(r[idxSender] || '').trim());
    const senderKey = senderRaw || '\u0000__NO_SENDER__';
    const senderOut = senderRaw || FORECAST_SENDER_EMPTY_;

    const ready = parseDateStock_(r[idxReady]);
    let weekKey;
    let monthKey;
    let bucket;
    if (ready) {
      const weekStart = mondayOfWeek_(ready);
      weekKey = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      monthKey = Utilities.formatDate(ready, Session.getScriptTimeZone(), 'MM/yy');
      bucket = forecastBucket_(now, ready);
    } else {
      weekKey = FORECAST_READY_EMPTY_WEEK_;
      monthKey = FORECAST_READY_EMPTY_MONTH_;
      bucket = FORECAST_READY_EMPTY_BUCKET_;
    }

    const article = String(r[idxArticle] || '').trim();
    const spec = normalizeSpecStock_(r[idxSpec], dataDisplay[i][idxSpec]);
    const qty = toNumStock_(r[idxQty]);
    const vol = toNumStock_(r[idxVolume]);
    const wt = toNumStock_(r[idxWeight]);
    const status = String(r[idxStatus] || '').trim();
    const key = [senderKey, weekKey, monthKey, bucket, spec].join('|');
    const absRow = STOCK_CFG.DATA_START_ROW + i;

    if (!agg[key]) {
      agg[key] = {
        sender: senderOut,
        week: weekKey,
        month: monthKey,
        bucket: bucket,
        spec: spec,
        rowNumbers: [],
        qty: 0,
        vol: 0,
        wt: 0,
        rows: 0,
        articles: {},
        statuses: {}
      };
    }
    const a = agg[key];
    a.rowNumbers.push(absRow);
    a.qty += qty;
    a.vol += vol;
    a.wt += wt;
    a.rows += 1;
    if (article) a.articles[article] = true;
    if (status) a.statuses[status] = true;
  });

  const out = [[
    'Отгрузка через',
    'Спецификация',
    'Неделя (понедельник)',
    'Период (MM/YY)',
    'Окно планирования',
    'Строк',
    'Уник. артикулов',
    'Сумма qty',
    'Сумма объема',
    'Сумма веса',
    'Статусы (сводно)',
    ''
  ]];

  Object.keys(agg)
    .sort(function (ka, kb) {
      return forecastAggSortCmp_(agg, ka, kb);
    })
    .forEach(function (k) {
    const a = agg[k];
    const nums = uniqueSortedInts_(a.rowNumbers);
    const minR = nums[0];
    const maxR = nums[nums.length - 1];
    const rangeA1 = summary.getRange(minR, 1, maxR, lastCol).getA1Notation();
    const labelRaw = String(a.spec || '').trim() || '(без номера)';
    const safeLabel = labelRaw.replace(/"/g, '""');
    const linkFormula = '=HYPERLINK("#gid=' + summaryGid + '&range=' + rangeA1 + '";"' + safeLabel + '")';
    const rowMeta = nums.join('|');
    out.push([
      a.sender,
      linkFormula,
      a.week,
      a.month,
      a.bucket,
      a.rows,
      Object.keys(a.articles).length,
      a.qty,
      +a.vol.toFixed(4),
      +a.wt.toFixed(2),
      Object.keys(a.statuses).sort().join('; '),
      rowMeta
    ]);
  });

  forecast.clearContents();
  forecast.getRange(1, 1, out.length, out[0].length).setValues(out);
  forecast.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  forecast.setFrozenRows(1);
  forecast.hideColumns(12, 12);
  SpreadsheetApp.getUi().alert('Планирование отгрузок обновлено. Строк: ' + (out.length - 1));
}

/** Просмотр позиций «Сводной» по выбранной строке «Планирование отгрузок» (кол. 12 — скрытые номера строк). */
function showShipmentForecastDetailForSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== STOCK_CFG.FORECAST_SHEET) {
    SpreadsheetApp.getUi().alert(
      'Откройте лист «' + STOCK_CFG.FORECAST_SHEET + '», выберите строку отчёта (со 2-й) и повторите.'
    );
    return;
  }
  const row = sh.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Выберите строку данных (начиная со 2-й).');
    return;
  }
  const meta = String(sh.getRange(row, 12).getValue() || '').trim();
  if (!meta) {
    SpreadsheetApp.getUi().alert(
      'Нет привязки к строкам «Сводной». Сначала: меню «Транзитный склад» → «Обновить \u00abПланирование отгрузок\u00bb».'
    );
    return;
  }
  const rowNums = meta.split('|').map(function (s) {
    return parseInt(String(s).trim(), 10);
  }).filter(function (n) {
    return n > 0;
  });
  if (!rowNums.length) {
    SpreadsheetApp.getUi().alert('Не удалось разобрать номера строк «Сводной».');
    return;
  }
  const summary = ss.getSheetByName(STOCK_CFG.SUMMARY_SHEET);
  if (!summary) {
    SpreadsheetApp.getUi().alert('Не найден лист «Сводная».');
    return;
  }
  const planLabels = sh.getRange(row, 1, row, 5).getDisplayValues()[0];
  const html = buildForecastDetailHtmlFromSummaryRows_(ss, summary, uniqueSortedInts_(rowNums), planLabels);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(960).setHeight(560),
    'Позиции: ' + String(planLabels[1] || planLabels[0] || 'строка ' + row).slice(0, 80)
  );
}

function uniqueSortedInts_(arr) {
  const seen = {};
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    const n = Number(arr[i]);
    if (!(n > 0) || seen[n]) continue;
    seen[n] = true;
    out.push(n);
  }
  out.sort(function (a, b) {
    return a - b;
  });
  return out;
}

function escapeHtmlForecast_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildForecastDetailHtmlFromSummaryRows_(ss, summary, rowNums, planLabels) {
  const lastCol = summary.getLastColumn();
  const hdr = summary.getRange(STOCK_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const hmap = indexMapStock_(hdr);
  const colDefs = [
    { title: 'Номер спецификации', names: ['Номер спецификации', 'Номер Спецификации'], fb: 12 },
    { title: 'Артикул ВБ', names: ['Артикул ВБ', 'Артикул WB'], fb: 1 },
    {
      title: 'Артикул поставщика',
      names: ['Артикул поставщика', 'Артикул Поставщика', 'Supplier SKU', 'Supplier article'],
      optional: true
    },
    { title: 'ШК', names: ['ШК', 'Barcode'], fb: 3 },
    { title: 'Кол-во', names: ['Итоговое количество', 'Количество'], fb: 7 },
    { title: 'Дата готовности', names: ['Дата готовности'], fb: 13 },
    { title: 'Поставщик', names: ['Поставщик'], fb: 9 },
    { title: 'Объём', names: ['Объем', 'Volume'], fb: 16 },
    { title: 'Вес', names: ['Вес', 'Weight'], fb: 17 }
  ];
  const usedIdx = {};
  const cols = [];
  for (let c = 0; c < colDefs.length; c++) {
    const d = colDefs[c];
    let idx;
    if (d.optional) {
      idx = findHeaderIdxOptional_(hmap, d.names);
      if (idx == null) continue;
    } else {
      idx = findHeaderStock_(hmap, d.names, d.fb);
    }
    if (usedIdx[idx]) continue;
    usedIdx[idx] = true;
    cols.push({ title: d.title, idx: idx });
  }

  const maxRows = 200;
  const list = rowNums.length > maxRows ? rowNums.slice(0, maxRows) : rowNums;
  const truncated = rowNums.length > maxRows;

  const want = {};
  for (let i = 0; i < list.length; i++) want[list[i]] = true;
  const minR = list[0];
  const maxR = list[list.length - 1];
  const block = summary.getRange(minR, 1, maxR - minR + 1, lastCol).getDisplayValues();
  const bodyRows = [];
  for (let i = 0; i < block.length; i++) {
    const rn = minR + i;
    if (!want[rn]) continue;
    const cells = block[i];
    const tds = cols.map(function (col) {
      return '<td>' + escapeHtmlForecast_(cells[col.idx]) + '</td>';
    });
    bodyRows.push('<tr><td style="color:#666">' + rn + '</td>' + tds.join('') + '</tr>');
  }

  const ths = cols.map(function (col) {
    return '<th>' + escapeHtmlForecast_(col.title) + '</th>';
  });
  const jumpRange = summary.getRange(minR, 1, maxR, lastCol).getA1Notation();
  const jumpUrl =
    'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + summary.getSheetId() + '&range=' + jumpRange;

  const ctx =
    '<p style="margin:0 0 8px 0;font-size:13px">' +
    '<b>Спецификация:</b> ' +
    escapeHtmlForecast_(planLabels[1]) +
    ' &nbsp;|&nbsp; <b>Неделя:</b> ' +
    escapeHtmlForecast_(planLabels[2]) +
    ' &nbsp;|&nbsp; <b>Период:</b> ' +
    escapeHtmlForecast_(planLabels[3]) +
    ' &nbsp;|&nbsp; <b>Окно:</b> ' +
    escapeHtmlForecast_(planLabels[4]) +
    '</p>';

  return (
    '<!DOCTYPE html><html><head><base target="_blank">' +
    '<meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:12px;margin:10px}' +
    'table{border-collapse:collapse;width:100%}' +
    'th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}' +
    'th{background:#f3f3f3;position:sticky;top:0}' +
    'tr:nth-child(even){background:#fafafa}' +
    '.wrap{max-height:420px;overflow:auto;border:1px solid #ddd}' +
    '</style></head><body>' +
    ctx +
    (truncated
      ? '<p style="color:#b45309">Показаны первые ' + maxRows + ' из ' + rowNums.length + ' строк.</p>'
      : '') +
    '<div class="wrap"><table><thead><tr><th>#</th>' +
    ths.join('') +
    '</tr></thead><tbody>' +
    bodyRows.join('') +
    '</tbody></table></div>' +
    '<p style="margin-top:12px"><a href="' +
    jumpUrl +
    '">Открыть этот диапазон на листе «Сводная»</a></p>' +
    '</body></html>'
  );
}

function parseDateStock_(v) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    const d = new Date(v);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const s = String(v || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function mondayOfWeek_(d) {
  const x = new Date(d);
  const day = x.getDay() || 7;
  x.setDate(x.getDate() - day + 1);
  x.setHours(0, 0, 0, 0);
  return x;
}

function forecastBucket_(today, readyDate) {
  const diff = Math.floor((readyDate - today) / 86400000);
  if (diff < 0) return 'Просрочено';
  if (diff <= 7) return '0-7 дней';
  if (diff <= 14) return '8-14 дней';
  if (diff <= 30) return '15-30 дней';
  return '31+ дней';
}

/** Timestamp понедельника недели из колонки «Неделя»; без даты — null (в конец списка). */
function forecastReadyWeekTs_(weekKeyStr) {
  if (!weekKeyStr || weekKeyStr === FORECAST_READY_EMPTY_WEEK_) return null;
  const m = String(weekKeyStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const t = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
  return isNaN(t) ? null : t;
}

function forecastAggSortCmp_(agg, ka, kb) {
  const a = agg[ka];
  const b = agg[kb];
  const ta = forecastReadyWeekTs_(a.week);
  const tb = forecastReadyWeekTs_(b.week);
  if (ta != null && tb != null && ta !== tb) return ta - tb;
  if (ta != null && tb == null) return -1;
  if (ta == null && tb != null) return 1;
  const cmpSender = String(a.sender).localeCompare(String(b.sender), 'ru');
  if (cmpSender !== 0) return cmpSender;
  return String(a.spec).localeCompare(String(b.spec), 'ru');
}
