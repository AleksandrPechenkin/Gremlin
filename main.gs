/**
 * Работа с кодом (договорённость с ассистентом / владельцем):
 * — В Google Apps Script каждый модуль переносится целиком: открыть файл .gs в этой папке
 *   проекта → выделить всё → вставить в соответствующий файл редактора скриптов → сохранить.
 *   Не собирать код из фрагментов в чате.
 * — Актуальные версии файлов — в папке проекта на диске; полный выклад файла в переписку
 *   не обязателен, если не попросили отдельно.
 * Подробнее для ассистента: PROJECT_CONTEXT.md
 */
const SHEET_NAME = 'Сводная';

const COL = {
  WB_ARTICLE: 0, SUPPLIER_ARTICLE: 1, BARCODE: 2, QTY: 3, QTY_PER_BOX: 4, BOXES: 5, TOTAL_QTY: 6,
  MANAGER: 7, SUPPLIER_NOTE: 8, DELIVERY_TYPE: 9, ORDER_STATUS: 10, SPEC_NUMBER: 11, READY_DATE: 12,
  PRICE: 13, AMOUNT: 14, VOLUME: 15, WEIGHT: 16, SUPPLIER_MS: 17, LABELING: 18,
  ADVANCE_SUM: 19, ADVANCE_PLAN: 20, ADVANCE_FACT: 21,
  BALANCE_SUM: 22, BALANCE_PLAN: 23, BALANCE_FACT: 24,
  DEFER_SUM: 25, DEFER_PLAN: 26, DEFER_FACT: 27,
  MS_ID: 28, MS_LINK: 29, STATUS: 30
};

const MS_ATTR = {
  SPEC_NUMBER: '31085725-4be3-11ee-0a80-0b1600104da1',
  ADVANCE_SUM: '306e1964-31cf-11f1-0a80-1d9d003965b2',
  ADVANCE_PLAN: '306e1b69-31cf-11f1-0a80-1d9d003965b3',
  ADVANCE_FACT: '306e1c6c-31cf-11f1-0a80-1d9d003965b4',
  BALANCE_SUM: '306e1d58-31cf-11f1-0a80-1d9d003965b5',
  BALANCE_PLAN: '306e1ecd-31cf-11f1-0a80-1d9d003965b6',
  BALANCE_FACT: '306e1fbe-31cf-11f1-0a80-1d9d003965b7',
  DEFER_SUM: '306e20a7-31cf-11f1-0a80-1d9d003965b8',
  DEFER_PLAN: '306e2196-31cf-11f1-0a80-1d9d003965b9',
  DEFER_FACT: '306e229e-31cf-11f1-0a80-1d9d003965ba'
};

const cache = { suppliers: {}, products: {}, currencyYuanMeta: null };

/** Совпадает с supplier_invoices: шапка в строке 2, данные с 3. */
const MANAGER_SUMMARY_SYNC_HEADER_ROW = 2;
const MANAGER_SUMMARY_SYNC_DATA_START_ROW = 3;
/** Фон строки-разделителя между блоками календарных месяцев (из имени вкладки Имя ММ/ГГ). */
const MANAGER_SUMMARY_MONTH_SEP_BG = '#D9E1F2';

/**
 * Флаг показа меню Sync Hub в книге 01.
 * - unset / 1 / true / yes / on: меню показываем
 * - 0 / false / no / off: меню скрываем (используется после переноса хаба в 04)
 */
function isSyncHubMenuEnabled_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('SYNC_HUB_MENU_ENABLED');
    if (raw == null || String(raw).trim() === '') return true;
    const v = String(raw).trim().toLowerCase();
    return !(v === '0' || v === 'false' || v === 'no' || v === 'off');
  } catch (e) {
    return true;
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 МойСклад')
    .addItem('🔄 Синхронизировать заказы (Создать / Обновить)', 'syncOrdersWithMS')
    .addSeparator()
    .addItem('💸 Обновить оплаты в таблице "Закуплено"', 'updateExternalPurchases')
    .addItem('🧩 Собрать "Сводная" из вкладок менеджеров', 'syncManagerTabsToSummary')
    .addToUi();
  if (typeof addSupplierInvoiceMenu_ === 'function') {
    addSupplierInvoiceMenu_(ui);
  }
  if (typeof addPaymentRegistryMenu_ === 'function') {
    addPaymentRegistryMenu_(ui);
  }
  if (typeof addSenderStockMenu_ === 'function') {
    addSenderStockMenu_(ui);
  }
  if (typeof addSyncHubMenu_ === 'function' && isSyncHubMenuEnabled_()) {
    addSyncHubMenu_(ui);
  }
  ui.createMenu('Планирование закупок')
    .addItem('Подтянуть планы продаж на лист «Планирование закупок»', 'refreshProcurementPlanningFromSalesSheets')
    .addItem('Обновить лист «Склады МС (остатки)» из МойСклад', 'syncMsStockStoresSheet')
    .addItem('Записать учётный остаток МС на «Планирование закупок»', 'updateProcurementPlanningMsAccountingStock')
    .addItem('Записать остаток WB на «Планирование закупок»', 'updateProcurementPlanningWbStock')
    .addItem('Рассчитать потребность закупки (остатки + в пути + продажи)', 'computeProcurementPurchasePlan')
    .addItem('Проверить сопоставление остатков (артикул/ШК)', 'checkProcurementPlanningStocksCoverage')
    .addToUi();
}

function syncOrdersWithMS() {
  try {
    getScriptPropOrThrow('MS_ORGANIZATION_ID');
  } catch (error) {
    return SpreadsheetApp.getUi().alert(`Ошибка настройки: ${error.message}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return;

  const updater = new BatchUpdater(sheet, data);
  const groups = {};
  
  const originalMsIds = new Set();

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const supplierName = safeString(row[COL.SUPPLIER_MS]); 
    const specNum = safeString(row[COL.SPEC_NUMBER]);
    const msId = safeString(row[COL.MS_ID]);

    if (msId && msId !== 'ID заказа МС') originalMsIds.add(msId);

    if (!supplierName) { 
      if (safeString(row[COL.WB_ARTICLE]) || safeString(row[COL.SUPPLIER_ARTICLE])) {
        updater.setStatus(i, '⚠️ Пропуск: нет поставщика в колонке R'); 
      }
      continue; 
    }

    const finalSpecNum = specNum || ('NO-SPEC-' + supplierName);
    const groupKey = supplierName + '|||' + finalSpecNum;

    if (!groups[groupKey]) groups[groupKey] = { supplierName: supplierName, specNum: specNum, candidateMsIds: [], rows: [] };
    groups[groupKey].rows.push({ rowIndex: i, rowData: row });

    if (msId && msId !== 'ID заказа МС') groups[groupKey].candidateMsIds.push(msId);
  }

  const msIdClaims = {};
  for (const key in groups) {
    const counts = {};
    groups[key].candidateMsIds.forEach(id => counts[id] = (counts[id] || 0) + 1);
    for (const id in counts) {
      if (!msIdClaims[id] || msIdClaims[id].count < counts[id]) msIdClaims[id] = { groupKey: key, count: counts[id] };
    }
  }

  const activeMsIds = new Set();
  for (const key in groups) groups[key].msId = null;
  for (const id in msIdClaims) {
    const winnerGroupKey = msIdClaims[id].groupKey;
    if (!groups[winnerGroupKey].msId) {
      groups[winnerGroupKey].msId = id;
      activeMsIds.add(id);
    }
  }

  let createdOrders = 0, updatedOrders = 0, deletedOrders = 0, errorCount = 0;
  
  const supplierMap = getSupplierCodeMap();

  for (const key in groups) {
    const group = groups[key];
    const lookupName = group.supplierName.toLowerCase();
    const msCode = supplierMap[lookupName];
    
    let supplier = cache.suppliers[group.supplierName];
    if (supplier === undefined) {
      supplier = msCode ? findCounterpartyByName(msCode) : findCounterpartyByName(group.supplierName);
      cache.suppliers[group.supplierName] = supplier;
    }

    if (!supplier) { 
      const searchTarget = msCode ? `Код: ${msCode}` : `Имя: ${group.supplierName} (нет в справочнике)`;
      group.rows.forEach(r => updater.setStatus(r.rowIndex, `❌ Не найден в МС: [${group.supplierName}] -> ${searchTarget}`)); 
      errorCount++; 
      continue; 
    }

    if (!group.msId && group.specNum) {
      const existingMsId = findExistingOrderInMS(supplier.id, group.specNum);
      if (existingMsId) { group.msId = existingMsId; activeMsIds.add(existingMsId); }
    }

    const positionMap = {};
    let hasErrors = false;

    for (const item of group.rows) {
      let product = getOrCreateProduct(item.rowData);
      if (!product) { updater.setStatus(item.rowIndex, '❌ Ошибка товара'); hasErrors = true; continue; }
      
      const qty = parseNumber(item.rowData[COL.TOTAL_QTY]) || parseNumber(item.rowData[COL.QTY]) || 1;
      const priceKopecks = Math.round((parseNumber(item.rowData[COL.PRICE]) || 0) * 100);
      
      const pId = product.id;
      if (!positionMap[pId]) {
        positionMap[pId] = { quantity: 0, price: priceKopecks, assortment: { meta: product.meta } };
      }
      positionMap[pId].quantity += qty;
    }

    const positions = Object.values(positionMap);
    if (hasErrors || positions.length === 0) { errorCount++; continue; }

    let mergedPaymentData = extractPaymentData(group.rows[0].rowData);
    for (let i = 1; i < group.rows.length; i++) mergedPaymentData = mergePayments(mergedPaymentData, extractPaymentData(group.rows[i].rowData));

    const payload = buildPurchaseOrderPayload(group.rows[0].rowData, supplier, positions, mergedPaymentData);

    let res;
    if (group.msId) {
      res = msFetch('/entity/purchaseorder/' + encodeURIComponent(group.msId), 'put', payload);
      if (res.success) {
        group.rows.forEach(item => {
          updater.setValue(item.rowIndex, COL.MS_ID, group.msId);
          updater.setValue(item.rowIndex, COL.MS_LINK, res.data ? res.data.meta.uuidHref : item.rowData[COL.MS_LINK]);
          updater.setStatus(item.rowIndex, '🔄 Обновлено в МС (Синхронизировано)');
        });
        updatedOrders++;
      } else {
        group.rows.forEach(item => updater.setStatus(item.rowIndex, '❌ Ошибка обн.: ' + res.error));
        errorCount++;
      }
    } else {
      res = msFetch('/entity/purchaseorder', 'post', payload);
      if (res.success && res.data) {
        group.rows.forEach(item => {
          updater.setValue(item.rowIndex, COL.MS_ID, res.data.id);
          updater.setValue(item.rowIndex, COL.MS_LINK, res.data.meta.uuidHref || '');
          updater.setStatus(item.rowIndex, '✅ Создан новый заказ в МС');
        });
        createdOrders++;
      } else {
        group.rows.forEach(item => updater.setStatus(item.rowIndex, '❌ Ошибка созд.: ' + res.error));
        errorCount++;
      }
    }
  }

  const abandonedIds = [...originalMsIds].filter(id => !activeMsIds.has(id));
  for (const abandonedId of abandonedIds) {
    const delRes = msFetch('/entity/purchaseorder/' + encodeURIComponent(abandonedId), 'delete');
    if (delRes.success) {
      deletedOrders++;
    } else {
      const clearPayload = { positions: [], description: '⚠️ ЗАКАЗ АННУЛИРОВАН И ОБЪЕДИНЕН СКРИПТОМ С ДРУГИМ ЗАКАЗОМ. Товары удалены во избежание дублей.' };
      msFetch('/entity/purchaseorder/' + encodeURIComponent(abandonedId), 'put', clearPayload);
      deletedOrders++; 
    }
  }
  
  updater.flush();
  SpreadsheetApp.getUi().alert(`Синхронизация завершена!\n\n🆕 Создано: ${createdOrders}\n🔄 Обновлено: ${updatedOrders}\n🗑️ Удалено/Очищено брошенных: ${deletedOrders}\n❌ Ошибок: ${errorCount}`);
}

/**
 * Заполнение листа «Сводная» только с вкладок вида «Имя ММ/ГГ» или «Имя ММ/ГГГГ»
 * (например «Нина 07/26» — менеджер Нина, заказы на июль 2026).
 * Месяц для сортировки и разделителей берётся из имени вкладки; в колонку «Менеджер» пишется имя без даты.
 */
function syncManagerTabsToSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySh = ss.getSheetByName(SHEET_NAME);
  if (!summarySh) {
    ui.alert('Не найден лист «' + SHEET_NAME + '».');
    return;
  }
  const ex = syncManagerSummaryExcludeSet_();
  const sheets = ss.getSheets();
  /** @type {{ sheet: GoogleAppsScript.Spreadsheet.Sheet, meta: Object }[]} */
  const sources = [];
  let i;
  let sh;
  for (i = 0; i < sheets.length; i++) {
    sh = sheets[i];
    const nameTrim = safeString(sh.getName()).trim();
    if (nameTrim === safeString(SHEET_NAME).trim()) continue;
    if (syncManagerSheetIsExcluded_(nameTrim, ex)) continue;
    const tabMeta = syncManagerSummaryParseManagerMonthTab_(nameTrim);
    if (!tabMeta) continue;
    if (sh.getLastRow() < MANAGER_SUMMARY_SYNC_DATA_START_ROW) continue;
    sources.push({ sheet: sh, meta: tabMeta });
  }

  sources.sort(function (a, b) {
    return a.sheet.getIndex() - b.sheet.getIndex();
  });

  if (!sources.length) {
    ui.alert(
      'Нет вкладок для сборки в формате «Имя ММ/ГГ» или «Имя ММ/ГГГГ», например «Нина 07/26» или «Никита 08/2026».\n' +
        'Нужна строка с данными начиная с ' +
        MANAGER_SUMMARY_SYNC_DATA_START_ROW +
        '. Исключены «Сводная», справочники и системные листы из списка исключений.'
    );
    return;
  }

  let maxCol = Math.max(summarySh.getLastColumn(), COL.STATUS + 1);
  let widest = sources[0].sheet;
  for (i = 0; i < sources.length; i++) {
    const sht = sources[i].sheet;
    maxCol = Math.max(maxCol, sht.getLastColumn());
    if (sht.getLastColumn() > widest.getLastColumn()) {
      widest = sht;
    }
  }

  const headerSheet = widest;
  const headerRaw =
    headerSheet
      .getRange(
        MANAGER_SUMMARY_SYNC_HEADER_ROW,
        1,
        MANAGER_SUMMARY_SYNC_HEADER_ROW,
        headerSheet.getLastColumn()
      )
      .getValues()[0] || [];
  const header = syncManagerSummaryPad_(headerRaw, maxCol);

  const row1Existing =
    summarySh.getLastRow() >= 1
      ? summarySh.getRange(1, 1, 1, Math.max(summarySh.getLastColumn(), 1)).getValues()[0]
      : [];
  const row1 = syncManagerSummaryPad_(row1Existing, maxCol);

  /** @type {{ row: Object[], monthKey: string, labelRus: string, sourceIndex: number, rowIdx: number }[]} */
  const collected = [];
  let sIdx;
  for (sIdx = 0; sIdx < sources.length; sIdx++) {
    const src = sources[sIdx].sheet;
    const tabMeta = sources[sIdx].meta;
    const srcIndex = src.getIndex();
    const lc = Math.max(src.getLastColumn(), 1);
    const lr = src.getLastRow();
    const block =
      lr >= MANAGER_SUMMARY_SYNC_DATA_START_ROW
        ? src.getRange(MANAGER_SUMMARY_SYNC_DATA_START_ROW, 1, lr, lc).getValues()
        : [];

    let rIdx;
    for (rIdx = 0; rIdx < block.length; rIdx++) {
      const rowSrc = block[rIdx];
      if (syncManagerSummaryRowIgnorable_(rowSrc)) continue;
      const row = syncManagerSummaryPad_(rowSrc, maxCol);
      row[COL.MANAGER] = tabMeta.manager;
      collected.push({
        row: row,
        monthKey: tabMeta.monthKey,
        labelRus: tabMeta.labelRus,
        sourceIndex: srcIndex,
        rowIdx: rIdx
      });
    }
  }

  collected.sort(function (a, b) {
    if (a.monthKey !== b.monthKey) {
      return a.monthKey < b.monthKey ? -1 : 1;
    }
    if (a.sourceIndex !== b.sourceIndex) {
      return a.sourceIndex - b.sourceIndex;
    }
    return a.rowIdx - b.rowIdx;
  });

  const out = [row1, header];
  const sepMask = [];
  sepMask.push(false);
  sepMask.push(false);

  let dataLines = collected.length;
  let sepLines = 0;
  let prevMonthKey = null;
  let cIdx;
  for (cIdx = 0; cIdx < collected.length; cIdx++) {
    const entry = collected[cIdx];
    const rowPad = entry.row;
    const monthKey = entry.monthKey;

    if (prevMonthKey !== null && monthKey !== prevMonthKey) {
      out.push(
        syncManagerSummarySeparatorRow_(maxCol, {
          hasDate: true,
          labelRus: entry.labelRus
        })
      );
      sepMask.push(true);
      sepLines++;
    }

    out.push(rowPad);
    sepMask.push(false);
    prevMonthKey = monthKey;
  }

  summarySh.clearContents();
  summarySh.getRange(1, 1, out.length, maxCol).setValues(out);

  syncManagerSummaryApplyColumnFormats_(summarySh, out.length);
  syncManagerSummaryApplySeparatorStyles_(summarySh, sepMask, maxCol);

  ui.alert(
    'Сводная обновлена.\nВкладок «Имя ММ/ГГ»: ' +
      sources.length +
      '\nСтрок данных: ' +
      dataLines +
      '\nРазделителей между месяцами: ' +
      sepLines +
      '\nКолонок: ' +
      maxCol +
      '.'
  );
}

function syncManagerSummaryApplyColumnFormats_(summarySh, totalRows) {
  const dataStart = MANAGER_SUMMARY_SYNC_DATA_START_ROW;
  const rows = Math.max(0, totalRows - dataStart + 1);
  if (!rows) return;

  // «Объем» и «Вес» должны быть числами; иначе при старых форматах листа отображаются как дата.
  const volumeCol1 = COL.VOLUME + 1;
  const weightCol1 = COL.WEIGHT + 1;
  try {
    summarySh.getRange(dataStart, volumeCol1, rows, 1).setNumberFormat('0.000');
  } catch (e) {}
  try {
    summarySh.getRange(dataStart, weightCol1, rows, 1).setNumberFormat('0.000');
  } catch (e) {}
}

/**
 * Имя вкладки менеджерского месяца: текст, пробелы, ММ/ГГ или ММ/ГГГГ в конце.
 * @returns {{ manager: string, monthKey: string, labelRus: string, normalizedName: string }|null}
 */
function syncManagerSummaryParseManagerMonthTab_(sheetNameTrim) {
  const normalizedName = sheetNameTrim.replace(/\s+/g, ' ').trim();
  const mm = normalizedName.match(/^(.+?)\s+(\d{1,2})\/(\d{2,4})$/);
  if (!mm) return null;

  let managerName = syncManagerNormalizeManagerToken_(safeString(mm[1]));
  const monthNum = parseInt(mm[2], 10);
  let yearNum = parseInt(mm[3], 10);

  if (!managerName || monthNum < 1 || monthNum > 12) return null;
  if (yearNum < 100) yearNum += 2000;
  if (yearNum < 1990 || yearNum > 2100) return null;

  const mk = monthNum < 10 ? yearNum + '-0' + monthNum : yearNum + '-' + monthNum;
  const MONTHS_RU = [
    'январь',
    'февраль',
    'март',
    'апрель',
    'май',
    'июнь',
    'июль',
    'август',
    'сентябрь',
    'октябрь',
    'ноябрь',
    'декабрь'
  ];
  const labelRus = MONTHS_RU[monthNum - 1] + ' ' + yearNum;

  return {
    manager: managerName,
    monthKey: mk,
    labelRus: labelRus,
    normalizedName: normalizedName
  };
}

function syncManagerNormalizeManagerToken_(displayName) {
  return safeString(displayName).replace(/^[\s-_]+/, '').replace(/[\s-_]+$/, '').trim();
}

function syncManagerSummarySeparatorRow_(maxCol, nextRowMonthMeta) {
  const row = [];
  let k;
  for (k = 0; k < maxCol; k++) {
    row.push('');
  }
  if (nextRowMonthMeta.hasDate && safeString(nextRowMonthMeta.labelRus)) {
    row[0] = '▸ ' + nextRowMonthMeta.labelRus;
  } else {
    row[0] = '▸ (месяц не распознан по вкладке)';
  }
  return row;
}

function syncManagerSummaryApplySeparatorStyles_(sh, sepMask, maxCol) {
  let r;
  let rowNum;
  for (r = 0; r < sepMask.length; r++) {
    if (!sepMask[r]) continue;
    rowNum = r + 1;
    const range = sh.getRange(rowNum, 1, rowNum, maxCol);
    range
      .setBackground(MANAGER_SUMMARY_MONTH_SEP_BG)
      .setFontWeight('bold')
      .setBorder(
        false,
        false,
        true,
        false,
        false,
        false,
        '#9BB4D7',
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
  }
}

function syncManagerSummaryExcludeSet_() {
  const exactRaw = ['Закуплено', 'SYNC_LOG', 'Stock_Movements', 'Транзитный склад', 'Планирование отгрузок'];
  const exact = {};
  let i;
  for (i = 0; i < exactRaw.length; i++) {
    exact[exactRaw[i]] = true;
  }
  exact[safeString(SHEET_NAME).trim()] = true;
  const prefixes = ['Справочник', 'Проставление планов', 'Склады МС'];
  return { exact: exact, prefixes: prefixes };
}

function syncManagerSheetIsExcluded_(nameTrim, ex) {
  if (ex.exact[nameTrim]) return true;
  let p;
  for (p = 0; p < ex.prefixes.length; p++) {
    if (nameTrim.indexOf(ex.prefixes[p]) === 0) return true;
  }
  return false;
}

function syncManagerSummaryPad_(row, len) {
  const out = [];
  let k;
  for (k = 0; k < len; k++) {
    out.push(k < row.length ? row[k] : '');
  }
  return out;
}

/** Пропуск только полностью пустых строк и «хвостов» без полей заказа. */
function syncManagerSummaryRowIgnorable_(row) {
  if (!row || !row.length) return true;

  // Явные ключи заказа
  if (
    !syncManagerBlankish_(row[COL.WB_ARTICLE]) ||
    !syncManagerBlankish_(row[COL.SUPPLIER_ARTICLE]) ||
    !syncManagerBlankish_(row[COL.BARCODE]) ||
    !syncManagerBlankish_(row[COL.SPEC_NUMBER])
  ) {
    return false;
  }

  // Статусы/типы/даты — тоже считаем «данными»
  if (
    !syncManagerBlankish_(row[COL.SUPPLIER_MS]) ||
    !syncManagerBlankish_(row[COL.ORDER_STATUS]) ||
    !syncManagerBlankish_(row[COL.DELIVERY_TYPE]) ||
    !syncManagerBlankish_(row[COL.READY_DATE])
  ) {
    return false;
  }

  // Числа: считаем пустым, если null/пусто/0.
  if (
    !syncManagerZeroish_(row[COL.TOTAL_QTY]) ||
    !syncManagerZeroish_(row[COL.QTY]) ||
    !syncManagerZeroish_(row[COL.PRICE]) ||
    !syncManagerZeroish_(row[COL.AMOUNT]) ||
    !syncManagerZeroish_(row[COL.VOLUME]) ||
    !syncManagerZeroish_(row[COL.WEIGHT])
  ) {
    return false;
  }

  // Финальный фолбэк: если где-то есть «осмысленное» значение (кроме колонки менеджера) — оставляем.
  let j;
  for (j = 0; j < row.length; j++) {
    if (j === COL.MANAGER) continue; // на вкладках часто проставляют менеджера формулой — не считаем это признаком данных
    if (!syncManagerBlankish_(row[j])) return false;
  }
  return true;
}

function syncManagerBlankish_(v) {
  if (v == null) return true;
  if (typeof v === 'number') return !isFinite(v) || v === 0;
  const s = String(v).trim();
  if (!s) return true;
  const t = s.replace(/\u00A0/g, ' ').trim().toLowerCase();
  return (
    t === '-' ||
    t === '—' ||
    t === '–' ||
    t === 'нет' ||
    t === 'n/a' ||
    t === 'na' ||
    t === '0'
  );
}

function syncManagerZeroish_(v) {
  const n = parseNumber(v);
  return n == null || n === 0;
}

function findExistingOrderInMS(supplierId, specNum) {
  if (!specNum || !supplierId) return null;
  const filter = encodeURIComponent(`agent=https://api.moysklad.ru/api/remap/1.2/entity/counterparty/${supplierId}`);
  const res = msFetch(`/entity/purchaseorder?filter=${filter}`);
  if (res.success && res.data && res.data.rows) {
    for (const order of res.data.rows) {
      if (order.attributes) {
        const specAttr = order.attributes.find(a => a.id === MS_ATTR.SPEC_NUMBER);
        if (specAttr && specAttr.value === specNum) return order.id;
      }
    }
  }
  return null;
}

function extractPaymentData(row) {
  return {
    advanceSum: row[COL.ADVANCE_SUM], advancePlan: row[COL.ADVANCE_PLAN], advanceFact: row[COL.ADVANCE_FACT],
    balanceSum: row[COL.BALANCE_SUM], balancePlan: row[COL.BALANCE_PLAN], balanceFact: row[COL.BALANCE_FACT],
    deferSum: row[COL.DEFER_SUM], deferPlan: row[COL.DEFER_PLAN], deferFact: row[COL.DEFER_FACT]
  };
}

function mergePayments(main, extra) {
  const result = { ...main };
  for (const key in extra) { if (!result[key] && extra[key]) result[key] = extra[key]; }
  return result;
}

function getOrCreateProduct(row) {
  const barcode = safeString(row[COL.BARCODE]);
  const wbArticle = safeString(row[COL.WB_ARTICLE]);
  const supArticle = safeString(row[COL.SUPPLIER_ARTICLE]); 
  
  const searchKey = barcode || wbArticle || supArticle;
  if (!searchKey) return null;
  if (cache.products[searchKey]) return cache.products[searchKey];

  let product = null;
  if (barcode) product = findProductInMS(barcode);
  if (!product && wbArticle) product = findProductInMS(wbArticle);
  if (!product && supArticle) product = findProductInMS(supArticle);

  if (!product) {
    const name = supArticle || wbArticle || barcode || 'Неизвестный товар';
    const payload = { name: name, article: wbArticle || '' };
    if (barcode) payload.barcodes = [{ [barcode.length === 13 ? 'ean13' : 'code128']: barcode }];
    logInfo('Создание товара в МойСклад', { name: name, article: payload.article, barcode: barcode || '' });
    const res = msPost('/entity/product', payload);
    if (res.success && res.data) product = res.data;
    else logWarn('Не удалось создать товар в МойСклад', { error: res.error || 'Неизвестная ошибка', name: name });
  }
  
  if (product) cache.products[searchKey] = product;
  return product;
}

function findProductInMS(query) {
  const res = msFetch('/entity/product?search=' + encodeURIComponent(query) + '&limit=1', 'get');
  return (res.success && res.data && res.data.rows && res.data.rows.length > 0) ? res.data.rows[0] : null;
}

function getYuanMeta() {
  if (cache.currencyYuanMeta) return cache.currencyYuanMeta;
  const res = msFetch('/entity/currency', 'get');
  if (res.success && res.data && res.data.rows) {
    const yuan = res.data.rows.find(c => c.isoCode === 'CNY' || c.name.toLowerCase().includes('юань'));
    if (yuan) { cache.currencyYuanMeta = buildMeta('currency', yuan.id).meta; return cache.currencyYuanMeta; }
  }
  return null;
}

function buildPurchaseOrderPayload(row, supplier, positions, paymentData) {
  const payload = {
    organization: buildMeta('organization', CONFIG.MS_ORGANIZATION_ID),
    agent: { meta: supplier.meta },
    description: buildDescription(row),
    positions: positions
  };

  const yuanMeta = getYuanMeta();
  if (yuanMeta) payload.rate = { currency: { meta: yuanMeta } };

  const attr = buildAttributes(row, paymentData);
  if (attr.length > 0) payload.attributes = attr;

  const moment = parseDateToMS(row[COL.READY_DATE]);
  if (moment) payload.moment = moment;
  if (CONFIG.MS_STORE_ID) payload.store = buildMeta('store', CONFIG.MS_STORE_ID);
  return payload;
}

function buildMeta(type, id) { return { meta: { href: `https://api.moysklad.ru/api/remap/1.2/entity/${type}/${id}`, type: type, mediaType: 'application/json' } }; }
function buildAttrMeta(id) { return { meta: { href: `https://api.moysklad.ru/api/remap/1.2/entity/purchaseorder/metadata/attributes/${id}`, type: 'attributemetadata', mediaType: 'application/json' } }; }

function buildDescription(row) {
  const parts = [];
  if (row[COL.MANAGER]) parts.push('Менеджер: ' + row[COL.MANAGER]);
  if (row[COL.ORDER_STATUS]) parts.push('Статус из таблицы: ' + row[COL.ORDER_STATUS]);
  if (row[COL.SUPPLIER_NOTE]) parts.push('Фабрика: ' + row[COL.SUPPLIER_NOTE]);
  return parts.join('\n');
}

function buildAttributes(row, data) {
  const attr = [];
  pushStringAttr(attr, MS_ATTR.SPEC_NUMBER, row[COL.SPEC_NUMBER]);
  pushNumberAttr(attr, MS_ATTR.ADVANCE_SUM, data.advanceSum);
  pushTimeAttr(attr, MS_ATTR.ADVANCE_PLAN, data.advancePlan);
  pushTimeAttr(attr, MS_ATTR.ADVANCE_FACT, data.advanceFact);
  pushNumberAttr(attr, MS_ATTR.BALANCE_SUM, data.balanceSum);
  pushTimeAttr(attr, MS_ATTR.BALANCE_PLAN, data.balancePlan);
  pushTimeAttr(attr, MS_ATTR.BALANCE_FACT, data.balanceFact);
  pushNumberAttr(attr, MS_ATTR.DEFER_SUM, data.deferSum);
  pushTimeAttr(attr, MS_ATTR.DEFER_PLAN, data.deferPlan);
  pushTimeAttr(attr, MS_ATTR.DEFER_FACT, data.deferFact);
  return attr;
}

function findCounterpartyByName(query) {
  if (!query) return null;
  const qStr = query.toString().trim();
  const q = encodeURIComponent(qStr);
  const qExact = encodeURIComponent(qStr.replace(/"/g, '\\"')); 
  
  let res = msFetch('/entity/counterparty?filter=code=' + q, 'get');
  if (res.success && res.data && res.data.rows && res.data.rows.length > 0) return res.data.rows[0];
  
  res = msFetch('/entity/counterparty?filter=name=' + qExact, 'get');
  if (res.success && res.data && res.data.rows && res.data.rows.length > 0) return res.data.rows[0];

  res = msFetch('/entity/counterparty?search=' + q + '&limit=1', 'get');
  return (res.success && res.data && res.data.rows && res.data.rows.length > 0) ? res.data.rows[0] : null;
}

class BatchUpdater {
  constructor(sheet, data) { this.sheet = sheet; this.numRows = data.length; this.updates = {}; }
  setValue(rowIndex, colIndex, value) {
    if (!this.updates[colIndex]) this.updates[colIndex] = new Array(this.numRows).fill('');
    this.updates[colIndex][rowIndex] = value;
  }
  setStatus(rowIndex, text) { this.setValue(rowIndex, COL.STATUS, text); }
  flush() {
    const headerCell = this.sheet.getRange(2, COL.STATUS + 1);
    if (!headerCell.getValue()) headerCell.setValue('Статус скрипта');
    for (const colIndex in this.updates) {
      const colData = this.updates[colIndex].slice(2).map(v => [v !== '' ? v : null]);
      if (colData.length === 0) continue;
      const range = this.sheet.getRange(3, parseInt(colIndex) + 1, colData.length, 1);
      const merged = range.getValues().map((row, i) => [colData[i][0] !== null ? colData[i][0] : row[0]]);
      range.setValues(merged);
    }
  }
}

function getSupplierCodeMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Справочник поставщики и условия работы');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  
  const headers = data[0].map(h => safeString(h).toLowerCase().replace(/\s+/g, ' '));
  
  const nameCol = headers.findIndex(h => h === 'внутреннее название' || h === 'ид' || h === 'id');
  const codeCol = headers.findIndex(h => h === 'код поставщика из мс');
  
  if (nameCol === -1 || codeCol === -1) return {};
  
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const internalName = safeString(data[i][nameCol]).toLowerCase(); 
    const msCode = safeString(data[i][codeCol]);
    if (internalName && msCode) {
      map[internalName] = msCode;
    }
  }
  return map;
}