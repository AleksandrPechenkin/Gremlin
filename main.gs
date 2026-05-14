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
  MS_ID: 28, MS_LINK: 29, STATUS: 30, SHIPMENT_ID: 31
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
/** Префикс листов с историческими отгрузками (один лист или несколько: «История отгрузок», «История отгрузок 05/26»). */
const HISTORY_SHIPMENTS_SHEET_PREFIX = 'История отгрузок';
/** Текст-предупреждение, который записывается в строку 1 «Сводной» при пересборке. */
const SUMMARY_AUTO_NOTICE = '⚠️ Лист собирается автоматически (Меню → 🧩 Собрать «Сводная»). Ручной ввод данных — только в "История отгрузок" или в манерные вкладки «Имя ММ/ГГ».';

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
    .addItem('🔍 Диагностика заголовков активного таба', 'diagnoseActiveTabHeadersForSummary')
    .addToUi();
  if (typeof addSupplierInvoiceMenu_ === 'function') {
    addSupplierInvoiceMenu_(ui);
  }
  if (typeof addPaymentRegistryMenu_ === 'function') {
    addPaymentRegistryMenu_(ui);
  }
  if (typeof addSyncHubMenu_ === 'function' && isSyncHubMenuEnabled_()) {
    addSyncHubMenu_(ui);
  }
  // Меню «Планирование закупок» живёт в книге 03 (см. main_03.gs).
  // В книге 01 его быть не должно: листов/свойств планирования здесь нет, и менеджеры
  // нажимая пункты в «не той» книге, получали бы ошибки. См. README.txt и PROJECT_CONTEXT.md.
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
 * Заполнение листа «Сводная» из двух источников:
 *   1. Вкладки вида «Имя ММ/ГГ» / «Имя ММ/ГГГГ» — текущие заказы менеджеров;
 *      месяц и менеджер берутся из имени вкладки.
 *   2. Листы с префиксом «История отгрузок» (один или несколько) — исторические/уже отгруженные
 *      позиции; менеджер берётся из самой строки (колонка «Менеджер»), месяц — из колонки «Период (MM/YY)»
 *      или ETA/«Дата готовности» при отсутствии.
 *
 * Колонка «Рейс» (SHIPMENT_ID) гарантированно появляется в заголовке Сводной — она нужна как связь
 * со строками «Партии_в_рейсе» книги 05.
 */
function syncManagerTabsToSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySh = ss.getSheetByName(SHEET_NAME);
  if (!summarySh) {
    ui.alert('Не найден лист «' + SHEET_NAME + '».');
    return;
  }

  // Создаём базовый лист «История отгрузок», если его ещё нет — пользователь сразу видит куда вносить старые отгрузки.
  if (typeof ensureHistoryShipmentsSheet_ === 'function') {
    try {
      ensureHistoryShipmentsSheet_(ss);
    } catch (e) {
      // Не блокируем сборку, если по какой-то причине не удалось создать.
    }
  }

  const ex = syncManagerSummaryExcludeSet_();
  const sheets = ss.getSheets();
  /** @type {{ sheet: GoogleAppsScript.Spreadsheet.Sheet, kind: 'manager_month'|'history', meta: Object|null }[]} */
  const sources = [];
  let i;
  let sh;
  for (i = 0; i < sheets.length; i++) {
    sh = sheets[i];
    const nameTrim = safeString(sh.getName()).trim();
    if (nameTrim === safeString(SHEET_NAME).trim()) continue;
    if (syncManagerSheetIsExcluded_(nameTrim, ex)) continue;
    if (sh.getLastRow() < MANAGER_SUMMARY_SYNC_DATA_START_ROW) continue;

    if (syncManagerIsHistoryShipmentsSheet_(nameTrim)) {
      sources.push({ sheet: sh, kind: 'history', meta: null });
      continue;
    }

    const tabMeta = syncManagerSummaryParseManagerMonthTab_(nameTrim);
    if (!tabMeta) continue;
    sources.push({ sheet: sh, kind: 'manager_month', meta: tabMeta });
  }

  sources.sort(function (a, b) {
    return a.sheet.getIndex() - b.sheet.getIndex();
  });

  if (!sources.length) {
    ui.alert(
      'Нет вкладок для сборки.\n' +
        'Ожидаются вкладки «Имя ММ/ГГ» (например «Нина 07/26») или листы с префиксом «' +
        HISTORY_SHIPMENTS_SHEET_PREFIX +
        '».\n' +
        'Данные должны начинаться со строки ' +
        MANAGER_SUMMARY_SYNC_DATA_START_ROW +
        '. Исключены «Сводная», справочники и системные листы.'
    );
    return;
  }

  let maxCol = Math.max(summarySh.getLastColumn(), COL.STATUS + 1, COL.SHIPMENT_ID + 1);
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
  let headerKeys = header.map(function (h) {
    const canon = syncManagerCanonHeader_(h);
    return syncManagerHeaderKey_(canon);
  });
  let headerKeyToIdx = syncManagerBuildHeaderIndex_(headerKeys);

  // Гарантируем, что в заголовке Сводной есть колонка «Рейс» (SHIPMENT_ID).
  // Если её нет ни в widest, ни в существующей Сводной — добавляем в конец.
  const ensured = syncManagerSummaryEnsureRequiredKeys_(
    header,
    headerKeys,
    headerKeyToIdx,
    sources,
    summarySh,
    [{ key: 'shipment_id', defaultDisplay: 'Рейс' }]
  );
  if (ensured.changed) {
    maxCol = Math.max(maxCol, header.length);
    headerKeyToIdx = ensured.headerKeyToIdx;
  }

  const row1Existing =
    summarySh.getLastRow() >= 1
      ? summarySh.getRange(1, 1, 1, Math.max(summarySh.getLastColumn(), 1)).getValues()[0]
      : [];
  const row1 = syncManagerSummaryPad_(row1Existing, maxCol);
  row1[0] = SUMMARY_AUTO_NOTICE;

  /** @type {{ row: Object[], monthKey: string, labelRus: string, sourceIndex: number, rowIdx: number, sourceKind: string }[]} */
  const collected = [];
  let historySheets = 0;
  let managerSheets = 0;
  let sIdx;
  for (sIdx = 0; sIdx < sources.length; sIdx++) {
    const srcEntry = sources[sIdx];
    const src = srcEntry.sheet;
    const kind = srcEntry.kind;
    if (kind === 'history') historySheets++;
    else managerSheets++;

    const srcIndex = src.getIndex();
    const lc = Math.max(src.getLastColumn(), 1);
    const lr = src.getLastRow();
    const srcHeaderRaw =
      src
        .getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, MANAGER_SUMMARY_SYNC_HEADER_ROW, lc)
        .getValues()[0] || [];
    const srcHeaderKeys = srcHeaderRaw.map(function (h) {
      const canon = syncManagerCanonHeader_(h);
      return syncManagerHeaderKey_(canon);
    });
    const srcMap = syncManagerBuildHeaderIndex_(srcHeaderKeys);
    const block =
      lr >= MANAGER_SUMMARY_SYNC_DATA_START_ROW
        ? src.getRange(MANAGER_SUMMARY_SYNC_DATA_START_ROW, 1, lr, lc).getValues()
        : [];

    let rIdx;
    for (rIdx = 0; rIdx < block.length; rIdx++) {
      const rowSrc = block[rIdx];
      if (syncManagerSummaryRowIgnorable_(rowSrc)) continue;
      const row = syncManagerAlignRowToHeader_(rowSrc, headerKeys, srcMap);

      let monthKey;
      let labelRus;
      if (kind === 'manager_month') {
        const managerIdx =
          headerKeyToIdx['manager'] != null ? headerKeyToIdx['manager'] : COL.MANAGER;
        if (managerIdx != null && managerIdx < row.length) row[managerIdx] = srcEntry.meta.manager;
        monthKey = srcEntry.meta.monthKey;
        labelRus = srcEntry.meta.labelRus;
      } else {
        const histMeta = syncManagerHistoryRowMeta_(row, headerKeyToIdx);
        monthKey = histMeta.monthKey;
        labelRus = histMeta.labelRus;
      }

      collected.push({
        row: row,
        monthKey: monthKey,
        labelRus: labelRus,
        sourceIndex: srcIndex,
        rowIdx: rIdx,
        sourceKind: kind
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
          hasDate: monthKey !== 'zzz-unknown',
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

  syncManagerSummaryApplyColumnFormats_(summarySh, out.length, headerKeyToIdx);
  syncManagerSummaryApplySeparatorStyles_(summarySh, sepMask, maxCol);

  ui.alert(
    'Сводная обновлена.\n' +
      'Вкладок «Имя ММ/ГГ»: ' +
      managerSheets +
      '\nЛистов «История отгрузок*»: ' +
      historySheets +
      '\nСтрок данных: ' +
      dataLines +
      '\nРазделителей между месяцами: ' +
      sepLines +
      '\nКолонок: ' +
      maxCol +
      '.'
  );
}

/**
 * Возвращает true, если имя листа соответствует префиксу «История отгрузок» —
 * это либо точное имя, либо имя с суффиксом (например «История отгрузок 05/26»).
 */
function syncManagerIsHistoryShipmentsSheet_(sheetNameTrim) {
  const name = safeString(sheetNameTrim).trim();
  if (!name) return false;
  if (name === HISTORY_SHIPMENTS_SHEET_PREFIX) return true;
  // Допускаем суффикс через пробел/подчёркивание/дефис/слэш
  if (name.indexOf(HISTORY_SHIPMENTS_SHEET_PREFIX + ' ') === 0) return true;
  if (name.indexOf(HISTORY_SHIPMENTS_SHEET_PREFIX + '_') === 0) return true;
  if (name.indexOf(HISTORY_SHIPMENTS_SHEET_PREFIX + '-') === 0) return true;
  return false;
}

/**
 * Если в заголовке Сводной нет нужных ключей (например, «shipment_id»), добавляет их в конец.
 * Заголовок берётся либо из первого источника, где такой ключ встретился, либо из defaultDisplay.
 * @returns {{ changed: boolean, headerKeyToIdx: Object }}
 */
function syncManagerSummaryEnsureRequiredKeys_(header, headerKeys, headerKeyToIdx, sources, summarySh, required) {
  let changed = false;
  let i;
  let s;
  for (i = 0; i < required.length; i++) {
    const req = required[i];
    if (headerKeyToIdx[req.key] != null) continue;

    // Поиск display-имени в источниках
    let display = null;
    for (s = 0; s < sources.length; s++) {
      const sht = sources[s].sheet;
      const lc = sht.getLastColumn();
      if (lc <= 0) continue;
      const hdr =
        sht
          .getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, lc)
          .getValues()[0] || [];
      let j;
      for (j = 0; j < hdr.length; j++) {
        const cell = hdr[j];
        const canon = syncManagerCanonHeader_(cell);
        if (syncManagerHeaderKey_(canon) === req.key) {
          display = String(cell);
          break;
        }
      }
      if (display) break;
    }

    // Запасной вариант — display из текущей Сводной (вдруг пользователь руками добавил)
    if (!display) {
      const lc = summarySh.getLastColumn();
      if (lc > 0) {
        const hdr =
          summarySh
            .getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, lc)
            .getValues()[0] || [];
        let j;
        for (j = 0; j < hdr.length; j++) {
          const cell = hdr[j];
          const canon = syncManagerCanonHeader_(cell);
          if (syncManagerHeaderKey_(canon) === req.key) {
            display = String(cell);
            break;
          }
        }
      }
    }

    if (!display) display = req.defaultDisplay || req.key;

    header.push(display);
    headerKeys.push(req.key);
    headerKeyToIdx[req.key] = header.length - 1;
    changed = true;
  }

  return { changed: changed, headerKeyToIdx: headerKeyToIdx };
}

/**
 * Для строк из «История отгрузок» определяет монтх-кей и русский лейбл по содержимому строки:
 *   приоритет — колонка «Период (MM/YY)», затем ETA, затем «Дата готовности».
 */
function syncManagerHistoryRowMeta_(row, headerKeyToIdx) {
  const periodIdx = headerKeyToIdx ? headerKeyToIdx['period'] : null;
  const etaIdx = headerKeyToIdx ? headerKeyToIdx['eta'] : null;
  const readyIdx = headerKeyToIdx ? headerKeyToIdx['ready_date'] : null;

  let parsed = null;
  if (periodIdx != null && periodIdx < row.length) {
    parsed = syncManagerParsePeriodMmYy_(row[periodIdx]);
  }
  if (!parsed && etaIdx != null && etaIdx < row.length) {
    parsed = syncManagerParseMonthFromDate_(row[etaIdx]);
  }
  if (!parsed && readyIdx != null && readyIdx < row.length) {
    parsed = syncManagerParseMonthFromDate_(row[readyIdx]);
  }

  if (!parsed) {
    return { monthKey: 'zzz-unknown', labelRus: 'без даты' };
  }
  return parsed;
}

function syncManagerParsePeriodMmYy_(v) {
  if (v == null || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return syncManagerMonthFromMonthYear_(v.getMonth() + 1, v.getFullYear());
  }
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[\/.\-](\d{2,4})$/);
  if (!m) return null;
  const month = parseInt(m[1], 10);
  let year = parseInt(m[2], 10);
  if (year < 100) year += 2000;
  return syncManagerMonthFromMonthYear_(month, year);
}

function syncManagerParseMonthFromDate_(v) {
  if (v == null || v === '') return null;
  let d = null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    d = v;
  } else if (typeof v === 'string' && v.trim()) {
    const parsed = new Date(v);
    if (!isNaN(parsed.getTime())) d = parsed;
  }
  if (!d) return null;
  return syncManagerMonthFromMonthYear_(d.getMonth() + 1, d.getFullYear());
}

function syncManagerMonthFromMonthYear_(monthNum, yearNum) {
  if (!Number.isFinite(monthNum) || !Number.isFinite(yearNum)) return null;
  if (monthNum < 1 || monthNum > 12) return null;
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
  return { monthKey: mk, labelRus: MONTHS_RU[monthNum - 1] + ' ' + yearNum };
}

/**
 * Создаёт лист «История отгрузок» если его ещё нет. Шапка во 2-й строке, идентичная COL-раскладке
 * Сводной (включая колонку «Рейс»). Если лист уже существует — ничего не делает (шапку не трогаем,
 * чтобы не перезаписать пользовательские правки).
 *
 * Пользователь может создавать дополнительные листы по тому же шаблону:
 *   «История отгрузок 05/26», «История отгрузок TR-2026-0010» и т.п. — сборщик их подмешает все.
 */
function ensureHistoryShipmentsSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const name = HISTORY_SHIPMENTS_SHEET_PREFIX;
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  const header = syncManagerHistorySheetDefaultHeader_();
  sh.getRange(1, 1, 1, 1).setValues([
    [
      'Ручной ввод исторических/уже отгруженных позиций. Сборщик «Сводная» автоматически подмешивает строки этого листа (и других с префиксом «' +
        HISTORY_SHIPMENTS_SHEET_PREFIX +
        '») в Сводную.'
    ]
  ]);
  sh.getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, header.length).setValues([header]);
  sh.getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, header.length)
    .setFontWeight('bold')
    .setBackground('#FFE4B5');
  try {
    sh.setFrozenRows(MANAGER_SUMMARY_SYNC_HEADER_ROW);
  } catch (e) {}
  return sh;
}

/**
 * Стандартная шапка листа «История отгрузок». Имена колонок выровнены под COL —
 * чтобы пользователю было привычно. Сборщик в любом случае матчит по ключам, не по позиции,
 * так что пользователь может переименовать/переставить колонки.
 */
function syncManagerHistorySheetDefaultHeader_() {
  const header = new Array(COL.SHIPMENT_ID + 1).fill('');
  header[COL.WB_ARTICLE] = 'Артикул ВБ';
  header[COL.SUPPLIER_ARTICLE] = 'Артикул поставщика';
  header[COL.BARCODE] = 'ШК';
  header[COL.QTY] = 'Количество';
  header[COL.QTY_PER_BOX] = 'Количество в коробке';
  header[COL.BOXES] = 'Коробки';
  header[COL.TOTAL_QTY] = 'Итоговое количество';
  header[COL.MANAGER] = 'Менеджер';
  header[COL.SUPPLIER_NOTE] = 'Поставщик';
  header[COL.DELIVERY_TYPE] = 'Отгрузка через';
  header[COL.ORDER_STATUS] = 'Статус заказа';
  header[COL.SPEC_NUMBER] = 'Номер спецификации';
  header[COL.READY_DATE] = 'Дата готовности';
  header[COL.PRICE] = 'Цена';
  header[COL.AMOUNT] = 'Сумма';
  header[COL.VOLUME] = 'Объем';
  header[COL.WEIGHT] = 'Вес';
  header[COL.SUPPLIER_MS] = 'Поставщик МС';
  header[COL.LABELING] = 'Маркировка';
  header[COL.ADVANCE_SUM] = 'Аванс сумма';
  header[COL.ADVANCE_PLAN] = 'Аванс план';
  header[COL.ADVANCE_FACT] = 'Аванс факт';
  header[COL.BALANCE_SUM] = 'Остаток сумма';
  header[COL.BALANCE_PLAN] = 'Остаток план';
  header[COL.BALANCE_FACT] = 'Остаток факт';
  header[COL.DEFER_SUM] = 'Отсрочка сумма';
  header[COL.DEFER_PLAN] = 'Отсрочка план';
  header[COL.DEFER_FACT] = 'Отсрочка факт';
  header[COL.MS_ID] = 'MS_ID';
  header[COL.MS_LINK] = 'MS_LINK';
  header[COL.STATUS] = 'Статус скрипта';
  header[COL.SHIPMENT_ID] = 'Рейс';
  // Период (MM/YY) и ETA — отдельная колонка с правой стороны, чтобы корректно определять
  // месяц прибытия для книги 03 (она читает period/eta/ready_date в этом приоритете).
  header.push('Период (MM/YY)');
  header.push('Плановая дата поступления');
  return header;
}

function syncManagerCanonHeader_(h) {
  return String(h == null ? '' : h)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/[№#]/g, ' ')
    .replace(/[.,:;()/\\\[\]{}'"“”«»]/g, ' ')
    // Дефисы и нижние подчёркивания приводим к пробелам, чтобы
    // «Кол-во в коробке» и «Кол_во в коробке» канонизировались одинаково.
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Приводит «сырые» заголовки к устойчивым ключам, чтобы разные варианты
 * («Номер спецификации» vs «Спецификация») сопоставлялись корректно.
 */
function syncManagerHeaderKey_(canonHeader) {
  const h = String(canonHeader || '');
  if (!h) return '';

  // Спецификация
  if (
    h === 'номер спецификации' ||
    h === 'номер спецификации ' ||
    h === 'спецификация' ||
    h === 'спецификации' ||
    h === 'номерспецификации' ||
    h === 'спец' ||
    h === 'spec' ||
    h === 'spec number' ||
    // Часто менеджеры пишут «Номер спецификации /инвойса» или с «№»: после канона
    // получится «номер спецификации инвойса» и подобные варианты.
    h === 'номер спецификации инвойса' ||
    h === 'номер спецификации инвойс' ||
    h === 'номер инвойса' ||
    h === 'спецификация инвойса' ||
    h === 'спец инвойс' ||
    h === 'спец инвойса'
  ) {
    return 'spec_number';
  }

  // Рейс / SHIPMENT_ID (связь с книгой 05)
  if (
    h === 'рейс' ||
    h === 'shipment id' ||
    h === 'shipment_id' ||
    h === 'shipmentid' ||
    h === 'id рейса' ||
    h === 'id_рейса' ||
    h === 'idрейса' ||
    h === 'tr' ||
    h === 'трип' ||
    h === 'trip'
  ) {
    return 'shipment_id';
  }

  // Артикулы/ШК
  if (h === 'артикул вб' || h === 'артикул wb' || h === 'wb' || h === 'артикул') return 'wb_article';
  if (h === 'артикул поставщика' || h === 'supplier article') return 'supplier_article';
  if (h === 'шк' || h === 'barcode' || h === 'штрихкод') return 'barcode';

  // Количество/сумма
  if (h === 'итоговое количество' || h === 'количество итог' || h === 'total qty') return 'total_qty';

  // Количество в коробке (per-box / штук в коробке).
  // ВАЖНО: проверяется ДО общего «количество», чтобы не сматчить раньше времени.
  if (
    h === 'количество в коробке' ||
    h === 'кол во в коробке' ||
    h === 'штук в коробке' ||
    h === 'шт в коробке' ||
    h === 'единиц в коробке' ||
    h === 'в коробке' ||
    h === 'итого в коробке' ||
    h === 'pcs per box' ||
    h === 'qty per box' ||
    h === 'pcs box' ||
    h === 'per box'
  ) {
    return 'qty_per_box';
  }

  // Коробки (число коробок / коробов).
  if (
    h === 'коробки' ||
    h === 'кол во коробок' ||
    h === 'количество коробок' ||
    h === 'число коробок' ||
    h === 'итого коробок' ||
    h === 'итого коробки' ||
    h === 'итог коробок' ||
    h === 'всего коробок' ||
    // Вариант с «коробов» (от слова «короб»)
    h === 'кол во коробов' ||
    h === 'количество коробов' ||
    h === 'число коробов' ||
    h === 'итого коробов' ||
    h === 'всего коробов' ||
    h === 'boxes' ||
    h === 'box count' ||
    h === 'total boxes' ||
    h === 'ctn' ||
    h === 'ctns' ||
    h === 'cartons'
  ) {
    return 'boxes';
  }

  if (h === 'количество' || h === 'qty') return 'qty';
  if (h === 'сумма' || h === 'amount') return 'amount';
  if (h === 'цена' || h === 'price') return 'price';

  // Логистика
  if (h === 'дата готовности' || h === 'ready date') return 'ready_date';
  if (h === 'отгрузка через' || h === 'доставка' || h === 'delivery type') return 'ship_via';
  if (h === 'статус заказа' || h === 'статус' || h === 'order status' || h === 'status') return 'order_status';

  // ETA / плановая дата поступления (нужно для месяца прибытия в книге 03)
  if (
    h === 'плановая дата поступления' ||
    h === 'плановая дата прибытия' ||
    h === 'дата поступления' ||
    h === 'дата прибытия' ||
    h === 'eta' ||
    h === 'plan eta'
  ) {
    return 'eta';
  }

  // Период плана (MM/YY) — используется для группировки по месяцам в Сводной
  if (
    h === 'период мм гг' ||
    h === 'период мм yy' ||
    h === 'период mm yy' ||
    h === 'период mm гг' ||
    h === 'период' ||
    h === 'period' ||
    h === 'period mm yy'
  ) {
    return 'period';
  }

  // Менеджер (на исторических вкладках берётся из самой строки, а не из имени вкладки)
  if (h === 'менеджер' || h === 'manager') return 'manager';

  // Параметры
  if (h === 'обьем' || h === 'объем' || h === 'volume') return 'volume';
  if (h === 'вес' || h === 'weight') return 'weight';

  // По умолчанию — используем сам канон (чтобы уникальные заголовки тоже совпадали, если одинаковы).
  return h;
}

function syncManagerBuildHeaderIndex_(canonHeaders) {
  const idx = {};
  let c;
  for (c = 0; c < canonHeaders.length; c++) {
    const key = canonHeaders[c];
    if (!key) continue;
    if (idx[key] == null) idx[key] = c;
  }
  return idx;
}

/**
 * Диагностический пункт меню: показывает шапку активного таба и сопоставление
 * каждой колонки с устойчивым ключом, который использует сборщик «Сводной».
 *
 * Помогает быстро понять, почему колонка («Кол-во в коробке», «Спецификация /инвойса»
 * и т.п.) не подтягивается в Сводную: если в правой колонке вывода стоит просто
 * канон вместо знакомого ключа (например, «количество в коробке» вместо «qty_per_box») —
 * значит, для этого варианта названия нет алиаса. Добавляем алиас в syncManagerHeaderKey_
 * или переименовываем колонку в табе.
 */
function diagnoseActiveTabHeadersForSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const name = sh.getName();
  const lc = Math.max(sh.getLastColumn(), 1);
  if (lc < 1 || sh.getMaxRows() < MANAGER_SUMMARY_SYNC_HEADER_ROW) {
    ui.alert('Лист пуст или шапка не на ожидаемой строке (' + MANAGER_SUMMARY_SYNC_HEADER_ROW + ').');
    return;
  }
  const headerRaw =
    sh.getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, lc).getValues()[0] || [];

  const importantKeys = {
    wb_article: 'Артикул ВБ',
    supplier_article: 'Артикул поставщика',
    barcode: 'ШК',
    qty: 'Количество',
    qty_per_box: 'Количество в коробке',
    boxes: 'Коробки',
    total_qty: 'Итоговое количество',
    manager: 'Менеджер',
    ship_via: 'Отгрузка через',
    order_status: 'Статус заказа',
    spec_number: 'Номер спецификации',
    ready_date: 'Дата готовности',
    price: 'Цена',
    amount: 'Сумма',
    volume: 'Объем',
    weight: 'Вес',
    shipment_id: 'Рейс',
    period: 'Период (MM/YY)',
    eta: 'Плановая дата поступления'
  };

  const lines = [];
  lines.push('Таб: «' + name + '»');
  lines.push('Шапка читается из строки ' + MANAGER_SUMMARY_SYNC_HEADER_ROW + ', данные с ' + MANAGER_SUMMARY_SYNC_DATA_START_ROW + '.');
  lines.push('');
  lines.push('# | Колонка | Канон | Ключ');
  lines.push('--+---------+-------+-----');

  const matchedKeys = {};
  let i;
  for (i = 0; i < headerRaw.length; i++) {
    const orig = String(headerRaw[i] == null ? '' : headerRaw[i]);
    const canon = syncManagerCanonHeader_(orig);
    const key = canon ? syncManagerHeaderKey_(canon) : '';
    if (key) matchedKeys[key] = true;
    const col = i + 1;
    lines.push(
      String(col).padStart(2, ' ') +
        ' | ' + (orig || '⟨пусто⟩') +
        ' | ' + (canon || '⟨—⟩') +
        ' | ' + (key || '⟨не сматчилось⟩')
    );
  }

  const missing = [];
  const keys = Object.keys(importantKeys);
  for (i = 0; i < keys.length; i++) {
    if (!matchedKeys[keys[i]]) missing.push(keys[i] + ' (' + importantKeys[keys[i]] + ')');
  }
  if (missing.length) {
    lines.push('');
    lines.push('❌ Не найдены важные ключи (их колонки не подтянутся в Сводную):');
    missing.forEach(function (m) { lines.push('  • ' + m); });
  } else {
    lines.push('');
    lines.push('✅ Все важные ключи на месте — колонки подтянутся в Сводную корректно.');
  }

  ui.alert('🔍 Диагностика заголовков', lines.join('\n'), ui.ButtonSet.OK);
}

function syncManagerAlignRowToHeader_(rowSrc, targetHeaderCanon, srcMap) {
  const out = new Array(targetHeaderCanon.length).fill('');
  let c;
  for (c = 0; c < targetHeaderCanon.length; c++) {
    const key = targetHeaderCanon[c];
    const srcIdx = key ? srcMap[key] : null;
    if (srcIdx == null) continue;
    out[c] = srcIdx < rowSrc.length ? rowSrc[srcIdx] : '';
  }
  return out;
}

function syncManagerSummaryApplyColumnFormats_(summarySh, totalRows, headerKeyToIdx) {
  const dataStart = MANAGER_SUMMARY_SYNC_DATA_START_ROW;
  const rows = Math.max(0, totalRows - dataStart + 1);
  if (!rows) return;

  const amt0 = headerKeyToIdx ? headerKeyToIdx['amount'] : null;
  const price0 = headerKeyToIdx ? headerKeyToIdx['price'] : null;

  // Денежные поля — с тысячными разделителями (формат зависит от локали таблицы).
  if (amt0 != null) {
    try {
      summarySh.getRange(dataStart, amt0 + 1, rows, 1).setNumberFormat('#,##0.00');
    } catch (e) {}
  }
  if (price0 != null) {
    try {
      summarySh.getRange(dataStart, price0 + 1, rows, 1).setNumberFormat('#,##0.00');
    } catch (e) {}
  }

  // «Объем» и «Вес» должны быть числами; ищем колонки по ключам заголовка (порядок в вкладках разный).
  const vol0 = headerKeyToIdx ? headerKeyToIdx['volume'] : null;
  const w0 = headerKeyToIdx ? headerKeyToIdx['weight'] : null;
  if (vol0 != null) {
    try {
      summarySh.getRange(dataStart, vol0 + 1, rows, 1).setNumberFormat('0.000');
    } catch (e) {}
  }
  if (w0 != null) {
    try {
      summarySh.getRange(dataStart, w0 + 1, rows, 1).setNumberFormat('0.000');
    } catch (e) {}
  }
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