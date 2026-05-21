/**
 * Книга 05 — Себестоимость.
 * Рабочий модуль: подгрузка справочников + dry-run + боевой пересчет.
 */

const COST_CFG = {
  PROPS: {
    MASTER_REF_SPREADSHEET_ID: 'MASTER_REF_SPREADSHEET_ID',
    SYNC_BATCH_SIZE: 'SYNC_BATCH_SIZE',
    MASTER_PRODUCTS_SHEET_NAME: 'MASTER_PRODUCTS_SHEET_NAME',
    MASTER_SUPPLIERS_SHEET_NAME: 'MASTER_SUPPLIERS_SHEET_NAME'
  },
  DEFAULTS: {
    SYNC_BATCH_SIZE: 800,
    MASTER_PRODUCTS_SHEET_NAME: 'Справочник товары',
    MASTER_SUPPLIERS_SHEET_NAME: 'Справочник поставщики и условия'
  },
  SHEETS: {
    TRIPS: 'Рейсы',
    BATCHES: 'Партии_в_рейсе',
    EVENTS: 'События_рейса',
    EVENT_TYPES: 'Типы_событий',
    TRIP_EXPENSES: 'Затраты рейса',
    CUSTOMS: 'Таможенные платежи',
    ALLOCATION: 'Аллокация затрат',
    COST_SKU: 'Себестоимость SKU',
    CUSTOMS_FEE_RATES: 'Справочник_таможсбор',
    PRODUCTS_REF: 'Справочник товаров (05)',
    SUPPLIERS_REF: 'Справочник поставщики и условия',
    DECL_LINES: 'Декларации_строки',
    DECL_JOURNAL: 'Декларации_журнал',
    COST_SKU_FACT: 'Факт_себестоимость SKU',
    PLAN_FACT_COMPARE: 'Сверка план vs факт',
    RECEIPT_JOURNAL: 'Приемка_журнал',
    RECEIPT_LINES: 'Приемка_строки',
    RECEIPT_DISC: 'Расхождения_приемки',
    RECEIPT_MAPPING: 'Справочник_маппинг_3PL'
  },
  DECL_LINES_HEADER: [
    'Загрузка_ID', 'Загрузка_время', 'SHIPMENT_ID', 'Номер_ГТД', 'Дата_ГТД',
    '№_строки', 'Код_ТНВЭД', 'Описание', 'Qty', 'Вес_нетто', 'Валюта',
    'Таможенная_стоимость', 'Пошлина', 'НДС', 'Сбор', 'Итого_платежи',
    'Артикул_ВБ', 'Статус_сопоставления', 'Комментарий'
  ],
  DECL_JOURNAL_HEADER: [
    'Загрузка_ID', 'Загрузка_время', 'SHIPMENT_ID', 'Имя_файла', 'Пресет',
    'Строк_товаров', 'Пошлина_итого', 'НДС_итого', 'Сбор_итого', 'Предупреждения'
  ]
};

function addCostingMenu_(ui) {
  const batchesMenu = ui.createMenu('Партии_в_рейсе ← Сводная (01)')
    .addItem('Безопасный (dry-run)', 'costingRebuildBatchesSafeDryRun')
    .addItem('Безопасный (БОЕВОЙ)', 'costingRebuildBatchesSafeLive')
    .addSeparator()
    .addItem('Обновить количества (dry-run)', 'costingRebuildBatchesRefreshQtyDryRun')
    .addItem('Обновить количества (БОЕВОЙ)', 'costingRebuildBatchesRefreshQtyLive')
    .addSeparator()
    .addItem('Полная пересборка (dry-run)', 'costingRebuildBatchesFullDryRun')
    .addItem('Полная пересборка (БОЕВОЙ, двойное подтверждение)', 'costingRebuildBatchesFullLiveWithConfirm')
    .addSeparator()
    .addItem('Создать снимок «Партии_в_рейсе» сейчас', 'costingSnapshotBatchesMenu');

  ui.createMenu('💰 Себестоимость')
    .addItem('Проверка данных', 'costingHealthCheck_')
    .addItem('Подтянуть справочники из 04', 'costingSyncRefsFrom04_')
    .addItem('Обновить статусы рейсов из событий', 'costingSyncTripStatusesFromEvents_')
    .addSeparator()
    .addItem('Dry-run по рейсу (ввести SHIPMENT_ID)', 'costingDryRunByShipmentPrompt_')
    .addItem('Пересчитать себестоимость (расширенный)', 'rebuildCosting_')
    .addItem('Пересчитать себестоимость (сокращенный)', 'rebuildCostingSummary_')
    .addSeparator()
    .addSubMenu(batchesMenu)
    .addSeparator()
    .addItem('🛠 Подставить курсы ЦБ в пустые Курс_к_RUB', 'costingFillEmptyFxRatesMenu_')
    .addSeparator()
    .addItem('Загрузить декларацию (XML)', 'costingCustomsXmlOpenUploadDialog_')
    .addItem('🔗 Сопоставить строки декларации с SKU', 'costingMatchDeclarationPrompt_')
    .addSeparator()
    .addItem('🧪 Dry-run факт (по рейсу)', 'costingDryRunFactPrompt_')
    .addItem('✅ Пересчитать фактическую себестоимость', 'rebuildCostingFact_')
    .addItem('📊 Обновить лист «Сверка план vs факт»', 'costingComparePlanFactPrompt_')
    .addSeparator()
    .addItem('Создать недостающие листы', 'costingEnsureSheets_')
    .addToUi();
}

function costingHealthCheck_() {
  try {
    costingHealthCheckCore_();
    SpreadsheetApp.getUi().alert(
      '✅ Проверка пройдена',
      'Структура книги 05 корректна.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Проверка не пройдена',
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function costingEnsureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const required = [
    COST_CFG.SHEETS.BATCHES,
    COST_CFG.SHEETS.TRIP_EXPENSES,
    COST_CFG.SHEETS.CUSTOMS,
    COST_CFG.SHEETS.COST_SKU,
    COST_CFG.SHEETS.PRODUCTS_REF,
    COST_CFG.SHEETS.SUPPLIERS_REF
  ];
  let created = 0;
  required.forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
      created++;
    }
  });
  costingEnsureDeclLinesSheet_(ss);
  costingEnsureDeclJournalSheet_(ss);
  costingEnsureCustomsFactColumns_(ss);
  costingEnsureFactCostSkuHeader_(ss);
  costingEnsurePlanFactCompareSheet_(ss);
  if (typeof costingReceiptEnsureSheets_ === 'function') {
    costingReceiptEnsureSheets_(ss);
  }
  SpreadsheetApp.getUi().alert(
    'Готово',
    created ? ('Создано листов: ' + created) : 'Все обязательные листы уже существуют.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function costingSyncRefsFrom04_() {
  costingHealthCheckCore_();
  const sourceId = costingGetRequiredProp_(COST_CFG.PROPS.MASTER_REF_SPREADSHEET_ID);
  const sourceSs = SpreadsheetApp.openById(sourceId);
  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
  const productsStats = costingMergeProductsBaseFrom04_(sourceSs, targetSs);
  SpreadsheetApp.getUi().alert(
    '✅ Справочники обновлены',
    productsStats,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function costingDryRun_() {
  try {
    costingHealthCheckCore_();
    const report = costingBuildSkuCostRows_(null, { writeTripExpenseFx: false });
    SpreadsheetApp.getUi().alert(
      '🧪 Dry-run себестоимости',
      'Будет записано строк: ' + report.rows.length + '\n' +
      'SKU: ' + report.skuCount + '\n' +
      'Рейсов: ' + report.shipmentCount,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Dry-run завершен с ошибкой',
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

function costingDryRunByShipmentPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Тестовый прогон', 'Введите SHIPMENT_ID (например TR-2026-0008):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  if (!shipmentId) throw new Error('SHIPMENT_ID не указан.');
  const report = costingBuildSkuCostRows_(shipmentId, { writeTripExpenseFx: false });
  ui.alert(
    '🧪 Dry-run по рейсу',
    'SHIPMENT_ID: ' + shipmentId + '\n' +
    'Будет записано строк: ' + report.rows.length + '\n' +
    'SKU: ' + report.skuCount + '\n' +
    'Рейсов: ' + report.shipmentCount,
    ui.ButtonSet.OK
  );
}

function rebuildCosting_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Пересчет себестоимости (расширенный)',
    'Введите SHIPMENT_ID для точечного пересчета или оставьте пустым для пересчета всей базы:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  if (shipmentId) {
    rebuildCostingByShipment_(shipmentId);
    return;
  }
  try {
    costingHealthCheckCore_();
    const report = costingBuildSkuCostRows_();
    const sh = costingGetSheetByRole_(SpreadsheetApp.getActiveSpreadsheet(), 'COST_SKU');
    sh.clearContents();
    const extraCols = report.additionalExpenseColumns || [];
    const header = [costingBuildFullOutputHeader_(extraCols)];
    sh.getRange(1, 1, 1, header[0].length).setValues(header);
    if (report.rows.length) {
      sh.getRange(2, 1, report.rows.length, header[0].length).setValues(report.rows);
    }
    const fxMsg = report.tripExpenseFxFilled
      ? ('\nПодставлено курсов в «Затраты рейса» (PLAN): ' + report.tripExpenseFxFilled)
      : '';
    SpreadsheetApp.getUi().alert(
      '✅ Пересчет завершен',
      'Записано строк: ' + report.rows.length + '\n' +
      'SKU: ' + report.skuCount + '\n' +
      'Рейсов: ' + report.shipmentCount + fxMsg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка пересчета',
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

function rebuildCostingByShipment_(shipmentId) {
  const id = String(shipmentId || '').trim();
  if (!id) throw new Error('SHIPMENT_ID не указан.');
  costingHealthCheckCore_();
  const report = costingBuildSkuCostRows_(id);
  const sh = costingGetSheetByRole_(SpreadsheetApp.getActiveSpreadsheet(), 'COST_SKU');
  const targetHeader = costingBuildFullOutputHeader_(report.additionalExpenseColumns || []);
  costingUpsertShipmentRowsToCostSku_(sh, id, report.rows, targetHeader);
  SpreadsheetApp.getUi().alert(
    '✅ Пересчет по рейсу завершен',
    'SHIPMENT_ID: ' + id + '\n' +
    'Обновлено строк: ' + report.rows.length + '\n' +
    'SKU: ' + report.skuCount,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function costingBuildFullOutputHeader_(extraCols) {
  return [
    'Расчет_время',
    'SHIPMENT_ID',
    'Артикул_ВБ',
    'Артикул_поставщика',
    'Код_ТНВЭД',
    'Qty',
    'Закупка_руб',
    'Таможенная_стоимость_база_RUB',
    'Расходы_до_границы_RUB',
    'Фрахт_до_РФ_RUB',
    'Прочая_логистика_RUB',
    'Пошлина_RUB',
    'НДС_RUB',
    'Таможенный_сбор_RUB',
    'Прочие_расходы_RUB'
  ].concat(extraCols || []).concat([
    'Итого_доп_расходы_RUB',
    'Итого_партия_руб',
    'Себестоимость_ед_руб',
    'Курс_средневзв_RUB',
    'FX_источник'
  ]);
}

function costingUpsertShipmentRowsToCostSku_(sheet, shipmentId, incomingRows, incomingHeader) {
  const data = sheet.getDataRange().getValues();
  const existingHeader = data.length ? data[0] : [];
  const targetHeader = existingHeader.length ? existingHeader.slice() : incomingHeader.slice();
  const targetNorm = targetHeader.map(function (h) { return costingNorm_(h); });

  // Якорь для вставки новых per-article колонок — перед «Итого_доп_расходы_RUB»,
  // чтобы новые статьи (например, «Услуги агента по оплате») вставали среди других
  // прочих расходов, а не уезжали в самый конец таблицы.
  const anchorCandidates = [
    'Итого_доп_расходы_RUB', 'Итого доп расходы RUB', 'Итого_доп_расходы', 'Итого доп расходы'
  ].map(function (h) { return costingNorm_(h); });
  function findAnchorIdx() {
    for (let i = 0; i < targetNorm.length; i++) {
      if (anchorCandidates.indexOf(targetNorm[i]) !== -1) return i;
    }
    return -1;
  }

  for (let i = 0; i < incomingHeader.length; i++) {
    const nh = costingNorm_(incomingHeader[i]);
    if (targetNorm.indexOf(nh) === -1) {
      const anchor = findAnchorIdx();
      if (anchor !== -1) {
        targetHeader.splice(anchor, 0, incomingHeader[i]);
        targetNorm.splice(anchor, 0, nh);
      } else {
        targetHeader.push(incomingHeader[i]);
        targetNorm.push(nh);
      }
    }
  }

  const oldHeaderNorm = existingHeader.map(function (h) { return costingNorm_(h); });
  const oldShipIdx = costingFirstIdx_(costingHeaderMap_(existingHeader), ['SHIPMENT_ID', 'Shipment_ID', 'Рейс']);
  const keptRows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const oldShipment = oldShipIdx != null ? String(row[oldShipIdx] || '').trim() : '';
    if (oldShipIdx != null && oldShipment === shipmentId) continue;
    const expanded = new Array(targetHeader.length).fill('');
    for (let c = 0; c < oldHeaderNorm.length; c++) {
      const toIdx = targetNorm.indexOf(oldHeaderNorm[c]);
      if (toIdx !== -1) expanded[toIdx] = row[c];
    }
    keptRows.push(expanded);
  }

  const inHeaderNorm = incomingHeader.map(function (h) { return costingNorm_(h); });
  const newRows = incomingRows.map(function (src) {
    const out = new Array(targetHeader.length).fill('');
    for (let c = 0; c < inHeaderNorm.length; c++) {
      const toIdx = targetNorm.indexOf(inHeaderNorm[c]);
      if (toIdx !== -1) out[toIdx] = src[c];
    }
    return out;
  });

  const allRows = [targetHeader].concat(keptRows, newRows);
  sheet.clearContents();
  sheet.getRange(1, 1, allRows.length, targetHeader.length).setValues(allRows);
  sheet.getRange(1, 1, 1, targetHeader.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function rebuildCostingSummary_() {
  try {
    costingHealthCheckCore_();
    const report = costingBuildSkuCostRows_();
    const sh = costingGetSheetByRole_(SpreadsheetApp.getActiveSpreadsheet(), 'COST_SKU');
    sh.clearContents();
    const header = [[
      'Расчет_время',
      'SHIPMENT_ID',
      'Артикул_ВБ',
      'Артикул_поставщика',
      'Qty',
      'Итого_партия_руб',
      'Себестоимость_ед_руб'
    ]];
    sh.getRange(1, 1, 1, header[0].length).setValues(header);
    if (report.rows.length) {
      const compact = report.rows.map(function (r) {
        return [r[0], r[1], r[2], r[3], r[5], r[r.length - 2], r[r.length - 1]];
      });
      sh.getRange(2, 1, compact.length, header[0].length).setValues(compact);
    }
    SpreadsheetApp.getUi().alert(
      '✅ Сокращенный расчет завершен',
      'Записано строк: ' + report.rows.length + '\n' +
      'SKU: ' + report.skuCount + '\n' +
      'Рейсов: ' + report.shipmentCount,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Ошибка сокращенного расчета',
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

function costingBuildSkuCostRows_(shipmentFilter, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = costingFindSheetByRole_(ss, 'ALLOCATION');
  if (!sh) {
    return costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter, options);
  }
  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { rows: [], skuCount: 0, shipmentCount: 0 };
  }
  const header = data[0];
  const idxShipment = costingFindCol_(header, ['SHIPMENT_ID', 'Shipment_ID', 'Рейс']);
  const idxSku = costingFindCol_(header, ['Артикул_ВБ', 'Артикул ВБ', 'WB_ARTICLE', 'Артикул']);
  const idxSupplierArticle = costingFindColOptional_(header, ['Артикул_поставщика', 'Артикул поставщика', 'Supplier Article']);
  const idxTnved = costingFindColOptional_(header, ['Код_ТНВЭД', 'Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  const idxQty = costingFindCol_(header, ['Qty', 'QTY', 'Количество']);
  const idxPurchase = costingFindCol_(header, ['Закупка_руб', 'Закупка руб', 'Закупка']);
    const idxCustomsBase = costingFindColOptional_(header, ['Таможенная_стоимость_база_RUB', 'Таможенная стоимость база RUB']);
    const idxFreight = costingFindColOptional_(header, ['Фрахт_до_РФ_RUB', 'Фрахт до РФ RUB', 'Логистика_PRE_RUB', 'Фрахт_RUB']);
  const idxLogOther = costingFindColOptional_(header, ['Прочая_логистика_RUB', 'Логистика_POST_RUB']);
  const idxDuty = costingFindColOptional_(header, ['Пошлина_RUB', 'Пошлина RUB', 'Таможня_RUB', 'Таможня RUB']);
  const idxVat = costingFindColOptional_(header, ['НДС_RUB', 'НДС RUB']);
  const idxFee = costingFindColOptional_(header, ['Таможенный_сбор_RUB', 'Таможенный сбор RUB']);
  const idxOther = costingFindColOptional_(header, ['Прочие_расходы_RUB', 'Прочие расходы RUB']);

  const now = new Date();
  const rows = [];
  const skuSet = {};
  const shipmentSet = {};

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const shipmentId = String(r[idxShipment] || '').trim();
    const sku = String(r[idxSku] || '').trim();
    if (!shipmentId || !sku) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;

    const qty = costingToNumber_(r[idxQty]);
    if (!qty) continue;

    const purchaseRub = costingToNumber_(r[idxPurchase]);
    const freightRub = idxFreight != null ? costingToNumber_(r[idxFreight]) : 0;
    const logOtherRub = idxLogOther != null ? costingToNumber_(r[idxLogOther]) : 0;
    const dutyRub = idxDuty != null ? costingToNumber_(r[idxDuty]) : 0;
    const vatRub = idxVat != null ? costingToNumber_(r[idxVat]) : 0;
    const feeRub = idxFee != null ? costingToNumber_(r[idxFee]) : 0;
    const otherRub = idxOther != null ? costingToNumber_(r[idxOther]) : 0;
    const customsBaseRub = idxCustomsBase != null ? costingToNumber_(r[idxCustomsBase]) : (purchaseRub + freightRub + otherRub);
    const extraTotal = freightRub + logOtherRub + dutyRub + vatRub + feeRub + otherRub;
    const total = purchaseRub + extraTotal;
    const unitCost = total / qty;

    rows.push([
      now,
      shipmentId,
      sku,
      idxSupplierArticle != null ? String(r[idxSupplierArticle] || '').trim() : '',
      idxTnved != null ? String(r[idxTnved] || '').trim() : '',
      qty,
      purchaseRub,
      customsBaseRub,
      freightRub,
      logOtherRub,
      dutyRub,
      vatRub,
      feeRub,
      otherRub,
      '',
      extraTotal,
      total,
      unitCost
    ]);
    skuSet[sku] = true;
    shipmentSet[shipmentId] = true;
  }

  return {
    rows: rows,
    skuCount: Object.keys(skuSet).length,
    shipmentCount: Object.keys(shipmentSet).length,
    additionalExpenseColumns: []
  };
}

function costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter, options) {
  options = options || {};
  const scenarioMode = String(options.scenarioMode || 'PLAN').toUpperCase();
  const isFact = scenarioMode === 'FACT';
  const writeTripExpenseFx = options.writeTripExpenseFx !== false;
  const batchesSh = costingGetSheetByRole_(ss, 'BATCHES');
  const expensesSh = costingGetSheetByRole_(ss, 'TRIP_EXPENSES');
  const batchesData = batchesSh.getDataRange().getValues();
  const expensesData = expensesSh.getDataRange().getValues();
  if (batchesData.length < 2) return { rows: [], skuCount: 0, shipmentCount: 0 };

  const bh = batchesData[0];
  const idxShipment = costingFindCol_(bh, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindCol_(bh, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const idxSupplierArticle = costingFindColOptional_(bh, ['Артикул поставщика', 'Артикул_поставщика']);
  const idxQty = costingFindCol_(bh, ['Количество', 'Qty', 'QTY']);
  const idxAllocQty = costingFindColOptional_(bh, [
    'Qty_для_аллокации', 'Qty для аллокации', 'Количество_для_аллокации'
  ]);
  const idxCost = costingFindColOptional_(bh, ['Стоимость', 'Закупка_руб', 'Закупка руб']);
  const idxPrice = costingFindColOptional_(bh, ['Цена', 'Цена_RUB', 'Цена руб']);
  const idxCurrency = costingFindColOptional_(bh, ['Валюта', 'Currency']);
  const idxRate = costingFindColOptional_(bh, ['Курс_к_RUB', 'Курс к RUB', 'Курс', 'Rate']);
  const idxSpec = costingFindColOptional_(bh, [
    'Номер спецификации', 'Номер_спецификации', 'Спецификация', '№ спецификации', '№спецификации',
    'Spec', 'Spec number', 'Номер инвойса', 'Инвойс'
  ]);
  const idxTnved = costingFindColOptional_(bh, ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  const idxShipVia = costingFindColOptional_(bh, ['Отгрузка через', 'Отгрузка_через', 'Поставщик', 'Supplier']);
  const idxVolume = costingFindColOptional_(bh, ['Объем', 'Объём', 'Volume']);
  const idxWeight = costingFindColOptional_(bh, ['Вес', 'Weight']);
  const dutyRulesByTnved = costingLoadDutyRulesByTnved_(ss);
  const vatRatesByTnved = costingLoadVatRatesByTnved_(ss);
  const customsFeeRules = costingLoadCustomsFeeRules_(ss);
  const skuToTnved = costingLoadSkuToTnvedMap_(ss);
  const paymentLoad = costingLoadPaymentTranchesMapSafe_(shipmentFilter);
  const paymentMap = paymentLoad.map || {};
  const fxCache = costingMakeFxCache_(new Date());
  const fxMissing = [];

  const expensePoolsByShipment = {};
  const expensePoolsByShipmentSupplier = {};
  const skippedFactExpenses = [];
  const tripExpenseFxUpdates = [];
  let tripExpenseFxFilled = 0;
  // Какие ALLOC_BASE реально использовались — для последующей валидации (где у SKU нулевая база — ошибка).
  const baseUsageByShipment = {};
  const baseUsageByShipmentSupplier = {};
  if (expensesData.length > 1) {
    const eh = expensesData[0];
    const eShipment = costingFindCol_(eh, ['SHIPMENT_ID', 'ID_рейса']);
    const eScenario = costingFindColOptional_(eh, ['Сценарий']);
    const eInSku = costingFindColOptional_(eh, ['IN_SKU_COST', 'IN SKU COST']);
    const eSumRub = costingFindColOptional_(eh, ['Сумма_RUB', 'Сумма RUB']);
    const eSum = costingFindColOptional_(eh, [
      'Сумма_в_валюте', 'Сумма в валюте', 'Сумма', 'Amount'
    ]);
    const eCurrency = costingFindColOptional_(eh, ['Валюта', 'Currency']);
    const eRate = costingFindColOptional_(eh, ['Курс_к_RUB', 'Курс к RUB', 'Курс', 'Rate']);
    const ePayDate = costingFindColOptional_(eh, [
      'Дата оплаты', 'Дата_оплаты', 'Дата оплаты (факт)', 'Дата_оплаты_факт'
    ]);
    const eExpenseDate = costingFindColOptional_(eh, [
      'Дата_расхода', 'Дата расхода', 'Дата_затраты', 'Дата затраты',
      'Дата_события', 'Дата события', 'Дата'
    ]);
    const eOnlyShipVia = costingFindColOptional_(eh, ['Только_отгрузка_через', 'Только отгрузка через']);
    const eAllocBase = costingFindColOptional_(eh, ['ALLOC_BASE', 'Alloc base', 'AllocBase', 'База_аллокации', 'База аллокации']);
    const expenseFxCols = {
      sumRub: eSumRub,
      sum: eSum,
      currency: eCurrency,
      rate: eRate,
      payDate: ePayDate,
      expenseDate: eExpenseDate
    };
    for (let i = 1; i < expensesData.length; i++) {
      const r = expensesData[i];
      const shipmentId = String(r[eShipment] || '').trim();
      if (!shipmentId) continue;
      if (shipmentFilter && shipmentId !== shipmentFilter) continue;
      const article = costingCellText_(r, costingFindColOptional_(eh, ['Статья_затрат', 'Статья затрат', 'Статья']));
      const eventType = costingCellText_(r, costingFindColOptional_(eh, ['Тип_событий', 'Тип события']));
      const scope = costingCellText_(r, costingFindColOptional_(eh, ['COST_SCOPE', 'Scope']));
      const cat = costingClassifyExpense_(article, eventType, scope);
      let expenseScenario = '';
      if (eScenario != null) {
        expenseScenario = String(r[eScenario] || '').trim().toUpperCase();
        if (isFact) {
          // Фактический пересчёт: только строки FACT. Таможня по SKU — из «Таможенные платежи» (XML).
          if (expenseScenario !== 'FACT') continue;
        } else if (expenseScenario === 'FACT') {
          continue;
        }
      }
      const fxDetail = costingResolveTripExpenseRubDetail_(r, expenseFxCols, fxCache, fxMissing, {
        shipmentId: shipmentId,
        sku: '(затраты рейса)',
        supplierArticle: article
      });
      if (!(fxDetail.sumRub > 0)) {
        if (isFact && expenseScenario === 'FACT') {
          skippedFactExpenses.push(
            shipmentId + ' / ' + (article || '(без статьи)') +
            ': не удалось получить сумму в RUB (Сумма_в_валюте, валюта, Дата_расхода; курс ЦБ, если Курс_к_RUB пуст)'
          );
        }
        continue;
      }
      const sumRub = fxDetail.sumRub;
      if (writeTripExpenseFx && (fxDetail.fillRate || fxDetail.fillSumRub)) {
        tripExpenseFxUpdates.push({
          row: i,
          rateCol: eRate,
          sumRubCol: eSumRub,
          rate: fxDetail.rate,
          sumRub: fxDetail.sumRub,
          fillRate: fxDetail.fillRate,
          fillSumRub: fxDetail.fillSumRub
        });
      }
      const onlyShipViaRaw = eOnlyShipVia != null ? String(r[eOnlyShipVia] || '') : '';
      const onlyShipViaKey = costingNormalizeSupplierKey_(onlyShipViaRaw);
      const allocBase = costingNormalizeAllocBase_(eAllocBase != null ? r[eAllocBase] : '');

      // Для логистики/прочих расходов учитываем только IN_SKU_COST=1.
      // Для таможенных компонентов (пошлина/НДС/таможсбор) берём факт даже без IN_SKU_COST.
      if (cat !== 'vat' && cat !== 'customsFee' && cat !== 'duty') {
        if (eInSku != null) {
          const inSkuRaw = String(r[eInSku] == null ? '' : r[eInSku]).trim().toUpperCase();
          if (!(inSkuRaw === '1' || inSkuRaw === 'TRUE' || inSkuRaw === 'ДА')) continue;
        }
      }

      if (!expensePoolsByShipment[shipmentId]) {
        expensePoolsByShipment[shipmentId] = costingMakeEmptyExpensePool_();
      }

      let targetPool = expensePoolsByShipment[shipmentId];
      let usageBucket = baseUsageByShipment[shipmentId] || (baseUsageByShipment[shipmentId] = {});
      if (onlyShipViaKey) {
        if (!expensePoolsByShipmentSupplier[shipmentId]) expensePoolsByShipmentSupplier[shipmentId] = {};
        if (!expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey]) {
          expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey] = costingMakeEmptyExpensePool_();
        }
        targetPool = expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey];
        if (!baseUsageByShipmentSupplier[shipmentId]) baseUsageByShipmentSupplier[shipmentId] = {};
        if (!baseUsageByShipmentSupplier[shipmentId][onlyShipViaKey]) baseUsageByShipmentSupplier[shipmentId][onlyShipViaKey] = {};
        usageBucket = baseUsageByShipmentSupplier[shipmentId][onlyShipViaKey];
      }
      usageBucket[allocBase] = true;

      if (cat === 'freight') costingAddToBaseMap_(targetPool.freightRub, allocBase, sumRub);
      else if (cat === 'duty') costingAddToBaseMap_(targetPool.dutyRub, allocBase, sumRub);
      else if (cat === 'vat') costingAddToBaseMap_(targetPool.vatRub, allocBase, sumRub);
      else if (cat === 'customsFee') costingAddToBaseMap_(targetPool.customsFeeRub, allocBase, sumRub);
      else if (cat === 'logisticsOther') costingAddToBaseMap_(targetPool.logisticsOtherRub, allocBase, sumRub);
      else if (costingIsPostBorder_(scope)) costingAddToBaseMap_(targetPool.otherPostRub, allocBase, sumRub);
      else costingAddToBaseMap_(targetPool.otherPreRub, allocBase, sumRub);
      if (cat !== 'freight' && cat !== 'duty' && cat !== 'vat' && cat !== 'customsFee' && cat !== 'logisticsOther') {
        const key = article || '(без статьи)';
        if (!targetPool.otherByArticle[key]) targetPool.otherByArticle[key] = {};
        targetPool.otherByArticle[key][allocBase] = (targetPool.otherByArticle[key][allocBase] || 0) + sumRub;
      }
    }
    if (writeTripExpenseFx && tripExpenseFxUpdates.length) {
      tripExpenseFxFilled = costingApplyTripExpenseFxFill_(expensesSh, tripExpenseFxUpdates);
    }
  }

  // Плановый контур: агрегат таможни по рейсу (без FACT). Факт — по SKU из «Таможенные платежи» ниже.
  if (!isFact) {
    const customsPoolsByShipment = costingLoadCustomsPoolsByShipment_(ss, shipmentFilter, 'PLAN');
    const customsPoolKeys = Object.keys(customsPoolsByShipment);
    for (let i = 0; i < customsPoolKeys.length; i++) {
      const shipmentId = customsPoolKeys[i];
      if (!expensePoolsByShipment[shipmentId]) {
        expensePoolsByShipment[shipmentId] = costingMakeEmptyExpensePool_();
      }
      const target = expensePoolsByShipment[shipmentId];
      const src = customsPoolsByShipment[shipmentId];
      if (src.dutyRub) costingAddToBaseMap_(target.dutyRub, 'QTY', src.dutyRub);
      if (src.vatRub) costingAddToBaseMap_(target.vatRub, 'QTY', src.vatRub);
      if (src.dutyRub || src.vatRub) {
        if (!baseUsageByShipment[shipmentId]) baseUsageByShipment[shipmentId] = {};
        baseUsageByShipment[shipmentId].QTY = true;
      }
    }
  }

  const rows = [];
  const skuSet = {};
  const shipmentSet = {};
  const staged = [];
  const additionalExpenseColumnSet = {};
  // Тоталы по каждой ALLOC_BASE — нужны для расчёта долей при аллокации.
  const baseTotalsByShipment = {};
  const baseTotalsByShipmentSupplier = {};

  for (let i = 1; i < batchesData.length; i++) {
    const r = batchesData[i];
    const shipmentId = String(r[idxShipment] || '').trim();
    const sku = String(r[idxSku] || '').trim();
    if (!shipmentId || !sku) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    let qty = costingToNumber_(r[idxQty]);
    if (idxAllocQty != null) {
      const allocQty = costingToNumber_(r[idxAllocQty]);
      if (allocQty > 0) qty = allocQty;
    }
    if (!qty) continue;
    const amountRaw = idxCost != null ? costingToNumber_(r[idxCost]) : costingToNumber_(r[idxPrice]) * qty;
    const currencyRaw = idxCurrency != null ? String(r[idxCurrency] || '').trim().toUpperCase() : 'RUB';
    const currency = currencyRaw || 'RUB';
    const spec = idxSpec != null ? String(r[idxSpec] || '').trim() : '';
    const batchKey = costingPaymentBatchKey_(shipmentId, sku, spec);
    const manualRate = idxRate != null ? r[idxRate] : '';
    const fxCtx = {
      shipmentId: shipmentId,
      sku: sku,
      supplierArticle: idxSupplierArticle != null ? String(r[idxSupplierArticle] || '').trim() : '',
      spec: spec,
      currency: currency
    };
    const fxResolved = costingResolveFxRateForBatch_(paymentMap, batchKey, manualRate, currency, fxCache, fxMissing, fxCtx);
    let fxRate = 0;
    let fxSource = '';
    if (fxResolved) {
      fxRate = fxResolved.rate;
      fxSource = fxResolved.source;
    }
    const purchaseRub = amountRaw * fxRate;
    const tnvedRaw = idxTnved != null ? r[idxTnved] : '';
    const tnved = costingNormalizeTnved_(tnvedRaw) || skuToTnved[sku] || '';
    const dutyRule = dutyRulesByTnved[tnved] || { type: '', rate: 0 };
    const dutyType = costingNorm_(dutyRule.type);
    // Для "весовой" пошлины ставка задаётся как EUR/ед — её нельзя делить на 100.
    // Для "стоимостной": '15' → 0.15, '0.15' → 0.15.
    const dutyRate = costingNormalizeDutyRate_(dutyRule.rate, dutyRule.type);
    const vatRate = costingNormalizeRate_(vatRatesByTnved[tnved] || 0);
    // Правило:
    // - стоимостная: пошлина = таможенная стоимость * ставка
    // - весовая: пошлина = ставка(EUR/ед) * количество * курс EUR
    // В текущей модели "таможенная стоимость" = закупка_руб по строке партии.
    let dutyRub = 0;
    if (!isFact) {
    if (dutyType.indexOf('стоимост') !== -1) {
      dutyRub = purchaseRub * dutyRate;
    } else if (dutyType.indexOf('весов') !== -1) {
      const eurRate = costingResolveFxRate_(fxCache, 'EUR');
      if (eurRate == null) {
        fxMissing.push({
          reason: 'cbr',
          shipmentId: shipmentId,
          sku: sku,
          supplierArticle: idxSupplierArticle != null ? String(r[idxSupplierArticle] || '').trim() : '',
          spec: spec,
          currency: 'EUR (для весовой пошлины)'
        });
        dutyRub = 0;
      } else {
        dutyRub = dutyRate * qty * eurRate;
      }
    } else {
      // fallback: если тип не заполнен, трактуем ставку как стоимостную.
      dutyRub = purchaseRub * dutyRate;
    }
    }

    const volume = idxVolume != null ? costingToNumber_(r[idxVolume]) : 0;
    const weight = idxWeight != null ? costingToNumber_(r[idxWeight]) : 0;
    const shipViaKey = costingNormalizeSupplierKey_(idxShipVia != null ? String(r[idxShipVia] || '').trim() : '');

    staged.push({
      shipmentId: shipmentId,
      sku: sku,
      supplierArticle: idxSupplierArticle != null ? String(r[idxSupplierArticle] || '').trim() : '',
      tnved: tnved,
      shipVia: idxShipVia != null ? String(r[idxShipVia] || '').trim() : '',
      shipViaKey: shipViaKey,
      volume: volume,
      weight: weight,
      qty: qty,
      purchaseRub: purchaseRub,
      dutyRub: dutyRub,
      vatRate: vatRate,
      fxRate: fxRate,
      fxSource: fxSource
    });

    if (!baseTotalsByShipment[shipmentId]) {
      baseTotalsByShipment[shipmentId] = { VOLUME: 0, WEIGHT: 0, VALUE: 0, QTY: 0 };
    }
    const tot = baseTotalsByShipment[shipmentId];
    tot.VOLUME += volume > 0 ? volume : 0;
    tot.WEIGHT += weight > 0 ? weight : 0;
    tot.VALUE += purchaseRub > 0 ? purchaseRub : 0;
    tot.QTY += qty;

    if (!baseTotalsByShipmentSupplier[shipmentId]) baseTotalsByShipmentSupplier[shipmentId] = {};
    if (!baseTotalsByShipmentSupplier[shipmentId][shipViaKey]) {
      baseTotalsByShipmentSupplier[shipmentId][shipViaKey] = { VOLUME: 0, WEIGHT: 0, VALUE: 0, QTY: 0 };
    }
    const tots = baseTotalsByShipmentSupplier[shipmentId][shipViaKey];
    tots.VOLUME += volume > 0 ? volume : 0;
    tots.WEIGHT += weight > 0 ? weight : 0;
    tots.VALUE += purchaseRub > 0 ? purchaseRub : 0;
    tots.QTY += qty;
  }

  if (fxMissing.length) {
    throw new Error(costingFormatFxMissingError_(fxMissing));
  }

  // Жёсткая валидация: если для рейса (или scoped-поставщика) использован ALLOC_BASE,
  // у каждого участвующего SKU должна быть ненулевая величина по этой базе. Иначе — стоп с пояснением.
  costingValidateAllocationBases_(staged, baseUsageByShipment, baseUsageByShipmentSupplier);

  // Доли по каждой базе для конкретного SKU (для глобального пула и для scoped пула).
  function sharesGlobalFor(x) {
    const t = baseTotalsByShipment[x.shipmentId] || { VOLUME: 0, WEIGHT: 0, VALUE: 0, QTY: 0 };
    return {
      VOLUME: t.VOLUME > 0 && x.volume > 0 ? x.volume / t.VOLUME : 0,
      WEIGHT: t.WEIGHT > 0 && x.weight > 0 ? x.weight / t.WEIGHT : 0,
      VALUE: t.VALUE > 0 && x.purchaseRub > 0 ? x.purchaseRub / t.VALUE : 0,
      QTY: t.QTY > 0 ? x.qty / t.QTY : 0
    };
  }
  function sharesScopedFor(x) {
    const t = ((baseTotalsByShipmentSupplier[x.shipmentId] || {})[x.shipViaKey])
      || { VOLUME: 0, WEIGHT: 0, VALUE: 0, QTY: 0 };
    return {
      VOLUME: t.VOLUME > 0 && x.volume > 0 ? x.volume / t.VOLUME : 0,
      WEIGHT: t.WEIGHT > 0 && x.weight > 0 ? x.weight / t.WEIGHT : 0,
      VALUE: t.VALUE > 0 && x.purchaseRub > 0 ? x.purchaseRub / t.VALUE : 0,
      QTY: t.QTY > 0 ? x.qty / t.QTY : 0
    };
  }

  // Подготовка таможенной стоимости по компаниям-поставщикам — для определения таможсбора.
  const supplierAggByShipment = {};
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    const byShipment = supplierAggByShipment[x.shipmentId] || {
      totalVolume: 0,
      totalQty: 0,
      bySupplier: {}
    };
    byShipment.totalVolume += x.volume > 0 ? x.volume : 0;
    byShipment.totalQty += x.qty;
    if (!byShipment.bySupplier[x.shipViaKey]) {
      byShipment.bySupplier[x.shipViaKey] = { customsBaseRub: 0, volume: 0, qty: 0, customsFeeRub: 0 };
    }
    supplierAggByShipment[x.shipmentId] = byShipment;
  }

  // Считаем таможенную стоимость (= закупка + PRE_BORDER расходы) для каждой компании.
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    const agg = supplierAggByShipment[x.shipmentId];
    const poolGlobal = expensePoolsByShipment[x.shipmentId] || costingMakeEmptyExpensePool_();
    const poolScoped = expensePoolsByShipmentSupplier[x.shipmentId] && expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      ? expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      : costingMakeEmptyExpensePool_();
    const sg = sharesGlobalFor(x);
    const sc = sharesScopedFor(x);
    const preBorderRowRub =
      costingPoolFieldAlloc_(poolGlobal.freightRub, sg) +
      costingPoolFieldAlloc_(poolGlobal.otherPreRub, sg) +
      costingPoolFieldAlloc_(poolScoped.freightRub, sc) +
      costingPoolFieldAlloc_(poolScoped.otherPreRub, sc);
    const customsBaseRowRub = x.purchaseRub + preBorderRowRub;
    const sup = agg.bySupplier[x.shipViaKey];
    sup.customsBaseRub += customsBaseRowRub;
    sup.volume += x.volume > 0 ? x.volume : 0;
    sup.qty += x.qty;
  }
  const shipIds = Object.keys(supplierAggByShipment);
  for (let si = 0; si < shipIds.length; si++) {
    const sid = shipIds[si];
    const bySupplier = supplierAggByShipment[sid].bySupplier;
    const keys = Object.keys(bySupplier);
    for (let k = 0; k < keys.length; k++) {
      const supplierKey = keys[k];
      if (!isFact) {
        bySupplier[supplierKey].customsFeeRub = costingResolveCustomsFeeRub_(bySupplier[supplierKey].customsBaseRub, customsFeeRules);
      }
    }
  }

  const customsFactBySku = isFact ? costingLoadCustomsFactByShipmentSku_(ss, shipmentFilter) : null;
  const now = new Date();
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    const poolGlobal = expensePoolsByShipment[x.shipmentId] || costingMakeEmptyExpensePool_();
    const poolScoped = expensePoolsByShipmentSupplier[x.shipmentId] && expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      ? expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      : costingMakeEmptyExpensePool_();
    const sg = sharesGlobalFor(x);
    const sc = sharesScopedFor(x);
    const freightRub = costingPoolFieldAlloc_(poolGlobal.freightRub, sg) + costingPoolFieldAlloc_(poolScoped.freightRub, sc);
    const logisticsOtherRub = costingPoolFieldAlloc_(poolGlobal.logisticsOtherRub, sg) + costingPoolFieldAlloc_(poolScoped.logisticsOtherRub, sc);
    const dutyFromExpensesRub = costingPoolFieldAlloc_(poolGlobal.dutyRub, sg) + costingPoolFieldAlloc_(poolScoped.dutyRub, sc);
    const vatFromExpensesRub = costingPoolFieldAlloc_(poolGlobal.vatRub, sg) + costingPoolFieldAlloc_(poolScoped.vatRub, sc);
    const otherPreRub = costingPoolFieldAlloc_(poolGlobal.otherPreRub, sg) + costingPoolFieldAlloc_(poolScoped.otherPreRub, sc);
    const otherPostRub = costingPoolFieldAlloc_(poolGlobal.otherPostRub, sg) + costingPoolFieldAlloc_(poolScoped.otherPostRub, sc);
    const otherRub = otherPreRub + otherPostRub;
    const customsBaseRub = x.purchaseRub + freightRub + otherPreRub;
    const preBorderExpensesRub = freightRub + otherPreRub;
    const supAgg = supplierAggByShipment[x.shipmentId] ? supplierAggByShipment[x.shipmentId].bySupplier[x.shipViaKey] : null;
    const supShare = supAgg
      ? (supAgg.volume > 0 ? (x.volume / supAgg.volume) : (supAgg.qty > 0 ? (x.qty / supAgg.qty) : 0))
      : 0;
    let customsFeeRub;
    let dutyRub;
    let vatRub;
    if (isFact) {
      const fc = customsFactBySku && customsFactBySku[x.shipmentId] && customsFactBySku[x.shipmentId][x.sku];
      customsFeeRub = fc ? fc.fee : 0;
      dutyRub = (fc ? fc.duty : 0) + dutyFromExpensesRub;
      vatRub = (fc ? fc.vat : 0) + vatFromExpensesRub;
    } else {
      customsFeeRub = supAgg ? (supAgg.customsFeeRub * supShare) : 0;
      dutyRub = x.dutyRub + dutyFromExpensesRub;
      const vatFromRateRub = (customsBaseRub + dutyRub) * x.vatRate;
      vatRub = vatFromRateRub + vatFromExpensesRub;
    }
    const articleBreakdown = {};
    costingPoolArticleAlloc_(poolGlobal.otherByArticle, sg, articleBreakdown);
    costingPoolArticleAlloc_(poolScoped.otherByArticle, sc, articleBreakdown);
    const articleKeys = Object.keys(articleBreakdown);
    for (let ak = 0; ak < articleKeys.length; ak++) additionalExpenseColumnSet[articleKeys[ak]] = true;
    const extraTotal = freightRub + logisticsOtherRub + dutyRub + vatRub + customsFeeRub + otherRub;
    const total = x.purchaseRub + extraTotal;
    const unit = x.qty > 0 ? total / x.qty : 0;
    rows.push({
      base: [
      now,
      x.shipmentId,
      x.sku,
      x.supplierArticle,
      x.tnved,
      x.qty,
      x.purchaseRub,
      customsBaseRub,
      preBorderExpensesRub,
      freightRub,
      logisticsOtherRub,
      dutyRub,
      vatRub,
      customsFeeRub,
      otherRub,
      extraTotal,
      total,
      unit,
      x.fxRate,
      x.fxSource || ''
      ],
      byArticle: articleBreakdown
    });
    skuSet[x.sku] = true;
    shipmentSet[x.shipmentId] = true;
  }

  const additionalExpenseColumns = Object.keys(additionalExpenseColumnSet).sort();
  const finalizedRows = rows.map(function (rowObj) {
    const base = rowObj.base.slice(0, 15);
    const totals = rowObj.base.slice(15); // итого + unit + FX (5 полей)
    for (let i = 0; i < additionalExpenseColumns.length; i++) {
      const col = additionalExpenseColumns[i];
      base.push(rowObj.byArticle[col] || 0);
    }
    return base.concat(totals);
  });

  return {
    rows: finalizedRows,
    skuCount: Object.keys(skuSet).length,
    shipmentCount: Object.keys(shipmentSet).length,
    additionalExpenseColumns: additionalExpenseColumns,
    skippedFactExpenses: skippedFactExpenses,
    tripExpenseFxFilled: tripExpenseFxFilled
  };
}

function costingFormatSkippedFactExpenses_(list, maxLines) {
  const cap = maxLines == null ? 8 : maxLines;
  if (!list || !list.length) return '';
  const lines = list.slice(0, cap);
  let out = '\n\n⚠️ Не попали в факт (нет суммы RUB):\n• ' + lines.join('\n• ');
  if (list.length > cap) out += '\n• … ещё ' + (list.length - cap);
  return out;
}

function costingLoadDutyRulesByTnved_(ss) {
  const sh = costingFindSheetByRole_(ss, 'PRODUCTS_REF');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxTnved = costingFindColOptional_(h, ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  const idxDuty = costingFindColOptional_(h, ['Ставка пошлины', 'Ставка_пошлины', 'Пошлина', 'Размер пошлины', 'Размер_пошлины', 'Пошлина_EUR', 'Пошлина EUR', 'Ставка_пошлины_EUR', 'Ставка пошлины EUR']);
  const idxDutyType = costingFindColOptional_(h, ['Тип пошлины', 'Тип_пошлины', 'Вид пошлины', 'Вид_пошлины']);
  if (idxTnved == null || idxDuty == null) return {};
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const tnved = costingNormalizeTnved_(data[i][idxTnved]);
    if (!tnved) continue;
    const dutyRateRaw = data[i][idxDuty];
    const dutyType = idxDutyType != null ? String(data[i][idxDutyType] || '') : '';
    if (String(dutyRateRaw || '').trim() === '') continue;
    out[tnved] = { type: dutyType, rate: dutyRateRaw };
  }
  return out;
}

function costingLoadVatRatesByTnved_(ss) {
  const sh = costingFindSheetByRole_(ss, 'PRODUCTS_REF');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxTnved = costingFindColOptional_(h, ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  const idxVat = costingFindColOptional_(h, ['Ставка НДС', 'НДС', 'VAT', 'VAT_RATE', 'Ставка_НДС']);
  if (idxTnved == null || idxVat == null) return {};
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const tnved = costingNormalizeTnved_(data[i][idxTnved]);
    if (!tnved) continue;
    const vatRaw = data[i][idxVat];
    if (String(vatRaw || '').trim() === '') continue;
    out[tnved] = vatRaw;
  }
  return out;
}

function costingLoadSkuToTnvedMap_(ss) {
  const sh = costingFindSheetByRole_(ss, 'PRODUCTS_REF');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxSku = costingFindColOptional_(h, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул']);
  const idxTnved = costingFindColOptional_(h, ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  if (idxSku == null || idxTnved == null) return {};
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const sku = String(data[i][idxSku] || '').trim();
    if (!sku) continue;
    const tnved = costingNormalizeTnved_(data[i][idxTnved]);
    if (tnved) out[sku] = tnved;
  }
  return out;
}

function costingNormalizeTnved_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number' && isFinite(v)) {
    const n = Math.round(v);
    if (n > 0) return String(n);
  }
  let s = String(v).trim();
  if (/e/i.test(s)) {
    const n = Number(s.replace(',', '.'));
    if (isFinite(n) && n > 0) return String(Math.round(n));
  }
  return s.replace(/[^\d]/g, '');
}

function costingTnvedFromCells_(value, displayValue) {
  const fromDisplay = costingNormalizeTnved_(displayValue);
  if (fromDisplay.length >= 8) return fromDisplay;
  return costingNormalizeTnved_(value);
}

function costingTnvedCodesMatch_(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.indexOf(b) === 0 || b.indexOf(a) === 0) return true;
  const len = Math.min(a.length, b.length, 10);
  return len >= 8 && a.slice(0, len) === b.slice(0, len);
}

/** Первые N цифр ТН ВЭД (для группировки субкодов в декларации). */
function costingTnvedHeading_(code, digits) {
  const d = digits == null ? 6 : digits;
  const c = costingNormalizeTnved_(code);
  if (!c) return '';
  return c.slice(0, Math.min(d, c.length));
}

/**
 * Коды ТН ВЭД партии для сопоставления с декларацией.
 * Справочник товаров приоритетнее колонки в «Партии_в_рейсе» (там часто устаревший код).
 */
function costingBatchTnvedsForDeclMatch_(batchValue, batchDisplay, sku, skuToTnved) {
  const fromRef = sku && skuToTnved[sku] ? costingNormalizeTnved_(skuToTnved[sku]) : '';
  const fromBatch = costingTnvedFromCells_(batchValue, batchDisplay);
  const out = [];
  if (fromRef) out.push(fromRef);
  if (fromBatch && out.indexOf(fromBatch) === -1) out.push(fromBatch);
  return out;
}

/* ===================== Декларации / фактическая себестоимость ===================== */

function costingEnsureDeclLinesHeaderColumns_(sh) {
  const header = COST_CFG.DECL_LINES_HEADER;
  if (sh.getLastRow() < 1) return;
  const ncol = Math.max(sh.getLastColumn(), 1);
  const row1 = sh.getRange(1, 1, 1, ncol).getValues()[0];
  const norm = row1.map(costingHeaderCanonForLookup_);
  const hasUpload = norm.indexOf(costingHeaderCanonForLookup_('Загрузка_ID')) !== -1;
  if (!hasUpload) return;
  for (let i = 0; i < header.length; i++) {
    const key = costingHeaderCanonForLookup_(header[i]);
    if (norm.indexOf(key) !== -1) continue;
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue(header[i]);
    norm.push(key);
  }
}

function costingEnsureDeclLinesSheet_(ss) {
  let sh = costingFindSheetByRole_(ss, 'DECL_LINES');
  if (!sh) sh = ss.insertSheet(COST_CFG.SHEETS.DECL_LINES);
  const header = COST_CFG.DECL_LINES_HEADER;
  const a1 = sh.getRange(1, 1).getValue();
  const a1c = costingHeaderCanonForLookup_(a1);
  const okFirst =
    a1c === costingHeaderCanonForLookup_('Загрузка_ID') ||
    a1c === costingHeaderCanonForLookup_('ID');
  if (!okFirst) {
    sh.clear();
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  } else {
    costingEnsureDeclLinesHeaderColumns_(sh);
  }
  return sh;
}

function costingEnsureDeclJournalSheet_(ss) {
  let sh = costingFindSheetByRole_(ss, 'DECL_JOURNAL');
  if (!sh) sh = ss.insertSheet(COST_CFG.SHEETS.DECL_JOURNAL);
  const header = COST_CFG.DECL_JOURNAL_HEADER;
  const a1c = costingHeaderCanonForLookup_(sh.getLastRow() < 1 ? '' : sh.getRange(1, 1).getValue());
  const okJournalFirst =
    a1c === costingHeaderCanonForLookup_('Загрузка_ID') ||
    a1c === costingHeaderCanonForLookup_('ID');
  if (sh.getLastRow() < 1 || !okJournalFirst) {
    sh.clear();
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    return sh;
  }
  const hRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), header.length)).getValues()[0];
  if (costingFindColOptional_(hRow, ['Сбор_итого', 'Сбор итого', 'Сбор_итог', 'Сбор итог']) == null) {
    const warnCol = costingFindColOptional_(hRow, ['Предупреждения']);
    const insertAt = warnCol != null ? warnCol + 2 : 9;
    sh.insertColumnBefore(insertAt);
    sh.getRange(1, insertAt).setValue('Сбор_итого');
    if (warnCol != null && sh.getLastRow() > 1) {
      const lastRow = sh.getLastRow();
      const warnVals = sh.getRange(2, warnCol + 1, lastRow, warnCol + 1).getValues();
      const feeVals = [];
      for (let r = 0; r < warnVals.length; r++) {
        const v = warnVals[r][0];
        const n = costingToNumber_(v);
        const s = String(v || '');
        feeVals.push([n > 50 && s.indexOf(';') < 0 && s.indexOf('Тамож') < 0 ? n : '']);
      }
      sh.getRange(2, insertAt, lastRow, insertAt).setValues(feeVals);
    }
  }
  return sh;
}

function costingEnsureCustomsFactColumns_(ss) {
  const sh = costingGetSheetByRole_(ss, 'CUSTOMS');
  const data = sh.getDataRange().getValues();
  const header = data.length ? data[0].slice() : ['SHIPMENT_ID'];
  const hmap = costingHeaderMap_(header);
  costingEnsureHeaderColumn_(sh, hmap, 'SHIPMENT_ID');
  costingEnsureHeaderColumn_(sh, hmap, 'Артикул_ВБ');
  costingEnsureHeaderColumn_(sh, hmap, 'Код_ТНВЭД');
  costingEnsureHeaderColumn_(sh, hmap, 'Сценарий');
  costingEnsureHeaderColumn_(sh, hmap, 'Источник');
  costingEnsureHeaderColumn_(sh, hmap, 'Номер_ГТД');
  costingEnsureHeaderColumn_(sh, hmap, 'Загрузка_ID');
  costingEnsureHeaderColumn_(sh, hmap, 'Пошлина_RUB');
  costingEnsureHeaderColumn_(sh, hmap, 'НДС_RUB');
  costingEnsureHeaderColumn_(sh, hmap, 'Таможенный_сбор_RUB');
}

function costingEnsureFactCostSkuHeader_(ss) {
  let sh = costingFindSheetByRole_(ss, 'COST_SKU_FACT');
  if (!sh) sh = ss.insertSheet(COST_CFG.SHEETS.COST_SKU_FACT);
  if (sh.getLastRow() < 1) {
    const header = costingBuildFactOutputHeader_([]);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function costingLoadCustomsFactByShipmentSku_(ss, shipmentFilter) {
  const sh = costingFindSheetByRole_(ss, 'CUSTOMS');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxShipment = costingFindColOptional_(h, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindColOptional_(h, ['Артикул_ВБ', 'Артикул ВБ', 'WB_ARTICLE', 'Артикул']);
  const idxScenario = costingFindColOptional_(h, ['Сценарий']);
  const idxDuty = costingFindColOptional_(h, ['Пошлина_RUB', 'Пошлина RUB', 'Пошлина']);
  const idxVat = costingFindColOptional_(h, ['НДС_RUB', 'НДС RUB', 'НДС']);
  const idxFee = costingFindColOptional_(h, ['Таможенный_сбор_RUB', 'Таможенный сбор RUB', 'Таможенный сбор']);
  if (idxShipment == null || idxSku == null) return {};
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (idxScenario != null && String(r[idxScenario] || '').trim().toUpperCase() !== 'FACT') continue;
    const shipmentId = String(r[idxShipment] || '').trim();
    const sku = String(r[idxSku] || '').trim();
    if (!shipmentId || !sku) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    if (!out[shipmentId]) out[shipmentId] = {};
    if (!out[shipmentId][sku]) out[shipmentId][sku] = { duty: 0, vat: 0, fee: 0 };
    const cell = out[shipmentId][sku];
    cell.duty += idxDuty != null ? costingToNumber_(r[idxDuty]) : 0;
    cell.vat += idxVat != null ? costingToNumber_(r[idxVat]) : 0;
    cell.fee += idxFee != null ? costingToNumber_(r[idxFee]) : 0;
  }
  return out;
}

function costingBuildFactOutputHeader_(extraCols) {
  return costingBuildFullOutputHeader_(extraCols || []).concat([
    'Пошлина_план_RUB',
    'НДС_план_RUB',
    'Сценарий'
  ]);
}

function costingMergePlanDutyVatIntoFactRows_(factRows, planRows) {
  const planMap = {};
  for (let i = 0; i < planRows.length; i++) {
    const r = planRows[i];
    const key = String(r[1] || '').trim() + '\t' + String(r[2] || '').trim();
    planMap[key] = { duty: costingToNumber_(r[11]), vat: costingToNumber_(r[12]) };
  }
  return factRows.map(function (r) {
    const key = String(r[1] || '').trim() + '\t' + String(r[2] || '').trim();
    const p = planMap[key] || { duty: 0, vat: 0 };
    return r.concat([p.duty, p.vat, 'FACT']);
  });
}

function costingBuildSkuCostRowsFact_(shipmentFilter, options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildOpts = {
    writeTripExpenseFx: options.writeTripExpenseFx !== false
  };
  const planReport = costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter, Object.assign({
    scenarioMode: 'PLAN'
  }, buildOpts));
  const factReport = costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter, Object.assign({
    scenarioMode: 'FACT'
  }, buildOpts));
  const rows = costingMergePlanDutyVatIntoFactRows_(factReport.rows, planReport.rows);
  return {
    rows: rows,
    skuCount: factReport.skuCount,
    shipmentCount: factReport.shipmentCount,
    additionalExpenseColumns: factReport.additionalExpenseColumns,
    skippedFactExpenses: factReport.skippedFactExpenses || [],
    tripExpenseFxFilled: factReport.tripExpenseFxFilled || 0,
    planReport: planReport
  };
}

function costingDryRunFactPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Dry-run факт',
    'Введите SHIPMENT_ID (пусто = все рейсы):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  try {
    costingHealthCheckCore_();
    costingEnsureCustomsFactColumns_(SpreadsheetApp.getActiveSpreadsheet());
    const report = costingBuildSkuCostRowsFact_(shipmentId || null, { writeTripExpenseFx: false });
    const compare = costingBuildPlanFactCompareFromReport_(report, shipmentId || null);
    ui.alert(
      '🧪 Dry-run факт',
      'Строк: ' + report.rows.length + '\nSKU: ' + report.skuCount + '\n\n' + compare +
      costingFormatSkippedFactExpenses_(report.skippedFactExpenses),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ Dry-run факт', e.message || String(e), ui.ButtonSet.OK);
    throw e;
  }
}

function rebuildCostingFact_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Фактическая себестоимость',
    'Введите SHIPMENT_ID для точечного пересчёта или оставьте пустым для всей базы:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  if (shipmentId) {
    rebuildCostingFactByShipment_(shipmentId);
    return;
  }
  try {
    costingHealthCheckCore_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    costingEnsureCustomsFactColumns_(ss);
    costingEnsureFactCostSkuHeader_(ss);
    const report = costingBuildSkuCostRowsFact_(null);
    const sh = costingGetSheetByRole_(ss, 'COST_SKU_FACT');
    const extraCols = report.additionalExpenseColumns || [];
    const header = [costingBuildFactOutputHeader_(extraCols)];
    sh.clearContents();
    sh.getRange(1, 1, 1, header[0].length).setValues(header);
    if (report.rows.length) {
      sh.getRange(2, 1, report.rows.length, header[0].length).setValues(report.rows);
    }
    const fxMsg = report.tripExpenseFxFilled
      ? ('\nПодставлено курсов в «Затраты рейса» (FACT): ' + report.tripExpenseFxFilled)
      : '';
    ui.alert(
      '✅ Фактический пересчёт завершён',
      'Записано строк: ' + report.rows.length + '\nSKU: ' + report.skuCount + fxMsg +
      costingFormatSkippedFactExpenses_(report.skippedFactExpenses),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ Ошибка фактического пересчёта', e.message || String(e), ui.ButtonSet.OK);
    throw e;
  }
}

function rebuildCostingFactByShipment_(shipmentId) {
  const id = String(shipmentId || '').trim();
  if (!id) throw new Error('SHIPMENT_ID не указан.');
  costingHealthCheckCore_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  costingEnsureCustomsFactColumns_(ss);
  costingEnsureFactCostSkuHeader_(ss);
  const report = costingBuildSkuCostRowsFact_(id);
  const sh = costingGetSheetByRole_(ss, 'COST_SKU_FACT');
  const targetHeader = costingBuildFactOutputHeader_(report.additionalExpenseColumns || []);
  costingUpsertShipmentRowsToCostSku_(sh, id, report.rows, targetHeader);
  SpreadsheetApp.getUi().alert(
    '✅ Факт по рейсу',
    'SHIPMENT_ID: ' + id + '\nОбновлено строк: ' + report.rows.length +
    costingFormatSkippedFactExpenses_(report.skippedFactExpenses),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function costingComparePlanFactPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Сверка план vs факт',
    'Введите SHIPMENT_ID (пусто = все рейсы на листе):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  try {
    const result = costingRebuildPlanFactCompareSheet_(shipmentId || null);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = costingGetSheetByRole_(ss, 'PLAN_FACT_COMPARE');
    ss.setActiveSheet(sh);
    ui.alert(
      '✅ Сверка обновлена',
      'Лист: «' + COST_CFG.SHEETS.PLAN_FACT_COMPARE + '»\n' +
      'Строк: ' + result.rowCount + '\n' +
      'С расхождением: ' + result.diffCount + '\n' +
      'Только в плане: ' + result.onlyPlan + '\n' +
      'Только в факте: ' + result.onlyFact + '\n\n' +
      result.summaryText,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ Сверка план vs факт', e.message || String(e), ui.ButtonSet.OK);
    throw e;
  }
}

/** Итоги себестоимости — сразу после Qty на листе сверки. */
function costingPlanFactCompareCostMetricNames_() {
  return ['Итого_партия_руб', 'Себестоимость_ед_руб'];
}

/** Детализация расходов — после блока себестоимости. */
function costingPlanFactCompareOtherMetricNames_() {
  return [
    'Закупка_руб',
    'Таможенная_стоимость_база_RUB',
    'Расходы_до_границы_RUB',
    'Фрахт_до_РФ_RUB',
    'Прочая_логистика_RUB',
    'Пошлина_RUB',
    'НДС_RUB',
    'Таможенный_сбор_RUB',
    'Прочие_расходы_RUB',
    'Итого_доп_расходы_RUB'
  ];
}

function costingPlanFactCompareMetricNames_() {
  return costingPlanFactCompareCostMetricNames_().concat(costingPlanFactCompareOtherMetricNames_());
}

function costingPlanFactCompareSkipHeaderNorms_() {
  return [
    'расчет время',
    'shipment id',
    'артикул вб',
    'артикул поставщика',
    'код тнвэд',
    'qty',
    'курс средневзв rub',
    'fx источник',
    'пошлина план rub',
    'ндс план rub',
    'сценарий',
    'обновлено время',
    'только план',
    'только факт',
    'есть расхождение'
  ];
}

function costingCollectCompareMetrics_(planHeader, factHeader) {
  const metrics = costingPlanFactCompareCostMetricNames_()
    .concat(costingPlanFactCompareOtherMetricNames_());
  const seen = {};
  metrics.forEach(function (m) {
    seen[costingHeaderCanonForLookup_(m)] = true;
  });
  const skip = costingPlanFactCompareSkipHeaderNorms_();
  const factByNorm = {};
  const extras = [];
  for (let i = 0; i < factHeader.length; i++) {
    const name = String(factHeader[i] || '').trim();
    if (!name) continue;
    factByNorm[costingHeaderCanonForLookup_(name)] = name;
  }
  for (let i = 0; i < planHeader.length; i++) {
    const name = String(planHeader[i] || '').trim();
    if (!name) continue;
    const norm = costingHeaderCanonForLookup_(name);
    if (skip.indexOf(norm) !== -1 || seen[norm]) continue;
    if (!factByNorm[norm]) continue;
    extras.push(name);
    seen[norm] = true;
  }
  return metrics.concat(extras);
}

function costingBuildPlanFactCompareHeader_(metrics) {
  const header = [
    'Обновлено_время',
    'SHIPMENT_ID',
    'Артикул_ВБ',
    'Артикул_поставщика',
    'Код_ТНВЭД',
    'Qty'
  ];
  for (let i = 0; i < metrics.length; i++) {
    const m = metrics[i];
    header.push(m + '_план', m + '_факт', m + '_Δ');
  }
  header.push('Только_план', 'Только_факт', 'Есть_расхождение', 'Σ|Δ|_RUB');
  return header;
}

function costingEnsurePlanFactCompareSheet_(ss) {
  let sh = costingFindSheetByRole_(ss, 'PLAN_FACT_COMPARE');
  if (!sh) sh = ss.insertSheet(COST_CFG.SHEETS.PLAN_FACT_COMPARE);
  if (sh.getLastRow() < 1) {
    const header = costingBuildPlanFactCompareHeader_(costingPlanFactCompareMetricNames_());
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Читает лист себестоимости (план или факт) в карту по ключу SHIPMENT_ID + Артикул_ВБ.
 */
function costingLoadCostSkuSheetMap_(sh) {
  const out = { header: [], byKey: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getDataRange().getValues();
  out.header = data[0];
  const h = out.header;
  const idxShip = costingFindCol_(h, ['SHIPMENT_ID', 'Shipment_ID', 'Рейс']);
  const idxSku = costingFindCol_(h, ['Артикул_ВБ', 'Артикул ВБ', 'WB_ARTICLE']);
  const idxSupplier = costingFindColOptional_(h, ['Артикул_поставщика', 'Артикул поставщика']);
  const idxTnved = costingFindColOptional_(h, ['Код_ТНВЭД', 'Код ТН ВЭД', 'ТНВЭД']);
  const idxQty = costingFindColOptional_(h, ['Qty', 'QTY', 'Количество']);
  const metricIdx = {};
  for (let c = 0; c < h.length; c++) {
    const name = String(h[c] || '').trim();
    if (!name) continue;
    const norm = costingHeaderCanonForLookup_(name);
    if (costingPlanFactCompareSkipHeaderNorms_().indexOf(norm) !== -1) continue;
    metricIdx[norm] = c;
  }
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const shipmentId = String(row[idxShip] || '').trim();
    const sku = String(row[idxSku] || '').trim();
    if (!shipmentId || !sku) continue;
    const key = costingPlanFactRowKey_(shipmentId, sku);
    const values = {};
    Object.keys(metricIdx).forEach(function (norm) {
      values[norm] = costingToNumber_(row[metricIdx[norm]]);
    });
    out.byKey[key] = {
      shipmentId: shipmentId,
      sku: sku,
      supplierArticle: idxSupplier != null ? String(row[idxSupplier] || '').trim() : '',
      tnved: idxTnved != null ? String(row[idxTnved] || '').trim() : '',
      qty: idxQty != null ? costingToNumber_(row[idxQty]) : 0,
      values: values
    };
  }
  return out;
}

function costingPlanFactRowKey_(shipmentId, wbArticle) {
  return String(shipmentId || '').trim() + '\t' + String(wbArticle || '').trim();
}

function costingMetricNorm_(metricName) {
  return costingHeaderCanonForLookup_(metricName);
}

function costingReadMetricValue_(entry, metricName) {
  if (!entry || !entry.values) return 0;
  const v = entry.values[costingMetricNorm_(metricName)];
  return v == null ? 0 : costingToNumber_(v);
}

function costingRebuildPlanFactCompareSheet_(shipmentFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSh = costingFindSheetByRole_(ss, 'COST_SKU');
  const factSh = costingFindSheetByRole_(ss, 'COST_SKU_FACT');
  if (!planSh || planSh.getLastRow() < 2) {
    throw new Error('Лист «Себестоимость SKU» (план) пуст. Сначала выполните плановый пересчёт.');
  }
  if (!factSh || factSh.getLastRow() < 2) {
    throw new Error('Лист «Факт_себестоимость SKU» пуст. Сначала выполните фактический пересчёт.');
  }
  const planMap = costingLoadCostSkuSheetMap_(planSh);
  const factMap = costingLoadCostSkuSheetMap_(factSh);
  const metrics = costingCollectCompareMetrics_(planMap.header, factMap.header);
  const header = costingBuildPlanFactCompareHeader_(metrics);
  const keys = {};
  Object.keys(planMap.byKey).forEach(function (k) { keys[k] = true; });
  Object.keys(factMap.byKey).forEach(function (k) { keys[k] = true; });
  const sortedKeys = Object.keys(keys).sort();
  const updatedAt = new Date();
  const rows = [];
  let diffCount = 0;
  let onlyPlan = 0;
  let onlyFact = 0;
  let dutyDiffSum = 0;
  let vatDiffSum = 0;
  const EPS = 0.01;

  for (let i = 0; i < sortedKeys.length; i++) {
    const key = sortedKeys[i];
    const p = planMap.byKey[key];
    const f = factMap.byKey[key];
    const shipmentId = (p && p.shipmentId) || (f && f.shipmentId) || '';
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    const inPlan = !!p;
    const inFact = !!f;
    if (inPlan && !inFact) onlyPlan++;
    if (!inPlan && inFact) onlyFact++;
    const meta = p || f;
    let absSum = 0;
    let hasDiff = false;
    const row = [
      updatedAt,
      shipmentId,
      meta.sku,
      meta.supplierArticle,
      meta.tnved,
      meta.qty
    ];
    for (let m = 0; m < metrics.length; m++) {
      const metric = metrics[m];
      const planVal = costingReadMetricValue_(p, metric);
      const factVal = costingReadMetricValue_(f, metric);
      const delta = factVal - planVal;
      row.push(planVal, factVal, delta);
      absSum += Math.abs(delta);
      if (Math.abs(delta) >= EPS) hasDiff = true;
      if (costingMetricNorm_(metric) === costingMetricNorm_('Пошлина_RUB')) dutyDiffSum += delta;
      if (costingMetricNorm_(metric) === costingMetricNorm_('НДС_RUB')) vatDiffSum += delta;
    }
    if (hasDiff) diffCount++;
    row.push(
      inPlan && !inFact ? 'ДА' : '',
      !inPlan && inFact ? 'ДА' : '',
      hasDiff ? 'ДА' : '',
      Math.round(absSum * 100) / 100
    );
    rows.push(row);
  }

  const sh = costingEnsurePlanFactCompareSheet_(ss);
  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  const summaryText =
    'Σ Δпошлина: ' + dutyDiffSum.toFixed(2) + '\nΣ ΔНДС: ' + vatDiffSum.toFixed(2);
  return {
    rowCount: rows.length,
    diffCount: diffCount,
    onlyPlan: onlyPlan,
    onlyFact: onlyFact,
    summaryText: summaryText
  };
}

function costingBuildPlanFactCompareFromReport_(report, shipmentFilter) {
  if (!report || !report.rows || !report.rows.length) {
    return 'Нет строк для сверки.';
  }
  let count = 0;
  let dutyDiffSum = 0;
  let vatDiffSum = 0;
  const lines = [];
  const n = report.rows[0].length;
  const idxDutyFact = 11;
  const idxVatFact = 12;
  const idxDutyPlan = n - 3;
  const idxVatPlan = n - 2;
  for (let i = 0; i < report.rows.length; i++) {
    const r = report.rows[i];
    const sid = String(r[1] || '').trim();
    if (shipmentFilter && sid !== shipmentFilter) continue;
    const dutyF = costingToNumber_(r[idxDutyFact]);
    const vatF = costingToNumber_(r[idxVatFact]);
    const dutyP = costingToNumber_(r[idxDutyPlan]);
    const vatP = costingToNumber_(r[idxVatPlan]);
    const dDuty = dutyF - dutyP;
    const dVat = vatF - vatP;
    if (Math.abs(dDuty) < 0.01 && Math.abs(dVat) < 0.01) continue;
    count++;
    dutyDiffSum += dDuty;
    vatDiffSum += dVat;
    if (lines.length < 12) {
      lines.push(sid + ' / ' + r[2] + ': Δпошлина ' + dDuty.toFixed(2) + ', ΔНДС ' + dVat.toFixed(2));
    }
  }
  let out = 'Строк с расхождением: ' + count + '\nΣ Δпошлина: ' + dutyDiffSum.toFixed(2) + '\nΣ ΔНДС: ' + vatDiffSum.toFixed(2);
  if (lines.length) out += '\n\nПримеры:\n• ' + lines.join('\n• ');
  return out;
}

function costingBuildPlanFactCompareReport_(shipmentFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const factSh = costingFindSheetByRole_(ss, 'COST_SKU_FACT');
  if (!factSh || factSh.getLastRow() < 2) {
    return 'Лист «Факт_себестоимость SKU» пуст. Сначала выполните фактический пересчёт.';
  }
  const data = factSh.getDataRange().getValues();
  const h = data[0];
  const idxShip = costingFindCol_(h, ['SHIPMENT_ID', 'Shipment_ID', 'Рейс']);
  const idxSku = costingFindCol_(h, ['Артикул_ВБ', 'Артикул ВБ', 'WB_ARTICLE']);
  const idxDuty = costingFindCol_(h, ['Пошлина_RUB', 'Пошлина RUB', 'Пошлина']);
  const idxVat = costingFindCol_(h, ['НДС_RUB', 'НДС RUB', 'НДС']);
  const idxDutyPlan = costingFindColOptional_(h, ['Пошлина_план_RUB', 'Пошлина план RUB']);
  const idxVatPlan = costingFindColOptional_(h, ['НДС_план_RUB', 'НДС план RUB']);
  if (idxDutyPlan == null || idxVatPlan == null) {
    return 'На листе факта нет колонок Пошлина_план_RUB / НДС_план_RUB. Запустите пересчёт факта заново.';
  }
  let count = 0;
  let dutyDiffSum = 0;
  let vatDiffSum = 0;
  const lines = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const sid = String(r[idxShip] || '').trim();
    if (shipmentFilter && sid !== shipmentFilter) continue;
    const dDuty = costingToNumber_(r[idxDuty]) - costingToNumber_(r[idxDutyPlan]);
    const dVat = costingToNumber_(r[idxVat]) - costingToNumber_(r[idxVatPlan]);
    if (Math.abs(dDuty) < 0.01 && Math.abs(dVat) < 0.01) continue;
    count++;
    dutyDiffSum += dDuty;
    vatDiffSum += dVat;
    if (lines.length < 15) {
      lines.push(sid + ' / ' + r[idxSku] + ': Δпошлина ' + dDuty.toFixed(2) + ', ΔНДС ' + dVat.toFixed(2));
    }
  }
  let out = 'Строк с расхождением: ' + count + '\nΣ Δпошлина: ' + dutyDiffSum.toFixed(2) + '\nΣ ΔНДС: ' + vatDiffSum.toFixed(2);
  if (lines.length) out += '\n\n' + lines.join('\n');
  return out;
}

function costingNormalizeRate_(v) {
  const n = costingToNumber_(v);
  if (!n) return 0;
  // Поддержка обоих форматов: 15 и 0.15
  return n > 1 ? (n / 100) : n;
}

function costingCellText_(row, idx) {
  if (idx == null || idx < 0 || idx >= row.length) return '';
  return String(row[idx] || '').trim();
}

function costingClassifyExpense_(article, eventType, scope) {
  const text = (article + ' ' + eventType + ' ' + scope)
    .toLowerCase()
    .replace(/ё/g, 'е');
  if (text.indexOf('пошлин') !== -1 || text.indexOf('duty') !== -1) return 'duty';
  if (text.indexOf('ндс') !== -1) return 'vat';
  if ((text.indexOf('сбор') !== -1 && text.indexOf('тамож') !== -1) || text.indexOf('таможсбор') !== -1) {
    return 'customsFee';
  }
  const isLogistics = (
    text.indexOf('фрахт') !== -1 ||
    text.indexOf('логист') !== -1 ||
    text.indexOf('достав') !== -1 ||
    text.indexOf('перевоз') !== -1
  );
  if (isLogistics) {
    if (text.indexOf('post_border') !== -1 || text.indexOf('после границ') !== -1 || text.indexOf('после границы') !== -1) {
      return 'logisticsOther';
    }
    return 'freight';
  }
  return 'other';
}

function costingIsPostBorder_(scope) {
  const s = String(scope || '').toLowerCase().replace(/ё/g, 'е');
  return (
    s.indexOf('post_border') !== -1 ||
    s.indexOf('после границ') !== -1 ||
    s.indexOf('после границы') !== -1
  );
}

/* ===================== FX: транши оплат (Сводная 01) + ЦБ по дате ===================== */

/** Ключ сопоставления партии ↔ строка «Сводной» с оплатами. */
function costingPaymentBatchKey_(shipmentId, wbArticle, specNumber) {
  return [
    String(shipmentId || '').trim(),
    String(wbArticle || '').trim(),
    costingNormalizeSpecForPaymentKey_(specNumber)
  ].join('\t');
}

/** Нормализация № спецификации — как в payments_sync.gs. */
function costingNormalizeSpecForPaymentKey_(val) {
  if (val == null || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    const d = val.getDate();
    const m = val.getMonth() + 1;
    return d + '/' + m;
  }
  let s = String(val).toLowerCase().trim().replace(/\s+/g, '').replace(/\.0$/, '');
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:t.*)?$/);
  if (m) return parseInt(m[3], 10) + '/' + parseInt(m[2], 10);
  m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (m) return parseInt(m[1], 10) + '/' + parseInt(m[2], 10);
  m = s.match(/^(\d{1,2})\/0(\d)$/);
  if (m) return parseInt(m[1], 10) + '/' + parseInt(m[2], 10);
  return s;
}

/** Является ли значение датой фактической оплаты (не URL). */
function costingPaymentFactDatePresent_(rawVal, displayVal) {
  if (rawVal === true) return true;
  if (rawVal === false) return false;
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) return true;
  const rawText = rawVal == null ? '' : String(rawVal).trim();
  const dispText = displayVal == null ? '' : String(displayVal).trim();
  if (/^https?:\/\//i.test(rawText) || /^https?:\/\//i.test(dispText)) return false;
  if (/^\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?$/.test(rawText)) return true;
  if (/^\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?$/.test(dispText)) return true;
  return false;
}

function costingParsePaymentAmount_(rawVal, displayVal) {
  const rawText = rawVal == null ? '' : String(rawVal).trim();
  const disp = displayVal == null ? '' : String(displayVal).trim();
  if (/^https?:\/\//i.test(rawText) || /^https?:\/\//i.test(disp)) return 0;
  if (typeof rawVal === 'number' && isFinite(rawVal)) return rawVal;
  if (/^-?\d+(?:[.,]\d+)?$/.test(rawText)) return parseFloat(rawText.replace(',', '.'));
  const cleaned = disp
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, '')
    .replace(/,/g, '.')
    .replace(/[^\d.-]/g, '');
  if (/^-?\d+(?:\.\d+)?$/.test(cleaned)) return parseFloat(cleaned);
  return 0;
}

function costingParsePaymentDate_(rawVal, displayVal) {
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) return rawVal;
  if (typeof rawVal === 'number' && isFinite(rawVal) && rawVal > 0) {
    const dSerial = new Date(Math.round((rawVal - 25569) * 86400 * 1000));
    if (!isNaN(dSerial.getTime())) return dSerial;
  }
  const rawText = rawVal == null ? '' : String(rawVal).trim();
  const dispText = displayVal == null ? '' : String(displayVal).trim();
  const tryText = dispText || rawText;
  if (!tryText) return null;
  if (/^https?:\/\//i.test(tryText)) return null;
  const mRu = tryText.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$/);
  if (mRu) {
    let y = parseInt(mRu[3], 10);
    if (y < 100) y += 2000;
    const d = new Date(y, parseInt(mRu[2], 10) - 1, parseInt(mRu[1], 10));
    return isNaN(d.getTime()) ? null : d;
  }
  const d2 = new Date(tryText);
  return isNaN(d2.getTime()) ? null : d2;
}

function costingPickPaymentDate_(factRaw, factDisp, planRaw, planDisp, preferFact) {
  if (preferFact && costingPaymentFactDatePresent_(factRaw, factDisp)) {
    return costingParsePaymentDate_(factRaw, factDisp);
  }
  const plan = costingParsePaymentDate_(planRaw, planDisp);
  if (plan) return plan;
  if (!preferFact && costingPaymentFactDatePresent_(factRaw, factDisp)) {
    return costingParsePaymentDate_(factRaw, factDisp);
  }
  return null;
}

/**
 * Транши из строки «Сводной»: аванс, остаток, отсрочка (дата операции для defer).
 * @returns {Array<{type:string, amount:number, date:Date}>}
 */
function costingBuildPaymentTranchesFromSummaryRow_(row, displayRow, cols) {
  const tranches = [];
  const add = function (type, sumCol, factCol, planCol, opCol) {
    if (sumCol == null) return;
    const amount = costingParsePaymentAmount_(row[sumCol], displayRow[sumCol]);
    if (!(amount > 0)) return;
    let date = null;
    if (opCol != null) {
      date = costingParsePaymentDate_(row[opCol], displayRow[opCol]);
    }
    if (!date) {
      date = costingPickPaymentDate_(
        factCol != null ? row[factCol] : '',
        factCol != null ? displayRow[factCol] : '',
        planCol != null ? row[planCol] : '',
        planCol != null ? displayRow[planCol] : '',
        true
      );
    }
    if (!date) {
      date = costingPickPaymentDate_(
        factCol != null ? row[factCol] : '',
        factCol != null ? displayRow[factCol] : '',
        planCol != null ? row[planCol] : '',
        planCol != null ? displayRow[planCol] : '',
        false
      );
    }
    if (!date) return;
    tranches.push({ type: type, amount: amount, date: date });
  };
  add('advance', cols.advSum, cols.advFact, cols.advPlan, null);
  add('balance', cols.balSum, cols.balFact, cols.balPlan, null);
  add('defer', cols.defSum, cols.defFact, cols.defPlan, cols.defOpDate);
  return tranches;
}

/**
 * Карта траншей оплат из «Сводной» (книга 01). Первая строка по ключу; дубликаты — в warnings.
 * @returns {{ map: Object, duplicateWarnings: string[] }}
 */
function costingLoadPaymentTranchesFromSummary_(summarySh, shipmentFilter) {
  const out = { map: {}, duplicateWarnings: [] };
  const lr = summarySh.getLastRow();
  const lc = summarySh.getLastColumn();
  if (lr < 3 || lc < 1) return out;

  const headerRow = summarySh.getRange(2, 1, 1, lc).getValues()[0] || [];
  const idxShipment = costingFindColOptional_(headerRow, ['Рейс', 'SHIPMENT_ID', 'Shipment_ID', 'ID_рейса']);
  const idxWb = costingFindCol_(headerRow, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const idxSpec = costingFindColOptional_(headerRow, [
    'Номер спецификации', 'Номер_спецификации', 'Спецификация', '№ спецификации', '№спецификации',
    'Spec', 'Spec number', 'Номер инвойса', 'Инвойс'
  ]);
  const cols = {
    advSum: costingFindColOptional_(headerRow, ['Аванс сумма', 'Аванс, сумма']),
    advPlan: costingFindColOptional_(headerRow, ['Аванс план', 'Аванс план дата', 'Дата план Аванс']),
    advFact: costingFindColOptional_(headerRow, ['Аванс факт', 'Аванс факт дата', 'Дата факт Аванс']),
    balSum: costingFindColOptional_(headerRow, ['Баланс сумма', 'Баланс, сумма', 'Остаток сумма', 'Остаток, сумма']),
    balPlan: costingFindColOptional_(headerRow, ['Баланс план', 'Баланс план дата', 'Дата план Баланс', 'Остаток план']),
    balFact: costingFindColOptional_(headerRow, ['Баланс факт', 'Баланс факт дата', 'Дата факт Баланс', 'Остаток факт']),
    defSum: costingFindColOptional_(headerRow, ['Отсрочка сумма', 'Отсрочка, сумма']),
    defPlan: costingFindColOptional_(headerRow, ['Отсрочка план', 'Отсрочка план дата', 'Дата План Отсрочка']),
    defFact: costingFindColOptional_(headerRow, ['Отсрочка факт', 'Отсрочка факт дата', 'Дата Факт Отсрочка']),
    defOpDate: costingFindColOptional_(headerRow, [
      'Дата операции', 'Дата операции отсрочки', 'Дата опер. отсрочки', 'Дата опер отсрочки'
    ])
  };
  if (idxShipment == null) return out;

  const data = summarySh.getRange(3, 1, lr - 2, lc).getValues();
  const dataDisplay = summarySh.getRange(3, 1, lr - 2, lc).getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const d = dataDisplay[i];
    if (String(r[0] || '').trim().indexOf('▸') === 0) continue;
    const shipmentId = String(r[idxShipment] || '').trim();
    const wb = String(r[idxWb] || '').trim();
    if (!shipmentId || !wb) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    const spec = idxSpec != null ? (String(r[idxSpec] || '').trim() || String(d[idxSpec] || '').trim()) : '';
    const tranches = costingBuildPaymentTranchesFromSummaryRow_(r, d, cols);
    if (!tranches.length) continue;
    const key = costingPaymentBatchKey_(shipmentId, wb, spec);
    if (out.map[key]) {
      if (out.duplicateWarnings.length < 20) {
        out.duplicateWarnings.push(shipmentId + ' / ' + wb + ' / спец ' + (spec || '(пусто)'));
      }
      continue;
    }
    out.map[key] = tranches;
  }
  return out;
}

/** Безопасная загрузка траншей из книги 01; при недоступности 01 — пустая map (только ручной курс в партии). */
function costingLoadPaymentTranchesMapSafe_(shipmentFilter) {
  try {
    const ss01 = costingOpenBook01_();
    const summarySh = ss01.getSheetByName('Сводная');
    if (!summarySh) return { map: {}, duplicateWarnings: [] };
    return costingLoadPaymentTranchesFromSummary_(summarySh, shipmentFilter);
  } catch (e) {
    return { map: {}, duplicateWarnings: [], book01Error: e.message || String(e) };
  }
}

function costingFxDateCacheKey_(currencyCode, dateObj) {
  const d = dateObj instanceof Date && !isNaN(dateObj.getTime()) ? dateObj : new Date();
  const yyyy = d.getFullYear();
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  return String(currencyCode || '').trim().toUpperCase() + '|' + yyyy + '-' + mm + '-' + dd;
}

/**
 * Средневзвешенный курс ЦБ по траншам.
 * @returns {{ rate: number, weightSum: number, trancheCount: number }|null}
 */
function costingComputeWeightedFxRate_(tranches, currencyCode, cache, missingOut) {
  if (!tranches || !tranches.length) return null;
  let weightSum = 0;
  let rubSum = 0;
  let used = 0;
  for (let i = 0; i < tranches.length; i++) {
    const t = tranches[i];
    const amount = costingToNumber_(t.amount);
    if (!(amount > 0) || !t.date) continue;
    const rate = costingResolveFxRateOnDate_(cache, currencyCode, t.date);
    if (rate == null) {
      if (missingOut) {
        missingOut.push({
          reason: 'cbr',
          currency: String(currencyCode || '').trim().toUpperCase(),
          date: t.date,
          tranche: t.type || ''
        });
      }
      return null;
    }
    weightSum += amount;
    rubSum += amount * rate;
    used++;
  }
  if (!(weightSum > 0) || !(used > 0)) return null;
  return { rate: rubSum / weightSum, weightSum: weightSum, trancheCount: used };
}

/**
 * Курс для строки партии: weighted → ручной Курс_к_RUB → ошибка.
 * @returns {{ rate: number, source: string }|null}
 */
function costingResolveFxRateForBatch_(paymentMap, batchKey, manualRateRaw, currencyCode, cache, missingOut, ctx) {
  const currency = String(currencyCode || '').trim().toUpperCase();
  if (!currency || currency === 'RUB') return { rate: 1, source: 'RUB' };

  const tranches = paymentMap && paymentMap[batchKey];
  const weightedMissing = [];
  const weighted = tranches ? costingComputeWeightedFxRate_(tranches, currency, cache, weightedMissing) : null;
  if (weighted && weighted.rate > 0) {
    return { rate: weighted.rate, source: 'weighted:' + weighted.trancheCount };
  }

  const manual = costingToNumber_(manualRateRaw);
  if (manual > 0) return { rate: manual, source: 'manual_batch' };

  if (weightedMissing.length && missingOut) {
    for (let i = 0; i < weightedMissing.length; i++) {
      missingOut.push(Object.assign({ reason: 'cbr' }, weightedMissing[i], ctx || {}));
    }
    return null;
  }

  if (missingOut) {
    missingOut.push(Object.assign({ reason: 'no_fx' }, ctx || {}));
  }
  return null;
}

function costingFormatFxMissingError_(fxMissing) {
  const cbrItems = [];
  const noFxItems = [];
  for (let i = 0; i < fxMissing.length; i++) {
    const reason = fxMissing[i].reason;
    if (reason === 'cbr') cbrItems.push(fxMissing[i]);
    else if (reason !== 'expense_no_date') noFxItems.push(fxMissing[i]);
  }
  const lines = [];

  if (cbrItems.length) {
    const byKey = {};
    for (let i = 0; i < cbrItems.length; i++) {
      const x = cbrItems[i];
      const dk = (x.currency || '?') + '|' + (x.date ? costingFxDateCacheKey_(x.currency, x.date).split('|')[1] : '?');
      if (!byKey[dk]) byKey[dk] = [];
      byKey[dk].push(x);
    }
    lines.push('Не удалось получить курс ЦБ на дату платежа (cbr.ru недоступен или валюта не найдена).');
    const keys = Object.keys(byKey).sort();
    for (let ki = 0; ki < keys.length; ki++) {
      const items = byKey[keys[ki]];
      const skuList = items.slice(0, 15).map(function (x) {
        return (x.shipmentId || '') + ' / ' + (x.sku || x.supplierArticle || '(пусто)');
      }).join(', ');
      const overflow = items.length > 15 ? ' …и ещё ' + (items.length - 15) : '';
      lines.push('• ' + keys[ki] + ' — ' + items.length + ' стр.: ' + skuList + overflow);
    }
  }

  const expenseNoDate = [];
  for (let i = 0; i < fxMissing.length; i++) {
    if (fxMissing[i].reason === 'expense_no_date') expenseNoDate.push(fxMissing[i]);
  }
  if (expenseNoDate.length) {
    if (lines.length) lines.push('');
    lines.push('Затраты рейса в валюте без даты (нужна «Дата оплаты» или «Дата расхода»):');
    const samples = expenseNoDate.slice(0, 10).map(function (x) {
      return (x.shipmentId || '') + ' / ' + (x.supplierArticle || x.sku || '');
    }).join(', ');
    lines.push('• ' + expenseNoDate.length + ' стр.: ' + samples);
  }

  if (noFxItems.length) {
    if (lines.length) lines.push('');
    lines.push('Нет траншей оплаты в «Сводной» и не заполнен «Курс_к_RUB» в «Партии_в_рейсе» (старые отгрузки).');
    const skuList = noFxItems.slice(0, 25).map(function (x) {
      const spec = x.spec ? (' / спец ' + x.spec) : '';
      return (x.shipmentId || '') + ' / ' + (x.sku || x.supplierArticle || '(пусто)') + spec;
    }).join(', ');
    const overflow = noFxItems.length > 25 ? ' …и ещё ' + (noFxItems.length - 25) : '';
    lines.push('• ' + noFxItems.length + ' стр.: ' + skuList + overflow);
  }

  lines.push('');
  lines.push('Что делать:');
  lines.push('  1) Проверить аванс/остаток/отсрочку в «Сводной» (книга 01) для этих позиций.');
  lines.push('  2) Для legacy-отгрузок — вписать «Курс_к_RUB» в «Партии_в_рейсе» вручную.');
  lines.push('  3) Если закупка в рублях — «Валюта» = RUB.');
  return lines.join('\n');
}

function costingGetEuroRateOnDate_(dateObj) {
  return costingGetFxRateOnDateWithFallback_('EUR', dateObj);
}

/**
 * Возвращает курс ЦБ для валюты на дату.
 * При ошибке (сеть, разбор XML, валюта не найдена) возвращает null —
 * каллер обязан явно решить, что делать. Тихого фолбэка на 1 больше нет:
 * раньше это приводило к тому, что строки с пустым Курс_к_RUB и неудачным
 * запросом к ЦБ молча сохранялись в «Закупка_руб» в юанях.
 */
function costingGetFxRateOnDate_(currencyCode, dateObj) {
  try {
    const target = String(currencyCode || '').trim().toUpperCase();
    if (!target) return null;
    if (target === 'RUB') return 1;
    const d = dateObj instanceof Date ? dateObj : new Date(dateObj);
    if (isNaN(d.getTime())) return null;
    const dd = ('0' + d.getDate()).slice(-2);
    const mm = ('0' + (d.getMonth() + 1)).slice(-2);
    const yyyy = d.getFullYear();
    const url = 'https://www.cbr.ru/scripts/XML_daily.asp?date_req=' + dd + '/' + mm + '/' + yyyy;
    const xml = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText('windows-1251');
    const doc = XmlService.parse(xml);
    const root = doc.getRootElement();
    const vals = root.getChildren('Valute');
    if (!vals || !vals.length) return null;
    for (let i = 0; i < vals.length; i++) {
      const codeEl = vals[i].getChild('CharCode');
      if (!codeEl) continue;
      if (String(codeEl.getText() || '').trim().toUpperCase() !== target) continue;
      const nominalEl = vals[i].getChild('Nominal');
      const valueEl = vals[i].getChild('Value');
      const nominal = nominalEl ? costingToNumber_(nominalEl.getText()) : 1;
      const value = valueEl ? costingToNumber_(String(valueEl.getText() || '').replace(',', '.')) : 0;
      if (nominal > 0 && value > 0) return value / nominal;
    }
  } catch (e) {
    return null;
  }
  return null;
}

/** Курс ЦБ на дату расхода; при отсутствии публикации (выходной) — до maxDaysBack дней назад. */
function costingGetFxRateOnDateWithFallback_(currencyCode, dateObj, maxDaysBack) {
  const cap = maxDaysBack == null ? 14 : maxDaysBack;
  const base = dateObj instanceof Date ? new Date(dateObj.getTime()) : new Date(dateObj);
  if (isNaN(base.getTime())) return null;
  for (let i = 0; i <= cap; i++) {
    const d = new Date(base.getTime());
    d.setDate(d.getDate() - i);
    const rate = costingGetFxRateOnDate_(currencyCode, d);
    if (rate != null && rate > 0) return rate;
  }
  return null;
}

/** Дата для FX по строке «Затраты рейса»: дата расхода → дата оплаты → fallbackDate (обычно сегодня). */
function costingTripExpenseFxDateFromRow_(row, cols, fallbackDate) {
  let d = null;
  if (cols.expenseDate != null) {
    d = costingParsePaymentDate_(row[cols.expenseDate], row[cols.expenseDate]);
  }
  if (!d && cols.payDate != null) {
    d = costingParsePaymentDate_(row[cols.payDate], row[cols.payDate]);
  }
  if (!d && fallbackDate instanceof Date && !isNaN(fallbackDate.getTime())) {
    d = fallbackDate;
  }
  return d;
}

function costingTripExpenseCellBlank_(v) {
  if (v == null || v === '') return true;
  if (typeof v === 'number') return !isFinite(v) || v === 0;
  return String(v).trim() === '';
}

/**
 * Сумма и курс для строки «Затраты рейса» + флаги, что записать в пустые ячейки листа.
 */
function costingResolveTripExpenseRubDetail_(row, cols, fxCache, missingOut, ctx) {
  const out = { sumRub: 0, rate: null, fillRate: false, fillSumRub: false };
  const amount = cols.sum != null ? costingToNumber_(row[cols.sum]) : 0;
  if (!(amount > 0)) return out;

  const currencyRaw = cols.currency != null ? String(row[cols.currency] || '').trim().toUpperCase() : '';
  const currency = currencyRaw || 'RUB';
  const existingSumRub = cols.sumRub != null ? costingToNumber_(row[cols.sumRub]) : 0;
  const rateBlank = cols.rate == null || costingTripExpenseCellBlank_(row[cols.rate]);
  const sumRubBlank = cols.sumRub == null || costingTripExpenseCellBlank_(row[cols.sumRub]);

  if (currency === 'RUB') {
    out.rate = 1;
    out.sumRub = sumRubBlank ? amount : existingSumRub;
    out.fillRate = rateBlank && cols.rate != null;
    out.fillSumRub = sumRubBlank && cols.sumRub != null;
    return out;
  }

  const manualRate = cols.rate != null ? costingToNumber_(row[cols.rate]) : 0;
  if (!rateBlank && manualRate > 0) {
    out.rate = manualRate;
    out.sumRub = sumRubBlank ? Math.round(amount * manualRate * 100) / 100 : existingSumRub;
    out.fillSumRub = sumRubBlank && cols.sumRub != null;
    return out;
  }

  if (!sumRubBlank && existingSumRub > 0 && rateBlank && amount > 0) {
    out.sumRub = existingSumRub;
    out.rate = existingSumRub / amount;
    out.fillRate = cols.rate != null;
    return out;
  }

  if (!sumRubBlank && existingSumRub > 0) {
    out.sumRub = existingSumRub;
    if (!rateBlank) out.rate = manualRate;
    return out;
  }

  const fxDate = costingTripExpenseFxDateFromRow_(row, cols, fxCache ? fxCache.date : null);
  if (!fxDate) {
    if (missingOut) {
      missingOut.push(Object.assign({
        reason: 'expense_no_date',
        currency: currency,
        amount: amount
      }, ctx || {}));
    }
    return out;
  }

  const rate = costingResolveFxRateOnDate_(fxCache, currency, fxDate);
  if (rate == null) {
    if (missingOut) {
      missingOut.push(Object.assign({
        reason: 'cbr',
        currency: currency,
        date: fxDate,
        tranche: 'затраты_рейса'
      }, ctx || {}));
    }
    return out;
  }

  out.rate = rate;
  out.sumRub = Math.round(amount * rate * 100) / 100;
  out.fillRate = rateBlank && cols.rate != null;
  out.fillSumRub = sumRubBlank && cols.sumRub != null;
  return out;
}

/** Записать в «Затраты рейса» рассчитанные Курс_к_RUB и Сумма_RUB (только пустые ячейки). */
function costingApplyTripExpenseFxFill_(sh, updates) {
  if (!sh || !updates || !updates.length) return 0;
  let filled = 0;
  for (let i = 0; i < updates.length; i++) {
    const u = updates[i];
    const sheetRow = u.row + 1;
    if (u.fillRate && u.rateCol != null && u.rate != null && u.rate > 0) {
      sh.getRange(sheetRow, u.rateCol + 1).setValue(Math.round(u.rate * 1000000) / 1000000);
      filled++;
    }
    if (u.fillSumRub && u.sumRubCol != null && u.sumRub > 0) {
      sh.getRange(sheetRow, u.sumRubCol + 1).setValue(Math.round(u.sumRub * 100) / 100);
    }
  }
  return filled;
}

/**
 * Кэш курсов на одно выполнение пересчёта.
 * Использование:
 *   const cache = costingMakeFxCache_(new Date());
 *   const rate = costingResolveFxRate_(cache, 'CNY');  // 10.93 или null
 *
 * Гарантирует ровно один запрос к ЦБ на валюту вне зависимости от того,
 * сколько строк используют этот курс.
 */
function costingMakeFxCache_(dateObj) {
  return {
    date: dateObj instanceof Date ? dateObj : new Date(),
    byCurrency: {},
    byCurrencyDate: {}
  };
}

function costingResolveFxRateOnDate_(cache, currencyCode, dateObj) {
  const target = String(currencyCode || '').trim().toUpperCase();
  if (!target) return null;
  if (target === 'RUB') return 1;
  const key = costingFxDateCacheKey_(target, dateObj);
  if (Object.prototype.hasOwnProperty.call(cache.byCurrencyDate, key)) {
    return cache.byCurrencyDate[key];
  }
  const rate = costingGetFxRateOnDateWithFallback_(target, dateObj);
  cache.byCurrencyDate[key] = rate;
  return rate;
}

function costingResolveFxRate_(cache, currencyCode) {
  return costingResolveFxRateOnDate_(cache, currencyCode, cache.date);
}

/**
 * Сумма затраты рейса в RUB (см. costingResolveTripExpenseRubDetail_).
 * @returns {number|null} null — пропустить строку
 */
function costingResolveTripExpenseSumRub_(row, cols, fxCache, missingOut, ctx) {
  const d = costingResolveTripExpenseRubDetail_(row, cols, fxCache, missingOut, ctx);
  return d.sumRub > 0 ? d.sumRub : null;
}

/**
 * Меню-обёртка: подставить курс в пустые «Курс_к_RUB» в «Партии_в_рейсе».
 * Средневзвешенный ЦБ по траншам из «Сводной» (01); уже заполненный ручной курс не трогаем.
 */
function costingFillEmptyFxRatesMenu_() {
  try {
    const result = costingFillEmptyFxRatesCore_();
    SpreadsheetApp.getUi().alert(
      result.title || '✅ Курсы подставлены',
      result.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Не удалось подставить курсы',
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

function costingFillEmptyFxRatesCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = costingGetSheetByRole_(ss, 'BATCHES');
  if (!sh) {
    throw new Error('Не нашёл лист «Партии_в_рейсе» в этой книге.');
  }
  if (sh.getLastRow() < 2) {
    return { title: 'ℹ️ Нет данных', message: 'В «Партии_в_рейсе» нет строк.' };
  }
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxCurrency = costingFindColOptional_(h, ['Валюта', 'Currency']);
  const idxRate = costingFindColOptional_(h, ['Курс_к_RUB', 'Курс к RUB', 'Курс', 'Rate']);
  if (idxCurrency == null || idxRate == null) {
    throw new Error('Не нашёл колонки «Валюта» и/или «Курс_к_RUB» в «Партии_в_рейсе».');
  }
  const idxShipment = costingFindColOptional_(h, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindColOptional_(h, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const idxSpec = costingFindColOptional_(h, [
    'Номер спецификации', 'Номер_спецификации', 'Спецификация', '№ спецификации', '№спецификации',
    'Spec', 'Spec number'
  ]);

  const paymentLoad = costingLoadPaymentTranchesMapSafe_(null);
  const paymentMap = paymentLoad.map || {};
  const fxCache = costingMakeFxCache_(new Date());
  let filledCount = 0;
  const ratesUsed = {};
  const failedNoTranches = [];
  const failedCbr = {};
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const curRaw = String(row[idxCurrency] || '').trim().toUpperCase();
    if (!curRaw || curRaw === 'RUB') continue;
    const rateCur = costingToNumber_(row[idxRate]);
    if (rateCur > 0) continue;

    const shipmentId = idxShipment != null ? String(row[idxShipment] || '').trim() : '';
    const sku = idxSku != null ? String(row[idxSku] || '').trim() : '';
    const spec = idxSpec != null ? String(row[idxSpec] || '').trim() : '';
    const batchKey = costingPaymentBatchKey_(shipmentId, sku, spec);
    const fxMissing = [];
    const resolved = costingResolveFxRateForBatch_(paymentMap, batchKey, '', curRaw, fxCache, fxMissing, {
      shipmentId: shipmentId,
      sku: sku,
      spec: spec
    });
    if (!resolved || !(resolved.rate > 0)) {
      const sample = shipmentId + ' / ' + (sku || '(пусто)');
      if (fxMissing.length && fxMissing[0].reason === 'cbr') {
        if (!failedCbr[curRaw]) failedCbr[curRaw] = [];
        if (failedCbr[curRaw].length < 5) failedCbr[curRaw].push(sample);
      } else {
        if (failedNoTranches.length < 8) failedNoTranches.push(sample);
      }
      continue;
    }
    ratesUsed[curRaw] = resolved.rate;
    filledCount++;
    updates.push({ rowIdx: i + 1, value: resolved.rate });
  }

  // Запись: пишем построчно. Для небольшого количества строк это нормально.
  // Если будут сотни — стоит переделать на блочную запись, но пока преждевременно.
  for (let i = 0; i < updates.length; i++) {
    sh.getRange(updates[i].rowIdx, idxRate + 1).setValue(updates[i].value);
  }

  const lines = [];
  if (filledCount === 0 && !failedNoTranches.length && !Object.keys(failedCbr).length) {
    return {
      title: 'ℹ️ Нечего обновлять',
      message: 'В «Партии_в_рейсе» нет строк с пустым «Курс_к_RUB» (и непустой «Валюта»).'
    };
  }
  if (filledCount) {
    lines.push('Заполнено строк (средневзв. по оплатам из «Сводной»): ' + filledCount);
    const usedCodes = Object.keys(ratesUsed).sort();
    for (let i = 0; i < usedCodes.length; i++) {
      const code = usedCodes[i];
      lines.push('  • ' + code + ' ≈ ' + (Math.round(ratesUsed[code] * 10000) / 10000) + ' ₽');
    }
  }
  if (failedNoTranches.length) {
    if (lines.length) lines.push('');
    lines.push('Без траншей в «Сводной» — заполни «Курс_к_RUB» вручную (legacy):');
    lines.push('  • ' + failedNoTranches.join(', '));
  }
  if (Object.keys(failedCbr).length) {
    if (lines.length) lines.push('');
    lines.push('ЦБ недоступен на дату платежа:');
    const codes = Object.keys(failedCbr).sort();
    for (let i = 0; i < codes.length; i++) {
      lines.push('  • ' + codes[i] + ': ' + failedCbr[codes[i]].join(', '));
    }
  }
  if (paymentLoad.book01Error) {
    lines.push('');
    lines.push('ℹ️ Книга 01 недоступна (' + paymentLoad.book01Error + ') — подставлены только строки с траншами, если 01 открылась частично.');
  }
  return {
    title: filledCount ? '✅ Курсы подставлены' : '⚠️ Курсы не получены',
    message: lines.join('\n')
  };
}

function costingLoadCustomsPoolsByShipment_(ss, shipmentFilter, scenarioMode) {
  const wantFact = String(scenarioMode || 'PLAN').toUpperCase() === 'FACT';
  const sh = costingFindSheetByRole_(ss, 'CUSTOMS');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxShipment = costingFindColOptional_(h, ['SHIPMENT_ID', 'ID_рейса']);
  if (idxShipment == null) return {};
  const idxScenario = costingFindColOptional_(h, ['Сценарий']);
  const idxDuty = costingFindColOptional_(h, ['Пошлина_RUB', 'Пошлина RUB', 'Пошлина']);
  const idxVat = costingFindColOptional_(h, ['НДС_RUB', 'НДС RUB', 'НДС']);
  const idxFee = costingFindColOptional_(h, ['Таможенный_сбор_RUB', 'Таможенный сбор RUB', 'Таможенный сбор']);

  const out = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const shipmentId = String(r[idxShipment] || '').trim();
    if (!shipmentId) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    if (idxScenario != null) {
      const sc = String(r[idxScenario] || '').trim().toUpperCase();
      if (wantFact ? sc !== 'FACT' : sc === 'FACT') continue;
    }
    if (!out[shipmentId]) {
      out[shipmentId] = { dutyRub: 0, vatRub: 0, customsFeeRub: 0 };
    }
    out[shipmentId].dutyRub += idxDuty != null ? costingToNumber_(r[idxDuty]) : 0;
    out[shipmentId].vatRub += idxVat != null ? costingToNumber_(r[idxVat]) : 0;
    out[shipmentId].customsFeeRub += idxFee != null ? costingToNumber_(r[idxFee]) : 0;
  }
  return out;
}

function costingLoadCustomsFeeRules_(ss) {
  const sh = costingFindSheetByRole_(ss, 'CUSTOMS_FEE_RATES');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxMin = costingFindColOptional_(h, ['От_ТС', 'Стоимость от', 'Таможенная стоимость от', 'От', 'Мин', 'Нижняя граница']);
  const idxMax = costingFindColOptional_(h, ['До_ТС', 'Стоимость до', 'Таможенная стоимость до', 'До', 'Макс', 'Верхняя граница']);
  const idxFee = costingFindColOptional_(h, ['Сбор_RUB', 'Размер сбора', 'Таможенный сбор', 'Сбор', 'Сумма сбора']);
  if (idxMin == null || idxFee == null) return [];
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const min = costingToNumber_(row[idxMin]);
    const max = idxMax != null ? costingToNumber_(row[idxMax]) : 0;
    const fee = costingToNumber_(row[idxFee]);
    if (fee <= 0) continue;
    rules.push({ min: min, max: max, fee: fee });
  }
  rules.sort(function (a, b) { return a.min - b.min; });
  return rules;
}

function costingResolveCustomsFeeRub_(customsBaseRub, rules) {
  const base = costingToNumber_(customsBaseRub);
  if (!base || !rules || !rules.length) return 0;
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    const minOk = base >= r.min;
    const maxOk = !r.max || base <= r.max;
    if (minOk && maxOk) return r.fee;
  }
  // Если не попали в диапазон, берём последний порог.
  return rules[rules.length - 1].fee;
}

function costingMakeEmptyExpensePool_() {
  // Все числовые поля стали map: { allocBase -> сумма }.
  // allocBase ∈ {'VOLUME', 'WEIGHT', 'VALUE', 'QTY'} — берётся из колонки ALLOC_BASE строки расхода.
  // otherByArticle хранит per-статья map по тем же базам: { article -> { base -> сумма } }.
  return {
    freightRub: {},
    logisticsOtherRub: {},
    dutyRub: {},
    vatRub: {},
    customsFeeRub: {},
    otherPreRub: {},
    otherPostRub: {},
    otherByArticle: {}
  };
}

function costingNormalizeAllocBase_(v) {
  const raw = String(v == null ? '' : v).trim().toUpperCase().replace(/Ё/g, 'Е');
  if (!raw) return 'VOLUME';
  if (raw === 'WEIGHT' || raw === 'ВЕС' || raw === 'WGT' || raw === 'KG' || raw === 'KGS') return 'WEIGHT';
  if (
    raw === 'VALUE' ||
    raw === 'COST' ||
    raw === 'СТОИМОСТЬ' ||
    raw === 'ПО СТОИМОСТИ' ||
    raw === 'CUSTOMS_VALUE' ||
    raw === 'CV' ||
    raw === 'VAL'
  ) {
    return 'VALUE';
  }
  if (raw === 'QTY' || raw === 'QUANTITY' || raw === 'КОЛИЧЕСТВО' || raw === 'ШТ' || raw === 'UNITS') return 'QTY';
  return 'VOLUME';
}

function costingAddToBaseMap_(map, base, sum) {
  if (!map || !sum) return;
  map[base] = (map[base] || 0) + sum;
}

function costingPoolFieldAlloc_(field, sharesByBase) {
  if (!field) return 0;
  let sum = 0;
  const bases = Object.keys(field);
  for (let i = 0; i < bases.length; i++) {
    const b = bases[i];
    sum += (field[b] || 0) * (sharesByBase[b] || 0);
  }
  return sum;
}

function costingPoolArticleAlloc_(otherByArticle, sharesByBase, out) {
  if (!otherByArticle) return out;
  const articles = Object.keys(otherByArticle);
  for (let i = 0; i < articles.length; i++) {
    const article = articles[i];
    const baseMap = otherByArticle[article] || {};
    const bases = Object.keys(baseMap);
    let allocated = 0;
    for (let j = 0; j < bases.length; j++) {
      const b = bases[j];
      allocated += (baseMap[b] || 0) * (sharesByBase[b] || 0);
    }
    out[article] = (out[article] || 0) + allocated;
  }
  return out;
}

function costingNormalizeDutyRate_(rateRaw, typeRaw) {
  // Для весовой ставки делить на 100 нельзя — она задаётся как EUR/ед (часто < 1, но может быть и > 1).
  // Для стоимостной: '15' → 0.15, '0.15' → 0.15.
  const t = String(typeRaw || '').toLowerCase().replace(/ё/g, 'е');
  if (t.indexOf('весов') !== -1) {
    return costingToNumber_(rateRaw);
  }
  return costingNormalizeRate_(rateRaw);
}

function costingValidateAllocationBases_(staged, baseUsageByShipment, baseUsageByShipmentSupplier) {
  if (!staged || !staged.length) return;

  const byShipment = {};
  const byShipmentSupplier = {};
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    if (!byShipment[x.shipmentId]) byShipment[x.shipmentId] = [];
    byShipment[x.shipmentId].push(x);
    if (!byShipmentSupplier[x.shipmentId]) byShipmentSupplier[x.shipmentId] = {};
    if (!byShipmentSupplier[x.shipmentId][x.shipViaKey]) byShipmentSupplier[x.shipmentId][x.shipViaKey] = [];
    byShipmentSupplier[x.shipmentId][x.shipViaKey].push(x);
  }

  function hasBaseValue(x, base) {
    if (base === 'VOLUME') return x.volume > 0;
    if (base === 'WEIGHT') return x.weight > 0;
    if (base === 'VALUE') return x.purchaseRub > 0;
    if (base === 'QTY') return x.qty > 0;
    return true;
  }

  function baseLabel(b) {
    if (b === 'VOLUME') return 'Объём (м³)';
    if (b === 'WEIGHT') return 'Вес (кг)';
    if (b === 'VALUE') return 'Закупка (₽)';
    if (b === 'QTY') return 'Количество (шт)';
    return b;
  }

  const errors = [];
  function check(skus, bases, scopeLabel) {
    if (!bases) return;
    const baseList = Object.keys(bases);
    for (let j = 0; j < baseList.length; j++) {
      const base = baseList[j];
      const missing = [];
      for (let s = 0; s < skus.length; s++) {
        const x = skus[s];
        if (!hasBaseValue(x, base)) {
          missing.push(x.sku || x.supplierArticle || '(пусто)');
        }
      }
      if (missing.length) {
        errors.push({ base: base, scope: scopeLabel, missing: missing });
      }
    }
  }

  const shipmentIds = Object.keys(byShipment);
  for (let i = 0; i < shipmentIds.length; i++) {
    const sid = shipmentIds[i];
    check(byShipment[sid], (baseUsageByShipment || {})[sid], 'рейс ' + sid);
  }
  const scopedShipIds = Object.keys(baseUsageByShipmentSupplier || {});
  for (let i = 0; i < scopedShipIds.length; i++) {
    const sid = scopedShipIds[i];
    const supKeys = Object.keys(baseUsageByShipmentSupplier[sid] || {});
    for (let j = 0; j < supKeys.length; j++) {
      const supKey = supKeys[j];
      const skus = (byShipmentSupplier[sid] || {})[supKey] || [];
      if (!skus.length) continue;
      check(
        skus,
        baseUsageByShipmentSupplier[sid][supKey],
        'рейс ' + sid + ', поставщик "' + (supKey || '(пусто)') + '"'
      );
    }
  }

  if (!errors.length) return;
  const lines = [
    'Не удалось распределить расходы: у части SKU отсутствуют данные по базе аллокации (ALLOC_BASE).',
    ''
  ];
  for (let i = 0; i < errors.length; i++) {
    const e = errors[i];
    const skuList = e.missing.slice(0, 20).join(', ');
    const overflow = e.missing.length > 20 ? ' …и ещё ' + (e.missing.length - 20) : '';
    lines.push(
      '• ' + e.scope + ' — база ' + e.base + ' (' + baseLabel(e.base) + '): ' +
      'не заполнено у ' + e.missing.length + ' SKU: ' + skuList + overflow
    );
  }
  lines.push('');
  lines.push('Заполните пропуски в «Партии_в_рейсе» (Объем/Вес/Закупка) — или поменяйте ALLOC_BASE в «Затраты рейса» на доступную базу — и запустите пересчёт заново.');
  throw new Error(lines.join('\n'));
}

function costingNormalizeSupplierKey_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/[\s\u00A0]+/g, ' ')
    .replace(/[.,;:]+$/g, '')
    .trim();
}

function costingHealthCheckCore_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredRoles = [
    'BATCHES',
    'TRIP_EXPENSES',
    'CUSTOMS',
    'COST_SKU'
  ];
  const missing = requiredRoles.filter(function (role) {
    return !costingFindSheetByRole_(ss, role);
  });
  if (missing.length) {
    const labels = missing.map(function (role) {
      return COST_CFG.SHEETS[role];
    });
    throw new Error('Отсутствуют обязательные листы: ' + labels.join(', '));
  }

  const allocation = costingFindSheetByRole_(ss, 'ALLOCATION');
  if (allocation) {
    const values = allocation.getDataRange().getValues();
    if (!values.length) throw new Error('Лист "Аллокация затрат" пуст.');
    const header = values[0];
    costingFindCol_(header, ['SHIPMENT_ID', 'Shipment_ID', 'Рейс']);
    costingFindCol_(header, ['Артикул_ВБ', 'Артикул ВБ', 'WB_ARTICLE', 'Артикул']);
    costingFindCol_(header, ['Qty', 'QTY', 'Количество']);
    costingFindCol_(header, ['Закупка_руб', 'Закупка руб', 'Закупка']);
    costingFindCol_(header, ['Распределенные_расходы_руб', 'Распределенные расходы руб', 'Аллокация_руб', 'Аллокация']);
    return;
  }

  const batches = costingGetSheetByRole_(ss, 'BATCHES').getDataRange().getValues();
  if (batches.length < 2) throw new Error('Лист "Партии_в_рейсе" пуст.');
  const bHeader = batches[0];
  costingFindCol_(bHeader, ['SHIPMENT_ID', 'ID_рейса']);
  costingFindCol_(bHeader, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  costingFindCol_(bHeader, ['Количество', 'Qty', 'QTY']);
}

function costingSyncTripStatusesFromEvents_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trips = costingGetSheetByRole_(ss, 'TRIPS');
  const events = costingGetSheetByRole_(ss, 'EVENTS');
  if (!trips) throw new Error('Не найден лист "' + COST_CFG.SHEETS.TRIPS + '".');
  if (!events) throw new Error('Не найден лист "' + COST_CFG.SHEETS.EVENTS + '".');
  if (trips.getLastRow() < 2) throw new Error('Лист "' + COST_CFG.SHEETS.TRIPS + '" пуст.');

  const eventsData = events.getDataRange().getValues();
  if (eventsData.length < 2) {
    SpreadsheetApp.getUi().alert(
      'События рейса',
      'Лист "События_рейса" пуст — обновлять нечего.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const emap = costingHeaderMap_(eventsData[0]);
  const idxEventShipment = costingFirstIdx_(emap, ['ID_рейса', 'SHIPMENT_ID']);
  const idxEventType = costingFirstIdx_(emap, ['Тип_события', 'Тип события']);
  const idxEventDate = costingFirstIdx_(emap, ['Дата']);
  const idxEventComment = costingFirstIdx_(emap, ['Комментарий']);
  if (idxEventShipment == null || idxEventType == null || idxEventDate == null) {
    throw new Error('В "События_рейса" нужны колонки: ID_рейса, Тип_события, Дата.');
  }

  const orderMap = costingLoadEventOrderMap_(ss);
  const byShipment = {};
  for (let i = 1; i < eventsData.length; i++) {
    const r = eventsData[i];
    const shipmentId = String(r[idxEventShipment] || '').trim();
    if (!shipmentId) continue;
    const eventType = String(r[idxEventType] || '').trim();
    const eventDate = costingParseDate_(r[idxEventDate]);
    if (!eventDate) continue;
    const comment = idxEventComment != null ? String(r[idxEventComment] || '').trim() : '';
    const order = orderMap[costingNorm_(eventType)] || 0;

    if (!byShipment[shipmentId]) {
      byShipment[shipmentId] = {
        count: 0,
        latestDate: null,
        latestType: '',
        latestComment: '',
        latestOrder: 0
      };
    }
    const rec = byShipment[shipmentId];
    rec.count++;
    const laterDate = !rec.latestDate || eventDate.getTime() > rec.latestDate.getTime();
    const sameDateHigherOrder =
      rec.latestDate &&
      eventDate.getTime() === rec.latestDate.getTime() &&
      order > rec.latestOrder;
    if (laterDate || sameDateHigherOrder) {
      rec.latestDate = eventDate;
      rec.latestType = eventType;
      rec.latestComment = comment;
      rec.latestOrder = order;
    }
  }

  const tripsData = trips.getDataRange().getValues();
  const tmap = costingHeaderMap_(tripsData[0]);
  const idxTripShipment = costingFirstIdx_(tmap, ['SHIPMENT_ID', 'ID_рейса']);
  if (idxTripShipment == null) throw new Error('В "Рейсы" нужна колонка SHIPMENT_ID.');

  const idxEventsCount = costingEnsureHeaderColumn_(trips, tmap, 'Событий_всего');
  const idxLastDate = costingEnsureHeaderColumn_(trips, tmap, 'Дата_последнего_события');
  const idxLastType = costingEnsureHeaderColumn_(trips, tmap, 'Последний_тип_события');
  const idxLastComment = costingEnsureHeaderColumn_(trips, tmap, 'Последний_комментарий_события');
  const idxLastOrder = costingEnsureHeaderColumn_(trips, tmap, 'Порядок_последнего_события');
  const idxStatusByEvents = costingEnsureHeaderColumn_(trips, tmap, 'Статус_по_событиям');

  const refreshed = trips.getDataRange().getValues();
  const writes = [];
  for (let i = 1; i < refreshed.length; i++) {
    const row = refreshed[i];
    const shipmentId = String(row[idxTripShipment] || '').trim();
    if (!shipmentId || !byShipment[shipmentId]) continue;
    const rec = byShipment[shipmentId];
    writes.push({
      row: i + 1,
      values: [rec.count, rec.latestDate, rec.latestType, rec.latestComment, rec.latestOrder, rec.latestType]
    });
  }

  for (let i = 0; i < writes.length; i++) {
    const w = writes[i];
    trips.getRange(w.row, idxEventsCount + 1).setValue(w.values[0]);
    trips.getRange(w.row, idxLastDate + 1).setValue(w.values[1]);
    trips.getRange(w.row, idxLastType + 1).setValue(w.values[2]);
    trips.getRange(w.row, idxLastComment + 1).setValue(w.values[3]);
    trips.getRange(w.row, idxLastOrder + 1).setValue(w.values[4]);
    trips.getRange(w.row, idxStatusByEvents + 1).setValue(w.values[5]);
  }

  SpreadsheetApp.getUi().alert(
    'Статусы рейсов обновлены',
    'Обновлено рейсов: ' + writes.length,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function costingLoadEventOrderMap_(ss) {
  const sh = costingFindSheetByRole_(ss, 'EVENT_TYPES');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const map = costingHeaderMap_(data[0]);
  const idxName = costingFirstIdx_(map, ['Наименование']);
  const idxCode = costingFirstIdx_(map, ['Код']);
  const idxOrder = costingFirstIdx_(map, ['Порядок']);
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const ord = idxOrder != null ? costingToNumber_(r[idxOrder]) : 0;
    if (idxName != null) {
      const n = costingNorm_(r[idxName]);
      if (n) out[n] = ord;
    }
    if (idxCode != null) {
      const c = costingNorm_(r[idxCode]);
      if (c) out[c] = ord;
    }
  }
  return out;
}

function costingEnsureHeaderColumn_(sheet, headerMap, headerName) {
  const key = costingNorm_(headerName);
  if (headerMap[key] != null) return headerMap[key];
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(headerName);
  headerMap[key] = col - 1;
  return col - 1;
}

function costingHeaderMap_(headerRow) {
  const out = {};
  for (let i = 0; i < headerRow.length; i++) {
    out[costingNorm_(headerRow[i])] = i;
  }
  return out;
}

function costingFirstIdx_(headerMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const k = costingNorm_(aliases[i]);
    if (headerMap[k] != null) return headerMap[k];
  }
  return null;
}

function costingParseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) return null;
  const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return isNaN(dt.getTime()) ? null : dt;
}

function costingNorm_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function costingCopyMappings_(sourceSs, targetSs, mappings) {
  const batchSize = Math.max(100, costingGetNumberProp_(COST_CFG.PROPS.SYNC_BATCH_SIZE, COST_CFG.DEFAULTS.SYNC_BATCH_SIZE));
  const parts = [];

  mappings.forEach(function (m) {
    const sourceSheet = costingFindSheetByNames_(sourceSs, [m.source].concat(m.sourceAliases || []));
    if (!sourceSheet) {
      throw new Error('Не найден исходный лист. Проверены имена: ' + [m.source].concat(m.sourceAliases || []).join(', '));
    }
    let targetSheet = targetSs.getSheetByName(m.target);
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet(m.target);
    }
    const rows = sourceSheet.getLastRow();
    const cols = sourceSheet.getLastColumn();
    if (!rows || !cols) {
      parts.push(m.target + ': источник пуст, пропущено');
      return;
    }
    targetSheet.clearContents();
    for (let start = 1; start <= rows; start += batchSize) {
      const count = Math.min(batchSize, rows - start + 1);
      const values = sourceSheet.getRange(start, 1, count, cols).getDisplayValues();
      targetSheet.getRange(start, 1, count, cols).setValues(values);
    }
    parts.push(m.target + ': ' + rows + ' строк, ' + cols + ' колонок');
  });

  return parts.join('; ');
}

function costingMergeProductsBaseFrom04_(sourceSs, targetSs) {
  const sourceName = costingGetProp_(COST_CFG.PROPS.MASTER_PRODUCTS_SHEET_NAME, COST_CFG.DEFAULTS.MASTER_PRODUCTS_SHEET_NAME);
  const sourceSheet = costingFindSheetByNames_(sourceSs, [sourceName, 'Справочник товары']);
  if (!sourceSheet) {
    throw new Error('Не найден лист товаров в книге 04.');
  }

  let targetSheet = costingFindSheetByRole_(targetSs, 'PRODUCTS_REF');
  if (!targetSheet) {
    targetSheet = targetSs.insertSheet(COST_CFG.SHEETS.PRODUCTS_REF);
  }

  const sourceDataRaw = sourceSheet.getDataRange().getValues();
  const sourceTable = costingDetectHeaderTable_(sourceDataRaw, ['Артикул ВБ', 'Артикул поставщика', 'Штрихкод', 'ШК']);
  if (!sourceTable || sourceTable.rows.length < 1) {
    return COST_CFG.SHEETS.PRODUCTS_REF + ': источник пуст, пропущено';
  }

  const batchesSkuSet = costingLoadSkuSetFromBatches_(targetSs);
  if (!Object.keys(batchesSkuSet).length) {
    return COST_CFG.SHEETS.PRODUCTS_REF + ': в "Партии_в_рейсе" нет SKU, синк пропущен';
  }

  const targetData = targetSheet.getDataRange().getValues();
  const sourceHeader = sourceTable.header;
  let targetHeader = targetData.length ? targetData[0] : [];

  if (!targetHeader.length || targetHeader.every(function (x) { return String(x || '').trim() === ''; })) {
    targetHeader = ['Артикул ВБ', 'Артикул поставщика', 'ШК', 'Код ТН ВЭД', 'Тип пошлины', 'Ставка пошлины', 'Ставка НДС'];
    targetSheet.getRange(1, 1, 1, targetHeader.length).setValues([targetHeader]);
  }

  const sourceMap = costingHeaderMap_(sourceHeader);
  const targetMap = costingHeaderMap_(targetHeader);
  const fieldMappings = [
    { target: 'Артикул ВБ', sourceAliases: ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул'], targetAliases: ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул'] },
    { target: 'Артикул поставщика', sourceAliases: ['Артикул поставщика', 'Артикул_поставщика'], targetAliases: ['Артикул поставщика', 'Артикул_поставщика'] },
    { target: 'ШК', sourceAliases: ['ШК', 'Штрихкод', 'Barcode'], targetAliases: ['ШК', 'Штрихкод', 'Barcode'] },
    { target: 'Код ТН ВЭД', sourceAliases: ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE'], targetAliases: ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE'] },
    { target: 'Тип пошлины', sourceAliases: ['Тип пошлины', 'Тип_пошлины', 'Вид пошлины', 'Вид_пошлины'], targetAliases: ['Тип пошлины', 'Тип_пошлины', 'Вид пошлины', 'Вид_пошлины'] },
    { target: 'Ставка пошлины', sourceAliases: ['Пошлина', 'Ставка пошлины', 'Ставка пошлин', 'Ставка_пошлины', 'Размер пошлины'], targetAliases: ['Ставка пошлины', 'Ставка пошлин', 'Ставка_пошлины', 'Размер пошлины', 'Пошлина'] },
    { target: 'Ставка спец.пош', sourceAliases: ['Ставка спец.пош', 'Ставка спец.пошл', 'Ставка_спец_пош'], targetAliases: ['Ставка спец.пош', 'Ставка спец.пошл', 'Ставка_спец_пош'] },
    { target: 'База спец.пош', sourceAliases: ['База спец.пош', 'База спец.пошл', 'База_спец_пош'], targetAliases: ['База спец.пош', 'База спец.пошл', 'База_спец_пош'] },
    { target: 'Режим комбини', sourceAliases: ['Режим комбини', 'Режим комбинир', 'Режим комбинированной'], targetAliases: ['Режим комбини', 'Режим комбинир', 'Режим комбинированной'] },
    { target: 'Ставка НДС', sourceAliases: ['Ставка НДС', 'Ставка_НДС', 'НДС'], targetAliases: ['Ставка НДС', 'Ставка_НДС', 'НДС'] },
    { target: 'Площадь_ед', sourceAliases: ['Площадь_ед', 'Площадь ед', 'Площадь'], targetAliases: ['Площадь_ед', 'Площадь ед', 'Площадь'] }
  ];
  const srcIdxSku = costingFirstIdx_(sourceMap, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул']);
  if (srcIdxSku == null) {
    throw new Error('Не найдена колонка "Артикул ВБ" в источнике (книга 04).');
  }

  // Гарантируем колонки в локальном справочнике (если их нет — добавляем).
  for (let i = 0; i < fieldMappings.length; i++) {
    const field = fieldMappings[i];
    const exists = costingFirstIdx_(targetMap, field.targetAliases);
    if (exists == null) {
      targetHeader.push(field.target);
      targetMap[costingNorm_(field.target)] = targetHeader.length - 1;
    }
  }
  targetSheet.getRange(1, 1, 1, targetHeader.length).setValues([targetHeader]);
  const tgtIdxSku = costingFirstIdx_(targetMap, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул']);

  const current = targetSheet.getDataRange().getValues();
  const rows = current.length > 1 ? current.slice(1) : [];
  const rowBySku = {};
  for (let i = 0; i < rows.length; i++) {
    const sku = String(rows[i][tgtIdxSku] || '').trim();
    if (sku) rowBySku[sku] = i;
  }

  let updated = 0;
  let added = 0;
  let skippedNotInBatches = 0;
  for (let i = 0; i < sourceTable.rows.length; i++) {
    const srcRow = sourceTable.rows[i];
    const sku = String(srcRow[srcIdxSku] || '').trim();
    if (!sku) continue;
    if (!batchesSkuSet[sku]) {
      skippedNotInBatches++;
      continue;
    }
    let trgRow;
    let trgIdx = rowBySku[sku];
    if (trgIdx == null) {
      trgRow = new Array(targetHeader.length).fill('');
      rows.push(trgRow);
      trgIdx = rows.length - 1;
      rowBySku[sku] = trgIdx;
      added++;
    } else {
      trgRow = rows[trgIdx];
      updated++;
    }

    for (let f = 0; f < fieldMappings.length; f++) {
      const fm = fieldMappings[f];
      const sIdx = costingFirstIdx_(sourceMap, fm.sourceAliases);
      const tIdx = costingFirstIdx_(targetMap, fm.targetAliases);
      if (sIdx == null || tIdx == null) continue;
      // В справочник 05 переносим значения из 04 без преобразований.
      trgRow[tIdx] = srcRow[sIdx];
    }
  }

  if (rows.length) {
    targetSheet.getRange(2, 1, rows.length, targetHeader.length).setValues(rows);
  }

  return COST_CFG.SHEETS.PRODUCTS_REF + ': обновлено=' + updated + ', добавлено=' + added + ', пропущено_вне_партий=' + skippedNotInBatches;
}

function costingDetectHeaderTable_(allRows, requiredAliases) {
  if (!allRows || !allRows.length) return null;
  for (let r = 0; r < allRows.length; r++) {
    const header = allRows[r];
    if (!header || !header.length) continue;
    const hmap = costingHeaderMap_(header);
    let hasAny = false;
    for (let i = 0; i < requiredAliases.length; i++) {
      if (costingFirstIdx_(hmap, [requiredAliases[i]]) != null) {
        hasAny = true;
        break;
      }
    }
    if (!hasAny) continue;
    const rows = allRows.slice(r + 1).filter(function (row) {
      return row.some(function (v) { return String(v || '').trim() !== ''; });
    });
    return { header: header, rows: rows, headerRowIndex: r + 1 };
  }
  return null;
}

function costingLoadSkuSetFromBatches_(ss) {
  const sh = costingFindSheetByRole_(ss, 'BATCHES');
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxSku = costingFindColOptional_(h, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE', 'Артикул']);
  if (idxSku == null) return {};
  const set = {};
  for (let i = 1; i < data.length; i++) {
    const sku = String(data[i][idxSku] || '').trim();
    if (sku) set[sku] = true;
  }
  return set;
}

function costingFindSheetByNames_(ss, names) {
  const uniq = {};
  const wantedNorm = {};
  for (let i = 0; i < names.length; i++) {
    const nm = String(names[i] || '').trim();
    if (!nm || uniq[nm]) continue;
    uniq[nm] = true;
    const exact = ss.getSheetByName(nm);
    if (exact) return exact;
    wantedNorm[costingNormSheetName_(nm)] = true;
  }
  const all = ss.getSheets();
  for (let j = 0; j < all.length; j++) {
    const sh = all[j];
    if (wantedNorm[costingNormSheetName_(sh.getName())]) {
      return sh;
    }
  }
  return null;
}

function costingNormSheetName_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function costingFindCol_(headerRow, aliases) {
  const norm = headerRow.map(costingHeaderCanonForLookup_);
  for (let i = 0; i < aliases.length; i++) {
    const key = costingHeaderCanonForLookup_(aliases[i]);
    const idx = norm.indexOf(key);
    if (idx !== -1) return idx;
  }
  throw new Error('Не найдена колонка. Ожидались варианты: ' + aliases.join(', '));
}

function costingFindColOptional_(headerRow, aliases) {
  const norm = headerRow.map(costingHeaderCanonForLookup_);
  for (let i = 0; i < aliases.length; i++) {
    const key = costingHeaderCanonForLookup_(aliases[i]);
    const idx = norm.indexOf(key);
    if (idx !== -1) return idx;
  }
  return null;
}

/**
 * Единая канонизация заголовков для поиска колонок в costing.gs.
 * Делает совместимость с разными вариантами написания: «№», «.», «,», «/», «\», «-», «_», скобки и т.п.
 * Согласована с syncManagerCanonHeader_ в main.gs, чтобы не получать
 * расхождения между Сводной и Партии_в_рейсе.
 */
function costingHeaderCanonForLookup_(h) {
  return String(h == null ? '' : h)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/[№#]/g, ' ')
    .replace(/[.,:;()/\\\[\]{}'"“”«»]/g, ' ')
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function costingFindSheetByRole_(ss, role) {
  const primary = COST_CFG.SHEETS[role];
  const aliasesByRole = {
    BATCHES: ['Партии_в_рейсе', 'Партии в рейсе'],
    TRIP_EXPENSES: ['Затраты_рейса', 'Затраты_рейса ', 'Затраты рейса'],
    CUSTOMS: ['Таможенный_расчет', 'Таможенные платежи', 'Таможенный расчет'],
    CUSTOMS_FEE_RATES: ['Справочник_таможсбор', 'Справочник таможсбор', 'Справочник таможенного сбора'],
    ALLOCATION: ['Аллокация затрат', 'Аллокация_затрат'],
    COST_SKU: ['Себестоимость SKU', 'Себестоимость_SKU'],
    DECL_LINES: ['Декларации_строки', 'Декларации строки'],
    DECL_JOURNAL: ['Декларации_журнал', 'Декларации журнал', 'Загрузки_ДТ', 'Загрузки ДТ'],
    COST_SKU_FACT: ['Факт_себестоимость SKU', 'Факт себестоимость SKU'],
    PLAN_FACT_COMPARE: ['Сверка план vs факт', 'Сверка план факт', 'План vs факт'],
    TRIPS: ['Рейсы'],
    EVENTS: ['События_рейса', 'События рейса'],
    EVENT_TYPES: ['Типы_событий', 'Типы событий'],
    PRODUCTS_REF: ['Справочник товаров (05)', 'Справочник товары', 'Справочник_товары'],
    SUPPLIERS_REF: ['Справочник поставщики и условия', 'Справочник поставщики и условия работы']
  };
  const candidates = [primary].concat(aliasesByRole[role] || []);
  const exactOrAlias = costingFindSheetByNames_(ss, candidates);
  if (exactOrAlias) return exactOrAlias;

  // Фолбэк для старых/переименованных листов: ищем по ключевым словам.
  const keywordHints = {
    BATCHES: ['партии', 'рейсе'],
    TRIP_EXPENSES: ['затраты', 'рейса'],
    CUSTOMS: ['тамож'],
    CUSTOMS_FEE_RATES: ['таможсбор'],
    ALLOCATION: ['аллокац'],
    COST_SKU: ['себестоим', 'sku'],
    DECL_LINES: ['декларац', 'строк'],
    DECL_JOURNAL: ['декларац', 'журнал', 'загрузк', 'дт'],
    COST_SKU_FACT: ['факт', 'себестоим'],
    PLAN_FACT_COMPARE: ['сверк', 'план', 'факт'],
    TRIPS: ['рейс'],
    EVENTS: ['событ', 'рейс'],
    EVENT_TYPES: ['типы', 'событ'],
    PRODUCTS_REF: ['справочник', 'товар'],
    SUPPLIERS_REF: ['справочник', 'поставщик']
  };
  const hints = keywordHints[role] || [];
  if (hints.length) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const nameNorm = costingNormSheetName_(sheets[i].getName());
      let ok = true;
      for (let j = 0; j < hints.length; j++) {
        if (nameNorm.indexOf(hints[j]) === -1) {
          ok = false;
          break;
        }
      }
      if (ok) return sheets[i];
    }
  }

  return null;
}

function costingGetSheetByRole_(ss, role) {
  const sh = costingFindSheetByRole_(ss, role);
  if (!sh) {
    const sheetNames = ss.getSheets().map(function (x) { return x.getName(); });
    throw new Error(
      'Не найден лист для роли ' + role + ' (' + COST_CFG.SHEETS[role] + '). ' +
      'Доступные листы: ' + sheetNames.join(', ')
    );
  }
  return sh;
}

function costingToNumber_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = Number(String(v || '').replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function costingGetRequiredProp_(key) {
  const v = costingGetProp_(key, '');
  if (!v) throw new Error('Не заполнено Script Property: ' + key);
  return v;
}

function costingGetProp_(key, fallback) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (raw == null || String(raw).trim() === '') return fallback;
  return String(raw).trim();
}

function costingGetNumberProp_(key, fallback) {
  const n = Number(costingGetProp_(key, fallback));
  return isFinite(n) ? n : fallback;
}

/**
 * Открывает книгу 01 (Сводная-источник).
 * 1) Если в проекте 05 есть sync_hub.gs (функция syncHubOpenSpreadsheetForBook_) — используем его.
 * 2) Иначе самостоятельно читаем Script Property ORDERS_SPREADSHEET_ID, чистим её
 *    (URL → id, обрезаем мусор) и открываем по ID.
 * Это позволяет держать sync_hub.gs только в книге 04, не дублируя его в 05.
 */
function costingOpenBook01_() {
  if (typeof syncHubOpenSpreadsheetForBook_ === 'function') {
    return syncHubOpenSpreadsheetForBook_('01');
  }
  const raw = costingGetProp_('ORDERS_SPREADSHEET_ID', '');
  if (!raw) {
    throw new Error(
      'Не задан Script Property ORDERS_SPREADSHEET_ID в книге 05.\n' +
      'Откройте Расширения → Apps Script → ⚙ Свойства проекта → Свойства скрипта и добавьте:\n' +
      '  ORDERS_SPREADSHEET_ID = <ID или URL книги 01>'
    );
  }
  const id = costingExtractSpreadsheetId_(raw);
  if (!id) {
    throw new Error('ORDERS_SPREADSHEET_ID указан, но не удалось извлечь корректный ID: ' + raw);
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      'Не удалось открыть книгу 01 по ORDERS_SPREADSHEET_ID = ' + id + '\n' +
      'Проверьте доступ исполнителя скрипта к этому файлу и корректность ID.\n' +
      'Исходная ошибка: ' + (e && e.message ? e.message : String(e))
    );
  }
}

function costingExtractSpreadsheetId_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  const mUrl = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (mUrl && mUrl[1]) return mUrl[1];
  const token = s.split(/[?#&\s]/)[0];
  const mToken = token.match(/^([a-zA-Z0-9-_]{25,})/);
  if (mToken && mToken[1]) return mToken[1];
  return s;
}

/* =================================================================================
 * Пересборка «Партии_в_рейсе» (книга 05) из «Сводной» (книга 01) по SHIPMENT_ID.
 *
 * Архитектурно: «Сводная» (01) — источник истины операционных данных, «Партии_в_рейсе»
 * (05) — производная для расчёта себестоимости. Эта функция переносит позиции с
 * заполненным SHIPMENT_ID из «Сводной» в «Партии_в_рейсе».
 *
 * Поддерживаемые режимы:
 *   - 'safe' (по умолчанию): добавляет в Партии_в_рейсе только те (SHIPMENT_ID, Артикул ВБ),
 *     которых там нет. Существующие строки НЕ ТРОГАЕТ (защита ручных правок цен/курсов/ТНВЭД).
 *   - 'refresh_qty': то же что safe + для существующих строк обновляет ТОЛЬКО Количество и Объем.
 *     Все остальные поля (Стоимость, Цена, Валюта, Курс, ТНВЭД и т.п.) сохраняются.
 *   - 'full_rebuild': удаляет все строки в Партии_в_рейсе и пишет заново. Требует двойного
 *     подтверждения через UI (см. costingRebuildBatchesFullLiveWithConfirm).
 *
 * Защита: при любом боевом запуске (dryRun=false) автоматически делается снимок
 * Партии_в_рейсе → лист «Партии_в_рейсе_бэкап_YYYYMMDD_hhmm». Лог пишется в книгу 04
 * через syncHubLog_ (если функция доступна).
 *
 * @param {{ mode?: 'safe'|'refresh_qty'|'full_rebuild', dryRun?: boolean, targetSs?: GoogleAppsScript.Spreadsheet.Spreadsheet }} opts
 * @returns {Object} отчёт по операции
 * ============================================================================== */
function costingRebuildBatchesFromSummary_(opts) {
  opts = opts || {};
  const mode = opts.mode || 'safe';
  const dryRun = !!opts.dryRun;
  if (['safe', 'refresh_qty', 'full_rebuild'].indexOf(mode) === -1) {
    throw new Error('Неизвестный режим: ' + mode);
  }

  const sourceSs = costingOpenBook01_();
  const summarySh = sourceSs.getSheetByName('Сводная');
  if (!summarySh) throw new Error('В книге 01 не найден лист «Сводная».');

  const targetSs = opts.targetSs || SpreadsheetApp.getActiveSpreadsheet();
  const batchesSh = costingGetSheetByRole_(targetSs, 'BATCHES');

  const summaryRows = costingReadSummaryShipmentRows_(summarySh);

  const batchesData = batchesSh.getDataRange().getValues();
  if (batchesData.length === 0) {
    throw new Error('В листе «' + batchesSh.getName() + '» отсутствует шапка. Заполните строку 1.');
  }
  const batchesHeader = batchesData[0];
  const bIdxShipment = costingFindCol_(batchesHeader, ['SHIPMENT_ID', 'ID_рейса']);
  const bIdxSku = costingFindCol_(batchesHeader, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const bIdxQty = costingFindColOptional_(batchesHeader, ['Количество', 'Qty', 'QTY']);
  const bIdxVolume = costingFindColOptional_(batchesHeader, ['Объем', 'Объём', 'Volume']);

  const existingByKey = {};
  for (let r = 1; r < batchesData.length; r++) {
    const sh = String(batchesData[r][bIdxShipment] || '').trim();
    const sku = String(batchesData[r][bIdxSku] || '').trim();
    if (!sh || !sku) continue;
    existingByKey[sh + '|' + sku] = { rowIdx: r, row: batchesData[r] };
  }

  const tripsRegistry = costingLoadTripsRegistry_(targetSs);

  const added = [];
  const updated = [];
  const unchanged = [];
  const unknownTrSet = {};
  const reportByShipment = {};

  function reportInc_(shipId, bucket) {
    if (!reportByShipment[shipId]) {
      reportByShipment[shipId] = { added: 0, updated: 0, unchanged: 0, overwritten: 0, total: 0 };
    }
    reportByShipment[shipId][bucket]++;
    reportByShipment[shipId].total++;
  }

  for (let i = 0; i < summaryRows.length; i++) {
    const sr = summaryRows[i];
    const key = sr.shipmentId + '|' + sr.wbArticle;

    if (Object.keys(tripsRegistry).length && !tripsRegistry[sr.shipmentId]) {
      unknownTrSet[sr.shipmentId] = true;
    }

    if (mode === 'full_rebuild') {
      added.push({ row: sr, batchRow: costingBuildBatchRowFromSummary_(sr, batchesHeader) });
      reportInc_(sr.shipmentId, 'overwritten');
      continue;
    }

    const existing = existingByKey[key];
    if (!existing) {
      added.push({ row: sr, batchRow: costingBuildBatchRowFromSummary_(sr, batchesHeader) });
      reportInc_(sr.shipmentId, 'added');
      continue;
    }

    if (mode === 'refresh_qty') {
      const oldQty = bIdxQty != null ? costingToNumber_(existing.row[bIdxQty]) : 0;
      const newQty = costingToNumber_(sr.qty);
      const oldVolume = bIdxVolume != null ? costingToNumber_(existing.row[bIdxVolume]) : 0;
      const newVolume = costingToNumber_(sr.volume);
      const qtyDiff = Math.abs(oldQty - newQty) > 1e-9;
      const volDiff = Math.abs(oldVolume - newVolume) > 1e-9;
      if (qtyDiff || volDiff) {
        updated.push({
          row: sr,
          rowIdx: existing.rowIdx,
          newQty: newQty,
          newVolume: newVolume,
          oldQty: oldQty,
          oldVolume: oldVolume
        });
        reportInc_(sr.shipmentId, 'updated');
      } else {
        unchanged.push({ row: sr });
        reportInc_(sr.shipmentId, 'unchanged');
      }
    } else {
      // safe mode
      unchanged.push({ row: sr });
      reportInc_(sr.shipmentId, 'unchanged');
    }
  }

  // Превью первых добавляемых строк — чтобы в dry-run сразу увидеть, какие поля
  // (ШК, спецификация, коробки, в коробке, валюта) реально считались из Сводной.
  const previewSize = Math.min(added.length, 5);
  const addedPreview = [];
  for (let p = 0; p < previewSize; p++) {
    const a = added[p].row;
    addedPreview.push({
      shipmentId: a.shipmentId,
      wbArticle: a.wbArticle,
      barcode: a.barcode,
      spec: a.spec,
      qty: costingToNumber_(a.qty),
      qtyPerBox: a.qtyPerBox === '' || a.qtyPerBox == null ? '' : costingToNumber_(a.qtyPerBox),
      boxes: a.boxes === '' || a.boxes == null ? '' : costingToNumber_(a.boxes),
      price: costingToNumber_(a.price),
      amount: costingToNumber_(a.amount),
      currency: a.currency || ''
    });
  }

  const summary = {
    mode: mode,
    dryRun: dryRun,
    summarySourceCount: summaryRows.length,
    addedCount: added.length,
    updatedCount: updated.length,
    unchangedCount: unchanged.length,
    overwrittenCount: mode === 'full_rebuild' ? added.length : 0,
    unknownTrCount: Object.keys(unknownTrSet).length,
    unknownTrList: Object.keys(unknownTrSet).sort(),
    reportByShipment: reportByShipment,
    addedPreview: addedPreview
  };

  if (dryRun) return summary;

  if (summary.addedCount === 0 && summary.updatedCount === 0 && summary.overwrittenCount === 0) {
    summary.snapshotName = null;
    return summary;
  }

  summary.snapshotName = costingSnapshotBatches_(targetSs);

  if (mode === 'full_rebuild') {
    const totalRows = batchesSh.getLastRow();
    if (totalRows > 1) {
      batchesSh.getRange(2, 1, totalRows - 1, batchesHeader.length).clearContent();
    }
    const newRows = added.map(function (a) { return a.batchRow; });
    if (newRows.length) {
      batchesSh.getRange(2, 1, newRows.length, batchesHeader.length).setValues(newRows);
    }
  } else {
    if (added.length) {
      const startRow = batchesSh.getLastRow() + 1;
      const newRows = added.map(function (a) { return a.batchRow; });
      batchesSh.getRange(startRow, 1, newRows.length, batchesHeader.length).setValues(newRows);
    }
    if (mode === 'refresh_qty' && updated.length) {
      for (let u = 0; u < updated.length; u++) {
        const upd = updated[u];
        if (bIdxQty != null) {
          batchesSh.getRange(upd.rowIdx + 1, bIdxQty + 1).setValue(upd.newQty);
        }
        if (bIdxVolume != null) {
          batchesSh.getRange(upd.rowIdx + 1, bIdxVolume + 1).setValue(upd.newVolume);
        }
      }
    }
  }

  costingSortBatches_(batchesSh, batchesHeader);
  return summary;
}

/**
 * Чтение строк «Сводной» (книга 01) с непустым SHIPMENT_ID.
 * Шапка в строке 2, данные с 3-й (см. MANAGER_SUMMARY_SYNC_HEADER_ROW в книге 01).
 */
function costingReadSummaryShipmentRows_(summarySh) {
  const lr = summarySh.getLastRow();
  const lc = summarySh.getLastColumn();
  if (lr < 3 || lc < 1) return [];

  const headerRow = summarySh.getRange(2, 1, 1, lc).getValues()[0] || [];
  const idxShipment = costingFindColOptional_(headerRow, ['Рейс', 'SHIPMENT_ID', 'Shipment_ID', 'ID_рейса']);
  if (idxShipment == null) {
    throw new Error(
      'В «Сводной» книги 01 не найдена колонка «Рейс» / «SHIPMENT_ID». ' +
        'Добавьте колонку в манерные вкладки или в «История отгрузок» и пересоберите Сводную.'
    );
  }
  const idxWb = costingFindCol_(headerRow, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const idxSupArticle = costingFindColOptional_(headerRow, ['Артикул поставщика', 'Артикул_поставщика']);
  const idxBarcode = costingFindColOptional_(headerRow, ['ШК', 'Barcode', 'Штрихкод']);
  const idxQty = costingFindColOptional_(headerRow, ['Итоговое количество', 'Количество', 'Qty', 'QTY']);
  const idxQtyPerBox = costingFindColOptional_(headerRow, [
    'Количество в коробке',
    'Количество_в_коробке',
    'Кол-во в коробке',
    'Кол_во в коробке',
    'Штук в коробке',
    'Шт в коробке',
    'Единиц в коробке',
    'В коробке',
    'Qty per box',
    'Pcs per box',
    'Pcs box',
    'Per box'
  ]);
  const idxBoxes = costingFindColOptional_(headerRow, [
    'Коробки',
    'Кол-во коробок',
    'Кол_во коробок',
    'Количество коробок',
    'Число коробок',
    'Итого коробок',
    // Вариант с «коробов» (от слова «короб»), часто встречается в шапках логистов
    'Количество коробов',
    'Кол-во коробов',
    'Кол_во коробов',
    'Число коробов',
    'Итого коробов',
    'Boxes',
    'Box count',
    'Total boxes',
    'CTN',
    'CTNS',
    'Cartons'
  ]);
  const idxPrice = costingFindColOptional_(headerRow, ['Цена', 'Price']);
  const idxAmount = costingFindColOptional_(headerRow, ['Сумма', 'Amount']);
  const idxVolume = costingFindColOptional_(headerRow, ['Объем', 'Объём', 'Volume']);
  const idxWeight = costingFindColOptional_(headerRow, ['Вес', 'Weight']);
  const idxSpec = costingFindColOptional_(headerRow, [
    'Номер спецификации',
    'Номер_спецификации',
    'Спецификация',
    'Спецификации',
    'Номер спецификации /инвойса',
    'Номер спецификации/инвойса',
    'Спецификация /инвойса',
    'Спецификация/инвойса',
    'Номер инвойса',
    'Инвойс',
    'Spec',
    'Spec number',
    '№ спецификации',
    '№спецификации'
  ]);
  const idxCurrency = costingFindColOptional_(headerRow, ['Валюта', 'Currency', 'Curr']);
  const idxSupplierNote = costingFindColOptional_(headerRow, ['Поставщик']);
  const idxShipVia = costingFindColOptional_(headerRow, ['Отгрузка через', 'Отгрузка_через']);

  const data = summarySh.getRange(3, 1, lr - 2, lc).getValues();
  const dataDisplay = summarySh.getRange(3, 1, lr - 2, lc).getDisplayValues();
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const d = dataDisplay[i];
    // Пропускаем строки-разделители месяцев (стартуют с «▸»)
    const firstCell = String(r[0] || '').trim();
    if (firstCell.indexOf('▸') === 0) continue;
    const shipId = String(r[idxShipment] || '').trim();
    const wb = String(r[idxWb] || '').trim();
    if (!shipId || !wb) continue;
    rows.push({
      shipmentId: shipId,
      wbArticle: wb,
      supplierArticle: idxSupArticle != null ? String(r[idxSupArticle] || '').trim() : '',
      barcode: idxBarcode != null ? String(r[idxBarcode] || '').trim() : '',
      qty: idxQty != null ? r[idxQty] : 0,
      qtyPerBox: idxQtyPerBox != null ? r[idxQtyPerBox] : '',
      boxes: idxBoxes != null ? r[idxBoxes] : '',
      price: idxPrice != null ? r[idxPrice] : 0,
      amount: idxAmount != null ? r[idxAmount] : 0,
      volume: idxVolume != null ? r[idxVolume] : 0,
      weight: idxWeight != null ? r[idxWeight] : 0,
      spec: idxSpec != null ? (String(r[idxSpec] || '').trim() || String(d[idxSpec] || '').trim()) : '',
      currency: idxCurrency != null ? String(r[idxCurrency] || '').trim() : '',
      supplier: idxSupplierNote != null ? String(r[idxSupplierNote] || '').trim() : '',
      shipVia: idxShipVia != null ? String(r[idxShipVia] || '').trim() : ''
    });
  }
  return rows;
}

/**
 * Строит строку «Партии_в_рейсе» из row-объекта «Сводной».
 *
 * Контракт по валютам:
 *   • Все денежные поля в «Партии_в_рейсе» (Цена, Сумма итого / Стоимость) остаются
 *     в валюте заказа (обычно CNY). Перевод в рубли НЕ делается — это задача листа
 *     «Себестоимость SKU», который использует «Курс к RUB» уже на этапе расчёта.
 *   • Валюта по умолчанию — CNY. Перегружается Script Property BATCHES_DEFAULT_CURRENCY
 *     (значения 'CNY' / 'RUB' / 'USD' и т.п.). Если в строке Сводной есть колонка «Валюта»
 *     с непустым значением — она имеет приоритет (per-row override).
 *   • Курс к RUB при сборке из Сводной не заполняется — пересчёт берёт курс из траншей оплат
 *     или ручной Курс_к_RUB (legacy).
 */
function costingBuildBatchRowFromSummary_(sr, batchesHeader) {
  const row = new Array(batchesHeader.length).fill('');
  const idxShipment = costingFindCol_(batchesHeader, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindCol_(batchesHeader, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  const idxSupArticle = costingFindColOptional_(batchesHeader, ['Артикул поставщика', 'Артикул_поставщика']);
  const idxBarcode = costingFindColOptional_(batchesHeader, ['ШК', 'Barcode', 'Штрихкод']);
  const idxQty = costingFindColOptional_(batchesHeader, ['Количество', 'Qty', 'QTY']);
  const idxQtyPerBox = costingFindColOptional_(batchesHeader, [
    'Количество в коробке',
    'Количество_в_коробке',
    'Кол-во в коробке',
    'Кол_во в коробке',
    'Штук в коробке',
    'Шт в коробке',
    'Единиц в коробке',
    'В коробке',
    'Qty per box',
    'Pcs per box',
    'Pcs box',
    'Per box'
  ]);
  const idxBoxes = costingFindColOptional_(batchesHeader, [
    'Коробки',
    'Кол-во коробок',
    'Кол_во коробок',
    'Количество коробок',
    'Число коробок',
    'Итого коробок',
    // Вариант с «коробов»
    'Количество коробов',
    'Кол-во коробов',
    'Кол_во коробов',
    'Число коробов',
    'Итого коробов',
    'Boxes',
    'Box count',
    'Total boxes',
    'CTN',
    'CTNS',
    'Cartons'
  ]);
  // «Сумма итого» в «Партии_в_рейсе» — это qty × price в валюте партии (НЕ в рублях).
  // Поддерживаем разные варианты названия колонки в реальных листах.
  const idxAmount = costingFindColOptional_(batchesHeader, [
    'Сумма итого',
    'Сумма_итого',
    'Сумма',
    'Total amount',
    'Total',
    'Стоимость',
    'Закупка_руб',
    'Закупка руб'
  ]);
  const idxPrice = costingFindColOptional_(batchesHeader, ['Цена', 'Цена_RUB', 'Цена руб', 'Price']);
  const idxCurrency = costingFindColOptional_(batchesHeader, ['Валюта', 'Currency']);
  const idxShipVia = costingFindColOptional_(batchesHeader, [
    'Отгрузка через',
    'Отгрузка_через',
    'Supplier'
  ]);
  const idxSupplier = costingFindColOptional_(batchesHeader, ['Поставщик']);
  const idxVolume = costingFindColOptional_(batchesHeader, ['Объем', 'Объём', 'Volume']);
  const idxWeight = costingFindColOptional_(batchesHeader, ['Вес', 'Weight']);
  const idxSpec = costingFindColOptional_(batchesHeader, [
    'Номер спецификации',
    'Номер_спецификации',
    'Спецификация',
    'Спецификации',
    'Номер спецификации /инвойса',
    'Номер спецификации/инвойса',
    'Спецификация /инвойса',
    'Спецификация/инвойса',
    'Номер инвойса',
    'Инвойс',
    'Spec',
    'Spec number',
    '№ спецификации',
    '№спецификации'
  ]);

  const defaultCurrency = costingGetProp_('BATCHES_DEFAULT_CURRENCY', 'CNY') || 'CNY';
  const rowCurrency = String(sr.currency || '').trim().toUpperCase() || defaultCurrency.toUpperCase();

  row[idxShipment] = sr.shipmentId;
  row[idxSku] = sr.wbArticle;
  if (idxSupArticle != null) row[idxSupArticle] = sr.supplierArticle;
  if (idxBarcode != null) row[idxBarcode] = sr.barcode;
  if (idxSpec != null) row[idxSpec] = sr.spec;
  if (idxQty != null) row[idxQty] = costingToNumber_(sr.qty);
  if (idxQtyPerBox != null && sr.qtyPerBox !== '' && sr.qtyPerBox !== null && sr.qtyPerBox !== undefined) {
    row[idxQtyPerBox] = costingToNumber_(sr.qtyPerBox);
  }
  if (idxBoxes != null && sr.boxes !== '' && sr.boxes !== null && sr.boxes !== undefined) {
    row[idxBoxes] = costingToNumber_(sr.boxes);
  }
  // Сумма итого = qty × price из Сводной, остаётся в валюте заказа.
  if (idxAmount != null) row[idxAmount] = costingToNumber_(sr.amount);
  if (idxPrice != null) row[idxPrice] = costingToNumber_(sr.price);
  if (idxCurrency != null) row[idxCurrency] = rowCurrency;
  // Курс к RUB заполняет пользователь (или отдельный механизм с курсами).
  // Не подставляем «1» по умолчанию, чтобы потом в «Себестоимость SKU» не было
  // случайной конверсии CNY → RUB по ложному курсу.
  if (idxShipVia != null) row[idxShipVia] = sr.shipVia;
  if (idxSupplier != null) row[idxSupplier] = sr.supplier;
  if (idxVolume != null) row[idxVolume] = costingToNumber_(sr.volume);
  if (idxWeight != null) row[idxWeight] = costingToNumber_(sr.weight);

  return row;
}

function costingLoadTripsRegistry_(ss) {
  const sh = costingFindSheetByRole_(ss, 'TRIPS');
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  const idx = costingFindColOptional_(data[0], ['SHIPMENT_ID', 'ID_рейса', 'Рейс']);
  if (idx == null) return {};
  const reg = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idx] || '').trim();
    if (id) reg[id] = true;
  }
  return reg;
}

/**
 * Делает копию «Партии_в_рейсе» с именем «Партии_в_рейсе_бэкап_YYYYMMDD_hhmm» в той же книге.
 * Возвращает имя созданного листа.
 */
function costingSnapshotBatches_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = costingGetSheetByRole_(ss, 'BATCHES');
  const now = new Date();
  const stamp =
    now.getFullYear() +
    String(now.getMonth() + 1).padStart(2, '0') +
    String(now.getDate()).padStart(2, '0') +
    '_' +
    String(now.getHours()).padStart(2, '0') +
    String(now.getMinutes()).padStart(2, '0');
  const base = COST_CFG.SHEETS.BATCHES + '_бэкап_' + stamp;
  let finalName = base;
  let suffix = 2;
  while (ss.getSheetByName(finalName)) {
    finalName = base + '_' + suffix;
    suffix++;
  }
  const copied = sh.copyTo(ss);
  copied.setName(finalName);
  try {
    ss.setActiveSheet(copied);
    ss.moveActiveSheet(ss.getNumSheets());
  } catch (e) {}
  return finalName;
}

function costingSortBatches_(batchesSh, batchesHeader) {
  const lr = batchesSh.getLastRow();
  if (lr < 3) return;
  const idxShipment = costingFindCol_(batchesHeader, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindCol_(batchesHeader, ['Артикул ВБ', 'Артикул_ВБ', 'WB_ARTICLE']);
  try {
    batchesSh.getRange(2, 1, lr - 1, batchesHeader.length).sort([
      { column: idxShipment + 1, ascending: true },
      { column: idxSku + 1, ascending: true }
    ]);
  } catch (e) {
    // Если сортировка не удалась — не критично, оставляем как есть.
  }
}

/* ------------------ Пункты меню для пересборки ------------------ */

function costingRebuildBatchesSafeDryRun() {
  costingRebuildBatchesMenuAction_({ mode: 'safe', dryRun: true });
}
function costingRebuildBatchesSafeLive() {
  costingRebuildBatchesMenuAction_({ mode: 'safe', dryRun: false });
}
function costingRebuildBatchesRefreshQtyDryRun() {
  costingRebuildBatchesMenuAction_({ mode: 'refresh_qty', dryRun: true });
}
function costingRebuildBatchesRefreshQtyLive() {
  costingRebuildBatchesMenuAction_({ mode: 'refresh_qty', dryRun: false });
}
function costingRebuildBatchesFullDryRun() {
  costingRebuildBatchesMenuAction_({ mode: 'full_rebuild', dryRun: true });
}
function costingRebuildBatchesFullLiveWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.alert(
    '⚠️ ПОЛНАЯ ПЕРЕСБОРКА «Партии_в_рейсе»',
    'Эта операция удалит ВСЕ строки в «Партии_в_рейсе» и запишет заново из «Сводной» книги 01.\n' +
      'Все ручные правки цен, валют, курсов, ТНВЭД для уже занесённых рейсов БУДУТ ПОТЕРЯНЫ.\n\n' +
      'Перед операцией будет создан автоматический снимок.\n\nПродолжить?',
    ui.ButtonSet.YES_NO
  );
  if (r1 !== ui.Button.YES) return;
  const r2 = ui.alert(
    '⚠️ Повторное подтверждение',
    'Точно полностью пересобрать «Партии_в_рейсе»? Восстановление возможно только из снимка.',
    ui.ButtonSet.YES_NO
  );
  if (r2 !== ui.Button.YES) return;
  costingRebuildBatchesMenuAction_({ mode: 'full_rebuild', dryRun: false });
}

function costingSnapshotBatchesMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const name = costingSnapshotBatches_(SpreadsheetApp.getActiveSpreadsheet());
    ui.alert('✅ Снимок создан', 'Лист: ' + name, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Не удалось создать снимок', e.message || String(e), ui.ButtonSet.OK);
  }
}

function costingRebuildBatchesMenuAction_(opts) {
  const ui = SpreadsheetApp.getUi();
  try {
    const report = costingRebuildBatchesFromSummary_(opts);
    const modeLabels = {
      safe: 'безопасный',
      refresh_qty: 'обновление количеств',
      full_rebuild: 'полная пересборка'
    };
    const lines = [];
    lines.push('Режим: ' + modeLabels[opts.mode] + (opts.dryRun ? ' (DRY-RUN)' : ' (БОЕВОЙ)'));
    lines.push('Строк в «Сводной» с SHIPMENT_ID: ' + report.summarySourceCount);
    lines.push('');
    if (opts.mode === 'full_rebuild') {
      lines.push('Будет перезаписано строк: ' + report.overwrittenCount);
    } else {
      lines.push('Добавится: ' + report.addedCount);
      if (opts.mode === 'refresh_qty') {
        lines.push('Обновится (qty/объём): ' + report.updatedCount);
      }
      lines.push('Без изменений: ' + report.unchangedCount);
    }
    lines.push('');
    lines.push('По рейсам:');
    const ships = Object.keys(report.reportByShipment).sort();
    if (!ships.length) {
      lines.push('  (нет рейсов в «Сводной»)');
    } else {
      for (let i = 0; i < ships.length; i++) {
        const r = report.reportByShipment[ships[i]];
        const parts = [];
        if (r.added) parts.push('добавится ' + r.added);
        if (r.updated) parts.push('обновится ' + r.updated);
        if (r.unchanged) parts.push('без изм. ' + r.unchanged);
        if (r.overwritten) parts.push('перезапишется ' + r.overwritten);
        lines.push('  ' + ships[i] + ': ' + parts.join(', '));
      }
    }
    if (report.unknownTrCount) {
      lines.push('');
      lines.push(
        '⚠️ Не найдены в листе «Рейсы»: ' +
          report.unknownTrList.join(', ') +
          ' — проверьте формат или заведите рейс в «Рейсы».'
      );
    }
    if (report.addedPreview && report.addedPreview.length) {
      lines.push('');
      lines.push('Превью первых добавляемых строк (что реально считалось из Сводной):');
      for (let p = 0; p < report.addedPreview.length; p++) {
        const a = report.addedPreview[p];
        lines.push(
          '  • ' + a.shipmentId + ' / ' + a.wbArticle +
            ' | ШК: ' + (a.barcode || '∅') +
            ' | Спец: ' + (a.spec || '∅') +
            ' | qty: ' + a.qty +
            ' | в коробке: ' + (a.qtyPerBox === '' ? '∅' : a.qtyPerBox) +
            ' | коробки: ' + (a.boxes === '' ? '∅' : a.boxes) +
            ' | цена: ' + a.price +
            ' | сумма: ' + a.amount +
            ' | вал: ' + (a.currency || '(default)')
        );
      }
      if (report.addedCount > report.addedPreview.length) {
        lines.push('  … и ещё ' + (report.addedCount - report.addedPreview.length) + ' строк');
      }
    }
    if (!opts.dryRun && report.snapshotName) {
      lines.push('');
      lines.push('Снимок создан: ' + report.snapshotName);
    } else if (!opts.dryRun && !report.snapshotName) {
      lines.push('');
      lines.push('Изменений не было — снимок не создавался.');
    }

    if (typeof syncHubLog_ === 'function') {
      try {
        syncHubLog_(
          'costing.rebuildBatchesFromSummary',
          opts.dryRun ? 'DRY-RUN' : 'OK',
          'mode=' +
            opts.mode +
            '; source=' +
            report.summarySourceCount +
            '; added=' +
            report.addedCount +
            '; updated=' +
            report.updatedCount +
            '; unchanged=' +
            report.unchangedCount +
            '; overwritten=' +
            report.overwrittenCount +
            '; unknown_tr=' +
            report.unknownTrCount +
            (report.snapshotName ? '; snapshot=' + report.snapshotName : ''),
          opts.dryRun
        );
      } catch (e) {}
    }

    ui.alert(
      opts.dryRun ? '🧪 Dry-run завершён' : '✅ Пересборка завершена',
      lines.join('\n'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    if (typeof syncHubLog_ === 'function') {
      try {
        syncHubLog_(
          'costing.rebuildBatchesFromSummary',
          'ERROR',
          'mode=' + opts.mode + '; error=' + (e.message || String(e)),
          opts.dryRun
        );
      } catch (logErr) {}
    }
    ui.alert('❌ Ошибка', e.message || String(e), ui.ButtonSet.OK);
    throw e;
  }
}
