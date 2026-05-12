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
    SUPPLIERS_REF: 'Справочник поставщики и условия'
  }
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
    const report = costingBuildSkuCostRows_();
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
  const report = costingBuildSkuCostRows_(shipmentId);
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
    SpreadsheetApp.getUi().alert(
      '✅ Пересчет завершен',
      'Записано строк: ' + report.rows.length + '\n' +
      'SKU: ' + report.skuCount + '\n' +
      'Рейсов: ' + report.shipmentCount,
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
    'Себестоимость_ед_руб'
  ]);
}

function costingUpsertShipmentRowsToCostSku_(sheet, shipmentId, incomingRows, incomingHeader) {
  const data = sheet.getDataRange().getValues();
  const existingHeader = data.length ? data[0] : [];
  const targetHeader = existingHeader.length ? existingHeader.slice() : incomingHeader.slice();
  const targetNorm = targetHeader.map(function (h) { return costingNorm_(h); });

  for (let i = 0; i < incomingHeader.length; i++) {
    const nh = costingNorm_(incomingHeader[i]);
    if (targetNorm.indexOf(nh) === -1) {
      targetHeader.push(incomingHeader[i]);
      targetNorm.push(nh);
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

function costingBuildSkuCostRows_(shipmentFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = costingFindSheetByRole_(ss, 'ALLOCATION');
  if (!sh) {
    return costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter);
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

function costingBuildSkuCostRowsFromBatchesAndExpenses_(ss, shipmentFilter) {
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
  const idxCost = costingFindColOptional_(bh, ['Стоимость', 'Закупка_руб', 'Закупка руб']);
  const idxPrice = costingFindColOptional_(bh, ['Цена', 'Цена_RUB', 'Цена руб']);
  const idxCurrency = costingFindColOptional_(bh, ['Валюта', 'Currency']);
  const idxRate = costingFindColOptional_(bh, ['Курс_к_RUB', 'Курс к RUB', 'Курс', 'Rate']);
  const idxTnved = costingFindColOptional_(bh, ['Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE', 'HS_CODE']);
  const idxShipVia = costingFindColOptional_(bh, ['Отгрузка через', 'Отгрузка_через', 'Поставщик', 'Supplier']);
  const idxVolume = costingFindColOptional_(bh, ['Объем', 'Объём', 'Volume']);
  const dutyRulesByTnved = costingLoadDutyRulesByTnved_(ss);
  const vatRatesByTnved = costingLoadVatRatesByTnved_(ss);
  const customsFeeRules = costingLoadCustomsFeeRules_(ss);
  const skuToTnved = costingLoadSkuToTnvedMap_(ss);
  const euroRate = costingGetEuroRateOnDate_(new Date());

  const expensePoolsByShipment = {};
  const expensePoolsByShipmentSupplier = {};
  if (expensesData.length > 1) {
    const eh = expensesData[0];
    const eShipment = costingFindCol_(eh, ['SHIPMENT_ID', 'ID_рейса']);
    const eScenario = costingFindColOptional_(eh, ['Сценарий']);
    const eInSku = costingFindColOptional_(eh, ['IN_SKU_COST']);
    const eSumRub = costingFindColOptional_(eh, ['Сумма_RUB', 'Сумма RUB']);
    const eSum = costingFindColOptional_(eh, ['Сумма']);
    const eOnlyShipVia = costingFindColOptional_(eh, ['Только_отгрузка_через', 'Только отгрузка через']);
    for (let i = 1; i < expensesData.length; i++) {
      const r = expensesData[i];
      const shipmentId = String(r[eShipment] || '').trim();
      if (!shipmentId) continue;
      if (shipmentFilter && shipmentId !== shipmentFilter) continue;
      if (eScenario != null && String(r[eScenario] || '').trim().toUpperCase() === 'FACT') continue;
      const sumRub = eSumRub != null ? costingToNumber_(r[eSumRub]) : costingToNumber_(r[eSum]);
      if (!sumRub) continue;
      const article = costingCellText_(r, costingFindColOptional_(eh, ['Статья_затрат', 'Статья затрат', 'Статья']));
      const eventType = costingCellText_(r, costingFindColOptional_(eh, ['Тип_событий', 'Тип события']));
      const scope = costingCellText_(r, costingFindColOptional_(eh, ['COST_SCOPE', 'Scope']));
      const cat = costingClassifyExpense_(article, eventType, scope);
      const onlyShipViaRaw = eOnlyShipVia != null ? String(r[eOnlyShipVia] || '') : '';
      const onlyShipViaKey = costingNormalizeSupplierKey_(onlyShipViaRaw);

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
      if (onlyShipViaKey) {
        if (!expensePoolsByShipmentSupplier[shipmentId]) expensePoolsByShipmentSupplier[shipmentId] = {};
        if (!expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey]) {
          expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey] = costingMakeEmptyExpensePool_();
        }
        targetPool = expensePoolsByShipmentSupplier[shipmentId][onlyShipViaKey];
      }

      if (cat === 'freight') targetPool.freightRub += sumRub;
      else if (cat === 'duty') targetPool.dutyRub += sumRub;
      else if (cat === 'vat') targetPool.vatRub += sumRub;
      else if (cat === 'customsFee') targetPool.customsFeeRub += sumRub;
      else if (cat === 'logisticsOther') targetPool.logisticsOtherRub += sumRub;
      else if (costingIsPostBorder_(scope)) targetPool.otherPostRub += sumRub;
      else targetPool.otherPreRub += sumRub;
      if (cat !== 'freight' && cat !== 'duty' && cat !== 'vat' && cat !== 'customsFee' && cat !== 'logisticsOther') {
        const key = article || '(без статьи)';
        targetPool.otherByArticle[key] = (targetPool.otherByArticle[key] || 0) + sumRub;
      }
    }
  }

  // Дополняем пулы таможенными значениями из отдельного листа "Таможенный_расчет".
  const customsPoolsByShipment = costingLoadCustomsPoolsByShipment_(ss, shipmentFilter);
  const customsPoolKeys = Object.keys(customsPoolsByShipment);
  for (let i = 0; i < customsPoolKeys.length; i++) {
    const shipmentId = customsPoolKeys[i];
    if (!expensePoolsByShipment[shipmentId]) {
      expensePoolsByShipment[shipmentId] = costingMakeEmptyExpensePool_();
    }
    const target = expensePoolsByShipment[shipmentId];
    const src = customsPoolsByShipment[shipmentId];
    target.dutyRub += src.dutyRub || 0;
    target.vatRub += src.vatRub || 0;
    // Таможенный сбор считаем по справочнику "Справочник_таможсбор" (ниже), не подмешиваем сюда.
  }

  const rows = [];
  const skuSet = {};
  const shipmentSet = {};
  const qtyByShipment = {};
  const qtyByShipmentSupplier = {};
  const staged = [];
  const additionalExpenseColumnSet = {};

  for (let i = 1; i < batchesData.length; i++) {
    const r = batchesData[i];
    const shipmentId = String(r[idxShipment] || '').trim();
    const sku = String(r[idxSku] || '').trim();
    if (!shipmentId || !sku) continue;
    if (shipmentFilter && shipmentId !== shipmentFilter) continue;
    const qty = costingToNumber_(r[idxQty]);
    if (!qty) continue;
    const amountRaw = idxCost != null ? costingToNumber_(r[idxCost]) : costingToNumber_(r[idxPrice]) * qty;
    const currencyRaw = idxCurrency != null ? String(r[idxCurrency] || '').trim().toUpperCase() : 'RUB';
    const currency = currencyRaw || 'RUB';
    const rateRaw = idxRate != null ? costingToNumber_(r[idxRate]) : 0;
    const fxRate = currency === 'RUB'
      ? 1
      : (rateRaw > 0 ? rateRaw : costingGetFxRateOnDate_(currency, new Date()));
    const purchaseRub = amountRaw * fxRate;
    qtyByShipment[shipmentId] = (qtyByShipment[shipmentId] || 0) + qty;
    const tnvedRaw = idxTnved != null ? r[idxTnved] : '';
    const tnved = costingNormalizeTnved_(tnvedRaw) || skuToTnved[sku] || '';
    const dutyRule = dutyRulesByTnved[tnved] || { type: '', rate: 0 };
    const dutyRate = costingNormalizeRate_(dutyRule.rate);
    const dutyType = costingNorm_(dutyRule.type);
    const vatRate = costingNormalizeRate_(vatRatesByTnved[tnved] || 0);
    // Правило:
    // - стоимостная: пошлина = таможенная стоимость * ставка
    // - весовая: пошлина = ставка * количество * курс EUR
    // В текущей модели "таможенная стоимость" = закупка_руб по строке партии.
    let dutyRub = 0;
    if (dutyType.indexOf('стоимост') !== -1) {
      dutyRub = purchaseRub * dutyRate;
    } else if (dutyType.indexOf('весов') !== -1) {
      dutyRub = dutyRate * qty * euroRate;
    } else {
      // fallback: если тип не заполнен, трактуем ставку как стоимостную.
      dutyRub = purchaseRub * dutyRate;
    }

    staged.push({
      shipmentId: shipmentId,
      sku: sku,
      supplierArticle: idxSupplierArticle != null ? String(r[idxSupplierArticle] || '').trim() : '',
      tnved: tnved,
      shipVia: idxShipVia != null ? String(r[idxShipVia] || '').trim() : '',
      shipViaKey: costingNormalizeSupplierKey_(idxShipVia != null ? String(r[idxShipVia] || '').trim() : ''),
      volume: idxVolume != null ? costingToNumber_(r[idxVolume]) : 0,
      qty: qty,
      purchaseRub: purchaseRub,
      dutyRub: dutyRub,
      vatRate: vatRate
    });
    const supplierQtyKey = shipmentId + '||' + costingNormalizeSupplierKey_(idxShipVia != null ? String(r[idxShipVia] || '').trim() : '');
    qtyByShipmentSupplier[supplierQtyKey] = (qtyByShipmentSupplier[supplierQtyKey] || 0) + qty;
  }

  // Подготовка базы распределения: до-границы распределяем по объему (fallback: Qty),
  // таможенный сбор считаем по каждой компании (поле "Отгрузка через") от таможенной стоимости.
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

  // Считаем таможенную стоимость по каждой компании и подбираем таможсбор из справочника.
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    const agg = supplierAggByShipment[x.shipmentId];
    const poolGlobal = expensePoolsByShipment[x.shipmentId] || costingMakeEmptyExpensePool_();
    const poolScoped = expensePoolsByShipmentSupplier[x.shipmentId] && expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      ? expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      : costingMakeEmptyExpensePool_();
    const shareGlobal = agg.totalVolume > 0 ? x.volume / agg.totalVolume : (agg.totalQty > 0 ? x.qty / agg.totalQty : 0);
    const supplierQtyKey = x.shipmentId + '||' + x.shipViaKey;
    const supplierQty = qtyByShipmentSupplier[supplierQtyKey] || 0;
    const shareScoped = supplierQty > 0 ? (x.qty / supplierQty) : 0;
    const preBorderRowRub =
      poolGlobal.freightRub * shareGlobal +
      poolGlobal.otherPreRub * shareGlobal +
      poolScoped.freightRub * shareScoped +
      poolScoped.otherPreRub * shareScoped;
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
      bySupplier[supplierKey].customsFeeRub = costingResolveCustomsFeeRub_(bySupplier[supplierKey].customsBaseRub, customsFeeRules);
    }
  }

  const now = new Date();
  for (let i = 0; i < staged.length; i++) {
    const x = staged[i];
    const poolGlobal = expensePoolsByShipment[x.shipmentId] || costingMakeEmptyExpensePool_();
    const poolScoped = expensePoolsByShipmentSupplier[x.shipmentId] && expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      ? expensePoolsByShipmentSupplier[x.shipmentId][x.shipViaKey]
      : costingMakeEmptyExpensePool_();
    const totalQty = qtyByShipment[x.shipmentId] || 0;
    const kGlobal = totalQty > 0 ? (x.qty / totalQty) : 0;
    const supplierQtyKey = x.shipmentId + '||' + x.shipViaKey;
    const supplierQty = qtyByShipmentSupplier[supplierQtyKey] || 0;
    const kScoped = supplierQty > 0 ? (x.qty / supplierQty) : 0;
    const freightRub = poolGlobal.freightRub * kGlobal + poolScoped.freightRub * kScoped;
    const logisticsOtherRub = poolGlobal.logisticsOtherRub * kGlobal + poolScoped.logisticsOtherRub * kScoped;
    const dutyFromExpensesRub = poolGlobal.dutyRub * kGlobal + poolScoped.dutyRub * kScoped;
    const vatFromExpensesRub = poolGlobal.vatRub * kGlobal + poolScoped.vatRub * kScoped;
    const otherPreRub = poolGlobal.otherPreRub * kGlobal + poolScoped.otherPreRub * kScoped;
    const otherPostRub = poolGlobal.otherPostRub * kGlobal + poolScoped.otherPostRub * kScoped;
    const otherRub = otherPreRub + otherPostRub;
    const customsBaseRub = x.purchaseRub + freightRub + otherPreRub;
    const preBorderExpensesRub = freightRub + otherPreRub;
    const supAgg = supplierAggByShipment[x.shipmentId] ? supplierAggByShipment[x.shipmentId].bySupplier[x.shipViaKey] : null;
    const supShare = supAgg
      ? (supAgg.volume > 0 ? (x.volume / supAgg.volume) : (supAgg.qty > 0 ? (x.qty / supAgg.qty) : 0))
      : 0;
    const customsFeeRub = supAgg ? (supAgg.customsFeeRub * supShare) : 0;
    const dutyRub = x.dutyRub + dutyFromExpensesRub;
    const vatFromRateRub = (customsBaseRub + dutyRub) * x.vatRate;
    const vatRub = vatFromRateRub + vatFromExpensesRub;
    const articleBreakdown = costingAllocateArticleMap_(
      poolGlobal.otherByArticle,
      kGlobal,
      poolScoped.otherByArticle,
      kScoped
    );
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
      unit
      ],
      byArticle: articleBreakdown
    });
    skuSet[x.sku] = true;
    shipmentSet[x.shipmentId] = true;
  }

  const additionalExpenseColumns = Object.keys(additionalExpenseColumnSet).sort();
  const finalizedRows = rows.map(function (rowObj) {
    const base = rowObj.base.slice(0, 15);
    const totals = rowObj.base.slice(15);
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
    additionalExpenseColumns: additionalExpenseColumns
  };
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
  return String(v == null ? '' : v).replace(/[^\d]/g, '');
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

function costingGetEuroRateOnDate_(dateObj) {
  return costingGetFxRateOnDate_('EUR', dateObj);
}

function costingGetFxRateOnDate_(currencyCode, dateObj) {
  try {
    const target = String(currencyCode || '').trim().toUpperCase();
    if (!target || target === 'RUB') return 1;
    const d = dateObj instanceof Date ? dateObj : new Date();
    const dd = ('0' + d.getDate()).slice(-2);
    const mm = ('0' + (d.getMonth() + 1)).slice(-2);
    const yyyy = d.getFullYear();
    const url = 'https://www.cbr.ru/scripts/XML_daily.asp?date_req=' + dd + '/' + mm + '/' + yyyy;
    const xml = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText('windows-1251');
    const doc = XmlService.parse(xml);
    const root = doc.getRootElement();
    const vals = root.getChildren('Valute');
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
    // fallback ниже
  }
  // Мягкий fallback: если курс не получили, считаем 1:1, чтобы не падал пересчет.
  return 1;
}

function costingLoadCustomsPoolsByShipment_(ss, shipmentFilter) {
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
    if (idxScenario != null && String(r[idxScenario] || '').trim().toUpperCase() === 'FACT') continue;
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
  return {
    freightRub: 0,
    logisticsOtherRub: 0,
    dutyRub: 0,
    vatRub: 0,
    customsFeeRub: 0,
    otherPreRub: 0,
    otherPostRub: 0,
    otherByArticle: {}
  };
}

function costingComposeExpensePool_(basePool, scopedPool) {
  const b = basePool || costingMakeEmptyExpensePool_();
  if (!scopedPool) return b;
  const out = {
    freightRub: (b.freightRub || 0) + (scopedPool.freightRub || 0),
    logisticsOtherRub: (b.logisticsOtherRub || 0) + (scopedPool.logisticsOtherRub || 0),
    dutyRub: (b.dutyRub || 0) + (scopedPool.dutyRub || 0),
    vatRub: (b.vatRub || 0) + (scopedPool.vatRub || 0),
    customsFeeRub: (b.customsFeeRub || 0) + (scopedPool.customsFeeRub || 0),
    otherPreRub: (b.otherPreRub || 0) + (scopedPool.otherPreRub || 0),
    otherPostRub: (b.otherPostRub || 0) + (scopedPool.otherPostRub || 0),
    otherByArticle: {}
  };
  const baseKeys = Object.keys(b.otherByArticle || {});
  for (let i = 0; i < baseKeys.length; i++) {
    const k = baseKeys[i];
    out.otherByArticle[k] = (out.otherByArticle[k] || 0) + costingToNumber_(b.otherByArticle[k]);
  }
  const scopedKeys = Object.keys(scopedPool.otherByArticle || {});
  for (let i = 0; i < scopedKeys.length; i++) {
    const k = scopedKeys[i];
    out.otherByArticle[k] = (out.otherByArticle[k] || 0) + costingToNumber_(scopedPool.otherByArticle[k]);
  }
  return out;
}

function costingNormalizeSupplierKey_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/[\s\u00A0]+/g, ' ')
    .replace(/[.,;:]+$/g, '')
    .trim();
}

function costingAllocateArticleMap_(globalMap, globalK, scopedMap, scopedK) {
  const out = {};
  const g = globalMap || {};
  const s = scopedMap || {};
  const gKeys = Object.keys(g);
  for (let i = 0; i < gKeys.length; i++) {
    const key = gKeys[i];
    out[key] = (out[key] || 0) + costingToNumber_(g[key]) * globalK;
  }
  const sKeys = Object.keys(s);
  for (let i = 0; i < sKeys.length; i++) {
    const key = sKeys[i];
    out[key] = (out[key] || 0) + costingToNumber_(s[key]) * scopedK;
  }
  return out;
}

function costingFormatAllocatedArticleBreakdown_(articleMap, k) {
  if (!articleMap) return '';
  const keys = Object.keys(articleMap);
  if (!keys.length || !k) return '';
  const parts = [];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const allocated = costingToNumber_(articleMap[key]) * k;
    if (!allocated) continue;
    parts.push({ key: key, value: allocated });
  }
  if (!parts.length) return '';
  parts.sort(function (a, b) { return Math.abs(b.value) - Math.abs(a.value); });
  return parts.map(function (x) {
    return x.key + ': ' + (Math.round(x.value * 100) / 100);
  }).join(' | ');
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
 * @param {{ mode?: 'safe'|'refresh_qty'|'full_rebuild', dryRun?: boolean }} opts
 * @returns {Object} отчёт по операции
 * ============================================================================== */
function costingRebuildBatchesFromSummary_(opts) {
  const mode = (opts && opts.mode) || 'safe';
  const dryRun = !!(opts && opts.dryRun);
  if (['safe', 'refresh_qty', 'full_rebuild'].indexOf(mode) === -1) {
    throw new Error('Неизвестный режим: ' + mode);
  }

  const sourceSs = costingOpenBook01_();
  const summarySh = sourceSs.getSheetByName('Сводная');
  if (!summarySh) throw new Error('В книге 01 не найден лист «Сводная».');

  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
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
 *   • Курс к RUB не заполняется автоматически — пользователь проставляет его руками
 *     для конкретной партии (это конверсионный курс к моменту закупки/оплаты).
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
