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