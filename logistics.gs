/**
 * Книга 05 — учёт логистики: режимы доставки, этапы рейса, нормативы, ETA.
 * Зависит от costing.gs (общие хелперы листов/дат).
 */

const LOGISTICS_CFG = {
  PROPS: {
    ETA_ANCHOR_CODE: 'LOGISTICS_ETA_ANCHOR_CODE',
    ETA_TARGET_CODE: 'LOGISTICS_ETA_TARGET_CODE',
    COST_SPREADSHEET_ID: 'COST_SPREADSHEET_ID'
  },
  DEFAULTS: {
    ETA_ANCHOR_CODE: 'READY',
    ETA_TARGET_CODE: 'ARR_WAREHOUSE_MOSCOW'
  },
  DELIVERY_MODES: ['Авто', 'Море', 'Море + ЖД', 'ЖД', 'Сборный груз'],
  SHIPPED_STATUS: '4. В пути в Москву',
  SHEETS: {
    TRIPS: 'Рейсы',
    EVENTS: 'События_рейса',
    EVENT_TYPES: 'Типы_событий',
    DELIVERY_NORMS: 'Нормативы_доставки',
    DIAG: 'Логистика_диагностика'
  }
};

function addLogisticsMenu_(ui) {
  ui.createMenu('🚚 Логистика')
    .addItem('Создать/обновить листы логистики', 'logisticsEnsureSheetsMenu')
    .addItem('Заполнить шаблон типов событий и нормативов', 'logisticsSeedTemplatesMenu')
    .addSeparator()
    .addItem('Записать событие (по выделенной строке «Рейсы»)', 'logisticsRecordEventPrompt')
    .addItem('Обновить статусы и ETA рейсов', 'logisticsSyncTripsFromEventsMenu')
    .addItem('Проверка цепочки этапов', 'logisticsValidateTripChainMenu')
    .addSeparator()
    .addItem('Записать ETA в «Сводную» (01), только пустые', 'logisticsApplyEtaToSummaryMenu')
    .addToUi();
}

function logisticsEnsureSheetsMenu() {
  try {
    const r = logisticsEnsureSheets_(SpreadsheetApp.getActiveSpreadsheet(), { seedIfEmpty: false });
    SpreadsheetApp.getUi().alert(
      'Готово',
      r.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка', e.message || String(e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function logisticsSeedTemplatesMenu() {
  try {
    const r = logisticsEnsureSheets_(SpreadsheetApp.getActiveSpreadsheet(), { seedIfEmpty: true });
    SpreadsheetApp.getUi().alert(
      'Шаблоны',
      r.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка', e.message || String(e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function logisticsSyncTripsFromEventsMenu() {
  logisticsSyncTripsFromEvents_({ uiAlert: true });
}

function logisticsValidateTripChainMenu() {
  const report = logisticsValidateTripChain_(SpreadsheetApp.getActiveSpreadsheet(), { writeSheet: true });
  SpreadsheetApp.getUi().alert(
    'Проверка цепочки',
    report.summary,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function logisticsApplyEtaToSummaryMenu() {
  try {
    const r = logisticsApplyEtaToSummaryLines_({ onlyEmpty: true, dryRun: false });
    SpreadsheetApp.getUi().alert(
      'ETA в Сводной',
      'Обновлено строк: ' + r.updated + '\nПропущено (уже заполнено): ' + r.skippedFilled,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка', e.message || String(e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function logisticsRecordEventPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const trips = logisticsGetTripsSheet_(ss);
  const active = ss.getActiveSheet();
  if (!active || active.getSheetId() !== trips.getSheetId()) {
    ui.alert('Откройте лист «' + LOGISTICS_CFG.SHEETS.TRIPS + '» и выделите строку рейса.');
    return;
  }
  const row = active.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Выделите строку данных рейса (со 2-й).');
    return;
  }
  const data = trips.getDataRange().getValues();
  const tmap = logisticsHeaderMap_(data[0]);
  const idxShip = logisticsFirstIdx_(tmap, ['SHIPMENT_ID', 'ID_рейса']);
  const idxMode = logisticsFirstIdx_(tmap, ['Режим_доставки', 'Режим доставки']);
  if (idxShip == null) throw new Error('В «Рейсы» нужна колонка SHIPMENT_ID.');
  const shipmentId = String(data[row - 1][idxShip] || '').trim();
  if (!shipmentId) {
    ui.alert('В выделенной строке пустой SHIPMENT_ID.');
    return;
  }
  const mode = idxMode != null ? String(data[row - 1][idxMode] || '').trim() : 'Авто';
  const catalog = logisticsLoadEventCatalog_(ss);
  const options = logisticsEventOptionsForMode_(catalog, mode);
  if (!options.length) {
    ui.alert('Нет типов событий для режима «' + mode + '». Запустите «Заполнить шаблон…».');
    return;
  }
  const list = options.map(function (o, i) {
    return (i + 1) + '. ' + o.label + ' [' + o.code + ']';
  }).join('\n');
  const pick = ui.prompt(
    'Событие для ' + shipmentId,
    'Режим: ' + mode + '\n\n' + list + '\n\nВведите номер или код/название:',
    ui.ButtonSet.OK_CANCEL
  );
  if (pick.getSelectedButton() !== ui.Button.OK) return;
  const raw = String(pick.getResponseText() || '').trim();
  let chosen = null;
  const num = parseInt(raw, 10);
  if (num >= 1 && num <= options.length) chosen = options[num - 1];
  if (!chosen) {
    const key = logisticsNorm_(raw);
    for (let i = 0; i < options.length; i++) {
      if (options[i].key === key) {
        chosen = options[i];
        break;
      }
    }
  }
  if (!chosen) {
    ui.alert('Не удалось сопоставить тип события: ' + raw);
    return;
  }
  const dateResp = ui.prompt('Дата события', 'ГГГГ-ММ-ДД или ДД.ММ.ГГГГ:', ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const eventDate = logisticsParseDate_(dateResp.getResponseText());
  if (!eventDate) {
    ui.alert('Некорректная дата.');
    return;
  }
  const commResp = ui.prompt('Комментарий (необязательно)', '', ui.ButtonSet.OK_CANCEL);
  const comment =
    commResp.getSelectedButton() === ui.Button.OK ? String(commResp.getResponseText() || '').trim() : '';
  logisticsAppendEvent_(ss, shipmentId, chosen.label, eventDate, comment);
  ui.alert('Записано', 'Рейс ' + shipmentId + '\n' + chosen.label + '\n' + Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'dd.MM.yyyy'));
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {{ seedIfEmpty?: boolean }} opt
 * @returns {string[]}
 */
function logisticsEnsureSheets_(ss, opt) {
  const seedIfEmpty = !!(opt && opt.seedIfEmpty);
  const lines = [];
  const trips = logisticsGetTripsSheet_(ss);
  const tmap = logisticsHeaderMap_(trips.getRange(1, 1, 1, Math.max(trips.getLastColumn(), 1)).getValues()[0]);
  const idxMode = logisticsEnsureHeaderColumn_(trips, tmap, 'Режим_доставки');
  logisticsApplyDeliveryModeValidation_(trips, idxMode + 1);
  logisticsEnsureHeaderColumn_(trips, tmap, 'ETA_план');
  logisticsEnsureHeaderColumn_(trips, tmap, 'ETA_факт');
  logisticsEnsureHeaderColumn_(trips, tmap, 'Просрочка_дней');
  lines.push('«Рейсы»: колонки режима и ETA проверены.');

  if (trips.getLastRow() >= 2) {
    const refreshed = trips.getDataRange().getValues();
    const rmap = logisticsHeaderMap_(refreshed[0]);
    const idxShip = logisticsFirstIdx_(rmap, ['SHIPMENT_ID', 'ID_рейса']);
    const idxDm = logisticsFirstIdx_(rmap, ['Режим_доставки']);
    if (idxShip != null && idxDm != null) {
      let filled = 0;
      for (let i = 1; i < refreshed.length; i++) {
        if (!String(refreshed[i][idxDm] || '').trim()) {
          trips.getRange(i + 1, idxDm + 1).setValue('Авто');
          filled++;
        }
      }
      if (filled) lines.push('Проставлен режим «Авто» в пустых строках: ' + filled);
    }
  }

  let normsSh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.DELIVERY_NORMS);
  if (!normsSh) {
    normsSh = ss.insertSheet(LOGISTICS_CFG.SHEETS.DELIVERY_NORMS);
    normsSh
      .getRange(1, 1, 1, 5)
      .setValues([['Режим_доставки', 'От_кода', 'До_кода', 'Дней', 'Примечание']])
      .setFontWeight('bold');
    normsSh.setFrozenRows(1);
    lines.push('Создан лист «' + LOGISTICS_CFG.SHEETS.DELIVERY_NORMS + '».');
  }

  const etSh = logisticsGetEventTypesSheet_(ss);
  const eh = etSh.getRange(1, 1, 1, Math.max(etSh.getLastColumn(), 1)).getValues()[0];
  const emap = logisticsHeaderMap_(eh);
  logisticsEnsureHeaderColumn_(etSh, emap, 'Режим_доставки');
  logisticsEnsureHeaderColumn_(etSh, emap, 'Обязательный');
  logisticsEnsureHeaderColumn_(etSh, emap, 'Влияет_на_ETA');
  lines.push('«Типы_событий»: расширенные колонки проверены.');

  if (seedIfEmpty) {
    const seeded = logisticsSeedEventTypesAndNorms_(ss);
    lines.push(seeded);
  }
  return lines;
}

function logisticsSeedEventTypesAndNorms_(ss) {
  const etSh = logisticsGetEventTypesSheet_(ss);
  const catalog = logisticsLoadEventCatalog_(ss);
  if (catalog.rows.length < 3) {
    const seedTypes = logisticsDefaultEventTypesSeed_();
    if (etSh.getLastRow() < 2) {
      etSh
        .getRange(1, 1, 1, 6)
        .setValues([['Код', 'Наименование', 'Порядок', 'Режим_доставки', 'Обязательный', 'Влияет_на_ETA']]);
      etSh.getRange(2, 1, seedTypes.length, seedTypes[0].length).setValues(seedTypes);
    } else {
      const start = etSh.getLastRow() + 1;
      etSh.getRange(start, 1, seedTypes.length, seedTypes[0].length).setValues(seedTypes);
    }
  }

  const normsSh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.DELIVERY_NORMS);
  if (normsSh && normsSh.getLastRow() < 2) {
    const seedNorms = logisticsDefaultNormsSeed_();
    normsSh.getRange(2, 1, seedNorms.length, seedNorms[0].length).setValues(seedNorms);
  }
  return 'Шаблоны типов событий и нормативов дополнены (без перезаписи существующих строк).';
}

function logisticsDefaultEventTypesSeed_() {
  return [
    ['READY', 'Готовность к отгрузке', 10, '', 'да', 'да'],
    ['LOADED_CN', 'Погрузка у поставщика', 20, '', 'нет', 'да'],
    ['DEPART_ORIGIN', 'Выезд / отправка', 30, 'Авто', 'нет', 'да'],
    ['ARR_BORDER', 'Прибытие на границу', 50, '', 'да', 'да'],
    ['CUSTOMS', 'Таможня', 60, '', 'да', 'да'],
    ['ARR_WAREHOUSE_MOSCOW', 'Склад Москва', 80, '', 'да', 'да'],
    ['DEPART_PORT', 'Отгрузка с порта', 25, 'Море', 'да', 'да'],
    ['ARR_PORT_RF', 'Порт РФ', 45, 'Море', 'да', 'да'],
    ['RAIL_LOAD', 'Погрузка на ЖД', 35, 'ЖД', 'да', 'да'],
    ['RAIL_DEPART', 'Отправление по ЖД', 40, 'ЖД', 'да', 'да'],
    ['SEA_RAIL_TRANSSHIP', 'Перегруз море→ЖД', 48, 'Море + ЖД', 'да', 'да'],
    ['CONSOLIDATION', 'Консолидация (сборный)', 15, 'Сборный груз', 'да', 'да'],
    ['CONSOL_BATCH', 'Отгрузка партии сборного', 22, 'Сборный груз', 'да', 'да'],
    ['RECEIPT_ACCEPTED', 'Приёмка завершена', 85, '', 'нет', 'нет']
  ];
}

function logisticsDefaultNormsSeed_() {
  return [
    ['Авто', 'READY', 'ARR_BORDER', 12, ''],
    ['Авто', 'ARR_BORDER', 'CUSTOMS', 5, ''],
    ['Авто', 'CUSTOMS', 'ARR_WAREHOUSE_MOSCOW', 7, ''],
    ['Море', 'READY', 'DEPART_PORT', 5, ''],
    ['Море', 'DEPART_PORT', 'ARR_PORT_RF', 35, ''],
    ['Море', 'ARR_PORT_RF', 'ARR_WAREHOUSE_MOSCOW', 14, ''],
    ['Море + ЖД', 'READY', 'DEPART_PORT', 5, ''],
    ['Море + ЖД', 'DEPART_PORT', 'ARR_PORT_RF', 30, ''],
    ['Море + ЖД', 'ARR_PORT_RF', 'SEA_RAIL_TRANSSHIP', 3, ''],
    ['Море + ЖД', 'SEA_RAIL_TRANSSHIP', 'ARR_BORDER', 10, ''],
    ['Море + ЖД', 'ARR_BORDER', 'ARR_WAREHOUSE_MOSCOW', 10, ''],
    ['ЖД', 'READY', 'RAIL_LOAD', 3, ''],
    ['ЖД', 'RAIL_LOAD', 'ARR_BORDER', 18, ''],
    ['ЖД', 'ARR_BORDER', 'ARR_WAREHOUSE_MOSCOW', 12, ''],
    ['Сборный груз', 'READY', 'CONSOLIDATION', 7, ''],
    ['Сборный груз', 'CONSOLIDATION', 'CONSOL_BATCH', 5, ''],
    ['Сборный груз', 'CONSOL_BATCH', 'ARR_WAREHOUSE_MOSCOW', 20, '']
  ];
}

function logisticsSyncTripsFromEvents_(opt) {
  const uiAlert = !!(opt && opt.uiAlert);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trips = logisticsGetTripsSheet_(ss);
  const events = logisticsGetEventsSheet_(ss);
  if (trips.getLastRow() < 2) throw new Error('Лист «' + LOGISTICS_CFG.SHEETS.TRIPS + '» пуст.');

  const eventsData = events.getDataRange().getValues();
  if (eventsData.length < 2) {
    if (uiAlert) {
      SpreadsheetApp.getUi().alert('События рейса', 'Лист «События_рейса» пуст — обновлять нечего.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return { updated: 0 };
  }

  const emap = logisticsHeaderMap_(eventsData[0]);
  const idxEventShipment = logisticsFirstIdx_(emap, ['ID_рейса', 'SHIPMENT_ID']);
  const idxEventType = logisticsFirstIdx_(emap, ['Тип_события', 'Тип события']);
  const idxEventDate = logisticsFirstIdx_(emap, ['Дата']);
  const idxEventComment = logisticsFirstIdx_(emap, ['Комментарий']);
  if (idxEventShipment == null || idxEventType == null || idxEventDate == null) {
    throw new Error('В «События_рейса» нужны колонки: ID_рейса, Тип_события, Дата.');
  }

  const catalog = logisticsLoadEventCatalog_(ss);
  const orderMap = catalog.orderMap;
  const byShipment = {};
  for (let i = 1; i < eventsData.length; i++) {
    const r = eventsData[i];
    const shipmentId = String(r[idxEventShipment] || '').trim();
    if (!shipmentId) continue;
    const eventType = String(r[idxEventType] || '').trim();
    const eventDate = logisticsParseDate_(r[idxEventDate]);
    if (!eventDate) continue;
    const comment = idxEventComment != null ? String(r[idxEventComment] || '').trim() : '';
    const code = catalog.typeToCode[eventType] || logisticsNorm_(eventType);
    const order = orderMap[logisticsNorm_(eventType)] || orderMap[logisticsNorm_(code)] || 0;

    if (!byShipment[shipmentId]) {
      byShipment[shipmentId] = {
        count: 0,
        latestDate: null,
        latestType: '',
        latestComment: '',
        latestOrder: 0,
        events: []
      };
    }
    const rec = byShipment[shipmentId];
    rec.count++;
    rec.events.push({ type: eventType, code: code, date: eventDate, order: order });
    const laterDate = !rec.latestDate || eventDate.getTime() > rec.latestDate.getTime();
    const sameDateHigherOrder =
      rec.latestDate && eventDate.getTime() === rec.latestDate.getTime() && order > rec.latestOrder;
    if (laterDate || sameDateHigherOrder) {
      rec.latestDate = eventDate;
      rec.latestType = eventType;
      rec.latestComment = comment;
      rec.latestOrder = order;
    }
  }

  const tripsData = trips.getDataRange().getValues();
  const tmap = logisticsHeaderMap_(tripsData[0]);
  const idxTripShipment = logisticsFirstIdx_(tmap, ['SHIPMENT_ID', 'ID_рейса']);
  if (idxTripShipment == null) throw new Error('В «Рейсы» нужна колонка SHIPMENT_ID.');

  const idxEventsCount = logisticsEnsureHeaderColumn_(trips, tmap, 'Событий_всего');
  const idxLastDate = logisticsEnsureHeaderColumn_(trips, tmap, 'Дата_последнего_события');
  const idxLastType = logisticsEnsureHeaderColumn_(trips, tmap, 'Последний_тип_события');
  const idxLastComment = logisticsEnsureHeaderColumn_(trips, tmap, 'Последний_комментарий_события');
  const idxLastOrder = logisticsEnsureHeaderColumn_(trips, tmap, 'Порядок_последнего_события');
  const idxStatusByEvents = logisticsEnsureHeaderColumn_(trips, tmap, 'Статус_по_событиям');
  const idxEtaPlan = logisticsEnsureHeaderColumn_(trips, tmap, 'ETA_план');
  const idxEtaFact = logisticsEnsureHeaderColumn_(trips, tmap, 'ETA_факт');
  const idxOverdue = logisticsEnsureHeaderColumn_(trips, tmap, 'Просрочка_дней');
  const idxMode = logisticsFirstIdx_(tmap, ['Режим_доставки']);

  const anchorCode = logisticsGetProp_(LOGISTICS_CFG.PROPS.ETA_ANCHOR_CODE, LOGISTICS_CFG.DEFAULTS.ETA_ANCHOR_CODE);
  const targetCode = logisticsGetProp_(LOGISTICS_CFG.PROPS.ETA_TARGET_CODE, LOGISTICS_CFG.DEFAULTS.ETA_TARGET_CODE);
  const normsByMode = logisticsLoadNormsByMode_(ss);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const refreshed = trips.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < refreshed.length; i++) {
    const row = refreshed[i];
    const shipmentId = String(row[idxTripShipment] || '').trim();
    if (!shipmentId || !byShipment[shipmentId]) continue;
    const rec = byShipment[shipmentId];
    const mode = idxMode != null ? String(row[idxMode] || '').trim() || 'Авто' : 'Авто';
    const eta = logisticsComputeTripEta_(rec.events, mode, anchorCode, targetCode, normsByMode[mode] || []);

    trips.getRange(i + 1, idxEventsCount + 1).setValue(rec.count);
    trips.getRange(i + 1, idxLastDate + 1).setValue(rec.latestDate);
    trips.getRange(i + 1, idxLastType + 1).setValue(rec.latestType);
    trips.getRange(i + 1, idxLastComment + 1).setValue(rec.latestComment);
    trips.getRange(i + 1, idxLastOrder + 1).setValue(rec.latestOrder);
    trips.getRange(i + 1, idxStatusByEvents + 1).setValue(rec.latestType);
    if (eta.plan) trips.getRange(i + 1, idxEtaPlan + 1).setValue(eta.plan);
    if (eta.fact) trips.getRange(i + 1, idxEtaFact + 1).setValue(eta.fact);
    if (eta.plan) {
      const ref = eta.fact || today;
      const overdue = Math.round((ref.getTime() - eta.plan.getTime()) / 86400000);
      trips.getRange(i + 1, idxOverdue + 1).setValue(overdue > 0 ? overdue : '');
    }
    updated++;
  }

  if (uiAlert) {
    SpreadsheetApp.getUi().alert('Статусы и ETA', 'Обновлено рейсов: ' + updated, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return { updated: updated };
}

/**
 * @param {Object[]} events
 * @param {string} mode
 * @param {string} anchorCode
 * @param {string} targetCode
 * @param {Object[]} norms
 * @returns {{ plan: Date|null, fact: Date|null }}
 */
function logisticsComputeTripEta_(events, mode, anchorCode, targetCode, norms) {
  const anchorNorm = logisticsNorm_(anchorCode);
  const targetNorm = logisticsNorm_(targetCode);
  let anchorDate = null;
  let factDate = null;
  for (let i = 0; i < events.length; i++) {
    const e = events[i];
    const c = logisticsNorm_(e.code || e.type);
    if (c === targetNorm || logisticsNorm_(e.type) === targetNorm) {
      if (!factDate || e.date.getTime() > factDate.getTime()) factDate = e.date;
    }
    if (c === anchorNorm || logisticsNorm_(e.type) === anchorNorm) {
      if (!anchorDate || e.date.getTime() < anchorDate.getTime()) anchorDate = e.date;
    }
  }
  if (!anchorDate) {
    let minD = null;
    for (let j = 0; j < events.length; j++) {
      if (!minD || events[j].date.getTime() < minD.getTime()) minD = events[j].date;
    }
    anchorDate = minD;
  }
  if (!anchorDate) return { plan: null, fact: factDate };

  const days = logisticsSumNormDays_(anchorCode, targetCode, norms);
  const plan = new Date(anchorDate.getTime());
  plan.setDate(plan.getDate() + days);
  return { plan: plan, fact: factDate };
}

function logisticsSumNormDays_(fromCode, toCode, norms) {
  if (!norms || !norms.length) return 0;
  const graph = {};
  norms.forEach(function (n) {
    const a = logisticsNorm_(n.fromCode);
    const b = logisticsNorm_(n.toCode);
    if (!graph[a]) graph[a] = [];
    graph[a].push({ to: b, days: n.days });
  });
  const start = logisticsNorm_(fromCode);
  const end = logisticsNorm_(toCode);
  if (start === end) return 0;
  const dist = {};
  dist[start] = 0;
  const q = [start];
  while (q.length) {
    const cur = q.shift();
    const edges = graph[cur] || [];
    for (let i = 0; i < edges.length; i++) {
      const nd = dist[cur] + edges[i].days;
      if (dist[edges[i].to] == null || nd < dist[edges[i].to]) {
        dist[edges[i].to] = nd;
        q.push(edges[i].to);
      }
    }
  }
  return dist[end] != null ? dist[end] : 0;
}

function logisticsValidateTripChain_(ss, opt) {
  const writeSheet = !!(opt && opt.writeSheet);
  const trips = logisticsGetTripsSheet_(ss);
  const catalog = logisticsLoadEventCatalog_(ss);
  const events = logisticsGetEventsSheet_(ss);
  const eventsData = events.getLastRow() >= 2 ? events.getDataRange().getValues() : [[]];
  const emap = eventsData.length ? logisticsHeaderMap_(eventsData[0]) : {};
  const idxShip = logisticsFirstIdx_(emap, ['ID_рейса', 'SHIPMENT_ID']);
  const idxType = logisticsFirstIdx_(emap, ['Тип_события', 'Тип события']);

  const byShipEvents = {};
  if (idxShip != null && idxType != null && eventsData.length > 1) {
    for (let i = 1; i < eventsData.length; i++) {
      const sid = String(eventsData[i][idxShip] || '').trim();
      if (!sid) continue;
      if (!byShipEvents[sid]) byShipEvents[sid] = [];
      byShipEvents[sid].push(String(eventsData[i][idxType] || '').trim());
    }
  }

  const tripsData = trips.getDataRange().getValues();
  const tmap = logisticsHeaderMap_(tripsData[0]);
  const idxTripShip = logisticsFirstIdx_(tmap, ['SHIPMENT_ID', 'ID_рейса']);
  const idxMode = logisticsFirstIdx_(tmap, ['Режим_доставки']);
  const issues = [];
  for (let i = 1; i < tripsData.length; i++) {
    const sid = idxTripShip != null ? String(tripsData[i][idxTripShip] || '').trim() : '';
    if (!sid) continue;
    const mode = idxMode != null ? String(tripsData[i][idxMode] || '').trim() || 'Авто' : 'Авто';
    const mandatory = (catalog.mandatoryByMode[mode] || []).concat(catalog.mandatoryByMode[''] || []);
    const had = byShipEvents[sid] || [];
    const hadNorm = had.map(function (x) {
      return logisticsNorm_(x);
    });
    mandatory.forEach(function (code) {
      const cn = logisticsNorm_(code);
      const label = catalog.codeToLabel[code] || code;
      let ok = false;
      for (let h = 0; h < hadNorm.length; h++) {
        if (hadNorm[h] === cn || hadNorm[h] === logisticsNorm_(label)) {
          ok = true;
          break;
        }
      }
      if (!ok) issues.push([sid, mode, 'Нет обязательного этапа: ' + label + ' [' + code + ']']);
    });
    const allowed = logisticsEventOptionsForMode_(catalog, mode);
    const allowedKeys = {};
    allowed.forEach(function (o) {
      allowedKeys[o.key] = true;
    });
    had.forEach(function (t) {
      const k = logisticsNorm_(t);
      const c = catalog.typeToCode[t] || '';
      if (allowedKeys[k] || (c && allowedKeys[logisticsNorm_(c)])) return;
      if (catalog.byCode[c] && !catalog.byCode[c].mode) return;
      issues.push([sid, mode, 'Событие не для режима «' + mode + '»: ' + t]);
    });
  }

  if (writeSheet && issues.length) {
    let diag = ss.getSheetByName(LOGISTICS_CFG.SHEETS.DIAG);
    if (!diag) diag = ss.insertSheet(LOGISTICS_CFG.SHEETS.DIAG);
    diag.clearContents();
    const out = [['SHIPMENT_ID', 'Режим', 'Проблема']].concat(issues);
    diag.getRange(1, 1, out.length, 3).setValues(out);
    diag.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  return {
    issueCount: issues.length,
    summary: issues.length ? 'Найдено замечаний: ' + issues.length : 'Цепочки этапов в порядке.',
    issues: issues
  };
}

/**
 * @param {{ onlyEmpty?: boolean, dryRun?: boolean }} opt
 */
function logisticsApplyEtaToSummaryLines_(opt) {
  const onlyEmpty = opt && opt.onlyEmpty !== false;
  const dryRun = !!(opt && opt.dryRun);
  const ss05 = SpreadsheetApp.getActiveSpreadsheet();
  const ss01 = typeof costingOpenBook01_ === 'function' ? costingOpenBook01_() : null;
  if (!ss01) throw new Error('Не удалось открыть книгу 01 (ORDERS_SPREADSHEET_ID).');

  const trips = logisticsGetTripsSheet_(ss05);
  const data = trips.getDataRange().getValues();
  const tmap = logisticsHeaderMap_(data[0]);
  const idxShip = logisticsFirstIdx_(tmap, ['SHIPMENT_ID', 'ID_рейса']);
  const idxEtaPlan = logisticsFirstIdx_(tmap, ['ETA_план']);
  const idxEtaFact = logisticsFirstIdx_(tmap, ['ETA_факт']);
  if (idxShip == null) throw new Error('В «Рейсы» нет SHIPMENT_ID.');

  const etaByShip = {};
  for (let i = 1; i < data.length; i++) {
    const sid = String(data[i][idxShip] || '').trim();
    if (!sid) continue;
    const d = idxEtaFact != null && data[i][idxEtaFact] ? logisticsParseDate_(data[i][idxEtaFact]) : null;
    const p = idxEtaPlan != null && data[i][idxEtaPlan] ? logisticsParseDate_(data[i][idxEtaPlan]) : null;
    etaByShip[sid] = d || p;
  }

  const summary = ss01.getSheetByName('Сводная');
  if (!summary || summary.getLastRow() < 2) throw new Error('В книге 01 нет листа «Сводная» с данными.');

  const headerRow = logisticsDetectHeaderRow_(summary, ['Артикул ВБ', 'Рейс', 'Статус заказа']) || 2;
  const lastCol = summary.getLastColumn();
  const hdr = summary.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const hmap = logisticsHeaderMap_(hdr);
  const idxSummaryShip = logisticsFirstIdx_(hmap, ['Рейс', 'SHIPMENT_ID', 'ID_рейса']);
  let idxEta = logisticsFirstIdx_(hmap, ['Плановая дата поступления', 'ETA', 'Дата поступления']);
  if (idxSummaryShip == null) throw new Error('В «Сводной» нет колонки «Рейс».');
  if (idxEta == null) {
    idxEta = lastCol;
    summary.getRange(headerRow, lastCol + 1).setValue('Плановая дата поступления');
  }

  const start = headerRow + 1;
  const lr = summary.getLastRow();
  if (lr < start) return { updated: 0, skippedFilled: 0 };
  const range = summary.getRange(start, idxEta + 1, lr, idxEta + 1);
  const vals = range.getValues();
  const ships = summary.getRange(start, idxSummaryShip + 1, lr, idxSummaryShip + 1).getValues();
  let updated = 0;
  let skippedFilled = 0;
  for (let i = 0; i < vals.length; i++) {
    const sid = String(ships[i][0] || '').trim();
    if (!sid || !etaByShip[sid]) continue;
    if (onlyEmpty && logisticsParseDate_(vals[i][0])) {
      skippedFilled++;
      continue;
    }
    if (!dryRun) vals[i][0] = etaByShip[sid];
    updated++;
  }
  if (!dryRun && updated) range.setValues(vals);
  return { updated: updated, skippedFilled: skippedFilled };
}

/** @returns {Object.<string,string>} SHIPMENT_ID → режим доставки */
function logisticsLoadTripDeliveryModeMap_(ss05) {
  const trips = logisticsGetTripsSheet_(ss05);
  const data = trips.getDataRange().getValues();
  if (data.length < 2) return {};
  const map = logisticsHeaderMap_(data[0]);
  const idxShip = logisticsFirstIdx_(map, ['SHIPMENT_ID', 'ID_рейса']);
  const idxMode = logisticsFirstIdx_(map, ['Режим_доставки', 'Режим доставки']);
  if (idxShip == null) return {};
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const sid = String(data[i][idxShip] || '').trim();
    if (!sid) continue;
    out[sid] = idxMode != null ? String(data[i][idxMode] || '').trim() || 'Авто' : 'Авто';
  }
  return out;
}

function logisticsLoadValidShipmentIds_(ss05) {
  const map = logisticsLoadTripDeliveryModeMap_(ss05);
  const out = {};
  Object.keys(map).forEach(function (k) {
    out[k] = true;
  });
  return out;
}

function logisticsOpenBook05_() {
  if (typeof syncHubOpenSpreadsheetForBook_ === 'function') {
    try {
      return syncHubOpenSpreadsheetForBook_('05');
    } catch (e) {}
  }
  const raw = logisticsGetProp_(LOGISTICS_CFG.PROPS.COST_SPREADSHEET_ID, '');
  if (!raw && typeof costingGetProp_ === 'function') {
    const alt = costingGetProp_('COST_SPREADSHEET_ID', '');
    if (alt) return SpreadsheetApp.openById(typeof costingExtractSpreadsheetId_ === 'function' ? costingExtractSpreadsheetId_(alt) : alt);
  }
  if (!raw) throw new Error('Не задан COST_SPREADSHEET_ID для книги 05.');
  const id = typeof costingExtractSpreadsheetId_ === 'function' ? costingExtractSpreadsheetId_(raw) : raw;
  return SpreadsheetApp.openById(id);
}

function logisticsAppendEvent_(ss, shipmentId, eventType, eventDate, comment) {
  const events = logisticsGetEventsSheet_(ss);
  const row = [shipmentId, eventType, eventDate, comment || ''];
  const hdr = events.getRange(1, 1, 1, Math.max(events.getLastColumn(), 4)).getValues()[0];
  const need = ['ID_рейса', 'Тип_события', 'Дата', 'Комментарий'];
  if (!hdr[0] || logisticsNorm_(hdr[0]).indexOf('id') === -1) {
    events.getRange(1, 1, 1, need.length).setValues([need]);
  }
  events.appendRow(row);
}

/** Запись события по коду из «Типы_событий» (например RECEIPT_ACCEPTED). */
function logisticsAppendEventByCode_(ss, shipmentId, eventCode, eventDate, comment) {
  const catalog = logisticsLoadEventCatalog_(ss);
  const code = String(eventCode || '').trim();
  const label = catalog.codeToLabel[code] || code;
  logisticsAppendEvent_(ss, shipmentId, label, eventDate, comment || '');
}

function logisticsLoadEventCatalog_(ss) {
  const sh = logisticsGetEventTypesSheet_(ss);
  const data = sh.getLastRow() >= 1 ? sh.getDataRange().getValues() : [];
  const header = data.length ? data[0] : [];
  const map = logisticsHeaderMap_(header);
  const idxCode = logisticsFirstIdx_(map, ['Код']);
  const idxName = logisticsFirstIdx_(map, ['Наименование']);
  const idxOrder = logisticsFirstIdx_(map, ['Порядок']);
  const idxMode = logisticsFirstIdx_(map, ['Режим_доставки', 'Режим доставки']);
  const idxReq = logisticsFirstIdx_(map, ['Обязательный']);
  const rows = [];
  const orderMap = {};
  const byCode = {};
  const codeToLabel = {};
  const typeToCode = {};
  const mandatoryByMode = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const code = idxCode != null ? String(r[idxCode] || '').trim() : '';
    const name = idxName != null ? String(r[idxName] || '').trim() : '';
    const mode = idxMode != null ? String(r[idxMode] || '').trim() : '';
    const ord = idxOrder != null ? logisticsToNumber_(r[idxOrder]) : 0;
    const req = idxReq != null ? logisticsTruthy_(r[idxReq]) : false;
    if (!code && !name) continue;
    const label = name || code;
    if (name) orderMap[logisticsNorm_(name)] = ord;
    if (code) {
      orderMap[logisticsNorm_(code)] = ord;
      byCode[code] = { code: code, label: label, mode: mode, order: ord };
      codeToLabel[code] = label;
    }
    if (name) typeToCode[name] = code;
    if (req && code) {
      const mk = mode || '';
      if (!mandatoryByMode[mk]) mandatoryByMode[mk] = [];
      mandatoryByMode[mk].push(code);
    }
    rows.push({ code: code, label: label, mode: mode, order: ord, key: logisticsNorm_(label) });
  }
  return {
    rows: rows,
    orderMap: orderMap,
    byCode: byCode,
    codeToLabel: codeToLabel,
    typeToCode: typeToCode,
    mandatoryByMode: mandatoryByMode
  };
}

function logisticsEventOptionsForMode_(catalog, mode) {
  const m = String(mode || '').trim() || 'Авто';
  const mn = logisticsNorm_(m);
  return catalog.rows
    .filter(function (row) {
      if (!row.mode) return true;
      return logisticsNorm_(row.mode) === mn;
    })
    .sort(function (a, b) {
      return a.order - b.order;
    })
    .map(function (row) {
      return { code: row.code, label: row.label, key: row.key };
    });
}

function logisticsLoadNormsByMode_(ss) {
  const sh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.DELIVERY_NORMS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getDataRange().getValues();
  const map = logisticsHeaderMap_(data[0]);
  const idxMode = logisticsFirstIdx_(map, ['Режим_доставки', 'Режим доставки']);
  const idxFrom = logisticsFirstIdx_(map, ['От_кода', 'От кода']);
  const idxTo = logisticsFirstIdx_(map, ['До_кода', 'До кода']);
  const idxDays = logisticsFirstIdx_(map, ['Дней', 'Дни']);
  if (idxFrom == null || idxTo == null || idxDays == null) return out;
  for (let i = 1; i < data.length; i++) {
    const mode = idxMode != null ? String(data[i][idxMode] || '').trim() || 'Авто' : 'Авто';
    if (!out[mode]) out[mode] = [];
    out[mode].push({
      fromCode: String(data[i][idxFrom] || '').trim(),
      toCode: String(data[i][idxTo] || '').trim(),
      days: logisticsToNumber_(data[i][idxDays])
    });
  }
  return out;
}

function logisticsApplyDeliveryModeValidation_(sheet, col1Based) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOGISTICS_CFG.DELIVERY_MODES, true)
    .setAllowInvalid(false)
    .build();
  const last = Math.max(sheet.getLastRow(), 500);
  sheet.getRange(2, col1Based, last, 1).setDataValidation(rule);
}

function logisticsGetTripsSheet_(ss) {
  if (typeof costingGetSheetByRole_ === 'function') return costingGetSheetByRole_(ss, 'TRIPS');
  const sh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.TRIPS);
  if (!sh) throw new Error('Не найден лист «' + LOGISTICS_CFG.SHEETS.TRIPS + '».');
  return sh;
}

function logisticsGetEventsSheet_(ss) {
  if (typeof costingGetSheetByRole_ === 'function') return costingGetSheetByRole_(ss, 'EVENTS');
  const sh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.EVENTS);
  if (!sh) throw new Error('Не найден лист «' + LOGISTICS_CFG.SHEETS.EVENTS + '».');
  return sh;
}

function logisticsGetEventTypesSheet_(ss) {
  if (typeof costingGetSheetByRole_ === 'function') return costingGetSheetByRole_(ss, 'EVENT_TYPES');
  const sh = ss.getSheetByName(LOGISTICS_CFG.SHEETS.EVENT_TYPES);
  if (!sh) throw new Error('Не найден лист «' + LOGISTICS_CFG.SHEETS.EVENT_TYPES + '».');
  return sh;
}

function logisticsDetectHeaderRow_(sheet, markers) {
  const lastRow = Math.min(sheet.getLastRow(), 15);
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return null;
  const block = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (let r = 0; r < block.length; r++) {
    const rowNorm = block[r].map(function (c) {
      return logisticsNorm_(c);
    });
    let hits = 0;
    for (let m = 0; m < markers.length; m++) {
      const mk = logisticsNorm_(markers[m]);
      for (let c = 0; c < rowNorm.length; c++) {
        if (rowNorm[c].indexOf(mk) !== -1) {
          hits++;
          break;
        }
      }
    }
    if (hits >= Math.min(2, markers.length)) return r + 1;
  }
  return null;
}

function logisticsHeaderMap_(headerRow) {
  if (typeof costingHeaderMap_ === 'function') return costingHeaderMap_(headerRow);
  const out = {};
  for (let i = 0; i < headerRow.length; i++) out[logisticsNorm_(headerRow[i])] = i;
  return out;
}

function logisticsFirstIdx_(headerMap, aliases) {
  if (typeof costingFirstIdx_ === 'function') return costingFirstIdx_(headerMap, aliases);
  for (let i = 0; i < aliases.length; i++) {
    const k = logisticsNorm_(aliases[i]);
    if (headerMap[k] != null) return headerMap[k];
  }
  return null;
}

function logisticsEnsureHeaderColumn_(sheet, headerMap, headerName) {
  if (typeof costingEnsureHeaderColumn_ === 'function') return costingEnsureHeaderColumn_(sheet, headerMap, headerName);
  const key = logisticsNorm_(headerName);
  if (headerMap[key] != null) return headerMap[key];
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(headerName);
  headerMap[key] = col - 1;
  return col - 1;
}

function logisticsParseDate_(v) {
  if (typeof costingParseDate_ === 'function') return costingParseDate_(v);
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function logisticsNorm_(v) {
  if (typeof costingNorm_ === 'function') return costingNorm_(v);
  return String(v || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\s+/g, ' ')
    .trim();
}

function logisticsToNumber_(v) {
  if (typeof costingToNumber_ === 'function') return costingToNumber_(v);
  const n = Number(String(v || '').replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function logisticsTruthy_(v) {
  const s = String(v == null ? '' : v)
    .trim()
    .toLowerCase();
  return s === '1' || s === 'true' || s === 'да' || s === 'yes';
}

function logisticsGetProp_(key, fallback) {
  if (typeof costingGetProp_ === 'function') return costingGetProp_(key, fallback);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  return raw == null || String(raw).trim() === '' ? fallback : String(raw).trim();
}
