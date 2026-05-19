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

/** Статус «Сводной», при котором строка создаётся/обновляется в МС (см. MS_SYNC_ORDER_STATUSES). */
const MS_SYNC_ORDER_STATUS_DEFAULT = '03. Создание заказа';

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
    .addSeparator()
    .addItem('🚛 Отгрузка → Списать выделенные строки в рейс…', 'assignSelectedRowsToShipment')
    .addSeparator()
    .addItem('🏢 Организации МС (для MS_ORGANIZATION_ID)', 'showMsOrganizationsForSetup')
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
  if (typeof gremlinScheduleAddMenu01_ === 'function') {
    gremlinScheduleAddMenu01_(ui);
  }
  // Меню «Планирование закупок» живёт в книге 03 (см. main_03.gs).
  // В книге 01 его быть не должно: листов/свойств планирования здесь нет, и менеджеры
  // нажимая пункты в «не той» книге, получали бы ошибки. См. README.txt и PROJECT_CONTEXT.md.
}

/**
 * Для синка с МС обязательны «Отгрузка через» или «Поставщик МС».
 * Одного «Поставщик» (фабрика/ФИО) недостаточно — иначе контрагент в МС не определён.
 */
function isMsSyncLogisticsReady_(row) {
  return (
    !syncManagerBlankish_(row[COL.SUPPLIER_MS]) || !syncManagerBlankish_(row[COL.DELIVERY_TYPE])
  );
}

function isMsSyncSpecReady_(row) {
  return !syncManagerBlankish_(row[COL.SPEC_NUMBER]);
}

/** Кандидаты на контрагента в МС (порядок: «Поставщик МС» → «Отгрузка через» → «Поставщик»). */
function collectSupplierNameCandidates_(row) {
  const out = [];
  const add = function (v) {
    const s = safeString(v);
    if (!s) return;
    let i;
    for (i = 0; i < out.length; i++) {
      if (out[i].toLowerCase() === s.toLowerCase()) return;
    }
    out.push(s);
  };
  add(row[COL.SUPPLIER_MS]);
  add(row[COL.DELIVERY_TYPE]);
  if (isMsSyncLogisticsReady_(row)) add(row[COL.SUPPLIER_NOTE]);
  return out;
}

function findSummaryHeaderRowIndex_(data) {
  const limit = Math.min(data.length, 15);
  let bestIdx = MANAGER_SUMMARY_SYNC_HEADER_ROW - 1;
  let bestScore = 0;
  let r;
  for (r = 0; r < limit; r++) {
    const row = data[r] || [];
    let score = 0;
    let c;
    for (c = 0; c < row.length; c++) {
      const key = syncManagerHeaderKey_(syncManagerCanonHeader_(row[c]));
      if (key === 'wb_article') score += 3;
      if (key === 'barcode') score += 2;
      if (key === 'supplier_note' || key === 'supplier_ms') score += 2;
      if (key === 'spec_number') score += 2;
      if (key === 'total_qty' || key === 'qty') score += 1;
      if (key === 'script_status' || key === 'ms_id') score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestIdx = r;
    }
  }
  return bestIdx;
}

function buildMsSyncKeyToCol_(headerRow) {
  const keyToCol = {};
  let c;
  for (c = 0; c < headerRow.length; c++) {
    const key = syncManagerHeaderKey_(syncManagerCanonHeader_(headerRow[c]));
    if (key && keyToCol[key] == null) keyToCol[key] = c;
  }
  return keyToCol;
}

function buildMsSyncLayout_(data) {
  const headerRowIdx = findSummaryHeaderRowIndex_(data);
  const headerRow = data[headerRowIdx] || [];
  const keyToCol = buildMsSyncKeyToCol_(headerRow);
  return {
    headerRowIdx: headerRowIdx,
    keyToCol: keyToCol,
    colStatus: keyToCol.script_status != null ? keyToCol.script_status : COL.STATUS,
    colMsId: keyToCol.ms_id != null ? keyToCol.ms_id : COL.MS_ID,
    colMsLink: keyToCol.ms_link != null ? keyToCol.ms_link : COL.MS_LINK
  };
}

/** Строка листа → канонический массив по COL (по заголовкам, а не по фиксированным буквам колонок). */
function rowToCanonicalForMsSync_(rawRow, keyToCol) {
  const canonHeaders = syncManagerSummaryCanonicalHeader_();
  const canonKeys = canonHeaders.map(function (h) {
    return syncManagerHeaderKey_(syncManagerCanonHeader_(h));
  });
  const out = new Array(canonKeys.length).fill('');
  let i;
  for (i = 0; i < canonKeys.length; i++) {
    const key = canonKeys[i];
    if (!key) continue;
    const src = keyToCol[key];
    if (src != null && src < rawRow.length) out[i] = rawRow[src];
  }
  return out;
}

function lookupMsSupplier_(candidates, supplierMap) {
  let i;
  for (i = 0; i < candidates.length; i++) {
    const name = candidates[i];
    const cacheKey = '__name__' + name.toLowerCase();
    if (cache.suppliers[cacheKey] !== undefined) {
      if (cache.suppliers[cacheKey]) return { supplier: cache.suppliers[cacheKey], name: name };
      continue;
    }
    const lookupName = name.toLowerCase();
    const msCode = supplierMap[lookupName];
    const supplier = msCode ? findCounterpartyByName(msCode) : findCounterpartyByName(name);
    cache.suppliers[cacheKey] = supplier || null;
    if (supplier) return { supplier: supplier, name: name };
  }
  return { supplier: null, name: candidates[0] || '' };
}

function normalizeOrderStatusForMsSync_(val) {
  return safeString(val)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Список допустимых статусов из MS_SYNC_ORDER_STATUSES или значение по умолчанию. */
function getMsSyncOrderStatusPatterns_() {
  const raw = getScriptProp('MS_SYNC_ORDER_STATUSES', '');
  const parts = raw ? String(raw).split(/[,;|]/) : [MS_SYNC_ORDER_STATUS_DEFAULT];
  const out = [];
  let i;
  for (i = 0; i < parts.length; i++) {
    const n = normalizeOrderStatusForMsSync_(parts[i]);
    if (n && out.indexOf(n) < 0) out.push(n);
  }
  return out.length ? out : [normalizeOrderStatusForMsSync_(MS_SYNC_ORDER_STATUS_DEFAULT)];
}

function orderStatusEligibleForMsSync_(status) {
  const norm = normalizeOrderStatusForMsSync_(status);
  if (!norm) return false;
  const patterns = getMsSyncOrderStatusPatterns_();
  let i;
  for (i = 0; i < patterns.length; i++) {
    const p = patterns[i];
    if (!p) continue;
    if (norm === p) return true;
    const codeM = p.match(/^(\d{2})\./);
    if (codeM && norm.indexOf(codeM[1] + '.') === 0) return true;
    const phrase = p.replace(/^\d{2}\.\s*/, '');
    if (phrase.length >= 6 && norm.indexOf(phrase) >= 0) return true;
  }
  return false;
}

function getMsSyncOrderStatusHint_() {
  return getMsSyncOrderStatusPatterns_()
    .map(function (p) {
      return '«' + p + '»';
    })
    .join(', ');
}

function getCanonicalHeaderKeys_() {
  return syncManagerSummaryCanonicalHeader_().map(function (h) {
    return syncManagerHeaderKey_(syncManagerCanonHeader_(h));
  });
}

/** Ключ сопоставления строки Сводной ↔ вкладки-источника (без MS_ID — он как раз пишется). */
function buildMsRowMatchKey_(row) {
  const supplier =
    safeString(row[COL.SUPPLIER_MS]) ||
    safeString(row[COL.DELIVERY_TYPE]) ||
    safeString(row[COL.SUPPLIER_NOTE]);
  return [
    safeString(row[COL.WB_ARTICLE]),
    safeString(row[COL.BARCODE]),
    safeString(row[COL.SUPPLIER_ARTICLE]),
    safeString(row[COL.SPEC_NUMBER]),
    supplier,
    syncManagerNormalizeManagerToken_(safeString(row[COL.MANAGER]))
  ]
    .join('||')
    .toLowerCase();
}

function readSourceSheetLayout_(sheet) {
  const lc = Math.max(sheet.getLastColumn(), 1);
  const headerRaw =
    sheet.getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, lc).getValues()[0] || [];
  const srcHeaderKeys = headerRaw.map(function (h) {
    return syncManagerHeaderKey_(syncManagerCanonHeader_(h));
  });
  const srcMap = syncManagerBuildHeaderIndex_(srcHeaderKeys);
  return { srcMap: srcMap, lastCol: lc };
}

/**
 * Колонки для записи результатов синка. Существующие — по заголовку;
 * недостающие MS_ID / MS_LINK / «Статус скрипта» — только в конец листа (оплаты не сдвигаются).
 */
function ensureMsWriteColumnsOnSource_(sheet, srcMap) {
  const writeCols = {
    msId: srcMap.ms_id != null ? srcMap.ms_id : null,
    msLink: srcMap.ms_link != null ? srcMap.ms_link : null,
    scriptStatus: srcMap.script_status != null ? srcMap.script_status : null,
    msOrderSheet: srcMap.ms_order_sheet != null ? srcMap.ms_order_sheet : null
  };
  let nextCol = Math.max(sheet.getLastColumn(), 1) + 1;
  const headerRow = MANAGER_SUMMARY_SYNC_HEADER_ROW;
  if (writeCols.msId == null) {
    sheet.getRange(headerRow, nextCol).setValue('MS_ID');
    writeCols.msId = nextCol - 1;
    nextCol++;
  }
  if (writeCols.msLink == null) {
    sheet.getRange(headerRow, nextCol).setValue('MS_LINK');
    writeCols.msLink = nextCol - 1;
    nextCol++;
  }
  if (writeCols.scriptStatus == null) {
    sheet.getRange(headerRow, nextCol).setValue('Статус скрипта');
    writeCols.scriptStatus = nextCol - 1;
    nextCol++;
  }
  return writeCols;
}

function buildSheetRowKeyIndex_(sheet, layout, headerKeys) {
  const index = {};
  const lr = sheet.getLastRow();
  if (lr < MANAGER_SUMMARY_SYNC_DATA_START_ROW) return index;
  const block = sheet
    .getRange(MANAGER_SUMMARY_SYNC_DATA_START_ROW, 1, lr, layout.lastCol)
    .getValues();
  let r;
  for (r = 0; r < block.length; r++) {
    const aligned = syncManagerAlignRowToHeader_(block[r], headerKeys, layout.srcMap);
    if (syncManagerSummaryRowIgnorable_(aligned)) continue;
    const key = buildMsRowMatchKey_(aligned);
    if (!key || key === '||||||') continue;
    const sheetRow = MANAGER_SUMMARY_SYNC_DATA_START_ROW + r;
    if (!index[key]) index[key] = [];
    index[key].push(sheetRow);
  }
  return index;
}

/** Индекс вкладок-источников для обратной записи MS_ID (менеджерские + история). */
function buildMsSourceWriteContext_(ss) {
  const headerKeys = getCanonicalHeaderKeys_();
  const ex = syncManagerSummaryExcludeSet_();
  const entries = [];
  const sheets = ss.getSheets();
  let i;
  for (i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const nameTrim = safeString(sh.getName()).trim();
    if (nameTrim === safeString(SHEET_NAME).trim()) continue;
    if (syncManagerSheetIsExcluded_(nameTrim, ex)) continue;
    if (sh.getLastRow() < MANAGER_SUMMARY_SYNC_DATA_START_ROW) continue;

    let kind = '';
    let meta = null;
    if (syncManagerIsHistoryShipmentsSheet_(nameTrim)) {
      kind = 'history';
    } else {
      meta = syncManagerSummaryParseManagerMonthTab_(nameTrim);
      if (!meta) continue;
      kind = 'manager_month';
    }

    const layout = readSourceSheetLayout_(sh);
    layout.writeCols = ensureMsWriteColumnsOnSource_(sh, layout.srcMap);
    layout.lastCol = Math.max(sh.getLastColumn(), layout.lastCol);
    layout.rowIndex = buildSheetRowKeyIndex_(sh, layout, headerKeys);

    entries.push({
      sheet: sh,
      name: nameTrim,
      kind: kind,
      meta: meta,
      layout: layout,
      managerKey: meta
        ? syncManagerNormalizeManagerToken_(meta.manager).toLowerCase()
        : ''
    });
  }
  return { headerKeys: headerKeys, entries: entries, written: 0, notFound: 0 };
}

function findSourceRowForCanonical_(ctx, canonicalRow) {
  const matchKey = buildMsRowMatchKey_(canonicalRow);
  if (!matchKey || matchKey === '||||||') return null;

  const manager = syncManagerNormalizeManagerToken_(safeString(canonicalRow[COL.MANAGER])).toLowerCase();
  let pass;
  for (pass = 0; pass < 2; pass++) {
    let e;
    for (e = 0; e < ctx.entries.length; e++) {
      const entry = ctx.entries[e];
      if (pass === 0 && manager) {
        if (entry.kind === 'manager_month' && entry.managerKey !== manager) continue;
      }
      const rows = entry.layout.rowIndex[matchKey];
      if (!rows || !rows.length) continue;
      return {
        sheet: entry.sheet,
        sheetName: entry.name,
        sheetRow: rows[0],
        writeCols: entry.layout.writeCols
      };
    }
    if (!manager) break;
  }
  return null;
}

/** Только ячейки MS — остальные колонки (оплаты, даты и т.д.) не трогаем. */
function writeMsFieldsToSourceRow_(hit, patch) {
  if (!hit || !patch) return false;
  const sh = hit.sheet;
  const r = hit.sheetRow;
  const c = hit.writeCols;
  if (patch.msId != null && patch.msId !== '' && c.msId != null) {
    sh.getRange(r, c.msId + 1).setValue(patch.msId);
    if (c.msOrderSheet != null) sh.getRange(r, c.msOrderSheet + 1).setValue(patch.msId);
  }
  if (patch.msLink != null && patch.msLink !== '' && c.msLink != null) {
    sh.getRange(r, c.msLink + 1).setValue(patch.msLink);
  }
  if (patch.status != null && patch.status !== '' && c.scriptStatus != null) {
    sh.getRange(r, c.scriptStatus + 1).setValue(patch.status);
  }
  return true;
}

function resolveSourceTabName_(ctx, canonicalRow) {
  const hit = findSourceRowForCanonical_(ctx, canonicalRow);
  return hit && hit.sheetName ? hit.sheetName : '';
}

/** Уникальные имена вкладок-источников для строк одной группы заказа. */
function resolveSourceTabNamesForGroup_(ctx, groupRows) {
  const names = [];
  const seen = {};
  let i;
  for (i = 0; i < groupRows.length; i++) {
    const n = resolveSourceTabName_(ctx, groupRows[i].rowData);
    if (!n) continue;
    const k = n.toLowerCase();
    if (seen[k]) continue;
    seen[k] = true;
    names.push(n);
  }
  return names;
}

function msSyncWriteToSourceTab_(ctx, canonicalRow, patch) {
  const hit = findSourceRowForCanonical_(ctx, canonicalRow);
  if (!hit) {
    ctx.notFound++;
    return false;
  }
  if (writeMsFieldsToSourceRow_(hit, patch)) {
    ctx.written++;
    return true;
  }
  return false;
}

/** Сводная + вкладка-источник: пишем только MS_ID / MS_LINK / статус скрипта. */
function msSyncApplyRowPatch_(ctx, updater, layout, rowIndex, canonicalRow, patch) {
  if (patch.msId != null && patch.msId !== '') updater.setValue(rowIndex, layout.colMsId, patch.msId);
  if (patch.msLink != null && patch.msLink !== '') updater.setValue(rowIndex, layout.colMsLink, patch.msLink);
  if (patch.status != null && patch.status !== '') updater.setStatus(rowIndex, patch.status);
  msSyncWriteToSourceTab_(ctx, canonicalRow, patch);
}

/** UUID из id, ссылки meta.href или текста в MS_ORGANIZATION_ID. */
function normalizeMsEntityId_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  const m = s.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i);
  return m ? m[1] : s;
}

/** Проверка MS_ORGANIZATION_ID до массовых запросов в МС. */
function validateMsOrganization_() {
  const orgId = normalizeMsEntityId_(CONFIG.MS_ORGANIZATION_ID);
  if (!orgId) {
    return { ok: false, orgId: '', error: 'Пустой MS_ORGANIZATION_ID' };
  }
  const res = msGet('/entity/organization/' + encodeURIComponent(orgId));
  if (res.success && res.data && res.data.id) {
    return { ok: true, orgId: res.data.id, name: safeString(res.data.name) || orgId };
  }
  return {
    ok: false,
    orgId: orgId,
    error: res.error || 'Организация не найдена в аккаунте МойСклад (проверьте токен и id).'
  };
}

/** Список организаций аккаунта (меню настройки MS_ORGANIZATION_ID). */
function fetchMsOrganizationsList_() {
  const res = msGet('/entity/organization?limit=100');
  if (!res.success || !res.data || !res.data.rows) {
    return { ok: false, error: res.error || 'Не удалось получить список организаций', rows: [] };
  }
  return { ok: true, rows: res.data.rows, error: '' };
}

/** Показать id организаций аккаунта МС — для свойства MS_ORGANIZATION_ID. */
function showMsOrganizationsForSetup() {
  const ui = SpreadsheetApp.getUi();
  let currentId = '';
  try {
    currentId = normalizeMsEntityId_(CONFIG.MS_ORGANIZATION_ID);
  } catch (e) {
    currentId = '(не задано)';
  }

  const list = fetchMsOrganizationsList_();
  if (!list.ok) {
    ui.alert('Не удалось загрузить организации МойСклад', String(list.error || ''), ui.ButtonSet.OK);
    return;
  }
  if (!list.rows.length) {
    ui.alert('В аккаунте МойСклад нет организаций или нет доступа по токену MS_TOKEN.', ui.ButtonSet.OK);
    return;
  }

  const lines = [
    'Скопируйте id нужной организации в Script Property MS_ORGANIZATION_ID',
    '(Расширения → Apps Script → Свойства скрипта / Project Settings → Script properties).',
    '',
    'Сейчас в MS_ORGANIZATION_ID: ' + currentId,
    ''
  ];
  list.rows.forEach(function (org, idx) {
    const mark = org.id === currentId ? ' ← сейчас в свойствах' : '';
    lines.push((idx + 1) + '. ' + safeString(org.name) + '\n   id: ' + org.id + mark);
  });

  ui.alert('Организации МойСклад', lines.join('\n'), ui.ButtonSet.OK);
}

function syncOrdersWithMS() {
  try {
    getScriptPropOrThrow('MS_ORGANIZATION_ID');
  } catch (error) {
    return SpreadsheetApp.getUi().alert(`Ошибка настройки: ${error.message}`);
  }

  const orgCheck = validateMsOrganization_();
  if (!orgCheck.ok) {
    return SpreadsheetApp.getUi().alert(
      'Синхронизация остановлена: организация из MS_ORGANIZATION_ID не найдена в МойСклад.\n\n' +
        'id в свойствах: ' +
        orgCheck.orgId +
        '\n' +
        orgCheck.error +
        '\n\n' +
        'Частые причины:\n' +
        '• id от другого аккаунта МС или устаревший после пересоздания организации;\n' +
        '• MS_TOKEN от другого аккаунта, чем id организации.\n\n' +
        'Меню → 📦 МойСклад → 🏢 Организации МС — скопируйте верный id в свойства скрипта.'
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return SpreadsheetApp.getUi().alert(`Не найден лист «${SHEET_NAME}».`);
  }

  const data = sheet.getDataRange().getValues();
  const layout = buildMsSyncLayout_(data);
  const headerRowIdx = layout.headerRowIdx;
  const dataStartIdx = headerRowIdx + 1;
  if (data.length <= dataStartIdx) {
    return SpreadsheetApp.getUi().alert(
      'На листе «Сводная» нет строк данных под шапкой (шапка найдена в строке ' +
        (headerRowIdx + 1) +
        ').\n\n' +
        'Сначала: 📦 МойСклад → 🧩 Собрать «Сводная» из вкладок менеджеров.'
    );
  }

  const updater = new BatchUpdater(sheet, data, layout);
  const sourceCtx = buildMsSourceWriteContext_(ss);
  const groups = {};
  let skippedByStatus = 0;
  let skippedByShipVia = 0;
  let skippedBySpec = 0;
  
  const originalMsIds = new Set();
  const statusHint = getMsSyncOrderStatusHint_();

  for (let i = dataStartIdx; i < data.length; i++) {
    const row = rowToCanonicalForMsSync_(data[i], layout.keyToCol);
    const candidates = collectSupplierNameCandidates_(row);
    const supplierName = candidates.length ? candidates[0] : '';
    const specNum = safeString(row[COL.SPEC_NUMBER]);
    const msId = safeString(row[COL.MS_ID]);
    const orderStatus = safeString(row[COL.ORDER_STATUS]);
    const hasLine =
      safeString(row[COL.WB_ARTICLE]) ||
      safeString(row[COL.SUPPLIER_ARTICLE]) ||
      safeString(row[COL.BARCODE]);
    const statusOk = orderStatusEligibleForMsSync_(orderStatus);

    if ((hasLine || supplierName) && !statusOk) {
      skippedByStatus++;
      msSyncApplyRowPatch_(sourceCtx, updater, layout, i, row, {
        status: '⏭ Пропуск: статус «' + (orderStatus || 'пусто') + '» (нужен ' + statusHint + ')'
      });
      continue;
    }

    if (statusOk && hasLine && !isMsSyncLogisticsReady_(row)) {
      skippedByShipVia++;
      msSyncApplyRowPatch_(sourceCtx, updater, layout, i, row, {
        status: '⚠️ Пропуск: не заполнено «Отгрузка через» (или «Поставщик МС»)'
      });
      continue;
    }

    if (statusOk && hasLine && !isMsSyncSpecReady_(row)) {
      skippedBySpec++;
      msSyncApplyRowPatch_(sourceCtx, updater, layout, i, row, {
        status: '⚠️ Пропуск: не заполнен номер спецификации / инвойса'
      });
      continue;
    }

    if (msId && msId !== 'ID заказа МС') originalMsIds.add(msId);

    if (!supplierName) { 
      if (hasLine && statusOk) {
        msSyncApplyRowPatch_(sourceCtx, updater, layout, i, row, {
          status: '⚠️ Пропуск: не заполнено «Отгрузка через» (или «Поставщик МС»)'
        });
      }
      continue; 
    }

    const groupKey = supplierName + '|||' + specNum;

    if (!groups[groupKey]) {
      groups[groupKey] = {
        supplierName: supplierName,
        supplierCandidates: candidates,
        specNum: specNum,
        candidateMsIds: [],
        rows: []
      };
    }
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

  if (!Object.keys(groups).length) {
    updater.flush();
    const activeTab = safeString(ss.getActiveSheet().getName());
    const onSummary = activeTab === SHEET_NAME;
    return SpreadsheetApp.getUi().alert(
      'Синхронизация завершена: не найдено строк для заказов в МС.\n\n' +
        'Скрипт обрабатывает только лист «Сводная» (не вкладку менеджера).\n' +
        (onSummary
          ? '• Заполните «Поставщик МС», «Отгрузка через» или «Поставщик» (и артикулы/ШК).\n' +
            '  Шапка распознана в строке ' +
            (headerRowIdx + 1) +
            '.\n'
          : '• Сейчас открыта вкладка «' + activeTab + '» — соберите Сводную:\n' +
            '  📦 МойСклад → 🧩 Собрать «Сводная» из вкладок менеджеров.\n' +
            '• Затем снова «Синхронизировать заказы».\n') +
        '• Контрагент в МС ищется по «Внутреннее название» в справочнике\n' +
        '  (часто это «Отгрузка через», а не ФИО в «Поставщик»).\n' +
        '• В МС уходят только строки со статусом: ' +
        statusHint +
        '.'
    );
  }
  
  const supplierMap = getSupplierCodeMap();

  for (const key in groups) {
    const group = groups[key];
    const candidates = group.supplierCandidates && group.supplierCandidates.length
      ? group.supplierCandidates
      : [group.supplierName];
    const found = lookupMsSupplier_(candidates, supplierMap);
    const supplier = found.supplier;

    if (!supplier) { 
      const tried = candidates.join(' → ');
      group.rows.forEach(function (r) {
        msSyncApplyRowPatch_(sourceCtx, updater, layout, r.rowIndex, r.rowData, {
          status: '❌ Не найден в МС, пробовали: ' + tried
        });
      });
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
      if (!product) {
        msSyncApplyRowPatch_(sourceCtx, updater, layout, item.rowIndex, item.rowData, {
          status: '❌ Ошибка товара'
        });
        hasErrors = true;
        continue;
      }
      
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

    const sourceTabNames = resolveSourceTabNamesForGroup_(sourceCtx, group.rows);
    const payload = buildPurchaseOrderPayload(
      group.rows[0].rowData,
      supplier,
      positions,
      mergedPaymentData,
      sourceTabNames
    );

    let res;
    if (group.msId) {
      res = msFetch('/entity/purchaseorder/' + encodeURIComponent(group.msId), 'put', payload);
      if (res.success) {
        group.rows.forEach(function (item) {
          msSyncApplyRowPatch_(sourceCtx, updater, layout, item.rowIndex, item.rowData, {
            msId: group.msId,
            msLink: res.data ? res.data.meta.uuidHref : item.rowData[COL.MS_LINK],
            status: '🔄 Обновлено в МС (Синхронизировано)'
          });
        });
        updatedOrders++;
      } else {
        group.rows.forEach(function (item) {
          msSyncApplyRowPatch_(sourceCtx, updater, layout, item.rowIndex, item.rowData, {
            status: '❌ Ошибка обн.: ' + res.error
          });
        });
        errorCount++;
      }
    } else {
      res = msFetch('/entity/purchaseorder', 'post', payload);
      if (res.success && res.data) {
        group.rows.forEach(function (item) {
          msSyncApplyRowPatch_(sourceCtx, updater, layout, item.rowIndex, item.rowData, {
            msId: res.data.id,
            msLink: res.data.meta.uuidHref || '',
            status: '✅ Создан новый заказ в МС'
          });
        });
        createdOrders++;
      } else {
        group.rows.forEach(function (item) {
          msSyncApplyRowPatch_(sourceCtx, updater, layout, item.rowIndex, item.rowData, {
            status: '❌ Ошибка созд.: ' + res.error
          });
        });
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
  SpreadsheetApp.getUi().alert(
    'Синхронизация завершена!\n\n' +
      '🆕 Создано: ' +
      createdOrders +
      '\n🔄 Обновлено: ' +
      updatedOrders +
      '\n⏭ Пропущено (другой статус): ' +
      skippedByStatus +
      '\n⏭ Пропущено (нет «Отгрузка через»): ' +
      skippedByShipVia +
      '\n⏭ Пропущено (нет спецификации): ' +
      skippedBySpec +
      '\n🗑️ Удалено/Очищено брошенных: ' +
      deletedOrders +
      '\n❌ Ошибок: ' +
      errorCount +
      '\n\nВ МС синхронизируются только строки со статусом: ' +
      statusHint +
      '.\n' +
      '📋 Записано на вкладки-источники: ' +
      sourceCtx.written +
      (sourceCtx.notFound ? ' (не найдена строка: ' + sourceCtx.notFound + ')' : '') +
      '.'
  );
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
  syncManagerTabsToSummaryImpl_(SpreadsheetApp.getActiveSpreadsheet(), {});
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {{ silent?: boolean }} opt
 * @returns {string|void}
 */
function syncManagerTabsToSummaryImpl_(ss, opt) {
  const silent = !!(opt && opt.silent);
  const ui = silent ? null : SpreadsheetApp.getUi();
  if (!ss) {
    throw new Error('Не передана книга для сборки Сводной.');
  }
  const summarySh = ss.getSheetByName(SHEET_NAME);
  if (!summarySh) {
    const msg = 'Не найден лист «' + SHEET_NAME + '».';
    if (silent) throw new Error(msg);
    ui.alert(msg);
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
    const noSrcMsg =
      'Нет вкладок для сборки.\n' +
      'Ожидаются вкладки «Имя ММ/ГГ» (например «Нина 07/26») или листы с префиксом «' +
      HISTORY_SHIPMENTS_SHEET_PREFIX +
      '».\n' +
      'Данные должны начинаться со строки ' +
      MANAGER_SUMMARY_SYNC_DATA_START_ROW +
      '. Исключены «Сводная», справочники и системные листы.';
    if (silent) throw new Error(noSrcMsg.replace(/\n/g, ' '));
    ui.alert(noSrcMsg);
    return;
  }

  // Целевая раскладка Сводной — каноническая, по COL-константам (см.
  // syncManagerSummaryCanonicalHeader_). widest-эвристика убрана: она привязывала
  // структуру Сводной к произвольному порядку колонок в менеджерских табах,
  // ломала валидации/условное форматирование и приводила к подмене столбцов
  // (например, «Менеджер» затирал значение «Статус заказа» в колонке H, потому что
  // в «Никита 07/26» «Статус заказа» стоит на позиции 8, а «Менеджер» отсутствует).
  const header = syncManagerSummaryCanonicalHeader_();
  const headerKeys = header.map(function (h) {
    const canon = syncManagerCanonHeader_(h);
    return syncManagerHeaderKey_(canon);
  });
  const headerKeyToIdx = syncManagerBuildHeaderIndex_(headerKeys);

  // Колонки из источников, ключи которых не покрыты каноном (например, «Тип доставки»,
  // «Заказ в МС», «Ссылка на инвойс», «Аванс № заявки» и т.п.) — дописываем в конец заголовка,
  // чтобы пользовательские данные не терялись. Имя берётся из первого источника, где такой
  // ключ встретился.
  syncManagerSummaryAppendUnknownSourceKeys_(header, headerKeys, headerKeyToIdx, sources);

  let tripModeMap = {};
  let deliveryModeMismatch = 0;
  if (typeof logisticsLoadTripDeliveryModeMap_ === 'function' && typeof logisticsOpenBook05_ === 'function') {
    try {
      tripModeMap = logisticsLoadTripDeliveryModeMap_(logisticsOpenBook05_());
    } catch (e) {
      // Книга 05 недоступна — сборка Сводной без наследования режима.
    }
  }

  let maxCol = header.length;

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
        // managerIdx всегда === COL.MANAGER благодаря канонической раскладке. Доступ через
        // headerKeyToIdx — на случай, если COL когда-нибудь поменяется: запись всё равно
        // пойдёт в колонку с ключом 'manager', а не в случайную позицию.
        const managerIdx = headerKeyToIdx['manager'];
        if (managerIdx != null && managerIdx < row.length) row[managerIdx] = srcEntry.meta.manager;
        monthKey = srcEntry.meta.monthKey;
        labelRus = srcEntry.meta.labelRus;
      } else {
        const histMeta = syncManagerHistoryRowMeta_(row, headerKeyToIdx);
        monthKey = histMeta.monthKey;
        labelRus = histMeta.labelRus;
      }

      const shipIdx = headerKeyToIdx['shipment_id'];
      const modeIdx = headerKeyToIdx['delivery_mode'];
      if (shipIdx != null && modeIdx != null && modeIdx < row.length) {
        const sid = String(row[shipIdx] || '').trim();
        const fromTrip = sid ? tripModeMap[sid] : '';
        if (fromTrip) {
          const existing = String(row[modeIdx] || '').trim();
          if (!existing) {
            row[modeIdx] = fromTrip;
          } else if (
            typeof logisticsNorm_ === 'function' &&
            logisticsNorm_(existing) !== logisticsNorm_(fromTrip)
          ) {
            deliveryModeMismatch++;
          }
        }
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

  const doneMsg =
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
    '.' +
    (deliveryModeMismatch
      ? '\n⚠️ Расхождение «Тип доставки» со «Рейсы» (05): ' + deliveryModeMismatch + ' строк.'
      : '');
  if (!silent) {
    ui.alert(doneMsg);
  }
  return doneMsg;
}

/**
 * Меню: списать выделенные строки активной вкладки в рейс (статус «4. В пути в Москву» + колонка «Рейс»).
 */
function assignSelectedRowsToShipment() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!sheet || sheet.getName() === SHEET_NAME) {
    ui.alert('Откройте вкладку менеджера или «История отгрузок», выделите строки и повторите.');
    return;
  }
  if (sheet.getLastRow() < MANAGER_SUMMARY_SYNC_DATA_START_ROW) {
    ui.alert('На листе нет строк данных.');
    return;
  }

  let validShipments = null;
  if (typeof logisticsOpenBook05_ === 'function' && typeof logisticsLoadValidShipmentIds_ === 'function') {
    try {
      validShipments = logisticsLoadValidShipmentIds_(logisticsOpenBook05_());
    } catch (e) {
      ui.alert('Не удалось открыть книгу 05: ' + (e.message || String(e)));
      return;
    }
  }

  const pick = ui.prompt(
    'Рейс (SHIPMENT_ID)',
    'Введите ID рейса (например TR-2026-0011):',
    ui.ButtonSet.OK_CANCEL
  );
  if (pick.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(pick.getResponseText() || '').trim();
  if (!shipmentId) {
    ui.alert('Пустой SHIPMENT_ID.');
    return;
  }
  if (validShipments && !validShipments[shipmentId]) {
    const cont = ui.alert(
      'Рейс не найден',
      'ID «' +
        shipmentId +
        '» отсутствует на листе «Рейсы» в книге 05.\nВсё равно записать в строки?',
      ui.ButtonSet.YES_NO
    );
    if (cont !== ui.Button.YES) return;
  }

  const lc = sheet.getLastColumn();
  const headerRow = MANAGER_SUMMARY_SYNC_HEADER_ROW;
  const headerRaw = sheet.getRange(headerRow, 1, 1, lc).getValues()[0] || [];
  const srcKeys = headerRaw.map(function (h) {
    return syncManagerHeaderKey_(syncManagerCanonHeader_(h));
  });
  const srcMap = syncManagerBuildHeaderIndex_(srcKeys);
  const idxStatus = srcMap.order_status;
  const idxShip = srcMap.shipment_id;
  if (idxStatus == null || idxShip == null) {
    ui.alert('На листе нужны колонки «Статус заказа» и «Рейс» (или добавьте «Рейс» в шапку).');
    return;
  }

  const shippedStatus =
    typeof LOGISTICS_CFG !== 'undefined' && LOGISTICS_CFG.SHIPPED_STATUS
      ? LOGISTICS_CFG.SHIPPED_STATUS
      : '4. В пути в Москву';

  const range = ss.getActiveRange();
  const startRow = Math.max(range.getRow(), MANAGER_SUMMARY_SYNC_DATA_START_ROW);
  const endRow = Math.min(range.getLastRow(), sheet.getLastRow());
  if (endRow < startRow) {
    ui.alert('Выделите строки данных.');
    return;
  }

  let updated = 0;
  for (let r = startRow; r <= endRow; r++) {
    sheet.getRange(r, idxStatus + 1).setValue(shippedStatus);
    sheet.getRange(r, idxShip + 1).setValue(shipmentId);
    updated++;
  }
  ui.alert('Готово', 'Обновлено строк: ' + updated + '\nРейс: ' + shipmentId, ui.ButtonSet.OK);
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
/**
 * Дополняет канонический заголовок Сводной колонками из источников, ключи которых не покрыты
 * каноном (например, «Тип доставки», «Заказ в МС», «Аванс № заявки» и т.п.).
 * Имя колонки берётся из первой вкладки, где такой ключ встретился. Дубликаты не добавляются.
 *
 * Мутирует `header`, `headerKeys`, `headerKeyToIdx` на месте.
 */
function syncManagerSummaryAppendUnknownSourceKeys_(header, headerKeys, headerKeyToIdx, sources) {
  let s;
  let j;
  for (s = 0; s < sources.length; s++) {
    const sht = sources[s].sheet;
    const lc = sht.getLastColumn();
    if (lc <= 0) continue;
    const hdr =
      sht
        .getRange(MANAGER_SUMMARY_SYNC_HEADER_ROW, 1, 1, lc)
        .getValues()[0] || [];
    for (j = 0; j < hdr.length; j++) {
      const orig = hdr[j];
      const canon = syncManagerCanonHeader_(orig);
      const key = canon ? syncManagerHeaderKey_(canon) : '';
      if (!key) continue;
      if (headerKeyToIdx[key] != null) continue;
      header.push(String(orig));
      headerKeys.push(key);
      headerKeyToIdx[key] = header.length - 1;
    }
  }
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
 * Канонический заголовок Сводной: фиксированный порядок колонок по COL-константам, плюс
 * служебные «Период (MM/YY)» и «Плановая дата поступления» в конце.
 *
 * Это единственный источник истины и для:
 *   — целевой раскладки при сборке Сводной (`syncManagerTabsToSummary`),
 *   — шаблона нового листа «История отгрузок» (`ensureHistoryShipmentsSheet_`).
 *
 * Раньше Сводная брала шапку из «самой широкой» вкладки-источника (widest). Это привязывало
 * структуру Сводной к произвольному порядку колонок в менеджерских табах и ломало
 * пользовательские валидации/условное форматирование, а также приводило к подмене столбцов
 * (в «Никита 07/26» «Статус заказа» стоял на позиции H, «Менеджер» отсутствовал — имя
 * менеджера записывалось поверх значения статуса). Канон делает раскладку устойчивой.
 */
function syncManagerSummaryCanonicalHeader_() {
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
  // Период (MM/YY) и ETA — отдельные колонки с правой стороны, чтобы корректно определять
  // месяц прибытия для книги 03 (она читает period/eta/ready_date в этом приоритете).
  header.push('Период (MM/YY)');
  header.push('Плановая дата поступления');
  return header;
}

/**
 * Шаблон листа «История отгрузок». Делегирует в канон — раскладка Сводной и шаблона
 * исторических отгрузок гарантированно одинаковая. Пользователь может переименовать/
 * переставить колонки в своих менеджерских вкладках — сборщик матчит по ключам.
 */
function syncManagerHistorySheetDefaultHeader_() {
  return syncManagerSummaryCanonicalHeader_();
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
  if (
    h === 'тип доставки' ||
    h === 'режим доставки' ||
    h === 'delivery mode' ||
    h === 'delivery_mode' ||
    h === 'способ доставки'
  ) {
    return 'delivery_mode';
  }
  if (h === 'отгрузка через' || h === 'ship via' || h === 'ship_via') return 'ship_via';
  if (h === 'доставка' && h.indexOf('тип') === -1) return 'ship_via';
  if (h === 'статус заказа' || h === 'статус' || h === 'order status' || h === 'status') return 'order_status';

  if (h === 'поставщик мс' || h === 'supplier ms' || h === 'ms supplier') return 'supplier_ms';
  if (h === 'поставщик' || h === 'supplier' || h === 'фабрика' || h === 'factory') return 'supplier_note';

  if (h === 'ms id' || h === 'ms_id' || h === 'id заказа мс') return 'ms_id';
  if (h === 'ms link' || h === 'ms_link') return 'ms_link';
  if (h === 'статус скрипта' || h === 'script status' || h === 'статус синхронизации') return 'script_status';
  if (h === 'заказ в мс' || h === 'заказ мс' || h === 'order in ms') return 'ms_order_sheet';

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
    delivery_mode: 'Тип доставки',
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

function buildPurchaseOrderPayload(row, supplier, positions, paymentData, sourceTabNames) {
  const orgId = normalizeMsEntityId_(CONFIG.MS_ORGANIZATION_ID);
  const payload = {
    organization: buildMeta('organization', orgId),
    agent: { meta: supplier.meta },
    description: buildDescription(row, sourceTabNames),
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

function buildDescription(row, sourceTabNames) {
  const parts = [];
  const tabs = sourceTabNames || [];
  if (tabs.length === 1) {
    parts.push('Вкладка (план): ' + tabs[0]);
  } else if (tabs.length > 1) {
    parts.push('Вкладки (план): ' + tabs.join(', '));
  }
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
  constructor(sheet, data, layout) {
    this.sheet = sheet;
    this.numRows = data.length;
    this.updates = {};
    this.headerRowIdx = layout && layout.headerRowIdx != null ? layout.headerRowIdx : MANAGER_SUMMARY_SYNC_HEADER_ROW - 1;
    this.colStatus = layout && layout.colStatus != null ? layout.colStatus : COL.STATUS;
    this.dataStartIdx = this.headerRowIdx + 1;
    this.dataStartSheetRow = this.headerRowIdx + 2;
    this.headerSheetRow = this.headerRowIdx + 1;
  }
  setValue(rowIndex, colIndex, value) {
    if (!this.updates[colIndex]) this.updates[colIndex] = new Array(this.numRows).fill('');
    this.updates[colIndex][rowIndex] = value;
  }
  setStatus(rowIndex, text) { this.setValue(rowIndex, this.colStatus, text); }
  flush() {
    const headerCell = this.sheet.getRange(this.headerSheetRow, this.colStatus + 1);
    if (!headerCell.getValue()) headerCell.setValue('Статус скрипта');
    for (const colIndex in this.updates) {
      const colData = this.updates[colIndex].slice(this.dataStartIdx).map(function (v) {
        return [v !== '' ? v : null];
      });
      if (colData.length === 0) continue;
      const range = this.sheet.getRange(this.dataStartSheetRow, parseInt(colIndex, 10) + 1, colData.length, 1);
      const merged = range.getValues().map(function (row, i) {
        return [colData[i][0] !== null ? colData[i][0] : row[0]];
      });
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