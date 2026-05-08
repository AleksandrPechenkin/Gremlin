const SYNC_HUB_CFG = {
  PROPS: {
    MASTER_REF_SPREADSHEET_ID: 'MASTER_REF_SPREADSHEET_ID',
    TRANSIT_SPREADSHEET_ID: 'TRANSIT_SPREADSHEET_ID',
    PROCUREMENT_SPREADSHEET_ID: 'PROCUREMENT_SPREADSHEET_ID',
    COST_SPREADSHEET_ID: 'COST_SPREADSHEET_ID',
    SYNC_BATCH_SIZE: 'SYNC_BATCH_SIZE',
    SYNC_LOCK_TIMEOUT_MS: 'SYNC_LOCK_TIMEOUT_MS',
    SYNC_LOG_SHEET_NAME: 'SYNC_LOG_SHEET_NAME',
    MASTER_PRODUCTS_SHEET_NAME: 'MASTER_PRODUCTS_SHEET_NAME',
    MASTER_SUPPLIERS_SHEET_NAME: 'MASTER_SUPPLIERS_SHEET_NAME',
    TRANSIT_SNAPSHOT_SHEET_NAME: 'TRANSIT_SNAPSHOT_SHEET_NAME',
    SHIPMENT_PLAN_SHEET_NAME: 'SHIPMENT_PLAN_SHEET_NAME',
    STOCK_MOVEMENTS_SHEET_NAME: 'STOCK_MOVEMENTS_SHEET_NAME',
    PROC_PLAN_SNAPSHOT_SHEET_NAME: 'PROC_PLAN_SNAPSHOT_SHEET_NAME',
    PROC_CALC_SNAPSHOT_SHEET_NAME: 'PROC_CALC_SNAPSHOT_SHEET_NAME',
    PROC_MS_STORES_SHEET_NAME: 'PROC_MS_STORES_SHEET_NAME',
    COST_TRIP_EXPENSES_SHEET_NAME: 'COST_TRIP_EXPENSES_SHEET_NAME',
    COST_CUSTOMS_SHEET_NAME: 'COST_CUSTOMS_SHEET_NAME',
    COST_ALLOCATION_SHEET_NAME: 'COST_ALLOCATION_SHEET_NAME',
    COST_SKU_SHEET_NAME: 'COST_SKU_SHEET_NAME'
  },
  DEFAULTS: {
    SYNC_BATCH_SIZE: 800,
    SYNC_LOCK_TIMEOUT_MS: 30000,
    SYNC_LOG_SHEET_NAME: 'SYNC_LOG',
    MASTER_PRODUCTS_SHEET_NAME: 'Справочник товары',
    MASTER_SUPPLIERS_SHEET_NAME: 'Справочник поставщики и условия',
    TRANSIT_SNAPSHOT_SHEET_NAME: 'Транзитный склад',
    SHIPMENT_PLAN_SHEET_NAME: 'Планирование отгрузок',
    STOCK_MOVEMENTS_SHEET_NAME: 'Stock_Movements',
    PROC_PLAN_SNAPSHOT_SHEET_NAME: 'Планирование закупок',
    PROC_CALC_SNAPSHOT_SHEET_NAME: 'Планирование закупок (расчёт)',
    PROC_MS_STORES_SHEET_NAME: 'Склады МС (остатки)',
    COST_TRIP_EXPENSES_SHEET_NAME: 'Затраты рейса',
    COST_CUSTOMS_SHEET_NAME: 'Таможенные платежи',
    COST_ALLOCATION_SHEET_NAME: 'Аллокация затрат',
    COST_SKU_SHEET_NAME: 'Себестоимость SKU'
  }
};

function addSyncHubMenu_(ui) {
  ui.createMenu('🔁 Синхронизация книг')
    .addItem('Проверить настройки синхронизации', 'syncHubHealthCheck')
    .addItem('Показать листы книги 04 (диагностика)', 'syncHubShowMasterSheets_')
    .addSeparator()
    .addItem('Синхронизировать справочники (из 04)', 'syncMasterRefsFrom04_')
    .addItem('Пробный прогон (dry-run)', 'syncAllExternalBooksDryRun_')
    .addToUi();
}

function syncHubHealthCheck() {
  const ui = SpreadsheetApp.getUi();
  const required = [
    SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID
  ];
  const missing = required.filter(function (k) {
    return !syncHubGetProp_(k, '');
  });

  if (missing.length) {
    ui.alert(
      '⚠️ Не хватает Script Properties',
      'Заполните свойства:\n- ' + missing.join('\n- '),
      ui.ButtonSet.OK
    );
    return;
  }
  ui.alert('✅ Проверка пройдена', 'Базовые настройки синхронизации заполнены.', ui.ButtonSet.OK);
}

function syncAllExternalBooks_() {
  return syncAllExternalBooksImpl_(false);
}

function syncAllExternalBooksDryRun_() {
  return syncAllExternalBooksImpl_(true);
}

function syncAllExternalBooksImpl_(dryRun) {
  const lock = LockService.getScriptLock();
  const lockTimeoutMs = syncHubGetNumberProp_(
    SYNC_HUB_CFG.PROPS.SYNC_LOCK_TIMEOUT_MS,
    SYNC_HUB_CFG.DEFAULTS.SYNC_LOCK_TIMEOUT_MS
  );
  lock.waitLock(Math.max(1000, lockTimeoutMs));
  try {
    const report = [];
    report.push(syncMasterRefsFrom04_(dryRun));
    SpreadsheetApp.getUi().alert(
      dryRun ? '🧪 Dry-run завершен' : '✅ Синхронизация завершена',
      report.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return report.join('\n');
  } finally {
    lock.releaseLock();
  }
}

function syncMasterRefsFrom04_(dryRun) {
  try {
    const sourceId = syncHubGetRequiredProp_(SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID);
    const suppliersSourceSheet = syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_SUPPLIERS_SHEET_NAME,
      SYNC_HUB_CFG.DEFAULTS.MASTER_SUPPLIERS_SHEET_NAME
    );
    const mappings = [
      {
        source: syncHubGetProp_(SYNC_HUB_CFG.PROPS.MASTER_PRODUCTS_SHEET_NAME, SYNC_HUB_CFG.DEFAULTS.MASTER_PRODUCTS_SHEET_NAME),
        target: 'Справочник товары'
      },
      {
        source: suppliersSourceSheet,
        sourceAliases: [
          'Справочник поставщики и условия',
          'Справочник поставщики и условия работы'
        ],
        target: 'Справочник поставщики и условия'
      }
    ];
    const sourceSs = SpreadsheetApp.openById(sourceId);
    const targetSs = SpreadsheetApp.getActiveSpreadsheet();
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, { dryRun: !!dryRun });
    const msg = 'Справочники: ' + stats;
    syncHubLog_('04→01 Справочники', 'OK', msg, !!dryRun);
    return msg;
  } catch (e) {
    syncHubLog_('04→01 Справочники', 'ERROR', e.message || String(e), !!dryRun);
    throw e;
  }
}

function syncHubCopyMappings_(sourceSs, targetSs, mappings, options) {
  const dryRun = !!(options && options.dryRun);
  const batchSize = Math.max(100, syncHubGetNumberProp_(SYNC_HUB_CFG.PROPS.SYNC_BATCH_SIZE, SYNC_HUB_CFG.DEFAULTS.SYNC_BATCH_SIZE));
  const parts = [];

  mappings.forEach(function (m) {
    const sourceSheet = syncHubFindSheetByNames_(sourceSs, [m.source].concat(m.sourceAliases || []));
    if (!sourceSheet) {
      throw new Error(
        'Не найден исходный лист. Проверены имена: ' + [m.source].concat(m.sourceAliases || []).join(', ')
      );
    }

    let targetSheet = targetSs.getSheetByName(m.target);
    if (!targetSheet) {
      parts.push(m.target + ': целевой лист отсутствует в 01, пропущено');
      return;
    }

    const rows = sourceSheet.getLastRow();
    const cols = sourceSheet.getLastColumn();

    if (!rows || !cols) {
      parts.push(m.target + ': источник пуст (rows=' + rows + ', cols=' + cols + '), пропущено');
      return;
    }

    if (dryRun) {
      parts.push(m.target + ': dry-run, было бы скопировано ' + rows + ' строк, ' + cols + ' колонок');
      return;
    }

    // Безопасный режим: очищаем цель только если источник не пустой.
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

function syncHubFindSheetByNames_(ss, names) {
  const uniq = {};
  const wantedNorm = {};
  for (let i = 0; i < names.length; i++) {
    const nm = String(names[i] || '').trim();
    if (!nm || uniq[nm]) continue;
    uniq[nm] = true;
    const sh = ss.getSheetByName(nm);
    if (sh) return sh;
    wantedNorm[syncHubNormSheetName_(nm)] = true;
  }

  // Фолбэк: мягкое сравнение имен (лишние/неразрывные пробелы, регистр, ё/е).
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    const sh = all[i];
    const n = syncHubNormSheetName_(sh.getName());
    if (wantedNorm[n]) return sh;
  }
  return null;
}

function syncHubNormSheetName_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function syncHubShowMasterSheets_() {
  const masterId = syncHubGetRequiredProp_(SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID);
  const ss = SpreadsheetApp.openById(masterId);
  const names = ss.getSheets().map(function (sh) { return sh.getName(); });
  SpreadsheetApp.getUi().alert(
    'Листы книги 04',
    names.length ? names.join('\n') : '(листов не найдено)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncHubLog_(blockName, status, details, dryRun) {
  const masterId = syncHubGetProp_(
    SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID,
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );
  const ss = SpreadsheetApp.openById(masterId);
  const logName = syncHubGetProp_(SYNC_HUB_CFG.PROPS.SYNC_LOG_SHEET_NAME, SYNC_HUB_CFG.DEFAULTS.SYNC_LOG_SHEET_NAME);
  let sh = ss.getSheetByName(logName);
  if (!sh) {
    sh = ss.insertSheet(logName);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Block', 'Status', 'DryRun', 'User', 'Details']]);
    sh.setFrozenRows(1);
  }
  const userEmail = Session.getEffectiveUser().getEmail() || '';
  sh.appendRow([new Date(), blockName, status, dryRun ? 'YES' : 'NO', userEmail, details]);
}

function syncHubGetRequiredProp_(key) {
  const v = syncHubGetProp_(key, '');
  if (!v) {
    throw new Error('Не заполнено Script Property: ' + key);
  }
  return v;
}

function syncHubGetProp_(key, fallback) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (raw == null || String(raw).trim() === '') return fallback;
  return String(raw).trim();
}

function syncHubGetNumberProp_(key, fallback) {
  const n = Number(syncHubGetProp_(key, fallback));
  return isFinite(n) ? n : fallback;
}
