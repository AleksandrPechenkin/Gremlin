/**
 * Sync Hub: обмен между книгами 01–05. Журнал по умолчанию пишется в книгу 04 (MASTER_REF_SPREADSHEET_ID).
 *
 * Реестр книг (Script Properties проекта, куда вставлен sync_hub.gs):
 *   MASTER_REF_SPREADSHEET_ID — книга 04 (справочники + SYNC_LOG)
 *   ORDERS_SPREADSHEET_ID     — книга 01 (заказы); если пусто и активная книга ≠ 04 — берётся текущий файл
 *   TRANSIT_SPREADSHEET_ID   — книга 02
 *   PROCUREMENT_SPREADSHEET_ID — книга 03
 *   COST_SPREADSHEET_ID      — книга 05
 *
 * Операционные снимки (Script Property, опционально):
 *   ORDERS_SUMMARY_SHEET_NAME — лист заказной сводки в 01 для копирования в 02/03 (по умолчанию: Сводная)
 */
const SYNC_HUB_CFG = {
  PROPS: {
    MASTER_REF_SPREADSHEET_ID: 'MASTER_REF_SPREADSHEET_ID',
    ORDERS_SPREADSHEET_ID: 'ORDERS_SPREADSHEET_ID',
    ORDERS_SUMMARY_SHEET_NAME: 'ORDERS_SUMMARY_SHEET_NAME',
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
    ORDERS_SUMMARY_SHEET_NAME: 'Сводная',
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
    COST_SKU_SHEET_NAME: 'Себестоимость SKU',
    MASTER_REF_EXTERNAL_SPREADSHEET_ID: 'MASTER_REF_EXTERNAL_SPREADSHEET_ID',
    MASTER_REF_EXTERNAL_PRODUCTS_SHEET: 'MASTER_REF_EXTERNAL_PRODUCTS_SHEET',
    MASTER_REF_EXTERNAL_SUPPLIERS_SHEET: 'MASTER_REF_EXTERNAL_SUPPLIERS_SHEET'
  },
  EXTERNAL_DEFAULTS: {
    SPREADSHEET_ID: '1PXWd05ENcZGvPYYbAVvf-1EPevdwkxr4IvjRbbojOlg',
    PRODUCTS_SHEET: 'Справочник с названием товаров',
    SUPPLIERS_SHEET: 'Справочник поставщики и условия'
  }
};

function addSyncHubMenu_(ui) {
  ui.createMenu('🔁 Синхронизация книг')
    .addItem('Проверить настройки синхронизации', 'syncHubHealthCheck')
    .addItem('Реестр книг 01–05 (Script Properties)', 'syncHubShowBookRegistry_')
    .addItem('Показать листы книги 04 (диагностика)', 'syncHubShowMasterSheets_')
    .addSeparator()
    .addItem('Синхронизировать справочники (из 04)', 'syncMasterRefsFrom04_')
    .addItem('Собрать справочники в 04 (из 05)', 'syncCollectRefsTo04From05_')
    .addItem('Восстановить полный «Справочник товары» в 04 (из внешней книги)', 'syncRestoreFullProductsTo04WithConfirm_')
    .addSeparator()
    .addItem('Пробный прогон (dry-run все потоки)', 'syncAllExternalBooksDryRun_')
    .addItem('Полная синхронизация (все потоки)', 'syncAllExternalBooksWithConfirm_')
    .addSubMenu(
      ui
        .createMenu('Операционные потоки')
        .addItem('Dry-run: Сводная 01→02 и 03', 'syncOperationalSnapshotsDryRun_')
        .addItem('Сводная 01→02 и 03 (боевой)', 'syncOperationalSnapshotsWithConfirm_')
        .addSeparator()
        .addItem('Dry-run: Сводная 01→03', 'syncOperationalOrdersSummaryFrom01To03DryRun_')
        .addItem('Сводная 01→03 (боевой)', 'syncOperationalOrdersSummaryFrom01To03WithConfirm_')
    )
    .addSubMenu(
      ui
        .createMenu('⏱ Расписание')
        .addItem('Установить триггеры этой книги (04)', 'gremlinScheduleInstallTriggers04_')
        .addItem('Снять триггеры этой книги', 'gremlinScheduleRemoveTriggers_')
        .addSeparator()
        .addItem('Сейчас: справочники (внешние → 04 → 01)', 'gremlinScheduleRunHourlyRefsNow_')
        .addItem('Сейчас: Сводная 01→02 и 03', 'gremlinScheduleRunHourlySnapshotsNow_')
        .addItem('Сейчас: полный hourly-цикл 04', 'gremlinScheduleRunHourlyFull04Now_')
    )
    .addToUi();
}

/**
 * Восстановление мастер-справочника товаров в книге 04 из внешней книги.
 * Источник сейчас фиксированный (по вашему сообщению):
 * - Spreadsheet ID: 1PXWd05ENcZGvPYYbAVvf-1EPevdwkxr4IvjRbbojOlg
 * - Sheet: «Справочник с названием товаров»
 */
function syncRestoreFullProductsTo04WithConfirm_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Восстановление справочника товаров',
    'Перезаписать лист «Справочник товары» в книге 04 данными из внешней книги?\n\n' +
      'Источник: 1PXWd05ENcZGvPYYbAVvf-1EPevdwkxr4IvjRbbojOlg\n' +
      'Лист: «Справочник с названием товаров»\n\n' +
      'Операция полностью очистит целевой лист и запишет значения.',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;
  const msg = syncRestoreFullProductsTo04_(false);
  ui.alert('✅ Готово', msg, ui.ButtonSet.OK);
}

function syncRestoreFullProductsTo04DryRun_() {
  return syncRestoreFullProductsTo04_(true);
}

function syncHubGetExternalRefConfig_() {
  return {
    spreadsheetId: syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_REF_EXTERNAL_SPREADSHEET_ID,
      SYNC_HUB_CFG.EXTERNAL_DEFAULTS.SPREADSHEET_ID
    ),
    productsSheet: syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_REF_EXTERNAL_PRODUCTS_SHEET,
      SYNC_HUB_CFG.EXTERNAL_DEFAULTS.PRODUCTS_SHEET
    ),
    suppliersSheet: syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_REF_EXTERNAL_SUPPLIERS_SHEET,
      SYNC_HUB_CFG.EXTERNAL_DEFAULTS.SUPPLIERS_SHEET
    )
  };
}

function syncRestoreFullProductsTo04_(dryRun) {
  return syncRestoreProductsFromExternal_(dryRun);
}

/**
 * Товары: внешняя книга → «Справочник товары» в 04.
 */
function syncRestoreProductsFromExternal_(dryRun) {
  const cfg = syncHubGetExternalRefConfig_();
  try {
    const sourceSs = SpreadsheetApp.openById(cfg.spreadsheetId);
    const targetSs = syncHubOpenSpreadsheetForBook_('04');
    const targetName = syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_PRODUCTS_SHEET_NAME,
      SYNC_HUB_CFG.DEFAULTS.MASTER_PRODUCTS_SHEET_NAME
    );
    const mappings = [
      {
        source: cfg.productsSheet,
        sourceAliases: [cfg.productsSheet, SYNC_HUB_CFG.EXTERNAL_DEFAULTS.PRODUCTS_SHEET],
        target: targetName,
        targetAliases: [SYNC_HUB_CFG.DEFAULTS.MASTER_PRODUCTS_SHEET_NAME, 'Справочник товары'],
        required: true
      }
    ];
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
      dryRun: !!dryRun,
      createMissingTarget: true
    });
    const msg = 'Восстановление товаров в 04: ' + stats;
    syncHubLog_('RESTORE 04 Товары', 'OK', msg, !!dryRun);
    return msg;
  } catch (e) {
    syncHubLog_('RESTORE 04 Товары', 'ERROR', e.message || String(e), !!dryRun);
    throw e;
  }
}

/**
 * Поставщики: внешняя книга → «Справочник поставщики и условия» в 04.
 */
function syncRestoreSuppliersFromExternal_(dryRun) {
  const cfg = syncHubGetExternalRefConfig_();
  const suppliersSheet = String(cfg.suppliersSheet || '').trim();
  if (!suppliersSheet || suppliersSheet.toLowerCase() === 'off') {
    const msg = 'Пропуск поставщиков (MASTER_REF_EXTERNAL_SUPPLIERS_SHEET пуст или OFF)';
    syncHubLog_('RESTORE 04 Поставщики', 'SKIP', msg, !!dryRun);
    return msg;
  }
  try {
    const sourceSs = SpreadsheetApp.openById(cfg.spreadsheetId);
    const targetSs = syncHubOpenSpreadsheetForBook_('04');
    const targetName = syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.MASTER_SUPPLIERS_SHEET_NAME,
      SYNC_HUB_CFG.DEFAULTS.MASTER_SUPPLIERS_SHEET_NAME
    );
    const mappings = [
      {
        source: suppliersSheet,
        sourceAliases: [
          suppliersSheet,
          'Справочник поставщики и условия',
          'Справочник поставщики и условия работы'
        ],
        target: targetName,
        targetAliases: [
          SYNC_HUB_CFG.DEFAULTS.MASTER_SUPPLIERS_SHEET_NAME,
          'Справочник поставщики и условия',
          'Справочник поставщики и условия работы'
        ],
        required: true
      }
    ];
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
      dryRun: !!dryRun,
      createMissingTarget: true
    });
    const msg = 'Восстановление поставщиков в 04: ' + stats;
    syncHubLog_('RESTORE 04 Поставщики', 'OK', msg, !!dryRun);
    return msg;
  } catch (e) {
    syncHubLog_('RESTORE 04 Поставщики', 'ERROR', e.message || String(e), !!dryRun);
    throw e;
  }
}

/**
 * Импорт товаров и поставщиков из внешней книги в 04 (для расписания и пакетных прогонов).
 */
function syncRestoreMasterRefsFromExternalImpl_(dryRun, opt) {
  const parts = [];
  parts.push(syncRestoreProductsFromExternal_(dryRun));
  parts.push(syncRestoreSuppliersFromExternal_(dryRun));
  return parts.join('\n');
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
  const lines = syncHubFormatBookRegistryLines_(true);
  ui.alert(
    '✅ Проверка пройдена',
    'Базовые настройки синхронизации заполнены.\n\n' + lines.join('\n'),
    ui.ButtonSet.OK
  );
}

/**
 * Строки для UI: код книги, имя свойства, статус (заполнено / пусто / текущий файл как 01).
 * @param {boolean} shortIds если true — показывать только часть id
 */
function syncHubFormatBookRegistryLines_(shortIds) {
  const rows = syncHubGetBookRegistryMeta_();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const resolved = syncHubResolveSpreadsheetIdForBook_(r.code, { forDisplay: true });
    let idDisp = resolved.id || '(пусто)';
    if (shortIds && resolved.id && resolved.id.length > 18) {
      idDisp = resolved.id.slice(0, 6) + '…' + resolved.id.slice(-6);
    }
    out.push(r.code + ' ' + r.title + ' | ' + r.prop + ' | ' + resolved.note + (resolved.id ? ' | ' + idDisp : ''));
  }
  return out;
}

function syncHubGetBookRegistryMeta_() {
  return [
    { code: '01', title: 'Заказы', prop: SYNC_HUB_CFG.PROPS.ORDERS_SPREADSHEET_ID },
    { code: '02', title: 'Транзит', prop: SYNC_HUB_CFG.PROPS.TRANSIT_SPREADSHEET_ID },
    { code: '03', title: 'Закупки', prop: SYNC_HUB_CFG.PROPS.PROCUREMENT_SPREADSHEET_ID },
    { code: '04', title: 'Справочники/хаб', prop: SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID },
    { code: '05', title: 'Рейсы/себестоимость', prop: SYNC_HUB_CFG.PROPS.COST_SPREADSHEET_ID }
  ];
}

/**
 * Разрешить Spreadsheet ID по коду книги 01|02|03|04|05.
 * Для 01: если ORDERS_SPREADSHEET_ID пуст, подставляется текущая книга только если она не совпадает с 04 (иначе пусто — задайте ORDERS_SPREADSHEET_ID в хабе).
 */
function syncHubResolveSpreadsheetIdForBook_(bookCode, options) {
  const c = String(bookCode || '').trim();
  const forDisplay = !!(options && options.forDisplay);
  const masterId = syncHubGetProp_(SYNC_HUB_CFG.PROPS.MASTER_REF_SPREADSHEET_ID, '');
  const activeId = SpreadsheetApp.getActiveSpreadsheet().getId();

  if (c === '04') {
    if (!masterId) return { id: '', note: forDisplay ? 'НЕ ЗАДАН 04' : 'MISSING' };
    return { id: masterId, note: '04' };
  }
  if (c === '01') {
    const explicit = syncHubGetProp_(SYNC_HUB_CFG.PROPS.ORDERS_SPREADSHEET_ID, '');
    if (explicit) return { id: explicit, note: '01 явный' };
    if (masterId && activeId === masterId) {
      return { id: '', note: forDisplay ? 'Задайте ORDERS_SPREADSHEET_ID (хаб 04 ≠ книга 01)' : 'MISSING_01' };
    }
    return { id: activeId, note: '01 текущий файл' };
  }
  if (c === '02') {
    const id = syncHubGetProp_(SYNC_HUB_CFG.PROPS.TRANSIT_SPREADSHEET_ID, '');
    return { id: id, note: id ? '02' : (forDisplay ? '02 не задан' : 'MISSING') };
  }
  if (c === '03') {
    const id = syncHubGetProp_(SYNC_HUB_CFG.PROPS.PROCUREMENT_SPREADSHEET_ID, '');
    return { id: id, note: id ? '03' : (forDisplay ? '03 не задан' : 'MISSING') };
  }
  if (c === '05') {
    const id = syncHubGetProp_(SYNC_HUB_CFG.PROPS.COST_SPREADSHEET_ID, '');
    return { id: id, note: id ? '05' : (forDisplay ? '05 не задан' : 'MISSING') };
  }
  throw new Error('Неизвестный код книги: ' + bookCode);
}

function syncHubOpenSpreadsheetForBook_(bookCode) {
  const r = syncHubResolveSpreadsheetIdForBook_(bookCode, {});
  if (!r.id) {
    throw new Error('Не задан Spreadsheet ID для книги ' + bookCode + ' (' + r.note + ')');
  }
  return SpreadsheetApp.openById(r.id);
}

function syncHubShowBookRegistry_() {
  const ui = SpreadsheetApp.getUi();
  const lines = syncHubFormatBookRegistryLines_(false);
  ui.alert(
    'Реестр книг (Script Properties)',
    lines.join('\n'),
    ui.ButtonSet.OK
  );
}

function syncAllExternalBooks_() {
  return syncAllExternalBooksImpl_(false);
}

function syncAllExternalBooksDryRun_() {
  return syncAllExternalBooksImpl_(true);
}

/**
 * Боевой прогон: 04→01 справочники, 04→01/02/05 статусы, 05→04 справочники.
 * С подтверждением, чтобы не случайно перезаписать листы.
 */
function syncAllExternalBooksWithConfirm_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Полная синхронизация',
    'Запустить боевой прогон всех потоков?\n\n' +
      '• 04→01: товары и поставщики\n' +
      '• 04→01, 02, 05: справочник статусов (при необходимости создаются листы)\n' +
      '• 05→04: урезанный каталог из рейсов, таможсбор, типы событий\n\n' +
      'Журнал: лист SYNC_LOG в книге 04.',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) {
    return;
  }
  syncAllExternalBooks_();
}

function syncAllExternalBooksImpl_(dryRun, opt) {
  const silent = !!(opt && opt.silent);
  const lock = LockService.getScriptLock();
  const lockTimeoutMs = syncHubGetNumberProp_(
    SYNC_HUB_CFG.PROPS.SYNC_LOCK_TIMEOUT_MS,
    SYNC_HUB_CFG.DEFAULTS.SYNC_LOCK_TIMEOUT_MS
  );
  lock.waitLock(Math.max(1000, lockTimeoutMs));
  try {
    const report = [];
    report.push(syncMasterRefsFrom04_(dryRun, { silent: true }));
    report.push(syncStatusRefFrom04ToBooks_(dryRun));
    report.push(syncCollectRefsTo04From05_(dryRun));
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        dryRun ? '🧪 Dry-run завершен' : '✅ Синхронизация завершена',
        report.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return report.join('\n');
  } finally {
    lock.releaseLock();
  }
}

function syncOperationalSnapshotsDryRun_() {
  syncOperationalSnapshotsImpl_(true);
}

/**
 * Снимок «Сводная» (или ORDERS_SUMMARY_SHEET_NAME) из 01 в 02 и при заданной 03 — в 03.
 * Нужен для sender_stock (02) и блоков in-transit в procurement_planning (03).
 */
function syncOperationalSnapshotsWithConfirm_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Сводная 01→02 и 03',
    'Скопировать лист заказной сводки из книги 01 в книги 02 и 03?\n\n' +
      'Свойство ORDERS_SUMMARY_SHEET_NAME (по умолчанию «Сводная»).\n' +
      'Если PROCUREMENT_SPREADSHEET_ID пуст — шаг 01→03 пропускается.\n\n' +
      'Целевой лист в 02/03 будет полностью перезаписан значениями с 01.',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) {
    return;
  }
  syncOperationalSnapshotsImpl_(false);
}

function syncOperationalOrdersSummaryFrom01To03DryRun_() {
  syncOperationalOrdersSummaryFrom01To03_(true, {});
}

function syncOperationalOrdersSummaryFrom01To03WithConfirm_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Сводная 01→03',
    'Скопировать лист заказной сводки из книги 01 в книгу 03?\n\n' +
      'Проверьте, что в Script Properties книги 04 задан PROCUREMENT_SPREADSHEET_ID.\n' +
      'Свойство ORDERS_SUMMARY_SHEET_NAME задаёт имя листа-источника (по умолчанию «Сводная»).\n\n' +
      'Целевой лист в 03 будет полностью перезаписан значениями с 01.',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) {
    return;
  }
  syncOperationalOrdersSummaryFrom01To03_(false, {});
}

function syncOperationalSnapshotsImpl_(dryRun, opt) {
  const silent = !!(opt && opt.silent);
  const lock = LockService.getScriptLock();
  const lockTimeoutMs = syncHubGetNumberProp_(
    SYNC_HUB_CFG.PROPS.SYNC_LOCK_TIMEOUT_MS,
    SYNC_HUB_CFG.DEFAULTS.SYNC_LOCK_TIMEOUT_MS
  );
  lock.waitLock(Math.max(1000, lockTimeoutMs));
  try {
    const report = [];
    report.push(syncOperationalOrdersSummaryFrom01To02_(dryRun, { silent: true }));
    report.push(syncOperationalOrdersSummaryFrom01To03_(dryRun, { silent: true }));
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        dryRun ? '🧪 Dry-run операционные снимки' : '✅ Операционные снимки готовы',
        report.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return report.join('\n');
  } finally {
    lock.releaseLock();
  }
}

function syncOperationalOrdersSummaryFrom01To02_(dryRun, opt) {
  const silent = !!(opt && opt.silent);
  try {
    const summaryName = syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.ORDERS_SUMMARY_SHEET_NAME,
      SYNC_HUB_CFG.DEFAULTS.ORDERS_SUMMARY_SHEET_NAME
    );
    const sourceSs = syncHubOpenSpreadsheetForBook_('01');
    const targetSs = syncHubOpenSpreadsheetForBook_('02');
    const mappings = [
      {
        source: summaryName,
        sourceAliases: ['Сводная'],
        target: summaryName,
        targetAliases: ['Сводная'],
        required: true
      }
    ];
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
      dryRun: !!dryRun,
      createMissingTarget: true
    });
    const msg = '01→02: ' + stats;
    syncHubLog_('01→02 Сводная', 'OK', msg, !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        dryRun ? '🧪 Dry-run Сводная 01→02' : '✅ Сводная 01→02',
        msg,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return msg;
  } catch (e) {
    syncHubLog_('01→02 Сводная', 'ERROR', e.message || String(e), !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'Ошибка Сводная 01→02',
        e.message || String(e),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    throw e;
  }
}

function syncOperationalOrdersSummaryFrom01To03_(dryRun, opt) {
  const silent = !!(opt && opt.silent);
  try {
    const r3 = syncHubResolveSpreadsheetIdForBook_('03', {});
    if (!r3.id) {
      const msg = '01→03: PROCUREMENT_SPREADSHEET_ID не задан — пропущено';
      syncHubLog_('01→03 Сводная', 'OK', msg, !!dryRun);
      return msg;
    }
    const summaryName = syncHubGetProp_(
      SYNC_HUB_CFG.PROPS.ORDERS_SUMMARY_SHEET_NAME,
      SYNC_HUB_CFG.DEFAULTS.ORDERS_SUMMARY_SHEET_NAME
    );
    const sourceSs = syncHubOpenSpreadsheetForBook_('01');
    const targetSs = SpreadsheetApp.openById(r3.id);
    const mappings = [
      {
        source: summaryName,
        sourceAliases: ['Сводная'],
        target: summaryName,
        targetAliases: ['Сводная'],
        required: true
      }
    ];
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
      dryRun: !!dryRun,
      createMissingTarget: true
    });
    const msg = '01→03: ' + stats;
    syncHubLog_('01→03 Сводная', 'OK', msg, !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        dryRun ? '🧪 Dry-run Сводная 01→03' : '✅ Сводная 01→03',
        msg,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return msg;
  } catch (e) {
    syncHubLog_('01→03 Сводная', 'ERROR', e.message || String(e), !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'Ошибка Сводная 01→03',
        e.message || String(e),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    throw e;
  }
}

/**
 * Справочники из книги 04 → книга 01 (ORDERS_SPREADSHEET_ID).
 * Не использует активную книгу: иначе при запуске из 04 копирование шло бы в 04 и не попадало в 01.
 * @param {boolean} dryRun
 * @param {{ silent?: boolean }} opt если silent — без отдельного Ui.alert (для пакетного dry-run/синка)
 */
function syncMasterRefsFrom04_(dryRun, opt) {
  const silent = !!(opt && opt.silent);
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
        target: 'Справочник поставщики и условия',
        targetAliases: [
          'Справочник поставщики и условия',
          'Справочник поставщики и условия работы'
        ]
      }
    ];
    const sourceSs = SpreadsheetApp.openById(sourceId);
    const targetSs = syncHubOpenSpreadsheetForBook_('01');
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, { dryRun: !!dryRun });
    const msg = 'Справочники (04→01): ' + stats;
    syncHubLog_('04→01 Справочники', 'OK', msg, !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        dryRun ? '🧪 Dry-run справочники из 04 в 01' : '✅ Справочники: 04 → 01',
        msg,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return msg;
  } catch (e) {
    syncHubLog_('04→01 Справочники', 'ERROR', e.message || String(e), !!dryRun);
    if (!silent) {
      SpreadsheetApp.getUi().alert('Ошибка 04→01 справочники', e.message || String(e), SpreadsheetApp.getUi().ButtonSet.OK);
    }
    throw e;
  }
}

/**
 * Раздача справочника статусов из 04 в книги 01/02/05.
 * Книга 03 исключена (по бизнес-правилу статусы там не используются).
 */
function syncStatusRefFrom04ToBooks_(dryRun) {
  const targets = [
    { code: '01', title: '01' },
    { code: '02', title: '02' },
    { code: '05', title: '05' }
  ];
  const sourceSs = syncHubOpenSpreadsheetForBook_('04');
  const lines = [];

  for (let i = 0; i < targets.length; i++) {
    const t = targets[i];
    try {
      const targetSs = syncHubOpenSpreadsheetForBook_(t.code);
      const mappings = [
        {
          source: 'Справочник_статусов',
          sourceAliases: ['Справочник статусов', 'Статусы'],
          target: 'Справочник_статусов',
          targetAliases: ['Справочник_статусов', 'Справочник статусов', 'Статусы'],
          required: false
        }
      ];
      const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
        dryRun: !!dryRun,
        createMissingTarget: true
      });
      lines.push(t.title + ': ' + stats);
    } catch (e) {
      lines.push(t.title + ': ERROR ' + (e.message || String(e)));
    }
  }

  const msg = '04→(01,02,05) Статусы: ' + lines.join(' | ');
  syncHubLog_('04→01/02/05 Статусы', 'OK', msg, !!dryRun);
  return msg;
}

function syncCollectRefsTo04From05_(dryRun) {
  try {
    const sourceSs = syncHubOpenSpreadsheetForBook_('05');
    const targetSs = syncHubOpenSpreadsheetForBook_('04');
    const mappings = [
      {
        source: 'Справочник товаров (05)',
        sourceAliases: ['Справочник товаров (05)', 'Справочник товары', 'Справочник_товары'],
        // В 05 этот справочник урезанный (только позиции в рейсах), не перезаписываем мастер в 04.
        target: 'Справочник товаров (05 урезанный)',
        targetAliases: ['Справочник товаров (05 урезанный)']
      },
      {
        source: 'Справочник поставщики и условия',
        sourceAliases: ['Справочник поставщики и условия работы'],
        target: 'Справочник поставщики и условия',
        targetAliases: ['Справочник поставщики и условия', 'Справочник поставщики и условия работы'],
        required: false
      },
      {
        source: 'Справочник_таможсбор',
        sourceAliases: ['Справочник таможсбор', 'Справочник таможенного сбора'],
        target: 'Справочник_таможсбор',
        targetAliases: ['Справочник_таможсбор', 'Справочник таможсбор', 'Справочник таможенного сбора'],
        required: false
      },
      {
        source: 'Типы_событий',
        sourceAliases: ['Типы событий'],
        target: 'Типы_событий',
        targetAliases: ['Типы_событий', 'Типы событий'],
        required: false
      },
      {
        source: 'Нормативы_доставки',
        sourceAliases: ['Нормативы доставки'],
        target: 'Нормативы_доставки',
        targetAliases: ['Нормативы_доставки', 'Нормативы доставки'],
        required: false
      }
    ];
    const stats = syncHubCopyMappings_(sourceSs, targetSs, mappings, {
      dryRun: !!dryRun,
      createMissingTarget: true
    });
    const msg = '05→04 Справочники: ' + stats;
    syncHubLog_('05→04 Справочники', 'OK', msg, !!dryRun);
    return msg;
  } catch (e) {
    syncHubLog_('05→04 Справочники', 'ERROR', e.message || String(e), !!dryRun);
    throw e;
  }
}

function syncHubCopyMappings_(sourceSs, targetSs, mappings, options) {
  const dryRun = !!(options && options.dryRun);
  const createMissingTarget = !!(options && options.createMissingTarget);
  const batchSize = Math.max(100, syncHubGetNumberProp_(SYNC_HUB_CFG.PROPS.SYNC_BATCH_SIZE, SYNC_HUB_CFG.DEFAULTS.SYNC_BATCH_SIZE));
  const parts = [];

  mappings.forEach(function (m) {
    const sourceSheet = syncHubFindSheetByNames_(sourceSs, [m.source].concat(m.sourceAliases || []));
    if (!sourceSheet) {
      const checked = [m.source].concat(m.sourceAliases || []);
      if (m.required === false) {
        parts.push((m.target || m.source) + ': исходный лист не найден, пропущено (' + checked.join(', ') + ')');
        return;
      }
      throw new Error(
        'Не найден исходный лист. Проверены имена: ' + checked.join(', ')
      );
    }

    let targetSheet = syncHubFindSheetByNames_(targetSs, [m.target].concat(m.targetAliases || []));
    if (!targetSheet) {
      if (createMissingTarget) {
        if (dryRun) {
          parts.push(m.target + ': dry-run, целевой лист отсутствует, был бы создан');
          return;
        }
        targetSheet = targetSs.insertSheet(m.target);
      } else {
        parts.push(m.target + ': целевой лист отсутствует, пропущено');
        return;
      }
    }

    const rows = sourceSheet.getLastRow();
    const cols = sourceSheet.getLastColumn();

    if (!rows || !cols) {
      parts.push(m.target + ': источник пуст (rows=' + rows + ', cols=' + cols + '), пропущено');
      return;
    }

    if (
      sourceSs.getId() === targetSs.getId() &&
      sourceSheet.getSheetId() === targetSheet.getSheetId()
    ) {
      parts.push(
        m.target +
          ': отменено — источник и цель один и тот же лист ' +
          '(проверьте MASTER_REF_SPREADSHEET_ID и книгу получателя)'
      );
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
      // Копируем вычисленные значения (в т.ч. результат IMPORTRANGE), не формулы.
      const values = sourceSheet.getRange(start, 1, count, cols).getValues();
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
  const normalized = syncHubNormalizePropValue_(key, String(raw).trim());
  return normalized;
}

function syncHubGetNumberProp_(key, fallback) {
  const n = Number(syncHubGetProp_(key, fallback));
  return isFinite(n) ? n : fallback;
}

/**
 * Нормализует значения Script Properties.
 * Для *_SPREADSHEET_ID пытается извлечь "чистый" id из URL/строки и убрать мусор.
 */
function syncHubNormalizePropValue_(key, value) {
  const k = String(key || '');
  const v = String(value == null ? '' : value).trim();
  if (!v) return v;
  if (k.indexOf('SPREADSHEET_ID') === -1) return v;
  return syncHubExtractSpreadsheetId_(v);
}

function syncHubExtractSpreadsheetId_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  const mUrl = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (mUrl && mUrl[1]) return mUrl[1];
  const token = s.split(/[?#&\s]/)[0];
  const mToken = token.match(/^([a-zA-Z0-9-_]{25,})/);
  if (mToken && mToken[1]) return mToken[1];
  return s;
}
