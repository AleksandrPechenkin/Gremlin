/**
 * Расписание Gremlin: hourly справочники/сводная/снимки, daily оплаты из реестра.
 * Деплой: книга 01 (summary + payments), 04 (refs + snapshots), 05 (logistics payments).
 * Не автоматизирует: «Закуплено», syncOrdersWithMS.
 */
const GREMLIN_SCHEDULE_CFG = {
  ENABLED: 'SCHEDULE_ENABLED',
  REF_MINUTE: 'SCHEDULE_REF_MINUTE',
  SUMMARY_MINUTE: 'SCHEDULE_SUMMARY_MINUTE',
  SNAPSHOT_MINUTE: 'SCHEDULE_SNAPSHOT_MINUTE',
  PAYMENTS_HOUR: 'SCHEDULE_PAYMENTS_HOUR',
  PAYMENTS_MINUTE: 'SCHEDULE_PAYMENTS_MINUTE',
  TICK_MINUTES: 'SCHEDULE_TICK_MINUTES',
  BOOK_CODE: 'GREMLIN_BOOK_CODE',
  DEFAULTS: {
    ENABLED: 'on',
    REF_MINUTE: '0',
    SUMMARY_MINUTE: '5',
    SNAPSHOT_MINUTE: '10',
    PAYMENTS_HOUR: '7',
    PAYMENTS_MINUTE: '0',
    TICK_MINUTES: '5'
  },
  TRIGGER_HANDLERS: {
    TICK_04: 'gremlinScheduledTick04_',
    TICK_01: 'gremlinScheduledTick01_',
    DAILY_PAY_01: 'gremlinScheduledDailyPayments01_',
    DAILY_PAY_05: 'gremlinScheduledDailyPayments05_'
  }
};

function gremlinScheduleProp_(key, defaultValue) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (raw == null || String(raw).trim() === '') return defaultValue;
  return String(raw).trim();
}

function gremlinScheduleIsEnabled_() {
  const v = gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.ENABLED, GREMLIN_SCHEDULE_CFG.DEFAULTS.ENABLED)
    .toLowerCase();
  return !(v === '0' || v === 'false' || v === 'no' || v === 'off');
}

function gremlinScheduleInMinuteWindow_(propKey, defaultMin) {
  const target = parseInt(gremlinScheduleProp_(propKey, String(defaultMin)), 10);
  if (isNaN(target)) return false;
  const m = new Date().getMinutes();
  const width = parseInt(
    gremlinScheduleProp_('SCHEDULE_MINUTE_WINDOW', '5'),
    10
  ) || 5;
  return m >= target && m < target + width;
}

function gremlinScheduleLog_(block, status, details) {
  if (typeof syncHubLog_ === 'function') {
    try {
      syncHubLog_(block, status, details, false);
      return;
    } catch (e) {
      Logger.log('syncHubLog_ failed: ' + (e.message || e));
    }
  }
  try {
    const masterId = PropertiesService.getScriptProperties().getProperty('MASTER_REF_SPREADSHEET_ID');
    if (masterId) {
      const ss = SpreadsheetApp.openById(String(masterId).trim());
      const logName = PropertiesService.getScriptProperties().getProperty('SYNC_LOG_SHEET_NAME') || 'SYNC_LOG';
      let sh = ss.getSheetByName(logName);
      if (!sh) sh = ss.insertSheet(logName);
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Block', 'Status', 'DryRun', 'User', 'Details']]);
        sh.setFrozenRows(1);
      }
      const userEmail = Session.getEffectiveUser().getEmail() || '';
      sh.appendRow([new Date(), block, status, 'NO', userEmail, details]);
      return;
    }
  } catch (e2) {
    Logger.log('gremlinScheduleLog_ master failed: ' + (e2.message || e2));
  }
  Logger.log('[Schedule] ' + block + ' | ' + status + ' | ' + details);
}

function gremlinScheduleRemoveTriggers_() {
  const handlers = GREMLIN_SCHEDULE_CFG.TRIGGER_HANDLERS;
  const names = {};
  let k;
  for (k in handlers) {
    if (handlers.hasOwnProperty(k)) names[handlers[k]] = true;
  }
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(function (t) {
    if (names[t.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert('Триггеры расписания', 'Удалено триггеров: ' + removed, ui.ButtonSet.OK);
  }
}

function gremlinScheduleInstallTickTrigger_(handlerName) {
  const mins = parseInt(
    gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.TICK_MINUTES, GREMLIN_SCHEDULE_CFG.DEFAULTS.TICK_MINUTES),
    10
  ) || 5;
  ScriptApp.newTrigger(handlerName).timeBased().everyMinutes(mins).create();
}

function gremlinScheduleInstallDailyPaymentsTrigger_(handlerName) {
  const hour = parseInt(
    gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_HOUR, GREMLIN_SCHEDULE_CFG.DEFAULTS.PAYMENTS_HOUR),
    10
  );
  const minute = parseInt(
    gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_MINUTE, GREMLIN_SCHEDULE_CFG.DEFAULTS.PAYMENTS_MINUTE),
    10
  );
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .atHour(isNaN(hour) ? 7 : hour)
    .nearMinute(isNaN(minute) ? 0 : minute)
    .everyDays(1)
    .create();
}

/* ===================== Книга 04 ===================== */

function gremlinScheduleAddMenu04_(ui) {
  // Меню встроено в addSyncHubMenu_ (sync_hub.gs).
}

function gremlinScheduleInstallTriggers04_() {
  gremlinScheduleRemoveTriggers_();
  PropertiesService.getScriptProperties().setProperty(GREMLIN_SCHEDULE_CFG.BOOK_CODE, '04');
  gremlinScheduleInstallTickTrigger_(GREMLIN_SCHEDULE_CFG.TRIGGER_HANDLERS.TICK_04);
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Расписание (книга 04)',
    'Установлен триггер каждые ' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.TICK_MINUTES, '5') +
      ' мин.\n\n' +
      'В :' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.REF_MINUTE, '0') +
      '–:… — внешние справочники → 04 → 01\n' +
      'В :' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.SNAPSHOT_MINUTE, '10') +
      '–:… — Сводная 01→02/03\n\n' +
      'Также установите триггеры в книгах 01 и 05 (меню «⏱ Расписание»).',
    ui.ButtonSet.OK
  );
}

function gremlinScheduledTick04_() {
  if (!gremlinScheduleIsEnabled_()) return;
  if (gremlinScheduleInMinuteWindow_(GREMLIN_SCHEDULE_CFG.REF_MINUTE, 0)) {
    gremlinScheduleRunHourlyRefsSilent_();
  }
  if (gremlinScheduleInMinuteWindow_(GREMLIN_SCHEDULE_CFG.SNAPSHOT_MINUTE, 10)) {
    gremlinScheduleRunHourlySnapshotsSilent_();
  }
}

function gremlinScheduleRunHourlyRefsSilent_() {
  try {
    const parts = [];
    if (typeof syncRestoreMasterRefsFromExternalImpl_ === 'function') {
      parts.push(syncRestoreMasterRefsFromExternalImpl_(false, { silent: true }));
    }
    if (typeof syncMasterRefsFrom04_ === 'function') {
      parts.push(syncMasterRefsFrom04_(false, { silent: true }));
    }
    gremlinScheduleLog_('Hourly refs', 'OK', parts.join(' | '));
  } catch (e) {
    gremlinScheduleLog_('Hourly refs', 'ERROR', e.message || String(e));
  }
}

function gremlinScheduleRunHourlySnapshotsSilent_() {
  try {
    const msg =
      typeof syncOperationalSnapshotsImpl_ === 'function'
        ? syncOperationalSnapshotsImpl_(false, { silent: true })
        : 'syncOperationalSnapshotsImpl_ недоступен';
    gremlinScheduleLog_('Hourly snapshots', 'OK', msg);
  } catch (e) {
    gremlinScheduleLog_('Hourly snapshots', 'ERROR', e.message || String(e));
  }
}

function gremlinScheduleRunHourlyRefsNow_() {
  gremlinScheduleRunHourlyRefsSilent_();
  SpreadsheetApp.getUi().alert('Справочники', 'Прогон завершён. См. SYNC_LOG.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function gremlinScheduleRunHourlySnapshotsNow_() {
  gremlinScheduleRunHourlySnapshotsSilent_();
  SpreadsheetApp.getUi().alert('Снимки Сводной', 'Прогон завершён. См. SYNC_LOG.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function gremlinScheduleRunHourlyFull04Now_() {
  gremlinScheduleRunHourlyRefsSilent_();
  Utilities.sleep(2000);
  gremlinScheduleRunHourlySnapshotsSilent_();
  SpreadsheetApp.getUi().alert(
    'Hourly 04',
    'Справочники и снимки выполнены. Сборку Сводной запустите в книге 01.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ===================== Книга 01 ===================== */

function gremlinScheduleAddMenu01_(ui) {
  ui.createMenu('⏱ Расписание')
    .addItem('Установить триггеры этой книги (01)', 'gremlinScheduleInstallTriggers01_')
    .addItem('Снять триггеры этой книги', 'gremlinScheduleRemoveTriggers_')
    .addSeparator()
    .addItem('Сейчас: собрать «Сводная»', 'gremlinScheduleRunSummaryNow_')
    .addItem('Сейчас: синхронизировать оплаты из реестра', 'gremlinScheduleRunPayments01Now_')
    .addToUi();
}

function gremlinScheduleInstallTriggers01_() {
  gremlinScheduleRemoveTriggers_();
  PropertiesService.getScriptProperties().setProperty(GREMLIN_SCHEDULE_CFG.BOOK_CODE, '01');
  gremlinScheduleInstallTickTrigger_(GREMLIN_SCHEDULE_CFG.TRIGGER_HANDLERS.TICK_01);
  gremlinScheduleInstallDailyPaymentsTrigger_(GREMLIN_SCHEDULE_CFG.TRIGGER_HANDLERS.DAILY_PAY_01);
  SpreadsheetApp.getUi().alert(
    'Расписание (книга 01)',
    'Триггеры установлены:\n' +
      '• каждые ' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.TICK_MINUTES, '5') +
      ' мин — сбор «Сводной» (окно с :' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.SUMMARY_MINUTE, '5') +
      ')\n' +
      '• ежедневно ' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_HOUR, '7') +
      ':' +
      (gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_MINUTE, '0').length === 1
        ? '0' + gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_MINUTE, '0')
        : gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_MINUTE, '0')) +
      ' — оплаты из реестра',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function gremlinScheduledTick01_() {
  if (!gremlinScheduleIsEnabled_()) return;
  if (!gremlinScheduleInMinuteWindow_(GREMLIN_SCHEDULE_CFG.SUMMARY_MINUTE, 5)) return;
  gremlinScheduleRunSummarySilent_();
}

function gremlinScheduleRunSummarySilent_() {
  try {
    const msg = syncManagerTabsToSummaryImpl_(SpreadsheetApp.getActiveSpreadsheet(), { silent: true });
    gremlinScheduleLog_('Hourly summary', 'OK', String(msg || '').slice(0, 500));
  } catch (e) {
    gremlinScheduleLog_('Hourly summary', 'ERROR', e.message || String(e));
  }
}

function gremlinScheduleRunSummaryNow_() {
  syncManagerTabsToSummary();
}

function gremlinScheduledDailyPayments01_() {
  if (!gremlinScheduleIsEnabled_()) return;
  gremlinScheduleRunPayments01Silent_();
}

function gremlinScheduleRunPayments01Silent_() {
  try {
    const r =
      typeof paySyncPaidStatusesImpl_ === 'function'
        ? paySyncPaidStatusesImpl_({ silent: true })
        : { message: 'paySyncPaidStatusesImpl_ недоступен' };
    gremlinScheduleLog_('Daily payments 01', 'OK', r.message || String(r.updated));
  } catch (e) {
    gremlinScheduleLog_('Daily payments 01', 'ERROR', e.message || String(e));
  }
}

function gremlinScheduleRunPayments01Now_() {
  paySyncPaidStatuses();
}

/* ===================== Книга 05 ===================== */

function gremlinScheduleAddMenu05_(ui) {
  ui.createMenu('⏱ Расписание')
    .addItem('Установить триггеры этой книги (05)', 'gremlinScheduleInstallTriggers05_')
    .addItem('Снять триггеры этой книги', 'gremlinScheduleRemoveTriggers_')
    .addSeparator()
    .addItem('Сейчас: синхронизировать оплаты логистики', 'gremlinScheduleRunPayments05Now_')
    .addToUi();
}

function gremlinScheduleInstallTriggers05_() {
  gremlinScheduleRemoveTriggers_();
  PropertiesService.getScriptProperties().setProperty(GREMLIN_SCHEDULE_CFG.BOOK_CODE, '05');
  gremlinScheduleInstallDailyPaymentsTrigger_(GREMLIN_SCHEDULE_CFG.TRIGGER_HANDLERS.DAILY_PAY_05);
  SpreadsheetApp.getUi().alert(
    'Расписание (книга 05)',
    'Установлен ежедневный триггер оплат логистики в ' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_HOUR, '7') +
      ':' +
      gremlinScheduleProp_(GREMLIN_SCHEDULE_CFG.PAYMENTS_MINUTE, '0'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function gremlinScheduledDailyPayments05_() {
  if (!gremlinScheduleIsEnabled_()) return;
  gremlinScheduleRunPayments05Silent_();
}

function gremlinScheduleRunPayments05Silent_() {
  try {
    const r =
      typeof payLogSyncPaidStatusesImpl_ === 'function'
        ? payLogSyncPaidStatusesImpl_({ silent: true })
        : { message: 'payLogSyncPaidStatusesImpl_ недоступен' };
    gremlinScheduleLog_('Daily payments 05', 'OK', r.message || String(r.updated));
  } catch (e) {
    gremlinScheduleLog_('Daily payments 05', 'ERROR', e.message || String(e));
  }
}

function gremlinScheduleRunPayments05Now_() {
  payLogSyncPaidStatuses();
}
