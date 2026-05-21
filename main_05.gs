function onOpen() {
  try {
    gremlinBuildMenus05_();
  } catch (e) {
    try {
      SpreadsheetApp.getUi().alert(
        'Ошибка построения меню',
        (e && e.message) ? e.message : String(e),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (ignore) {}
  }
}

/** Запуск из редактора скриптов: обновить меню без перезагрузки книги */
function gremlinRefreshMenus05() {
  gremlinBuildMenus05_();
  SpreadsheetApp.getUi().alert(
    'Меню обновлено',
    'Если пунктов не видно — обновите страницу таблицы (F5).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function gremlinBuildMenus05_() {
  const ui = SpreadsheetApp.getUi();
  if (typeof addCostingMenu_ === 'function') {
    addCostingMenu_(ui);
  }
  if (typeof addCostingCustomsMenu_ === 'function') {
    addCostingCustomsMenu_(ui);
  }
  if (typeof addCostingReceiptMenu_ === 'function') {
    addCostingReceiptMenu_(ui);
  }
  if (typeof addPaymentLogisticsMenu_ === 'function') {
    addPaymentLogisticsMenu_(ui);
  }
  if (typeof addLogisticsMenu_ === 'function') {
    addLogisticsMenu_(ui);
  }
  if (typeof gremlinScheduleAddMenu05_ === 'function') {
    gremlinScheduleAddMenu05_(ui);
  }
}
