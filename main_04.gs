/**
 * Точка входа книги 04 (хаб синхронизации).
 * В этой книге держим меню синка, логи и оркестрацию между 01-05.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (typeof addSyncHubMenu_ === 'function') {
    addSyncHubMenu_(ui);
  }
}
