/**
 * Точка входа книги 02 (транзитный склад).
 * В этой книге держим меню и операции sender_stock.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (typeof addSenderStockMenu_ === 'function') {
    addSenderStockMenu_(ui);
  }
}
