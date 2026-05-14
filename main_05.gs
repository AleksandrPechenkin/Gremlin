function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (typeof addCostingMenu_ === 'function') {
    addCostingMenu_(ui);
  }
  if (typeof addPaymentLogisticsMenu_ === 'function') {
    addPaymentLogisticsMenu_(ui);
  }
}
