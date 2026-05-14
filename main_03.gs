/**
 * Точка входа книги 03 (планирование закупок).
 * В этой книге держим меню «Планирование закупок» и операции procurement_planning.
 * Состав модулей книги 03: procurement_planning.gs, wildberries_stocks.gs, wb_sales.gs,
 * ozon_stocks.gs, ozon_sales.gs, moysklad_api.gs, helpers.gs.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Планирование закупок')
    .addItem('Подтянуть планы продаж на лист «Планирование закупок»', 'refreshProcurementPlanningFromSalesSheets')
    .addItem('Диагностика книги планов (какие колонки реально нашёл скрипт)', 'diagnoseProcurementPlanningSourceSheets')
    .addItem('Обновить лист «Склады МС (остатки)» из МойСклад', 'syncMsStockStoresSheet')
    .addItem('Записать учётный остаток МС на «Планирование закупок»', 'updateProcurementPlanningMsAccountingStock')
    .addItem('Записать остаток WB на «Планирование закупок»', 'updateProcurementPlanningWbStock')
    .addItem('Записать остатки Ozon (FBO/FBS) на «Планирование закупок»', 'updateProcurementPlanningOzonStock')
    .addItem('Рассчитать потребность закупки (остатки + в пути + продажи)', 'computeProcurementPurchasePlan')
    .addItem('Проверить сопоставление остатков (артикул/ШК)', 'checkProcurementPlanningStocksCoverage')
    .addItem('Диагностика «Сводной» (что попало во «в пути»)', 'diagnoseProcurementInboundFromSummary')
    .addToUi();
}
