Структура проекта:
- sync_hub.gs: синхронизация между книгами 01–05; меню «Синхронизация книг»; лист SYNC_LOG в книге 04 (MASTER_REF_SPREADSHEET_ID). Реестр: ORDERS_SPREADSHEET_ID (01), TRANSIT_SPREADSHEET_ID (02), PROCUREMENT_SPREADSHEET_ID (03), COST_SPREADSHEET_ID (05); если скрипт запускается из книги 01 и ORDERS_SPREADSHEET_ID пуст — для 01 подставляется текущий файл (но не когда активна книга 04: тогда задайте ORDERS_SPREADSHEET_ID явно)
- main.gs: основная синхронизация заказов с МойСклад
- payments_sync.gs: перенос оплат в таблицу "Закуплено"
- moysklad_api.gs: авторизация, HTTP-методы, обработка ошибок API
- helpers.gs: общие функции (даты, числа, строки, атрибуты)
- procurement_planning.gs: свод планов продаж («Заказали, шт») по «Артикул ВБ» на лист «Планирование закупок»; ШК и наименование (артикул поставщика) из справочника товаров; месяцы — округление вверх до коробки (шт/кор, по умолчанию колонка T); лист «Склады МС (остатки)» (syncMsStockStoresSheet); учётный остаток МС на планировании (updateProcurementPlanningMsAccountingStock — в расчёт только артикулы ВБ и ШК из строк плана, поле stock, склады с «Использовать»); остаток WB (updateProcurementPlanningWbStock — см. wildberries_stocks.gs); проверка сопоставления остатков по артикулу/ШК (checkProcurementPlanningStocksCoverage)
- procurement_planning.gs: свод планов продаж («Заказали, шт») по «Артикул ВБ» на лист «Планирование закупок»; ШК и наименование (артикул поставщика) из справочника товаров; месяцы — округление вверх до коробки (шт/кор, по умолчанию колонка T); лист «Склады МС (остатки)» (syncMsStockStoresSheet); учётный остаток МС на планировании (updateProcurementPlanningMsAccountingStock — в расчёт только артикулы ВБ и ШК из строк плана, поле stock, склады с «Использовать»); остаток WB (updateProcurementPlanningWbStock — см. wildberries_stocks.gs); проверка сопоставления остатков по артикулу/ШК (checkProcurementPlanningStocksCoverage); расчёт потребности закупки с выгрузкой детализированной таблицы (computeProcurementPurchasePlan)
- wildberries_stocks.gs: остатки Wildberries для планирования (отчёт warehouse_remains на seller-analytics-api или supplier/stocks на statistics-api; токен WB_API_TOKEN; свойства WB_STOCK_SOURCE, базовые URL — в комментарии в начале файла)
- wb_sales.gs: факт продаж WB текущего месяца (statistics API /api/v1/supplier/sales) для корректировки плана месяца в расчёте закупки

Обязательные Script Properties:
- MS_TOKEN
- MS_ORGANIZATION_ID
- EXTERNAL_SPREADSHEET_ID

Для остатков WB на листе планирования (по желанию):
- WB_API_TOKEN — токен продавца WB («Аналитика» для warehouse_remains, «Статистика» для supplier/stocks; в режиме auto сначала analytics)

Опциональные Script Properties:
- ORDERS_SPREADSHEET_ID — ID книги 01 (заказы); для хаба в 04 обязательно; в книге 01 можно не задавать, если совпадает с текущим файлом
- MASTER_REF_SPREADSHEET_ID, TRANSIT_SPREADSHEET_ID, PROCUREMENT_SPREADSHEET_ID, COST_SPREADSHEET_ID — см. sync_hub.gs (синхронизация между книгами)
- SYNC_HUB_MENU_ENABLED — показ меню «Синхронизация книг» в книге 01 (по умолчанию включено; поставьте 0/false/no/off, чтобы скрыть меню после переноса хаба в 04)
- MS_STORE_ID
- EXTERNAL_SHEET_NAME (по умолчанию: Закуплено)
- SALES_PLANS_SPREADSHEET_ID (книга с «Проставление планов …»; в коде есть ID по умолчанию)
- SALES_PLANS_SHEET_GREMLIN, SALES_PLANS_SHEET_OBSHCHIY, SALES_PLANS_SHEET_DEPT3 — не трогать, пока листы в книге отделов не переименовывают (на имена завязаны другие скрипты); свойства — только при осознанном переименовании везде
- SALES_PLANS_EXTRA_SPREADSHEET_ID — вторая книга с планом (в коде ID по умолчанию); пусто или OFF — не подмешивать
- SALES_PLANS_EXTRA_SHEET_NAME — лист во второй книге (по умолчанию: Проставление планов)
- PROCUREMENT_PLANNING_SHEET (куда писать свод; по умолчанию: Планирование закупок)
- PROCUREMENT_PLANNING_YEAR, PROCUREMENT_PLANNING_MONTHS (например 2026 и «5,6,7,8» для мая–августа)
- PRODUCT_REFERENCE_SPREADSHEET_ID, PRODUCT_REFERENCE_SHEET_NAME (как в инвойсах; в коде есть значения по умолчанию)
- PRODUCT_REFERENCE_PCS_PER_BOX_COL (номер колонки «шт в коробке» в справочнике; по умолчанию 20 = T)
- MS_STOCK_STORES_SHEET (имя листа перечня складов для остатков; по умолчанию: Склады МС (остатки))
- MS_STORE_SYNC_INCLUDE_ARCHIVED (1 — подтягивать и архивные склады МС; по умолчанию только неархивные)
- WB_STOCK_SOURCE (analytics | statistics | auto — по умолчанию auto)
- WB_ANALYTICS_API_BASE, WB_STATISTICS_API_BASE (редко нужны; см. wildberries_stocks.gs)
- WB_STOCK_USE_QUANTITY_FULL (1 — в режиме statistics суммировать quantityFull вместо quantity)
- PROCUREMENT_SHIPPED_STATUS_CODE (по умолчанию S13_IN_TRANSIT_MOSCOW — статус «отгружен» для гибридной логики in-transit)
- PROCUREMENT_PURCHASE_REPORT_SHEET (лист детализации расчёта закупки; по умолчанию: Планирование закупок (расчёт))
- PROCUREMENT_TARGET_CONTROL_MONTH (контрольный месяц для двух срезов, формат YYYY-MM; по умолчанию последний месяц горизонта)

---
Синхронизация с облаком

1) OneDrive (уже)
Проект лежит в папке OneDrive (Документы). При включённом клиенте OneDrive и
зелёной галочке у файлов копии уходят в облако Microsoft автоматически.
Перед выключением ПК: сохраните файлы в Cursor (Ctrl+S) и дождитесь синхронизации
(иконка OneDrive в трее — без «в ожидании»).
Папка в браузере (OneDrive в облаке):
https://onedrive.live.com/?id=%2Fpersonal%2Fd0b2710d08b281b9%2FDocuments%2FGremlin%2F%D0%90%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F%20Gremlin&viewid=37fd5091%2D1e00%2D48a9%2Dadea%2D0ae2cd53bd3d&view=0

2) Git + GitHub (рекомендуется для истории версий и копии вне OneDrive)
Удалённый репозиторий: https://github.com/AleksandrPechenkin/Gremlin
Ветка по умолчанию: main. Локальная история: git log --oneline

После правок в папке проекта:
  git add -A
  git commit -m "кратко что сделано"
  git push
В репозиторий не кладите токены МойСклад — только код из .gs и PROJECT_CONTEXT.md.
Файл .gitignore уже в проекте (в т.ч. .clasp.json не попадёт в коммит).
