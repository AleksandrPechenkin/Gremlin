Структура проекта:
- sync_hub.gs: синхронизация между книгами 01–05; меню «Синхронизация книг»; лист SYNC_LOG в книге 04 (MASTER_REF_SPREADSHEET_ID). Реестр: ORDERS_SPREADSHEET_ID (01), TRANSIT_SPREADSHEET_ID (02), PROCUREMENT_SPREADSHEET_ID (03), COST_SPREADSHEET_ID (05); если скрипт запускается из книги 01 и ORDERS_SPREADSHEET_ID пуст — для 01 подставляется текущий файл (но не когда активна книга 04: тогда задайте ORDERS_SPREADSHEET_ID явно). Импорт справочников из внешней книги: MASTER_REF_EXTERNAL_SPREADSHEET_ID / MASTER_REF_EXTERNAL_PRODUCTS_SHEET / MASTER_REF_EXTERNAL_SUPPLIERS_SHEET (поставщики: OFF — пропуск)
- gremlin_scheduled.gs: расписание (триггеры). Каждый час (окна минут, по умолчанию :00 справочники в 04, :05 сбор «Сводной» в 01, :10 снимок 01→02/03 в 04). Ежедневно 07:00 — синхронизация «Оплачено» из реестра в 01 и 05. Меню «⏱ Расписание» в 01/04/05; в 04 также подменю в «Синхронизация книг». Не автоматизирует: «Закуплено» (updateExternalPurchases), МойСклад (syncOrdersWithMS). Свойства: SCHEDULE_ENABLED, SCHEDULE_REF_MINUTE, SCHEDULE_SUMMARY_MINUTE, SCHEDULE_SNAPSHOT_MINUTE, SCHEDULE_PAYMENTS_HOUR, SCHEDULE_PAYMENTS_MINUTE, SCHEDULE_TICK_MINUTES (по умолчанию 5)
- main.gs: основная синхронизация заказов с МойСклад; сборщик «Сводной» (syncManagerTabsToSummary) — подмешивает в Сводную манерные вкладки «Имя ММ/ГГ» и листы с префиксом «История отгрузок» (лист с уже отгруженными машинами; авто-создаётся при первом запуске сборщика). В «Сводной» есть колонка «Рейс» (SHIPMENT_ID, формат TR-YYYY-NNNN) — связь со строками «Партии_в_рейсе» книги 05; для строк «История отгрузок» месяц определяется из «Период (MM/YY)» → ETA → «Дата готовности». В строку 1 «Сводной» при сборке записывается предупреждение «Лист собирается автоматически» — ручной ввод только в «История отгрузок» или в манерные вкладки. Канонизация заголовков (`syncManagerCanonHeader_`) приводит «№», «/», «\», точки, запятые, скобки, дефисы и подчёркивания к пробелам — за счёт чего «Кол-во в коробке», «Кол_во в коробке», «Номер_спецификации» сматчатся с однотипными ключами. В `syncManagerHeaderKey_` явные ключи для `qty_per_box` (количество в коробке), `boxes` (коробки/коробов/коробок), `spec_number`, `shipment_id` и др. — алиасы покрывают типичные написания у менеджеров («Итого коробок», «Количество коробов», «№ спецификации» и т.п.). Меню «📦 МойСклад → 🔍 Диагностика заголовков активного таба» — показывает по каждой колонке имя→канон→ключ и помечает важные ключи, которые не нашлись (быстрая ловля расхождений названий).
- main_02.gs: точка входа книги 02 (транзитный склад). Содержит onOpen, подключает меню sender_stock.
- payments_sync.gs: перенос оплат в таблицу "Закуплено"
- moysklad_api.gs: авторизация, HTTP-методы, обработка ошибок API
- helpers.gs: общие функции (даты, числа, строки, атрибуты)
- procurement_planning.gs: свод планов продаж («Заказали, шт») по «Артикул ВБ» на лист «Планирование закупок»; ШК и наименование (артикул поставщика) из справочника товаров; месяцы — округление вверх до коробки (шт/кор, по умолчанию колонка T); лист «Склады МС (остатки)» (syncMsStockStoresSheet); учётный остаток МС на планировании (updateProcurementPlanningMsAccountingStock — в расчёт только артикулы ВБ и ШК из строк плана, поле stock, склады с «Использовать»); остаток WB (updateProcurementPlanningWbStock — см. wildberries_stocks.gs); проверка сопоставления остатков по артикулу/ШК (checkProcurementPlanningStocksCoverage)
- procurement_planning.gs: свод планов продаж («Заказали, шт») по «Артикул ВБ» на лист «Планирование закупок»; ШК и наименование (артикул поставщика) из справочника товаров; месяцы — округление вверх до коробки (шт/кор, по умолчанию колонка T); лист «Склады МС (остатки)» (syncMsStockStoresSheet); учётный остаток МС на планировании (updateProcurementPlanningMsAccountingStock — в расчёт только артикулы ВБ и ШК из строк плана, поле stock, склады с «Использовать»); остаток WB (updateProcurementPlanningWbStock — см. wildberries_stocks.gs); проверка сопоставления остатков по артикулу/ШК (checkProcurementPlanningStocksCoverage); расчёт потребности закупки с выгрузкой детализированной таблицы (computeProcurementPurchasePlan)
- wildberries_stocks.gs: остатки Wildberries для планирования (отчёт warehouse_remains на seller-analytics-api или supplier/stocks на statistics-api; токен WB_API_TOKEN; свойства WB_STOCK_SOURCE, базовые URL — в комментарии в начале файла)
- wb_sales.gs: факт продаж WB текущего месяца для корректировки плана месяца в расчёте закупки. По умолчанию — seller-analytics-api /api/analytics/v3/sales-funnel/products (мягкий лимит 3 запроса/мин, burst 3; метрика выбирается через WB_SALES_ANALYTICS_METRIC, по умолчанию orderCount; barcode из этого эндпоинта не возвращается, матчинг по nmId/vendorCode). Старый GET /api/v1/supplier/sales оставлен как ручной fallback через WB_SALES_SOURCE=statistics или WB_SALES_SOURCE=auto. Кеш в DocumentCache (WB_SALES_CACHE_MIN), пропуск целиком — WB_SALES_SKIP.
- ozon_stocks.gs: остатки Ozon для книги 03 (Seller API POST /v4/product/info/stocks + /v3/product/info/list для barcodes; Ozon в 2026 перевёл stocks с v3 на v4 — v3 теперь отдаёт 404, схема запроса/ответа осталась прежней); сопоставление с планом: offer_id ↔ canon(«Артикул поставщика»), barcode ↔ canon(«ШК»). Меню «Записать остатки Ozon…» создаёт/обновляет колонки «Остаток Ozon FBO, шт» и «Остаток Ozon FBS, шт» на «Планирование закупок».
- ozon_sales.gs: факт продаж Ozon текущего месяца (FBO + FBS) — POST /v2/posting/fbo/list (для FBO у Ozon актуальна v2; v3 для FBO нет) и /v3/posting/fbs/list, отмены (status=cancelled) не учитываются. Матчинг с планом — по offer_id ↔ «Артикул поставщика». Подмешивается в computeProcurementPurchasePlan: текущий месяц = max(0, план − WB − Ozon).
- Диагностика книги 03:
  - «Диагностика книги планов» (diagnoseProcurementPlanningSourceSheets) — лист «Планирование закупок (диагностика)»: какие колонки реально нашлись в книгах планов (anchor / «Артикул ВБ» или «Артикул поставщика» / «Заказали, шт»), сколько строк под шапкой, сколько с qty>0 (для каждой строки — пример «артикул | raw=… | parsed=N»). Не зависит от helpers.gs (свой парсер ppParseQtyForPlan_).
  - «Диагностика «Сводной»» (diagnoseProcurementInboundFromSummary) — лист «Планирование закупок (диагностика Сводной)»: что попало в колонку «В пути до конца горизонта, шт». Показывает найденную шапку «Сводной» (сканит первые 10 строк), какие колонки распознаны (Артикул ВБ, ШК, Итоговое количество, Статус, ETA, Период, Дата готовности), причины отбраковки строк (qty<=0 / нет даты / месяц вне горизонта / не найдено в плане) с примерами, раскладку учтённых по месяцам и покрытие плана. Предусловие: лист «Сводная» в книге 03 уже наполнен — его создаёт «Операционные потоки → Сводная 01→03 (боевой)» из книги 04.
- payment_requests_05.gs: книга 05; заявки на оплату услуг логистики по аналогии с payment_requests.gs (книга 01). Меню «💳 Оплаты логистики»: «Авторизация / 1) Подать заявку (по выделенным строкам «Затраты рейса») / 2) Проверка-отправка в реестр / 3) Синхронизировать оплаченные / 🛠 Добавить недостающие колонки оплат». Источник — выделенные строки на листе «Затраты рейса»; в одной заявке могут участвовать несколько строк, но у всех должен совпадать Контрагент и Валюта, а у каждой быть непустые SHIPMENT_ID и Сумма (>0). Очередь — лист `Payment_Link_Map_Logistics` в самой книге 05 (17 колонок: queue_id, created_at, created_by, expense_rows_json, shipment_ids, articles, counterparty, amount, currency, due_date, payment_purpose, folder_url, file_links_json, status, registry_request_no, registry_row, reject_reason). При первом запуске «Подать заявку» автоматически добавляет в шапку «Затраты рейса» 4 колонки оплат, если их нет: «№ заявки», «Статус оплаты» (На проверке / На согласовании / Отклонено / Оплачено), «Дата оплаты», «Ссылка на папку». Реестр — тот же, что у книги 01 (PAYMENT_REGISTRY_SPREADSHEET_ID / PAYMENT_REGISTRY_SHEET_NAME, по умолчанию `1BUFF4-…` / «Реестр»); инициатор «Печёнкин А.А.», отдел «Снабжение», префикс заявки общий «ЗАЯВКА-YYYY-NNNN». Связующее звено — SHIPMENT_ID: попадает в назначение (колонка F) и в комментарий (колонка M Реестра, шаблон «Передано из логистической заявки <queue_id>. Рейс: <TR-…>»). Согласующий — общий с книгой 01 (PAYMENT_APPROVER_EMAIL). Файлы грузятся чанками (CHUNK=450000) — те же три точки входа `payLogDialogPrepare/UploadChunk/Finalize`. Дублирующиеся заявки на проверке (та же сигнатура counterparty|currency|shipment_ids|expense_rows|purpose) при одобрении одной автоматически помечаются «Отклонено: дубликат». Возврат факта оплаты — `payLogSyncPaidStatuses`: по «№ заявки» сматчивает строку очереди с Реестром, и если в Реестре «Оплачено» + есть дата факта, проставляет на тех же строках «Затраты рейса» «Статус оплаты = Оплачено» и «Дата оплаты». Связка товар↔машина↔расход↔оплата замыкается: «Сводная»/«Партии_в_рейсе» дают товар по SHIPMENT_ID, «Затраты рейса» (с заявкой/датой оплаты) — расходы по тому же SHIPMENT_ID, «Себестоимость SKU» уже распределяет эти расходы на SKU.
- logistics.gs: книги 05 (меню «🚚 Логистика») и 01 (наследование «Тип доставки» + отгрузка в рейс); учёт логистики и этапов рейса. Меню «🚚 Логистика»: создание листов/шаблонов, запись событий, обновление статусов и ETA, проверка цепочки, запись ETA в «Сводную» (01, только пустые ячейки). Режим доставки — на листе «Рейсы» (`Режим_доставки`: Авто / Море / Море + ЖД / ЖД / Сборный груз); строки «Сводной» наследуют «Тип доставки» при сборке по `SHIPMENT_ID`. Справочники: «Типы_событий» (код, порядок, режим, обязательность, влияние на ETA), «Нормативы_доставки» (дни между кодами этапов), журнал «События_рейса». Свойства: LOGISTICS_ETA_ANCHOR_CODE (по умолчанию READY), LOGISTICS_ETA_TARGET_CODE (по умолчанию ARR_WAREHOUSE_MOSCOW), COST_SPREADSHEET_ID. Синк 05→04: нормативы и типы событий (sync_hub). Книга 01: меню «🚛 Отгрузка → Списать выделенные строки в рейс…»; ключ заголовка `delivery_mode` отделён от `ship_via` («Отгрузка через»). Книга 03: PROCUREMENT_USE_TRIP_ETA=1 — для «в пути» учитывать только «Плановая дата поступления» у отгруженных, без фоллбэка на период/дату готовности.
- costing.gs + costing_customs_xml.gs: книга 05; себестоимость по SHIPMENT_ID. Плановый пересчёт — лист «Себестоимость SKU». Фактический контур: «Загрузить декларацию (XML)», «Сопоставить строки декларации с SKU», «Пересчитать фактическую себестоимость» → «Факт_себестоимость SKU» (колонки Пошлина_план_RUB / НДС_план_RUB для diff); «Сверка план vs факт» и сводка в dry-run; листы «Декларации_строки», «Декларации_журнал»; FACT на «Таможенные платежи» (Сценарий=FACT). XML-парсер: пресет fts_esadout_cu (ED_Container / ESADout_CUGoods, коды 1010/2010/5010); протестировано на ESAD0001–0004. Пересборка «Партии_в_рейсе» из «Сводной» книги 01: меню «💰 Себестоимость → Партии_в_рейсе ← Сводная (01)». Три режима — safe (только добавление новых строк, существующие не трогает), refresh_qty (новые + обновление только Количество/Объем у существующих), full_rebuild (полная пересборка с двойным подтверждением). Перед любым боевым запуском автоматически создаётся снимок «Партии_в_рейсе_бэкап_YYYYMMDD_hhmm» в той же книге; лог в книгу 04 через syncHubLog_. Также есть пункт «Создать снимок сейчас» для ручной защиты. Локальный fallback `costingOpenBook01_()` открывает книгу 01 без `sync_hub.gs` в проекте 05 — читает Script Property ORDERS_SPREADSHEET_ID (URL/id), чистит её через `costingExtractSpreadsheetId_`. В dry-run печатается превью первых 5 добавляемых строк со всеми ключевыми полями (ШК, спец, qty, в коробке, коробки, цена, сумма, валюта) — сразу видно, что реально считалось из Сводной. Канонизация заголовков `costingHeaderCanonForLookup_` согласована с `syncManagerCanonHeader_` (main.gs). Валюты: «Партии_в_рейсе» держит суммы в валюте партии (CNY/USD/…), перевод в RUB — задача листа «Себестоимость SKU»; валюта по умолчанию для новых строк — CNY (Script Property BATCHES_DEFAULT_CURRENCY). Курс ЦБ при пересчёте: один запрос на валюту на весь прогон (fxCache, cbr.ru XML), все строки одной валюты получают одинаковый курс; ручной «Курс_к_RUB» в строке пересчётом игнорируется. Утилита «🛠 Подставить курсы ЦБ в пустые Курс_к_RUB» заполняет это поле в листе только для аудита (на расчёт не влияет). Аллокация расходов из «Затраты_рейса» уважает колонку `ALLOC_BASE` (VOLUME/WEIGHT/VALUE/QTY); если для рейса/поставщика выбран ALLOC_BASE, у каждого SKU должна быть ненулевая величина по этой базе — иначе пересчёт падает с детальным списком.

Обязательные Script Properties:
- MS_TOKEN
- MS_ORGANIZATION_ID
- EXTERNAL_SPREADSHEET_ID

Для остатков WB на листе планирования (по желанию):
- WB_API_TOKEN — токен продавца WB («Аналитика» для warehouse_remains, «Статистика» для supplier/stocks; в режиме auto сначала analytics)

Для остатков и продаж Ozon в книге 03 (по желанию):
- OZON_API_TOKEN — строка формата «<Client-Id>:<Api-Key>» из ЛК Ozon (Настройки → Seller API). Если свойство пустое — остатки/продажи Ozon в расчёт не подмешиваются, остальная логика работает как раньше.

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
- WB_SALES_SOURCE (книга 03; analytics | statistics | auto — по умолчанию analytics. «analytics» — POST на seller-analytics-api /api/analytics/v3/sales-funnel/products (мягкий лимит 3 запроса/мин, burst 3); «statistics» — старый GET /api/v1/supplier/sales (после серии 429 банит токен надолго, оставлен только как fallback); «auto» — сначала analytics, при ошибке — statistics. Для analytics WB_API_TOKEN должен иметь права на категорию «Аналитика».)
- WB_SALES_ANALYTICS_API_BASE (книга 03; опционально, по умолчанию https://seller-analytics-api.wildberries.ru — менять только при тестовом прокси)
- WB_SALES_ANALYTICS_METRIC (книга 03; orderCount | buyoutCount | netOrderCount — по умолчанию orderCount; согласован с колонкой плана «Заказали, шт». netOrderCount = max(0, orderCount - cancelCount), «чистые заказы». Применяется только к analytics; для statistics берётся фактическая quantity по продажам.)
- WB_SALES_RETRY_MS (книга 03; базовая пауза перед повтором WB sales при 429/too many requests; по умолчанию 65000)
- WB_SALES_RETRY_ATTEMPTS (книга 03; число повторов при 429; по умолчанию 2; паузы растут линейно: retryMs, 2*retryMs)
- WB_SALES_TIME_BUDGET_MS (книга 03; максимум времени на сбор WB sales; по умолчанию 240000 = 4 минуты — оставляет место для остальной части расчёта в 6-минутном лимите Apps Script. Если ближайший retry не вписывается в бюджет, скрипт его пропустит и вернёт текущий результат вместо падения по таймауту)
- WB_SALES_SKIP (книга 03; «1»/«true»/«yes» — временно пропустить WB sales целиком, не делать запросов к WB; полезно, когда токен в долгом 429-окне)
- WB_SALES_CACHE_MIN (книга 03; сколько минут держать удачную выборку WB sales в кеше документа; по умолчанию 30 — повторные «Рассчитать потребность закупки» в течение этого окна не бьют WB. Ключ кеша включает source и metric, поэтому переключение WB_SALES_SOURCE/WB_SALES_ANALYTICS_METRIC не путает выборки.)
- PROCUREMENT_SHIPPED_STATUS_CODE (по умолчанию S13_IN_TRANSIT_MOSCOW — статус «отгружен» для гибридной логики in-transit)
- PROCUREMENT_PURCHASE_REPORT_SHEET (лист детализации расчёта закупки; по умолчанию: Планирование закупок (расчёт))
- PROCUREMENT_TARGET_CONTROL_MONTH (контрольный месяц для двух срезов, формат YYYY-MM; по умолчанию последний месяц горизонта)
- BATCHES_DEFAULT_CURRENCY (книга 05; валюта по умолчанию для новых строк «Партии_в_рейсе»; по умолчанию CNY; per-row override — колонкой «Валюта» в Сводной)
- PAYMENT_REGISTRY_SPREADSHEET_ID (общий для книг 01 и 05; id таблицы реестра платежей; по умолчанию `1BUFF4-tlLg-H8faxgKUcnlcaL8aJS84x9bZGza9_K_M`)
- PAYMENT_REGISTRY_SHEET_NAME (общий; имя листа реестра; по умолчанию «Реестр»)
- PAYMENT_REGISTRY_ROOT_FOLDER_ID (общий; id корневой папки на Drive, где создаются папки заявок; по умолчанию `1wMYUKNsNixmNcI1HdsmScn4V-G5Vy6wL`)
- PAYMENT_APPROVER_EMAIL (общий; email согласующего; по умолчанию banych83@gmail.com)
- PAYMENT_DRAFT_FOLDER_NAME_LOG (книга 05; имя подпапки черновиков логистических заявок; по умолчанию `_DRAFT_PAYMENT_REQUESTS_LOG`)
- PAYMENT_INITIATOR_NAME (книга 05; ФИО инициатора в строке Реестра; по умолчанию «Печёнкин А.А.»)
- PAYMENT_LOG_DEPARTMENT (книга 05; отдел в строке Реестра; по умолчанию «Снабжение»)
- OZON_API_BASE (книга 03; по умолчанию https://api-seller.ozon.ru — менять только при тестовом прокси)
- OZON_FETCH_BARCODES (книга 03; «0» отключает дозагрузку barcodes Ozon через /v3/product/info/list — тогда матчинг строк плана с Ozon идёт только по offer_id↔«Артикул поставщика»; по умолчанию включено)
- OZON_STOCKS_PAGE_LIMIT (книга 03; размер страницы /v4/product/info/stocks; по умолчанию 1000)
- OZON_INFO_LIST_BATCH (книга 03; размер пачки offer_id для /v3/product/info/list; по умолчанию 500)
- OZON_SALES_PAGE_LIMIT, OZON_SALES_MAX_PAGES (книга 03; параметры пагинации /v2/posting/fbo/list и /v3/posting/fbs/list; лимит страницы — максимум 1000)

Книга 03 (планирование закупок) — состав файлов для деплоя в её GAS-проект:
- main_03.gs, procurement_planning.gs, wildberries_stocks.gs, wb_sales.gs,
- ozon_stocks.gs, ozon_sales.gs, moysklad_api.gs, helpers.gs.
В проекте 03 не должно быть main.gs (иначе будет второй onOpen).
В книге 04 для синка «Сводная 01→03» должно быть заполнено PROCUREMENT_SPREADSHEET_ID = id книги 03.

Книга 01 (заказы) — для расписания добавьте gremlin_scheduled.gs; после заливки: меню «⏱ Расписание → Установить триггеры этой книги (01)». В свойствах книги 01 задайте MASTER_REF_SPREADSHEET_ID (id книги 04) — для записи журнала в SYNC_LOG.

Книга 04 (хаб) — main_04.gs, sync_hub.gs, gremlin_scheduled.gs. Меню «Синхронизация книг → ⏱ Расписание → Установить триггеры этой книги (04)».

Книга 05 (себестоимость + оплаты логистики) — состав файлов для деплоя в её GAS-проект:
- main_05.gs (один onOpen, который подключает оба меню), costing.gs, payment_requests_05.gs, gremlin_scheduled.gs.
В проекте 05 не должно быть main.gs/main_01.gs (иначе будет второй onOpen).
Для работы «💳 Оплаты логистики» нужно: PAYMENT_REGISTRY_SPREADSHEET_ID / PAYMENT_REGISTRY_SHEET_NAME / PAYMENT_REGISTRY_ROOT_FOLDER_ID / PAYMENT_APPROVER_EMAIL (общие с книгой 01). Опционально — PAYMENT_DRAFT_FOLDER_NAME_LOG, PAYMENT_INITIATOR_NAME, PAYMENT_LOG_DEPARTMENT. Расписание: «⏱ Расписание → Установить триггеры этой книги (05)»; MASTER_REF_SPREADSHEET_ID — для журнала.

---
Резервная копия кода (локально + GitHub)

OneDrive для папки проекта не используется — только локальная копия на диске и GitHub.
Как отключить синхронизацию OneDrive / перенести папку: scripts/onedrive-unsync-instructions.txt

GitHub: https://github.com/AleksandrPechenkin/Gremlin
Ветка: main. История: git log --oneline

Ручной push после правок:
  git add -A
  git commit -m "кратко что сделано"
  git push

Ежедневный автоматический push (Планировщик Windows, по умолчанию 20:00):
  powershell -ExecutionPolicy Bypass -File scripts\install-git-daily-task.ps1
  powershell -ExecutionPolicy Bypass -File scripts\git-sync-push.ps1
Логи: scripts\logs\git-sync-YYYY-MM-DD.log
Задача в планировщике: Gremlin-GitHub-Daily-Push

В репозиторий не кладите токены — только код .gs, README, PROJECT_CONTEXT.md.
.gitignore: .clasp.json, __pycache__, scripts/logs/
