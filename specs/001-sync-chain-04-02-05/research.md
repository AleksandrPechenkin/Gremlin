# Research: Цепочка отгрузки 04→02→05

**Date**: 2026-05-19  
**Feature**: [spec.md](./spec.md)

## R1: Где размещать оркестратор

**Decision**: Точка входа и меню — в `sync_hub.gs` (книга 04), подменю «Операционные потоки».

**Rationale**: Книга 04 уже открывает 01/02/05 по Script Properties (`syncHubOpenSpreadsheetForBook_`), пишет `SYNC_LOG`, использует `LockService` для снимков. `main_04.gs` только подключает меню хаба.

**Alternatives considered**:
- Оркестратор в книге 01 — отклонено: оператор после отгрузки логичнее запускает из хаба; 01 не знает про транзит/партии как единый поток.
- Отдельный триггер по расписанию — отклонено для v1: нужен явный запуск после ручной отгрузки.

## R2: Вызов шагов 2 и 3 из проекта скриптов книги 04

**Decision**: Рефакторинг существующих функций на работу с явным объектом `Spreadsheet` + **расширение состава GAS-проекта книги 04**: добавить в деплой `sender_stock.gs` и `costing.gs` (те же файлы, что в 02/05).

**Rationale**: В Apps Script код проекта 04 **не может** вызвать функции, объявленные только в проекте 02 или 05. `rebuildSenderStockData()` и `costingRebuildBatchesFromSummary_()` сегодня используют `SpreadsheetApp.getActiveSpreadsheet()` — при запуске из 04 active = книга 04, что сломает шаги 2–3.

**Alternatives considered**:
- Apps Script Library — отклонено: нет общей библиотеки в репозитории, лишняя настройка для владельца.
- Дублировать логику пересчёта в `sync_hub.gs` — отклонено: два источника правды, высокий риск расхождения.
- UrlFetch к внешнему runner — отклонено: избыточно.

**Refactor contract**:
- `rebuildSenderStockDataImpl_(ss, opt)` — `opt.silent`, `opt.dryRun` (при dryRun: построить события и статистику без записи в `Stock_Movements` / баланс).
- `costingRebuildBatchesFromSummary_(opts)` — добавить `opts.targetSs`; если задан, использовать вместо active; snapshot и запись только при `dryRun: false`.

## R3: Шаг 1 — только 01→02

**Decision**: Вызывать существующую `syncOperationalOrdersSummaryFrom01To02_(dryRun, { silent: true })`, **не** `syncOperationalSnapshotsImpl_` (который также копирует в 03).

**Rationale**: Прямое соответствие FR-003 и Assumptions в spec.

## R4: Блокировки и частичный успех

**Decision**: Одна блокировка `LockService.getScriptLock()` на всю цепочку (как в `syncOperationalSnapshotsImpl_` / `syncAllExternalBooksImpl_`). При ошибке шага N — `throw`, шаги N+1 не вызываются; финальный `syncHubLog_` с префиксом `Цепочка 04→02→05` и номером шага в Details.

**Rationale**: Соответствует FR-006, FR-007; снижает гонки при параллельном запуске.

**Alternatives considered**:
- Откат шага 1 при падении шага 2 — отклонено для v1: копирование сводной обратимо только вручную; достаточно явного лога и сообщения.

## R5: Dry-run цепочки

**Decision**:
- Шаг 1: `dryRun: true` в `syncHubCopyMappings_` (уже есть).
- Шаг 2: новый путь в `rebuildSenderStockDataImpl_` без записи листов.
- Шаг 3: существующий `dryRun` в `costingRebuildBatchesFromSummary_` (без snapshot при dryRun).

**Rationale**: FR-002; согласовано с одиночными dry-run в меню себестоимости.

## R6: UI и подтверждение

**Decision**: Два пункта меню в «Операционные потоки»:
- `Dry-run: цепочка отгрузки 04→02→05` → `syncShipmentChain04DryRun_`
- `Цепочка отгрузки 04→02→05 (боевой)` → confirm → `syncShipmentChain04Live_`

Текст confirm перечисляет 3 шага и напоминает, что 01→03 не выполняется.

**Rationale**: FR-001, FR-009; паттерн как у других боевых пунктов хаба.

## R7: Лимит времени GAS

**Decision**: v1 без разбиения на триггеры; целевой объём — типичная отгрузка (десятки–сотни строк сводной). В `quickstart.md` зафиксировать проверку на реальном объёме; при timeout — follow-up (отдельная фича).

**Rationale**: Assumption в spec; hourly snapshot уже гоняет 01→02+03 отдельно.

## R8: Деплой и обратная совместимость

**Decision**: Обновить `README.txt` (состав книги 04). Меню в 02/05 не менять; обёртки `rebuildSenderStockData` / menu costing вызывают refactored impl с active spreadsheet.

**Rationale**: FR-008, Constitution V (деплой целым файлом).
