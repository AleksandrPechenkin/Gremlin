# Implementation Plan: Цепочка отгрузки 04→02→05

**Branch**: `001-sync-chain-04` | **Date**: 2026-05-19 | **Spec**: [spec.md](./spec.md)

**Input**: Feature specification from `specs/001-sync-chain-04-02-05/spec.md`

## Summary

Один запуск из книги **04** выполняет три уже существующих бизнес-шага: копирование
«Сводной» 01→02, пересчёт транзитных движений в 02, safe-добавление строк в
«Партии_в_рейсе» в 05. Технически: оркестратор в `sync_hub.gs`, рефакторинг
`sender_stock.gs` и `costing.gs` для работы с явным `Spreadsheet` (не active),
расширение деплоя GAS-проекта книги 04 двумя дополнительными `.gs` файлами.
Исследование: [research.md](./research.md).

## Technical Context

**Language/Version**: Google Apps Script (V8), ES5-style в существующих `.gs`

**Primary Dependencies**: SpreadsheetApp, LockService, PropertiesService, UI alerts

**Storage**: Google Sheets (книги 01, 02, 04, 05); журнал `SYNC_LOG` в 04

**Testing**: Ручной dry-run + боевой прогон по [quickstart.md](./quickstart.md);
сверка с пошаговым ручным запуском (SC-002)

**Target Platform**: Google Sheets + привязанные GAS-проекты (04 — оркестратор)

**Project Type**: Multi-spreadsheet automation (Gremlin), не web app

**Performance Goals**: Полная цепочка укладывается в лимит GAS ~6 мин для типичной
отгрузки (сотни строк сводной, не тысячи)

**Constraints**:
- Constitution: safe-only для шага 3; бэкап перед записью в 05; минимальный дифф
- Шаг 01→03 **не** входит в цепочку
- Секреты только в Script Properties, не в коде
- Деплой: целые файлы в соответствующие GAS-проекты

**Scale/Scope**: v1 — интерактивное меню в 04; без нового расписания

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Status | Notes |
|-----------|--------|-------|
| I. Таблицы — операционка | ✅ Pass | Цепочка усиливает работу в таблицах |
| II. МойСклад осознанно | ✅ Pass | МС не затрагивается |
| III. Источник правды | ✅ Pass | 01 — сводная; 05 — safe-производная |
| IV. Минимальный дифф | ✅ Pass | Рефакторинг + оркестратор, без новых markdown вне specs |
| V. Деплой GAS | ⚠️ Note | Книга 04 получает +2 файла в деплое — документировать в README |
| VI. Безопасность данных | ✅ Pass | safe + snapshot на шаге 3 |
| VII. Русский UI | ✅ Pass | Меню и alert на русском |

**Post-design re-check**: ✅ Все gates pass. Расширение деплоя 04 обосновано в
[research.md R2](./research.md) (ограничение GAS — нет кросс-вызова между проектами).

## Project Structure

### Documentation (this feature)

```text
specs/001-sync-chain-04-02-05/
├── plan.md              # This file
├── research.md          # Phase 0
├── data-model.md        # Phase 1
├── quickstart.md        # Phase 1
├── contracts/
│   └── sync-chain-menu.md
├── checklists/
│   └── requirements.md
└── tasks.md             # /speckit-tasks (next)
```

### Source Code (repository root)

```text
sync_hub.gs           # NEW: syncShipmentChain04*_, menu items
sender_stock.gs       # REFACTOR: rebuildSenderStockDataImpl_(ss, opt)
costing.gs            # REFACTOR: opts.targetSs for batches rebuild
main_04.gs            # unchanged (onOpen → sync hub menu)
README.txt            # UPDATE: book 04 deploy list
PROJECT_CONTEXT.md    # OPTIONAL: shorten checkpoint after ship
```

**Structure Decision**: Gremlin — плоские `.gs` в корне репозитория; логика книг
разделена по деплою, не по папкам `src/`. Оркестрация сосредоточена в `sync_hub.gs`.

## Implementation Phases (for /speckit-tasks)

### Phase A — Refactor prerequisites

1. `sender_stock.gs`: выделить `rebuildSenderStockDataImpl_(ss, opt)`; dry-run без
   записи; `rebuildSenderStockData()` → wrapper с `getActiveSpreadsheet()`.
2. `costing.gs`: параметр `targetSs` в `costingRebuildBatchesFromSummary_`; открытие
   05 через `COST_SPREADSHEET_ID` при вызове из хаба.

### Phase B — Orchestrator (sync_hub.gs)

3. `syncShipmentChain04Impl_(dryRun, opt)` — lock, 3 шага, лог, throw on error.
4. Menu: dry-run + live with confirm ([contract](./contracts/sync-chain-menu.md)).
5. Форматирование итогового alert / silent mode.

### Phase C — Docs & deploy

6. `README.txt`: состав GAS книги 04 (+ `sender_stock.gs`, `costing.gs`).
7. Прогон [quickstart.md](./quickstart.md).

## Complexity Tracking

> No constitution violations requiring justification beyond deploy note (documented in research R2).

| Item | Why Needed | Simpler Alternative Rejected Because |
|------|------------|-------------------------------------|
| costing.gs in book 04 project | GAS cannot call 05's script remotely | Duplicate batch logic in sync_hub |

## Artifacts Generated

| Artifact | Path |
|----------|------|
| Research | [research.md](./research.md) |
| Data model | [data-model.md](./data-model.md) |
| Contract | [contracts/sync-chain-menu.md](./contracts/sync-chain-menu.md) |
| Quickstart | [quickstart.md](./quickstart.md) |

**Next command**: `/speckit-tasks`
