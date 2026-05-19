# Data Model: Цепочка отгрузки 04→02→05

**Date**: 2026-05-19

## Spreadsheet books (logical)

| Code | Role | Script project files (repo) |
|------|------|-----------------------------|
| 01 | Источник «Сводная» | `main.gs`, … |
| 02 | Копия «Сводная» + транзит | `main_02.gs`, `sender_stock.gs` |
| 04 | Хаб, журнал, оркестратор | `main_04.gs`, `sync_hub.gs`, `gremlin_scheduled.gs`, **`sender_stock.gs`**, **`costing.gs`** (после фичи) |
| 05 | «Партии_в_рейсе» | `main_05.gs`, `costing.gs`, … |

## Script Properties (chain prerequisites)

| Property | Used for |
|----------|----------|
| `ORDERS_SPREADSHEET_ID` | Книга 01 (источник сводной) |
| `TRANSIT_SPREADSHEET_ID` | Книга 02 (шаг 1 target, шаг 2) |
| `COST_SPREADSHEET_ID` | Книга 05 (шаг 3 target) |
| `MASTER_REF_SPREADSHEET_ID` | Книга 04 / SYNC_LOG |
| `ORDERS_SUMMARY_SHEET_NAME` | Имя листа (default: `Сводная`) |
| `SYNC_LOCK_TIMEOUT_MS` | Таймаут lock (existing) |

## Entities

### SummaryRow (01 → 02)

Операционная строка заказной сводки. Ключевые поля для цепочки:

| Field | Description |
|-------|-------------|
| `wb_article` | Артикул ВБ |
| `order_status` | Статус (в т.ч. «4. В пути в Москву») |
| `shipment_id` | `TR-YYYY-NNNN` |
| `qty`, `boxes`, `spec_number`, … | Для шага 3 (партии) |

**Validation**: лист «Сводная» существует в 01; при пустом листе — ошибка шага 1.

### StockMovement (02)

Событие движения на транзите, производное от строк «Сводной» (02) и карты статусов.

| Field | Description |
|-------|-------------|
| `movement_key` | Уникальность (dedupe) |
| `wb_article`, `qty_signed`, … | Как в `Stock_Movements` |

**State**: пересчёт полностью перезаписывает лист движений и пересобирает «Транзитный склад».

### BatchRow (05)

Строка «Партии_в_рейсе».

| Field | Description |
|-------|-------------|
| `SHIPMENT_ID` | Рейс |
| `Артикул ВБ` | SKU |
| `Количество`, `Объем`, цены, валюта, … | Из сводной при добавлении |

**Composite key**: `(SHIPMENT_ID, Артикул ВБ)` — unique для safe-режима.

**Rules**:
- safe: insert only if key not in `existingByKey`
- live step 3: auto backup sheet `Партии_в_рейсе_бэкап_YYYYMMDD_hhmm` before write

### SyncLogEntry (04)

| Column | Value |
|--------|--------|
| Timestamp | run time |
| Block | `Цепочка 04→02→05` / `Цепочка … шаг 2` |
| Status | `OK` / `ERROR` |
| DryRun | YES/NO |
| User | email |
| Details | step stats or error message |

## Chain state machine

```text
[Idle]
  → run(dryRun?)
  → Step1_CopySummary_01_to_02
       → fail → [Logged Error] (02 may be unchanged if dryRun; if live partial copy — see log)
       → ok → Step2_RebuildTransit_02
            → fail → [Logged Error] (02 summary already updated if step1 live)
            → ok → Step3_SafeBatches_05
                 → fail → [Logged Error]
                 → ok → [Success]
```

## Relationships

```text
Сводная (01) --copy--> Сводная (02) --derive--> Stock_Movements / Транзитный склад
Сводная (01) --read-------------------------> Партии_в_рейсе (05)  [safe append]
```

Книги 02 и 05 **не** связаны напрямую; связь только через 01 и оркестратор 04.
