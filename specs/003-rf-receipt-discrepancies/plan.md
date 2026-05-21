# Implementation Plan: 003-rf-receipt-discrepancies

## Code

| Файл | Изменения |
|------|-----------|
| `costing_receipt_3pl.gs` | Новый модуль: листы, импорт, сверка, apply |
| `costing.gs` | COST_CFG, ensure sheets, `Qty_для_аллокации` в пересчёте |
| `logistics.gs` | Событие `RECEIPT_ACCEPTED` в шаблоне |
| `main_05.gs` | Меню приёмки |

## Deploy

Книга **05**: добавить `costing_receipt_3pl.gs` в GAS-проект вместе с `costing.gs`, `main_05.gs`, `logistics.gs`.
