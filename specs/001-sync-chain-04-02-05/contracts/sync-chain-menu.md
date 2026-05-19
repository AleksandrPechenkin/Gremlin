# Contract: Menu & orchestration API (book 04)

**Version**: 1.0.0  
**Feature**: 001-sync-chain-04-02-05

## Menu entries

Parent: `🔁 Синхронизация книг` → `Операционные потоки`

| Menu label (RU) | Handler | Writes data |
|-----------------|---------|-------------|
| Dry-run: цепочка отгрузки 04→02→05 | `syncShipmentChain04DryRun_` | No |
| Цепочка отгрузки 04→02→05 (боевой) | `syncShipmentChain04LiveWithConfirm_` | Yes (02, 05) |

Existing entries MUST remain unchanged.

## Orchestrator

### `syncShipmentChain04Impl_(dryRun, opt)`

**Parameters**:
- `dryRun` (boolean): required
- `opt.silent` (boolean, optional): suppress final UI alert (for future scheduler)

**Returns**: `string` — multiline report (step summaries)

**Throws**: `Error` with message including step id on failure

**Steps** (sequential):

| Step | Id | Function | Book |
|------|-----|----------|------|
| 1 | `01→02` | `syncOperationalOrdersSummaryFrom01To02_` | 01→02 |
| 2 | `02-transit` | `rebuildSenderStockDataImpl_` | 02 |
| 3 | `05-batches-safe` | `costingRebuildBatchesFromSummary_` | 01→05 |

**Side effects**:
- `syncHubLog_('Цепочка 04→02→05', status, details, dryRun)` on completion or error
- `LockService.getScriptLock()` held for full duration

### Refactored dependencies

#### `rebuildSenderStockDataImpl_(ss, opt)`

| Field | Type | Default |
|-------|------|---------|
| `opt.silent` | boolean | false |
| `opt.dryRun` | boolean | false |

**Returns**: stats object `{ movementsCount, rowsScanned, … }` for chain report

#### `costingRebuildBatchesFromSummary_(opts)` (extended)

| Field | Type | Default |
|-------|------|---------|
| `opts.mode` | `'safe' \| …` | `'safe'` |
| `opts.dryRun` | boolean | false |
| `opts.targetSs` | Spreadsheet | active spreadsheet |

Chain MUST pass `mode: 'safe'` only.

## User-visible final alert (non-silent)

```
[🧪 Dry-run | ✅ Готово] Цепочка отгрузки 04→02→05

Шаг 1 (01→02): …
Шаг 2 (транзит 02): …
Шаг 3 (партии 05, safe): …
```

On error:

```
❌ Цепочка остановлена на шаге N (…)

Шаг 1: OK | …
Шаг 2: ERROR — …
```

## Confirm dialog (live only)

Title: `Цепочка отгрузки 04→02→05`

Body MUST include:
1. Copy «Сводная» 01 → 02 (not 03)
2. Recalculate transit movements in 02
3. Safe append to «Партии_в_рейсе» in 05 (with backup)

Buttons: YES / NO (cancel = no op)
