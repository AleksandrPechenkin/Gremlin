# Tasks: Цепочка отгрузки 04→02→05

**Input**: [plan.md](./plan.md), [spec.md](./spec.md)

## Phase 1: Refactor

- [x] T001 Refactor `rebuildSenderStockDataImpl_(ss, opt)` in `sender_stock.gs`
- [x] T002 Add `opts.targetSs` to `costingRebuildBatchesFromSummary_` in `costing.gs`

## Phase 2: Orchestrator

- [x] T003 [US1] Implement `syncShipmentChain04Impl_` in `sync_hub.gs`
- [x] T004 [US1] Add menu dry-run and live confirm in `sync_hub.gs`
- [x] T005 [US2] Error handling with step labels and `SYNC_LOG` in `sync_hub.gs`

## Phase 3: Docs

- [x] T006 Update book 04 deploy list in `README.txt`
- [x] T007 Note implementation in `PROJECT_CONTEXT.md`

## Manual verification (owner)

- [ ] T008 Run dry-run and live chain per [quickstart.md](./quickstart.md) in Google Sheets
