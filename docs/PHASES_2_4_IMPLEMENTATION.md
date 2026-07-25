# Phase 2–4 implementation report

## Phase 2 — Real streaming

| Item | Status | Implementation |
|---|---|---|
| Direct NDJSON generator | Completed | `JsonReader::iterateSheet()` reads one physical line at a time with record-size and row limits. |
| Direct XML generator | Completed | `XmlReader::iterateSheet()` uses `XMLReader` and yields selected rows without a workbook array. |
| Streaming JSON-array parser | Completed | `StreamingJsonArrayParser` tokenizes a top-level array from fixed-size chunks and decodes one item at a time. |
| Shared XLSX string-provider abstraction | Completed | `SharedStringProviderInterface`, in-memory provider, and SQLite/disk-backed `LargeSharedStringCache` are shared by normal and large readers. |
| Source-level column projection | Completed | `ColumnProjection` is used by CSV, JSON, XML, XLSX, ODS, and XLS adapters before session normalization where the source permits it. |

Complex workbook-shaped JSON objects still use materialized parsing because named-sheet selection requires document-level structure. This behavior is explicit rather than being presented as streaming.

## Phase 3 — Stable public API

| Item | Status | Implementation |
|---|---|---|
| Unified range terminology | Completed | `ReaderOptions::withRange()` and `ReadSession::range()` normalize `start_row`, `end_row`, `start_column`, and `end_column`. |
| Common normal/large read session | Completed | `ReadSession::normal()`, `streaming()`, `rows()`, `rowStates()`, and `chunks()` share one API. Legacy large sessions remain supported. |
| Typed `ReaderOptions` | Completed | Immutable options object plus `ReadMode` and `RowErrorPolicy` enums; arrays remain compatible. |
| Row state and progress | Completed | JSON-serializable `RowState` and `ReadProgress`; final completed progress event. |
| Row-level error policies | Completed | Throw, memory-free skip, collect, and callback/replacement behavior. |

## Phase 4 — Advanced compatibility

| Item | Status | Implementation |
|---|---|---|
| Merged cells | Completed | XLSX modes: `anchor`, `expand`, and metadata range reporting via `MergedCellMap`. |
| Richer formula results | Completed | `formula_cells => both` returns `FormulaResult` with expression, cached result, result type, and shared/array formula metadata. It does not calculate formulas. |
| XML schema mapping | Completed | Safe child paths, row attributes, text, defaults, required fields, and type conversions through `XmlSchemaMapping`. |
| Reader plugin registration | Completed | Instance-scoped `ReaderRegistry` and `SpreadsheetManager`; legacy static registration retained. |
| ODS support | Completed for reading | Native forward-only ODS reader with repeated row/column handling, values, formulas, ranges, and projection. |
| XLS support | Completed through adapter | Optional PhpSpreadsheet-backed binary XLS reader in a separate package. |

## Compatibility boundaries

- Formula calculation is not included; cached values are whatever the producing spreadsheet application saved.
- ODS writing is not part of 1.2.
- Binary XLS is intentionally not implemented natively because a safe, complete BIFF parser is a separate large project.
- Runtime tests for native XLSX/ODS/XML require ZIP/XMLReader extensions. The reader implementations use depth-aware subtree skipping rather than sibling-jumping `XMLReader::next()` calls.
- XLS adapter tests require the optional PhpSpreadsheet dependency.
