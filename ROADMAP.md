# MNB PHPExcel Roadmap

## v1.3.0 — Full Excel capability upgrade: completed

- Direct cell, multi-cell, range, cell-detail, style, rich-text, and image APIs.
- Optional true formula recalculation adapter.
- Semantic automatic header detection.
- Reusable validated import templates.
- Arbitrary freeze panes and advanced native filters.
- Native conditional formatting and chart generation.
- Template-driven pivot cache preservation, rebinding, and refresh-on-open.
- Full modular-package and backward-compatibility smoke coverage.

## v1.2.0 — Completed Phase 2, Phase 3, and Phase 4 foundation

### Phase 2 — Real streaming: completed

- Direct NDJSON generator.
- Direct XML generator.
- Streaming top-level JSON-array parser.
- Shared XLSX string-provider abstraction.
- Source-level column projection.

### Phase 3 — Stable public API: completed

- Unified range terminology.
- Common normal/streaming read session.
- Typed `ReaderOptions`.
- Typed row state and progress information.
- Throw, skip, collect, and callback row-error policies.

### Phase 4 — Advanced compatibility foundation: completed

- Merged-cell expansion and metadata.
- Formula expression plus cached-result objects.
- XML schema mapping.
- Instance-scoped reader plugin registration.
- Native ODS reading.
- Optional legacy XLS adapter.

### Next release hardening

- Run full XML/XLSX/ODS integration tests on all supported PHP versions with required extensions.
- Add real ODS and binary XLS fixtures.
- Publish generated package trees through split repositories or a Composer repository.
- Benchmark top-level JSON arrays, NDJSON, XML, XLSX memory strings, and disk-backed shared strings.
- Add ODS writing only after the read compatibility matrix is stable.

## v1.0.5 — Structured export and upload-safety hardening

Status: completed. Associative JSON/XML round trips, fluent header-row parsing, hostile XLSX upload checks, always-on smoke assertions, clean-checkout testing, and PHP 8.1–8.4 CI are included.

Next priorities:

- Run the complete XLSX and SQLite-backed streaming suite in CI on every supported PHP version.
- Add mutation/fuzz tests for malformed ZIP/XML/CSV/JSON inputs.
- Add static analysis and a documented backward-compatibility/deprecation policy.
- Split broad smoke scripts into focused unit and integration suites while retaining dependency-free smoke coverage.

## v1.0.4 — Application-Ready Excel Toolkit Release

Status: ready for public release. This release includes the rich small/normal workbook engine, large XLSX streaming import/export, SQL import/export workflows, application integration helpers, plugin/extensibility layer, benchmark/compatibility proof scaffolding, and release-quality smoke coverage.

## v1.0.0 — First Public Release

Status: completed for the supported rich small/normal-file scope. Large import/export support is available through separate streaming modes below.

### Included

- Array-first XLSX/CSV/JSON/XML engine.
- Small-file XLSX read/write.
- CSV dialects, encoding, locale parsing, and injection safety.
- Report builder.
- Import preview and validation.
- Failed rows report.
- Import summary report.
- SQL dry-run and simple SQL import/export helpers.
- Database connection resolution from `.env`, PHP config files, DSN strings, arrays, constants, runtime environment variables, or existing `PDO`.
- Formula and cell safety.
- Release readiness checks.
- JSON/XML conversion.
- JSON to XLSX/CSV/XML/SQL conversion.
- Multi-sheet JSON workbook import.
- XLSX rich text/comment/hyperlink metadata extraction.
- Template-based advanced XLSX object preservation for supported package parts.
- XLSX integrity validation and save-time corruption prevention.
- Error handling and atomic save guard for CSV, JSON, XML, XLSX, structured outputs, and SQL wrappers.
- Environment diagnostics for extension/runtime readiness.
- XLSX compatibility verification harness for generated and optional external fixtures.
- Programmatic XLSX comments/notes and hyperlinks writer.
- Smoke tests.

### Not included in rich workbook mode

- Large XLSX files should not use the normal full-array reader. Use the separate large import mode below.
- Parallel workers.
- NDJSON/SQLite staging.
- Built-in native formula calculation engine. True recalculation is available through the optional adapter.
- From-scratch pivot-layout design. Template-driven pivot workflows are supported.
- Native legacy `.xls` engine. Optional XLS compatibility is available through the XLS module.
- Macro execution. Macros are never executed; advanced package preservation is available only when explicitly requested from a trusted template.

## Completed foundation: Large File Import Engine

- Streaming worksheet reader.
- Row iterator.
- Chunk iterator.
- Row counter/preflight advisor.
- Memory guard.
- Shared-string memory/SQLite cache strategy.
- SQL batch import engine.
- Failed-row CSV export.
- Progress manifest and checkpoint/resume.

Still future:

- Safe NDJSON chunk converter.
- Temp disk guard.
- Queue framework adapters. Import dashboard helper is completed.
- Deeper arbitrary existing-workbook object-graph editing.


## Completed foundation: Large File Export Engine

- Streaming XLSX writer for generators/iterables.
- PDO cursor export without `fetchAll()`.
- Inline strings for low-memory writing.
- Auto sheet splitting.
- Basic column formats: integer, decimal, currency, date, datetime, percent, text.
- Progress callbacks and export events.
- Atomic save and XLSX integrity validation.
- CSV ZIP fallback with split CSV files and manifest.

## Completed application polish: Import Dashboard Helper

- Dashboard/API response from import manifests.
- Progress percentage and ETA.
- Failed-row download metadata.
- Resume button/action data.
- Safe admin messages.

## Future: Enterprise Import Engine

- Resume failed import.
- Failed chunk retry.
- Audit logs.
- Queue/worker orchestration.
- Import dashboards.
- Duplicate strategy policies.
- SQLite staging.


## First Public Release Completion Additions

- Structured output completion: `toStructuredArray()`, `toStructuredJson()`, and `saveStructuredJson()`.
- Header rename/alias support for read-time cleanup.
- Header case modes for snake, preserve, lower, and camel output.
- Composer metadata included for Packagist readiness.


## Public release structured output note

Workbook-level structured output is completed in v1.0.0 through `toStructuredArray()` / `toStructuredWorkbookArray()`. Selected-sheet output remains available through `toStructuredSheetArray()`.

## v1.0.2 — Structured JSON return hotfix

- `toStructuredJson()` returns JSON string for assignment, printing, controllers, and API responses.
- `saveStructuredJson()` remains file-save only.
- Composer version field removed so Packagist imports tags cleanly.

- Structured XML return support completed: `toStructuredXml()` returns XML string and `saveStructuredXml()` saves XML files.

## Unreleased — XLSX metadata and preservation upgrade

- `sheetMetadata()` for rich text runs, comments/notes, hyperlinks, and advanced object inventory.
- Structured XLSX output includes metadata by default.
- `preserveAdvancedObjectsFrom()` copies supported advanced workbook package parts from a trusted source template while regenerating row data.


## Unreleased — XLSX Integrity Validator

- Validate required XLSX package files before returning generated output.
- Validate every internal `.rels` target and skip only explicit external relationship targets.
- Validate content type coverage for workbook, sheet, styles, docProps, drawings, media, comments, and preserved package parts.
- Validate XML package-part structure, using strict XMLReader parsing when available.
- Validate workbook and worksheet relationship IDs so `r:id` references do not point to missing relationships.
- Delete failed XLSX output and throw before returning a corrupted file path.


## Unreleased — Error Handling and Atomic Save Guard

- Stable string error codes for package exceptions.
- Public-safe error responses for web/API use.
- Debug error reports for logs/admin screens.
- Atomic temp-file save workflow for writer outputs.
- Stronger write checks for CSV and XLSX zip entries.
- Partial/temp file cleanup on failure.
- PDO import/export error wrapping with previous exceptions preserved.


## Unreleased — XLSX Compatibility Verification and Comments/Hyperlinks Writer

- `environmentCheck()` reports PHP/XLSX/SQL/encoding capability status.
- `verifyXlsxCompatibility()` validates generated compatibility fixtures and optional real-world XLSX files exported from Excel, LibreOffice, Google Sheets, or WPS Office.
- Generated compatibility cases cover basic workbook structure, formulas, styles, merged cells, comments/notes, hyperlinks, and advanced-object template preservation.
- `hyperlink()` / `hyperlinks()` write XLSX hyperlink relationships.
- `comment()` / `comments()` write classic Excel note/comment parts with VML drawing support.
- Integrity validation remains the save-time guard for these generated package parts.

## Large-file import roadmap

Completed foundation:

- Large Excel preflight analyzer
- Import method advisor
- Decision matrix by rows/cells/server level
- Streaming chunk read session
- Memory and time-budget guards
- Shared strings memory/SQLite cache

Recommended next upgrades:

1. Database batch import adapter for streaming chunks.
2. Checkpoint/resume support for CLI/queue jobs.
3. Failed-row CSV export during streaming validation.
4. Progress manifest file for admin dashboards.
5. Large export streaming writer with auto sheet splitting.

## Completed — LargeExcelDatabaseImportEngine

- Streaming chunk validation.
- PDO batch insert.
- Failed-row CSV export.
- Progress manifest for admin dashboards.
- Checkpoint/resume for CLI/queue/shared-hosting jobs.

Completed stability upgrade:

- Environment alert helpers for missing XLSX/SQLite extensions.
- Numeric-string preservation for large streaming imports.
- Excel date serial conversion through `styles.xml` date styles.
- Idempotent DB import options and duplicate strategies.
- Human-readable failed-row CSV output.
- All-sheet large workbook import orchestrator.
- Stronger local XLSX runtime smoke tests.

Recommended next upgrades:

1. Large import admin/job dashboard helpers.
2. Failed-row re-import helper.
3. Database-driver benchmarking and tuning for MySQL/PostgreSQL/SQLite.
4. Large XLSX streaming writer with auto sheet splitting.


## Completed: RealApplicationAndPluginIntegrationLayer

The package now includes application-level storage, profiles, import status/resume helpers, upload validation, events, plugins, transformers, custom validators, and optional logger support. Future framework adapters should live in separate packages such as `mnb/mnb-phpexcel-laravel`, `mnb/mnb-phpexcel-codeigniter`, and `mnb/mnb-phpexcel-symfony` so the core remains lightweight.

## Next advanced spreadsheet roadmap

### Optional adapter and compatibility expansion

The v1.3 release completed the optional PhpSpreadsheet calculation adapter, native conditional formatting, chart generation, ODS reading, NDJSON streaming, and recalculate-on-open support. Future work can remain outside the lightweight core:

- Deep arbitrary editing of existing complex workbook object graphs.
- OpenSpout comparison adapter for alternative streaming workflows.
- ODS writing, HTML table adapters, and additional data-format plugins.
- Additional chart families and chart-template mutation.
- Optional formula dependency tracing and calculation diagnostics.

### Release proof roadmap

- Add real fixture files from Microsoft Excel, LibreOffice Calc, Google Sheets export, and WPS Office.
- Run and publish local benchmark tables for 100k, 500k, and 1M rows.
- Add live MySQL and PostgreSQL integration CI jobs when credentials/services are available.
- Build separate Laravel adapter package instead of adding Laravel as a core dependency.
