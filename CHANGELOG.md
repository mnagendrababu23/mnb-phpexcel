# Changelog

## 1.4.0 — Typed Domain Import APIs

### Added

- First-class presets for users, products, orders, inventory, students, attendance, marks, contacts, locations, blog posts, media paths, and categories.
- `DomainImportType`, `DomainImportPreset`, instance-scoped `DomainImportRegistry`, and streaming `DomainImporter` APIs.
- Domain schema inspection, dry-run previews, validated templates, alias-aware mapping, value normalization, cross-field rules, duplicate policies, progress callbacks, and failed-row reporting.
- High-level `MnbExcel::importUsers()`, `importProducts()`, `importOrders()`, `importInventory()`, `importStudents()`, `importAttendance()`, `importMarks()`, `importContacts()`, `importLocations()`, `importBlogPosts()`, `importImagesWithPaths()`, `importMedia()`, and `importCategories()` methods.
- `max_length` and `min_length` validation rules.
- Dedicated domain API and isolated modular-package tests.

### Changed

- Version increased to `1.4.0`.
- Split package dependencies now target the matching `^1.4` family.
- The database package owns domain presets/services; the application package exposes facade methods. No additional split repository is required.

### Compatibility

- Existing generic imports and v1.3 APIs remain available.
- Presets do not force a fixed database schema: table names, mappings, aliases, rules, defaults, normalizers, transformers, and unique keys are overridable.
- The same domain API can consume any installed reader format.

## 1.3.0 — Full Excel Cell, Presentation, Template, and Pivot Workflows

### Added

- Direct `cell()`, `cells()`, `rangeValues()`, `cellDetails()`, `cellStyle()`, and `rangeStyles()` APIs for native XLSX sessions.
- Typed `CellSnapshot`, `RichText`, `RichTextRun`, and `HeaderDetection` value objects.
- Complete XLSX style metadata for fonts, fills, borders, number formats, alignment, and protection.
- Image inventory and collision-safe embedded-image extraction with anchors, dimensions, MIME types, and descriptions.
- Optional PhpSpreadsheet formula evaluator for true workbook recalculation without increasing the native XLSX package's mandatory dependency footprint.
- Semantic header detection through `detectHeader()`, `autoDetectHeader()`, and `ReaderOptions::withAutoHeader()`.
- Reusable import-template generation with instructions, examples, comments, and native data-validation rules.
- Arbitrary row/column freeze panes and explicit top-left pane selection.
- Native value, custom, top/bottom, dynamic, and color auto-filter definitions.
- Native conditional formatting with differential styles, expressions, comparisons, color scales, data bars, icon sets, duplicate/unique rules, text rules, and time periods.
- Native column, bar, line, area, pie, doughnut, and scatter chart generation.
- Template-driven pivot-table preservation, workbook pivot-cache relationships, source-range rebinding, and refresh-on-open support.
- Format-specific XLSX convenience methods and a full-capability v1.3 smoke suite.

### Changed

- Version increased to `1.3.0`.
- Split packages now depend on the corresponding `^1.3` module family.
- The XLSX split package suggests PhpSpreadsheet only for optional true formula recalculation.
- Direct multi-cell reads share worksheet XML, style, date-system, and shared-string state instead of reopening the package for each cell.
- Cell style reads no longer scan comments, images, and unrelated worksheet metadata.

### Compatibility

- Existing v1.2 APIs and defaults remain valid.
- Native formula reading still has no extra dependency; only true recalculation uses the optional adapter.
- Pivot tables are fully supported through trusted templates; a from-scratch pivot-layout designer is not claimed.

## 1.2.0 — Modular Packages, Real Streaming, Stable Reader API, Advanced Compatibility

### Added

- Generated Composer package family for core, CSV, JSON, XML, XLSX, ODS, XLS, database, application, and all-module installs.
- Format-specific `Format\Csv`, `Format\Json`, `Format\Xml`, `Format\Xlsx`, `Format\Ods`, and `Format\Xls` entry points.
- Direct NDJSON generator and fixed-chunk top-level JSON-array parser.
- Direct XML row generator and safe XML schema mapping.
- Shared XLSX string-provider contract with in-memory and disk/SQLite providers.
- Source-level column projection and unified row/column ranges.
- Immutable `ReaderOptions`, `ReadMode`, and `RowErrorPolicy`.
- Common normal/streaming `ReadSession`, row states, progress objects, chunks, and row-error collection/replacement.
- XLSX merged-cell expansion/metadata and `FormulaResult` objects containing formula plus cached result.
- Native streaming ODS reader.
- Optional PhpSpreadsheet-backed binary XLS adapter.
- Instance-scoped reader registry and plugin registration.
- Modular package builder, archive manifest, hashes, and isolated package tests.

### Changed

- `ReadSession::toArray()` now consumes the common row pipeline so typed options, progress, row-error policies, projection, and streaming behavior are consistent.
- `MnbExcel::read()` accepts `ReaderOptions` and discovers installed reader modules through `ReaderRegistry`.
- Version increased to `1.2.0`.

### Fixed

- Alphabetic associative projection keys such as `name` are no longer mistaken for Excel column letters. Positional letters are explicitly uppercase `A` through `XFD`.
- XLSX shared-string providers are closed when worksheet XML cannot be opened.
- Depth-aware XML subtree skipping replaces `XMLReader::next()` in XML, XLSX, and large-XLSX readers, preventing selected rows or adjacent cells from being skipped accidentally.
- Expanded merged-cell values retain anchors even when the anchor row itself is outside the requested output range.
- `skip` row-error policy no longer retains errors in memory; only `collect` and `callback` retain details.
- `RowState` source and output row numbers are consistently one-based.
- Progress completion reports physical source rows, including header rows.

## v1.1.1 — Reader correctness phase 1

### Fixed

- Duplicate header renaming now guarantees collision-free final names, including cases such as `a,a,a_2`.
- Integer casting no longer silently clamps values outside PHP's integer range; oversized integers remain exact strings by default.
- Precision-sensitive decimal strings remain strings in safe numeric mode instead of being silently rounded.
- JSON Lines / NDJSON is detected from multiple independent JSON records even when the file has an unrelated suffix.
- Header selection no longer mixes physical and normalized row positions in the fluent API.

### Added

- Added `headerAtPhysicalRow()`, `headerAtDataRow()`, and `firstNonEmptyRowAsHeader()`.
- `withHeaderRow()` now explicitly uses normalized data-row semantics.
- Added `extra_columns` policies: `ignore`, `error`, `generate`, and `collect`.
- Added `missing_columns` policies: `null` and `error`.
- Added safe numeric controls: `numeric_precision`, `big_integer_mode`, `max_safe_decimal_digits`, and `invalid_cast`.
- Added `json_document_mode` and content-based NDJSON selection.
- Added regression tests for all five phase-one correctness areas.

### Compatibility

- Direct array option `header_row` without `header_row_mode` retains legacy selection behavior.
- Existing extra-column truncation and missing-column null padding remain the defaults (`ignore` and `null`).
- Callers that intentionally need native overflow/rounding behavior may set `numeric_precision => 'native'`.

## v1.1.0 — Reader functionality upgrade

### Added

- Added automatic reader format detection for XLSX, CSV/TSV, JSON/JSON Lines, and XML through `MnbExcel::read()`.
- Added explicit `MnbExcel::readXlsx()`, `MnbExcel::readXml()`, and `MnbExcel::detectFormat()` APIs.
- Added secure XML row/workbook reading compatible with the library XML writer.
- Added JSON Lines / NDJSON import with large-integer string preservation.
- Added forward-only row iteration for CSV and normal XLSX readers.
- Added fluent reader operations: `withOptions()`, `withoutHeaderRow()`, `skip()`, `limit()`, `selectColumns()`, `rows()`, `first()`, `countRows()`, `eachRow()`, and `chunk()`.

### Improved

- CSV reader can auto-detect comma, semicolon, tab, or pipe delimiters.
- CSV reader supports strict column counts, row/column/file guards, blank-line control, and source row ranges.
- XLSX reader supports source row ranges, projected columns, compact selected-column output, package validation, and worksheet/shared-string size guards.
- JSON/XML readers support file-size, row-count, source range, and source limit controls.
- Reader errors now include more precise file, row, column, and limit context.

### Compatibility

- Existing `readCsv()`, `readJson()`, `read()` XLSX usage, `withHeaderRow()`, and `toArray()` behavior remain available.
- `toArray()` uses the new iterator path only when fluent skip/limit/column-selection transforms are requested.

## v1.0.5 — Structured export and upload-safety hardening

### Fixed

- Preserved associative input keys in default JSON and XML workbook exports.
- Added `preserve_associative_rows => false` for callers that need the previous indexed-table representation.
- Added `data_only => true` to omit generated title/header/summary/footer presentation rows from structured exports.
- Added `ReadSession::withHeaderRow()` so CSV/XLSX header parsing is explicit and fluent.
- Made the smoke-test bootstrap work from a clean source checkout even before Composer creates `vendor/autoload.php`.

### Security

- Hardened upload validation with PHP upload-error handling, actual-size verification, symlink rejection, MIME inspection, and filename null-byte checks.
- Added XLSX ZIP entry-count, uncompressed-size, per-entry-size, compression-ratio, path-traversal, and encryption checks.
- Added configurable macro and external-link rejection for untrusted workbook uploads.

### Quality

- Added regression coverage for associative JSON/XML output, fluent header rows, MIME metadata, unsafe names, upload errors, and ZIP path traversal.
- Added a PHP 8.1–8.4 GitHub Actions test matrix.
- Forced `zend.assertions=1` in smoke-test subprocesses so legacy `assert()` checks cannot silently compile out.

## Unreleased — Shipped test suite and CI

- Added a real `tests/` directory (`*SmokeTest.php` files) so `composer test:smoke` /
  `composer test` no longer fail on a clean checkout — previously
  `tools/run-smoke-tests.php` expected `tests/*SmokeTest.php` but no such
  directory was included in the release package.
- Added `.github/workflows/ci.yml`: runs `composer test` (syntax check +
  smoke tests) on PHP 8.1–8.4, both with and without the optional
  `ext-zip`/`ext-xmlreader`/`ext-pdo` extensions installed, plus a job that
  publishes the `environmentCheck()` report as a build artifact/log.
- Initial smoke coverage: array↔XLSX round trip and integrity validation,
  CSV round trip and formula/CSV-injection scanning, JSON/XML conversion,
  array validation (`validateArray()`), formula-guard security behavior
  (including that a blocked formula actually prevents a save), the large
  streaming writer/reader round trip, and environment/release-readiness
  diagnostics. See `tests/README.md` for scope and what's intentionally not
  covered yet.

## v1.0.4 — Application-Ready Excel Toolkit Release

### Added

- Rich small/normal XLSX, CSV, JSON, XML, SQL, report, validation, safety, structured output, and integrity workflows.
- Large XLSX preflight analyzer, method advisor, streaming reader, streaming writer, PDO cursor export, CSV ZIP fallback, and benchmark/compatibility proof tools.
- Large Excel database import engine with chunk validation, batch insert, failed-row CSV, manifest, checkpoint/resume, idempotent duplicate strategies, and dashboard/status helpers.
- Real application integration layer with import profiles, storage paths, upload safety, event hooks, plugin manager, row transformers, custom validators, framework-neutral adapters, and optional logger support.
- Database configuration support through existing PDO, .env files, PHP config files, DSN strings, arrays, constants, and runtime environment variables.

### Release notes

- Composer/Packagist version is intentionally controlled by Git tag `v1.0.4`. `composer.json` does not include a hardcoded `version` field.
- The normal workbook reader/writer remains optimized for small and normal files; large files should use the streaming import/export APIs.

### Runtime smoke-test hardening patch

- Hardened large SQL import row counters against reader/manifest drift in runtime smoke tests.
- Improved duplicate `skip` insert counting for SQLite/MySQL/PostgreSQL-style idempotent imports.
- Treated `.git/` and `vendor/` as release-readiness warnings in developer working copies while keeping clean ZIP guidance.
- Made generated compatibility fixtures advisory warnings instead of hard failures when package integrity needs manual Excel/LibreOffice confirmation.


## Local composer-test stability fixes

- Fixed large database import fresh-run header handling so streaming import consumes the header in the main pass instead of pre-reading it separately. This keeps `rows_scanned`, validation, and insert counts accurate on Windows/XAMPP generated XLSX fixtures.
- Added `.gitignore` to the release package so `ReleaseReadiness` no longer fails on clean ZIP audits.
- Downgraded the optional advanced-object preservation compatibility fixture to a warning when package-level validation needs manual Excel/LibreOffice confirmation, while core generated fixtures still fail on integrity errors.

## Large XLSX streaming writer and dashboard helper upgrade

- Added `MnbExcel::largeExport()`, `MnbExcel::largeWrite()`, `MnbExcel::largeWriteCsvZip()`, `MnbExcel::largeExportFromSql()`, and `MnbExcel::largeWriteFromSql()`.
- Added `LargeXlsxStreamingWriter` and `LargeXlsxWriteSession` for generator/iterable/PDO-cursor exports without full workbook buffering.
- Added inline-string XLSX writing, auto sheet splitting, basic column formats, header freeze/autofilter, progress callbacks, atomic save, and XLSX integrity validation.
- Added CSV ZIP fallback for ultra-large table exports with split CSV parts and `manifest.json`.
- Added `ImportDashboardHelper`, `MnbExcel::importDashboard()`, and `MnbExcel::importStatusResponse()` for admin/API progress responses, failed-row download metadata, resume actions, safe messages, and ETA.
- Added smoke tests for the large writer API and dashboard helper.


## Database connection configuration upgrade

- Added framework-neutral database configuration resolution from `.env`, PHP config files, DSN strings, arrays, constants, and runtime environment variables.
- Added `MnbExcel::dbConfig()`, `MnbExcel::pdo()`, `MnbExcel::databaseConnection()`, and `MnbExcel::databaseConfigSummary()`.
- SQL helpers can now accept an existing `PDO`, `.env` path, PHP config file path, DSN string, config array, constants, or runtime env.
- Added database config validation and connection error codes.
- Added database connection config smoke test and example.

## Accuracy and option hardening

- Fixed XLSX numeric serialization so decimal values are no longer rounded to 6 places during write.
- Preserved explicit `CellValue::number()` numeric strings until XLSX serialization to avoid early PHP float coercion.
- Allowed formula cached numeric values to keep string-supplied precision.
- Added normal XLSX reader option `preserve_numeric_strings` to avoid PHP int/float coercion for IDs, long numbers, and high-precision decimals.
- Made normal XLSX shared string loading avoid direct `zip://...sharedStrings.xml` XMLReader paths for better Windows/XAMPP reliability.
- Hardened XML escaping with invalid UTF-8 substitution in XLSX/XML writers.

## Large import audit hardening

- Fixed large streaming reader partial-chunk flushing when row limits/time budgets stop before a full chunk.
- Added shared string XML size guard so large shared string tables do not silently fall back to unlimited memory mode when `uniqueCount` is missing.
- Made preflight shared string count parsing avoid direct `zip://` XMLReader usage for better Windows/XAMPP compatibility.


## Unreleased

- Fixed Windows/PowerShell Composer test compatibility by replacing Unix-only `find | xargs` and shell loops with PHP-based test runners.
- Fixed large streaming reader shared-string probing so workbooks that use inline strings no longer emit XMLReader warnings when `xl/sharedStrings.xml` is absent.

## Unreleased — XLSX Compatibility Verification and Comments/Hyperlinks Writer

### Added

- `MnbExcel::environmentCheck()` for PHP extension/runtime capability diagnostics.
- `Support\EnvironmentDiagnostics` for reporting XLSX, SQL, encoding, and atomic-save readiness.
- `MnbExcel::verifyXlsxCompatibility()` for generated and optional external XLSX fixture verification.
- `Support\XlsxCompatibilityVerifier` with generated cases for basic workbooks, formulas/styles/merged cells, comments/hyperlinks, and advanced-object preservation flow.
- `WorkbookBuilder::hyperlink()` and `WorkbookBuilder::hyperlinks()` for programmatic XLSX hyperlink writing.
- `WorkbookBuilder::comment()` and `WorkbookBuilder::comments()` for programmatic classic Excel note/comment writing.
- `tests/EnvironmentDiagnosticsSmokeTest.php`.
- `tests/XlsxCompatibilityVerificationSuiteSmokeTest.php`.
- `tests/CommentsHyperlinksWriterSmokeTest.php`.
- `examples/xlsx_compatibility_verification.php`.
- `examples/comments_hyperlinks_writer.php`.

### Improved

- XLSX writer now emits worksheet hyperlink nodes plus external hyperlink relationships.
- XLSX writer now emits comments XML, VML note drawing parts, worksheet relationships, and content type declarations for generated comments.
- Save-time XLSX integrity validation now covers the generated comments/hyperlinks relationship structure before returning the file path.
- README now includes an XLSX compatibility matrix and notes about manual UI-level repair-warning verification.

## Unreleased — Error Handling and Atomic Save Guard

### Added

- Stable package error codes through `Support\ErrorCode`.
- Backward-compatible `MnbExcelException` metadata: error code, category, context, safe message, and previous exception.
- `MnbExcel::safeError()` for frontend/API-safe error responses.
- `MnbExcel::errorReport()` for debug/admin/log error reports.
- `Support\AtomicFileWriter` for temp-file-first save operations.
- Smoke coverage in `tests/ErrorHandlingAndAtomicSaveGuardSmokeTest.php`.
- Example usage in `examples/error_handling_atomic_save.php`.

### Improved

- CSV, JSON, XML, structured JSON/XML, and XLSX saves now write through temporary files before replacing the final output.
- CSV `fwrite()` and close operations are checked and throw package exceptions on failure.
- XLSX `ZipArchive::addFromString()`, `addFile()`, and `close()` results are checked.
- Existing output files remain unchanged when export fails before final replacement.
- Temporary/partial files are deleted on failure.
- PDO SQL export/import errors are wrapped into package-level exceptions with previous exceptions preserved.


## Unreleased — XLSX Integrity Validator Upgrade

### Added

- `Support\XlsxIntegrityValidator` for XLSX package integrity validation before returning generated workbooks.
- `MnbExcel::validateXlsx()` to return a structured validation report.
- `MnbExcel::assertValidXlsx()` to throw on invalid/corrupted XLSX packages.
- Save-time validation for XLSX output through `WorkbookBuilder`; failed output is deleted and `MnbExcelException` is thrown.
- `WorkbookBuilder::xlsxIntegrityValidation()`, `strictXlsxIntegrityValidation()`, and `skipXlsxIntegrityValidation()` for explicit control.
- Smoke coverage in `tests/XlsxIntegrityValidatorSmokeTest.php`.
- Example usage in `examples/xlsx_integrity_validation.php`.

### Checks

- Required XLSX package parts: `[Content_Types].xml`, root relationships, workbook, workbook relationships, styles, and worksheets.
- `.rels` relationship targets, skipping explicit external targets.
- XML package-part structure, with strict parsing when `ext-xmlreader` is available.
- Content type defaults/overrides and package-part coverage.
- Workbook and worksheet `r:id` references against their relationship files.

### Improved

- XLSX corruption prevention now fails fast instead of returning a file that Excel may repair.

## Unreleased — XLSX Metadata and Preservation Upgrade

### Added

- `ReadSession::sheetMetadata()` for XLSX rich text runs, comments/notes, hyperlinks, and advanced object inventory.
- Structured workbook/sheet output now includes XLSX `metadata` by default. Pass `include_cell_metadata => false` to skip it.
- `WorkbookBuilder::preserveAdvancedObjectsFrom()` and alias `preserveAdvancedExcelObjects()` to copy supported advanced XLSX package parts from a source template while regenerating row data.
- Advanced object preservation covers worksheet relationship XML and common package parts such as drawings, comments, threaded comments, VML drawings, tables, charts, pivot/cache parts, embedded media, printer settings, query tables, slicers, timelines, and custom XML.

### Improved

- XLSX inspection messages now point users to metadata extraction and template preservation instead of marking comments/notes as unavailable.

## v1.0.2 — Structured JSON Return Hotfix

### Added

- `ReadSession::toStructuredJson()` now explicitly returns structured workbook JSON as a string for API responses, direct printing, or assignment to a variable.
- `ReadSession::saveStructuredJson()` remains file-save only and internally reuses `toStructuredJson()`.
- `Writer\JsonWriter::payloadToString()` for encoding prepared payloads without saving files.
- Smoke coverage for compact/non-pretty structured JSON output and saved structured JSON output.

### Improved

- Structured JSON now respects JSON writer options such as `pretty`, `trailing_newline`, and `preserve_zero_fraction`.
- Removed the hardcoded Composer `version` field so Packagist uses Git tags correctly.

### Usage

```php
$json = MnbExcel::read($file)->toStructuredJson([
    'header_row' => true,
]);

header('Content-Type: application/json; charset=UTF-8');
echo $json;
```

## v1.0.0 — Structured Output Completion

### Added

- Improved Excel/CSV/JSON read output structure through `toStructuredArray()`.
- Structured JSON output through `toStructuredJson()` and `saveStructuredJson()`.
- Source metadata in structured output: file, format, and size.
- Sheet metadata in structured output: selected sheet, resolved sheet details, dimension, and state when available.
- Header metadata with original header text, normalized key, column index, and Excel column letter.
- Data rows with original Excel/CSV row numbers and nested `values` payload.
- Summary metrics: source rows, processed rows, data rows, columns, header row number, and skipped empty rows.
- Read-time header rename support through `rename_headers`, `header_aliases`, and `header_map`.
- Header case modes: `snake`, `preserve`, `none`, `lower`, and `camel`.
- Unicode-aware header normalization for non-English headers.
- New smoke test: `tests/StructuredOutputSmokeTest.php`.
- Added `composer.json` to the release package metadata.

### Notes

- Version remains `1.0.0`; this is part of first public release completion.
- `toArray()` remains unchanged for backward compatibility.

## v1.0.0 — JSON Import and Conversion Completion

### Added

- JSON to XLSX/CSV/XML/SQL workflows through `MnbExcel::fromJson()`.
- JSON read session through `MnbExcel::readJson()`.
- New `Reader\JsonReader` and `Support\JsonArrayNormalizer`.
- Support for list-of-objects JSON, single-object JSON, `{"rows": [...]}`, sheet maps, `{"sheets": {...}}`, and `{"sheets": [{"name": "...", "rows": [...]}]}`.
- Multi-sheet JSON to workbook conversion.
- Nested JSON object flattening into column keys such as `address.city`.
- Mixed-key normalization across JSON rows.
- Big integer preservation using `JSON_BIGINT_AS_STRING`.
- JSON sheet listing through read-session `sheetNames()`.
- JSON inspection through read-session `inspect()`.
- New smoke test: `tests/JsonImportSmokeTest.php`.

### Notes

- Version remains `1.0.0`; this is still part of the first public release completion scope.

## v1.0.0 — JSON and XML Conversion Completion

### Added

- Excel/CSV selected sheet to JSON through read-session `toJson()` and `saveJson()`.
- Excel/CSV selected sheet to XML through read-session `toXml()` and `saveXml()`.
- Array/workbook JSON export through builder `toJson()` and `saveJson()`.
- Array/workbook XML export through builder `toXml()` and `saveXml()`.
- Direct `.json` and `.xml` support in `save()`.
- Safe export filename support for `json` and `xml` in `saveSafe()`.
- New writers: `Writer\JsonWriter` and `Writer\XmlWriter`.
- New example: `examples/excel_to_json_xml.php`.
- New smoke test: `tests/JsonXmlExportSmokeTest.php`.

### Notes

- Version remains `1.0.0`; this is still part of the first public release completion scope.

## v1.0.0 — Public Release Hardening and Accuracy Completion

### Added

- Release hardening pass kept under the first public `v1.0.0` version.
- `MnbExcel::releaseReadiness()` and `Support\ReleaseReadiness` for local Composer/package readiness checks.
- New validation rules: `nullable`, `required_if`, `url`, `alpha`, `alpha_num`, `phone_basic`, `starts_with`, `ends_with`, and `unique_in_file`.
- Detailed validation error metadata through `error_details` with row, column, value, rule, and message.
- Hidden column skip support for XLSX reads through `include_hidden_columns => false`.
- XLSX read guards for `max_rows` and `max_cells`.
- Extra XLSX inspection warnings for macros, external links, comments/notes, calc chain, and pivot table parts.
- Automatic XLSX text/date/number column format styles when using `textColumns()`, `dateColumns()`, and `numberColumns()`.
- `tests/PublicReleaseHardeningSmokeTest.php`.
- Composer scripts: `test`, `test:syntax`, and `test:smoke`.

### Improved

- Original row number preservation is more accurate when XLSX rows are sparse or filtered by hidden/empty row logic.
- Column style resolution now prioritizes associative array keys before treating alphabetic names as Excel column letters.
- Import duplicate detection defaults to `_mnb_original_row_number` when available.
- README and roadmap now describe the first public release scope instead of internal `v0.x` planning.

### Notes

- Version remains `1.0.0`; this is a same-release hardening pass, not a version bump.
- Large-file streaming remains out of scope for this public release.

## v1.0.0 — Workbook Structured Output Completion

### Improved

- `toStructuredArray()` now returns workbook-level output by default with `source`, `sheets`, global `summary`, `warnings`, and `errors`.
- Added `toStructuredWorkbookArray()` for explicit workbook-level structured output.
- Added `toStructuredSheetArray()` for selected-sheet-only structured output.
- Structured sheets now include `sheetname`, `headers`, `columns`, `rows`, per-sheet `summary`, warnings, and errors.
- Structured JSON export now uses the workbook-level structure by default.

### Notes

- Version remains `1.0.0`.
- Existing `toArray()` behavior is unchanged.

## v1.0.0 — First Public Release with Formula and Cell Safety

### Added

- First public Composer package version for `mnb/mnb-phpexcel`.
- Explicit typed cell helpers through `MnbExcel::text()`, `number()`, `bool()`, `date()`, `formula()`, and `blank()`.
- `CellValue` object for array-first typed cells.
- Safe explicit formula writing through `formulaPolicy('safe')`.
- Formula blocking policy through `formulaPolicy('block')` and full override through `formulaPolicy('allow')`.
- Formula guard for unsafe formulas including external workbook links, URLs, and risky functions such as `HYPERLINK`, `WEBSERVICE`, `CALL`, `EXEC`, and `RTD`.
- Cell safety configuration through `cellSafety()`.
- `maxCellTextLength()` for Excel's 32,767-character cell text limit.
- `controlCharPolicy()` for XML-invalid control characters.
- `MnbExcel::scanCells()` and builder `scanCellSafety()` for pre-export safety reports.
- Long numeric text detection to avoid Excel precision loss.
- Advanced cell array support using `['type' => 'formula', 'formula' => 'SUM(A1:A2)']` and related typed definitions.
- New example: `examples/formula_cell_safety.php`.
- New smoke test: `tests/FormulaCellSafetySmokeTest.php`.
- Advanced CSV/text encoding auto-detection through `MnbExcel::detectEncoding()`.
- New `EncodingDetector` helper with BOM detection, UTF-8 validation, UTF-16/UTF-32 NUL-pattern detection, Windows-1252 heuristic, and ISO-8859-1 fallback.
- CSV reader now defaults to `encoding => auto` and can normalize UTF-16/UTF-32 CSV files to temporary UTF-8 before parsing.
- CSV writer now converts full CSV lines for non-UTF-8 encodings, including delimiters and line endings.
- UTF-16/UTF-32 BOM writing support for encoded CSV exports.
- New example: `examples/encoding_auto_detection.php`.
- New smoke test: `tests/EncodingAutoDetectionSmokeTest.php`.

### Improved

- Raw formula-like text remains escaped by default.
- CSV writing now handles typed cell values safely.
- CSV reading is more resilient for legacy files from Excel, Windows systems, and UTF-16 exports.
- XLSX writing now supports explicit formula cells without requiring global formula-like text escaping to be disabled.
- Cell text is cleaned for XLSX XML safety before export.

### Notes

- This release intentionally stops internal `v0.x` bumps and marks the package as the first public release.
- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.

## v0.7.0 — Small File CSV and Locale Upgrade

### Added

- CSV dialect presets through `csvDialect()` and `MnbExcel::csvDialectOptions()`: `excel`, `semicolon`, `excel_tab`, `unix`, and `pipe`.
- CSV delimiter/enclosure/escape configuration through `csvDelimiter()`, `csvEnclosure()`, and `csvEscape()`.
- UTF-8 BOM control through `csvBom()`.
- CSV output encoding option through `csvEncoding()`.
- CSV injection policy modes through `csvInjectionPolicy()`: `escape`, `tab_escape`, `strip`, `block`, and `none`.
- CSV read defaults through `MnbExcel::readCsv($path, $options)`.
- CSV reader encoding normalization options.
- Locale-aware number parsing for CSV/array reads through `locale`, `number_columns`, and `integer_columns`.
- Locale-aware date parsing through `date_columns`, `date_input_formats`, and `date_output_format`.
- Empty column cleanup through `skip_empty_columns` / `drop_empty_columns`.
- New example: `examples/csv_locale.php`.
- New smoke test: `tests/CsvLocaleSmokeTest.php`.

### Improved

- CSV writer now passes delimiter, enclosure, escape, and line ending explicitly.
- CSV exports can be optimized for regional data where comma is used as a decimal separator.
- Formula-like values can now be blocked or stripped for safer CSV download workflows.
- Read-session normalization now supports value trimming, empty-string-to-null conversion, locale parsing, and empty-column cleanup.

### Notes

- This upgrade is still small-file focused. Large CSV/XLSX streaming remains planned for a later engine layer.
- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.

## v0.6.0 — Small File Export Polish Upgrade

### Added

- Auto-width estimation through `autoWidth()`.
- More export format helpers: `formatColumns()`, `integerColumns()`, `decimalColumns()`, `datetimeStyleColumns()`, and `textStyleColumns()`.
- Conditional row styling through `conditionalRowStyle()`.
- Workbook metadata API through `metadata()`, `creator()`, and `company()`.
- Safe export helpers: `saveSafe()`, `MnbExcel::safeFileName()`, and `MnbExcel::safeSheetName()`.
- Import summary workbook builder through `MnbExcel::fromImportSummary()`.
- Import summary sheet builder through `withImportSummarySheet()`.
- New built-in styles: `mnb.row.success`, `mnb.row.warning`, `mnb.row.danger`, `mnb.row.muted`, `mnb.integer`, `mnb.decimal`, `mnb.datetime`, and `mnb.text`.
- New examples: `examples/export_polish.php` and `examples/import_summary_report.php`.
- New smoke test: `tests/ExportPolishSmokeTest.php`.

### Improved

- XLSX writer now writes workbook core/app metadata properties.
- Manual column widths override auto-width estimates.
- Conditional row styles can use the original associative input row while keeping the array-first API simple.
- Failed-row reports can now include an additional import summary sheet in the same workbook.

### Notes

- This remains a small-file export polish layer. Large-file streaming is still planned later.
- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.

## v0.5.0 — Small File Import Quality Upgrade

### Added

- Import preview through `MnbExcel::previewImport()` and read-session `previewImport()`.
- Column mapping suggestions through `MnbExcel::suggestColumnMap()`.
- Required column detection.
- Unexpected column detection and strict column mode.
- Duplicate row detection through `MnbExcel::duplicateRows()` and read-session `duplicateRows()`.
- Original row-number preservation when reading with `preserve_original_row_numbers => true`.
- Import-focused validation alias: `MnbExcel::validateImport()` and read-session `validateImport()`.
- Validation options: `row_number_key`, `strict_columns`, `allowed_columns`, and `duplicate_by`.
- SQL dry-run import planning through `dryRunImportToSql()` and `dry_run => true`.
- Failed row exports now work cleanly with original row numbers.
- New example: `examples/import_quality.php`.
- New smoke test: `tests/ImportQualitySmokeTest.php`.

### Improved

- `WorkbookBuilder::importToSql()` now preserves associative source rows for SQL imports instead of importing normalized numeric worksheet rows.
- SQL import result now includes planned rows, planned batches, resolved columns, and batch size in dry-run mode.
- Import analysis reports per-column filled/empty/type statistics.

### Notes

- This version is still small-file/import-screen focused. Large-file streaming/chunk import remains planned for v2.0.0.
- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.

## v0.4.0 — Small File Report Builder

### Added

- Report-ready shortcut: `MnbExcel::report($rows, $template)`.
- Title rows through `title()` and custom pre-header rows through `titleRow()`.
- Summary rows through `summaryRows()`.
- Footer rows through `footerRows()`.
- Named style registry through `namedStyle()`.
- Built-in report styles: `mnb.title`, `mnb.subtitle`, `mnb.header`, `mnb.header.blue`, `mnb.header.green`, `mnb.summary`, `mnb.footer`, `mnb.currency`, `mnb.percent`, and `mnb.date`.
- Reusable report templates: `simple`, `business`, and `finance`.
- Per-column styles through `columnStyle()` and `columnStyles()`.
- Per-row styles through `rowStyle()` and `rowStyles()`.
- Per-cell styles through `cellStyle()` and `cellStyles()`.
- Per-range styles through `rangeStyle()` and `rangeStyles()`.
- Currency column formatting through `currencyColumns()`.
- Percentage column formatting through `percentageColumns()`.
- Date-style column formatting through `dateStyleColumns()`.
- Dynamic header row position for reports with title rows.
- Auto-filter and freeze-pane support when the header is not row 1.
- New example: `examples/report_builder.php`.
- New smoke test: `tests/ReportBuilderSmokeTest.php`.

### Improved

- XLSX style generation now supports multiple unique styles instead of a single header style.
- Report title/footer rows can be merged automatically across the report width.
- Column formats can combine with row styles, useful for summary rows that still need currency/percentage formatting.

### Notes

- This remains a small-file PHPExcel-style report builder. Large-file streaming/chunk reports are still planned later.
- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.

## v0.3.0 — Small File Reader Accuracy Upgrade

### Added

- Relationship-based worksheet lookup using `xl/workbook.xml` and `xl/_rels/workbook.xml.rels`.
- Sheet selection by name: `MnbExcel::read($file)->sheet('Students')`.
- Sheet name listing through `MnbExcel::sheetNames($file)` and read-session `sheetNames()`.
- XLSX inspection through `MnbExcel::inspect($file)` and read-session `inspect()`.
- Date style detection from `xl/styles.xml`.
- Excel serial date conversion with 1900 and 1904 workbook date-system support.
- Hidden row skipping option: `include_hidden_rows => false`.
- Hidden sheet, hidden row, and hidden column reporting through the inspector.
- Corrupted XLSX checks for missing workbook, relationship, worksheet, shared string, and style parts.
- Formula cell option: `formula_cells => formula|cached_value`.
- New reader accuracy example: `examples/xlsx_reader_accuracy.php`.
- XLSX reader feature smoke test with safe skip when XLSX extensions are unavailable.

### Improved

- XLSX reader no longer assumes `xl/worksheets/sheetN.xml` directly.
- Real-world XLSX compatibility improved for files with reordered sheets, custom sheet paths, hidden sheets, and workbook relationships.
- Date-like formatted numeric cells now return formatted date strings by default.

### Notes

- XLSX runtime still requires `ext-zip` and `ext-xmlreader`.
- This version remains focused on small/medium PHPExcel-style workflows, not large-file streaming yet.

## v0.2.0 — Small File XLSX Feature Completion

### Added

- Header style improvements: font color, fill color, alignment, border, font size.
- Merge cell support through `mergeCells()`.
- Column width support through `columnWidth()` and `columnWidths()`.
- Row height support through `rowHeight()` and `rowHeights()`.
- Image/logo insertion through `addImage()` for PNG, JPEG, and GIF files.
- Validation helper `MnbExcel::validateArray()`.
- Failed rows report builder `MnbExcel::fromFailedRows()`.
- More validation rules: `integer`, `string`, `boolean`, `date`, `min`, `max`, `length`, `in`, and `regex`.
- New examples for XLSX small-file features and validation error reports.

### Improved

- XLSX writer now supports worksheet relationships and drawing parts for images.
- Header styling now supports more professional report exports.
- README updated with v0.2.0 examples.

### Notes

- XLSX runtime requires `ext-zip` and `ext-xmlreader`.
- Large-file streaming/chunk processing is still planned for a later engine layer.

## v0.1.0 — Foundation Build

### Added

- Composer package structure.
- `MnbExcel::fromArray()`.
- `MnbExcel::fromWorkbookArray()`.
- `MnbExcel::read()`.
- `MnbExcel::readCsv()`.
- `MnbExcel::fromSql()`.
- Array to XLSX writer.
- XLSX to array reader.
- Array to CSV writer.
- CSV to array reader.
- SQL import helper.
- Header row support.
- Custom column heading support.
- Text/date/number column options.
- Formula-like text escaping.
- Basic validation class.
- Initial examples, README, roadmap, and license.

## v1.0.3 - Structured XML Return Hotfix

- Added `ReadSession::toStructuredXml()` to return structured workbook XML as a string.
- Added `ReadSession::saveStructuredXml()` to save structured workbook XML to a file.
- Added `XmlWriter::payloadToString()` for direct XML encoding of prepared structured payloads.
- Added `examples/structured_xml_response.php`.
- Updated structured output smoke test to cover XML string and file-save workflows.


## LargeExcelPreflightAndMethodAdvisor

- Added `MnbExcel::analyzeXlsxForImport()` for safe XLSX preflight without loading workbook rows.
- Added `MnbExcel::recommendImportMethod()` and `MnbExcel::recommendImportMethodFromProfile()`.
- Added `MnbExcel::autoImportPlan()` to return profile + advisor route + chunk size.
- Added `MnbExcel::largeRead()` streaming read session for chunk/row callbacks.
- Added method decision matrix for tiny/small/normal/medium/large/very-large Excel files.
- Added shared-string memory/SQLite cache strategy for large streaming imports.
- Added memory guard and soft execution-time guard for chunk imports.
- Added smoke tests and example for large-file method shifting.

## Unreleased — Large Excel Database Import Engine

- Added `MnbExcel::largeImportToSql()` for streaming XLSX-to-database imports.
- Added `MnbExcel::largeImportManifestPath()` for deterministic resumable manifest paths.
- Added `LargeExcelDatabaseImportEngine` with chunk validation and PDO batch inserts.
- Added JSON progress manifest with rows scanned, valid rows, failed rows, inserted rows, chunks completed, status, and last processed Excel row.
- Added checkpoint/resume support for CLI/queue/shared-hosting repeated jobs.
- Added failed-row CSV export containing Excel row number, validation errors, and source row JSON.
- Added transaction-per-chunk database safety.
- Added smoke test and example for large database import flow.



## Unreleased — LargeExcelImportStabilityUpgrade

- Added `MnbExcel::environmentAlert()` and `MnbExcel::environmentAlertMessage()` for clear missing-extension alerts covering `ext-zip`, `ext-xmlreader`, and `pdo_sqlite`.
- Added `LargeXlsxReadSession::preserveNumericStrings()` for large streaming imports where IDs/account numbers must not be converted to PHP integers/floats.
- Added `LargeXlsxReadSession::convertDates()` to convert Excel numeric date serials using `styles.xml` date formats.
- Added idempotent database import support through `idempotent`, `duplicate_strategy`, and `unique_by` options.
- Added duplicate strategies: `fail`, `skip`, and `update` with driver-specific SQL for SQLite/MySQL/PostgreSQL where supported.
- Added human-readable failed-row CSV format with original columns beside validation errors.
- Added `MnbExcel::largeImportWorkbookToSql()` for all-sheet workbook import orchestration.
- Added stronger local XLSX runtime smoke tests for environment alerts, numeric preservation, date conversion, duplicate handling, failed rows, and all-sheet import.

## RealApplicationAndPluginIntegrationLayer

- Added import profiles for reusable application import configurations.
- Added storage path manager for temp/upload/manifest/failed-row/report paths.
- Added import status reader and resumeImport() runner for CLI/queue/cron workflows.
- Added upload safety validator for XLSX/CSV imports.
- Added event hooks around large import lifecycle.
- Added plugin manager and plugin registry for profiles, validators, transformers, and events.
- Added row transformer pipeline for large reads and large SQL imports.
- Added custom validator registry integrated with ArrayValidator.
- Added optional logger bridge for PSR-3-like loggers or callable loggers.
- Added framework-neutral adapter interfaces for future framework packages.

## Benchmark and compatibility proof suite

- Added benchmark plan APIs for 100k / 500k / 1M row release testing.
- Added internal large XLSX writer benchmark runner.
- Added real-world XLSX fixture compatibility suite scaffolding for Excel, LibreOffice, Google Sheets, and WPS files.
- Added MySQL/PostgreSQL integration readiness checker.
- Added Packagist-ready comparison guide.
- Added advanced workbook capability matrix covering deep editing, multi-format support, advanced styles, formula calculation, and complex workbook manipulation.
- Added Slim and CodeIgniter-style integration examples.
- Added import dashboard UI example.
- Added release benchmark/performance table template.
