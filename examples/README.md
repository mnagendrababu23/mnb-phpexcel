# MNB PHPExcel Examples

Run examples after installing dependencies with Composer. Most file-reading examples accept the source path as their first CLI argument.

## Structured output

- `structured_array.php` — workbook-level `toStructuredArray()` output.
- `structured_sheet_array.php` — selected-sheet structure with headers, columns, summary, warnings, and metadata.
- `structured_flat_rows.php` — flat records with original source row numbers.
- `structured_output_files.php` — save the same structured payload as JSON and XML.
- `structured_json_response.php` — return structured JSON from an HTTP endpoint.
- `structured_xml_response.php` — return structured XML from an HTTP endpoint.

## Reading and large files

- `xlsx_file_info.php` — file size, package totals, document properties, sheet count, and sheet names without loading workbook rows.
- `xlsx_sheet_info.php` — fast worksheet dimensions plus optional streamed physical/filled row and cell counts.
- `xlsx_row_count.php` — filled, physical, last-used, and declared row counts for one or every worksheet.
- `read_to_array.php` — normal row-array reading.
- `read_header_projection.php` — detect headers, limit ranges, and project worksheet columns.
- `large_xlsx_chunk_reader.php` — preflight a large XLSX and process projected columns in chunks.
- `large_excel_preflight_method_advisor.php` — detailed method selection guidance.
- `large_excel_database_import.php` — database import with bounded-memory processing.
- `large_xlsx_streaming_writer.php` — stream a large XLSX export.

## Imports, validation, and application workflows

- `import_quality.php` — preview, map, validate, and export failures.
- `validation_error_report.php` — generate a correction workbook.
- `database_connection_config.php` — framework-neutral database configuration.
- `real_application_plugin_integration.php` — real application workflow integration.
- `frameworks/` — framework route/controller examples.
- `ui/` — import dashboard example.

## Security and verification

- `formula_cell_safety.php` — formula and untrusted-cell safety.
- `xlsx_integrity_validation.php` — validate XLSX package integrity.
- `xlsx_compatibility_verification.php` — compatibility verification workflow.
- `public_release_hardening.php` — release and deployment checks.

The remaining examples cover CSV/JSON/XML conversion, reporting, formatting, metadata, comments, hyperlinks, and SQL exports.
