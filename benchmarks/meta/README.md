# Meta Compatibility & Benchmark Suite

This suite measures metadata correctness, compatibility, time, and peak-memory behavior.

## Included
- generator targeting 100+ XLSX/XLS/CSV fixtures when the runtime can write all formats
- CSV dialect/BOM/multiline cases
- corrupt/truncated cases
- QUICK and FULL MNB metadata benchmarks
- JSON + CSV machine-readable results
- optional PhpSpreadsheet comparison runner
- fixture manifest for Excel, LibreOffice, Google Sheets, encryption, macros, pivots and security cases

## Run
`php benchmarks/meta/generate_fixtures.php`

`php benchmarks/meta/run_benchmarks.php`

Optional:
`php benchmarks/meta/compare_phpspreadsheet.php /path/to/vendor/autoload.php`

For real Excel/LibreOffice/Google Sheets compatibility, add exported fixtures to a directory and pass it:
`php benchmarks/meta/run_benchmarks.php /path/to/fixtures`

The repository does not fabricate Microsoft Excel/Google Sheets-produced fixtures. Those must be genuine exported files.

## Metrics
Per file: format, bytes, success/failure, profile, elapsed milliseconds, peak-memory delta, detected format and sheet count.

## Release gate
Recommended gate: no regression in known metadata fields, all valid fixtures readable, corrupt fixtures fail safely,
round-trip metadata preserved for writable formats, and benchmark deltas reviewed before release.
