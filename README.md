# MNB PHPExcel

**MNB PHPExcel** is an **application-ready PHP Excel engine** with a rich small/normal workbook mode plus separate streaming modes for large XLSX import and export.

```text
PHP Array ⇄ XLSX ⇄ CSV ⇄ JSON ⇄ XML ⇄ SQL
```

Current development release:

```text
v1.4.0 — Typed Domain Import APIs
```

This package keeps the rich workbook reader/writer optimized for **small and normal files**. Large XLSX files are handled through separate **preflight/advisor + streaming import/export** layers so applications do not load huge workbooks into PHP arrays.

## Install

Backward-compatible full package:

```bash
composer require mnb/mnb-phpexcel
```

Lightweight format packages can be published from the generated release trees:

```bash
# XLSX + shared core only
composer require mnb/mnb-phpexcel-xlsx

# XML + shared core only
composer require mnb/mnb-phpexcel-xml

# CSV and JSON only
composer require mnb/mnb-phpexcel-csv mnb/mnb-phpexcel-json

# Complete split family
composer require mnb/mnb-phpexcel-all
```

Format-specific modular entry points avoid loading the legacy all-in-one facade:

```php
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;

$options = ReaderOptions::defaults()
    ->withRange(2, 50000, 'A', 'F')
    ->withMode('streaming');

foreach (Xlsx::read('orders.xlsx', $options)->rows() as $row) {
    // process one row
}
```

See `docs/MODULAR_PACKAGES.md` for package ownership and release instructions.

## v1.4 typed domain import APIs

The database/application packages now provide first-class imports for users, products, orders, inventory, students, attendance, marks, contacts, locations, blog posts, media paths, and categories. Each preset includes canonical columns, common header aliases, validation, defaults, normalization, duplicate keys, template metadata, and cross-field rules.

```php
use Mnb\PHPExcel\MnbExcel;

$preview = MnbExcel::previewDomainImport('products', 'products.xlsx', [
    'sheet' => 'Products',
    'limit' => 25,
]);

$result = MnbExcel::importProducts('products.xlsx', $pdo, 'products', [
    'duplicate_strategy' => 'update',
    'unique_by' => ['sku'],
    'batch_size' => 1000,
    'failed_rows_csv' => __DIR__ . '/failed-products.csv',
]);
```

Generate a validated Excel template from the same preset:

```php
MnbExcel::domainImportTemplate('products')->save('product-import-template.xlsx');
```

The convenience layer remains configurable: aliases, mappings, rules, defaults, normalizers, transformers, unique keys, duplicate policies, progress callbacks, and row-error callbacks can be overridden. See `docs/DOMAIN_IMPORTS_1_4.md`.

## v1.3 full Excel APIs

The native XLSX module now includes direct cell/range access, complete style metadata, typed rich text, embedded-image extraction, semantic header detection, validated import templates, arbitrary freeze panes, advanced native filters, native conditional formatting, native charts, and template-driven pivot-table rebinding.

```php
$sheet = Xlsx::read('report.xlsx')->sheet('Sales');
$details = $sheet->cellDetails('D12');
$images = $sheet->extractImages(__DIR__ . '/images');

$rows = Xlsx::read('upload.xlsx')->autoDetectHeader()->toArray();
```

See `docs/EXCEL_FULL_CAPABILITIES_1_3.md` for complete examples and exact compatibility boundaries.

Required for core package features:

```text
PHP >= 8.1
ext-json
```

Required for XLSX read/write:

```text
ext-zip
ext-xmlreader
```

Required only for SQL import/export helpers:

```text
ext-pdo
```

Recommended for better legacy CSV encoding support:

```text
ext-mbstring
ext-iconv
```


## Large XLSX streaming export

Use the large writer when exporting database-sized reports. It streams rows to worksheet XML and uses inline strings, so the full workbook is not kept in memory.

```php
use Mnb\PHPExcel\MnbExcel;

MnbExcel::largeExport($rowGenerator)
    ->withHeader()
    ->sheetName('Orders')
    ->formatColumn('amount', 'currency')
    ->autoSplitSheets(true)
    ->progress(function (array $state) {
        // update CLI/admin progress
    }, 5000)
    ->save('orders.xlsx');
```

Export from a PDO cursor without `fetchAll()`:

```php
MnbExcel::largeExportFromSql(__DIR__ . '/.env', 'SELECT id, name, amount FROM orders')
    ->formatColumn('amount', 'decimal')
    ->save('orders.xlsx');
```

CSV ZIP fallback is available for ultra-large table exports:

```php
MnbExcel::largeExport($rowGenerator)
    ->withHeader()
    ->csvRowsPerFile(500000)
    ->saveCsvZip('orders-csv-parts.zip');
```

## Import dashboard helper

Turn a resumable import manifest into a frontend/admin response:

```php
$response = MnbExcel::importDashboard($manifestPath, [
    'download_base_url' => '/admin/imports/downloads',
]);

// status, progress percent, failed-row download info, resume action, safe messages
```

## 1.2 reader API

- Real line-by-line NDJSON reading.
- Streaming parser for top-level JSON arrays.
- Forward-only XML reading with optional schema mapping.
- Shared in-memory/disk XLSX string provider.
- Source-level column projection.
- One `ReadSession` for normal and streaming modes.
- Immutable `ReaderOptions`, typed progress/state objects, and row error policies.
- XLSX merged-cell expansion/metadata and formula-plus-cached-result objects.
- Native ODS reading and an optional legacy XLS adapter.
- Instance-scoped custom reader registration.

```php
$session = MnbExcel::read('orders.xlsx')
    ->range(2, 100000, 'A', 'H')
    ->projectColumns(['order_id', 'amount'])
    ->streaming()
    ->onRowError('collect')
    ->onProgress($progressCallback, 1000);

foreach ($session->chunks(1000) as $chunk) {
    // Same API for normal and low-memory reads.
}

$errors = $session->rowErrors();
```

## Main idea

Most PHP applications already use arrays from database queries, APIs, forms, and admin panels. MNB PHPExcel makes those arrays easy to export, import, validate, inspect, and convert.

```php
use Mnb\PHPExcel\MnbExcel;

MnbExcel::fromArray($rows)->save('students.xlsx');
$rows = MnbExcel::read('students.xlsx')->toArray();
```

## Core features

### Array-first API

- `MnbExcel::fromArray()`
- `MnbExcel::fromWorkbookArray()`
- `MnbExcel::read()` (auto-detects XLSX/CSV/JSON/XML)
- `MnbExcel::readXlsx()`
- `MnbExcel::readCsv()`
- `MnbExcel::readXml()`
- `MnbExcel::fromJson()`
- `MnbExcel::readJson()`
- `MnbExcel::fromSql()`
- `toArray()`
- `toStructuredArray()`
- `toStructuredJson()` / `saveStructuredJson()`
- `importToSql()`
- `toJson()` / `saveJson()`
- `toXml()` / `saveXml()`

### XLSX support

- Create XLSX files from arrays.
- Read XLSX files into arrays.
- Multiple worksheets.
- Sheet selection by index or name.
- Sheet name listing.
- Relationship-based sheet lookup using `workbook.xml.rels`.
- Shared strings and inline strings.
- Number, text, boolean, blank, error, date, and formula cells.
- Date style detection.
- Excel 1900 and 1904 date systems.
- Hidden sheet/row/column inspection.
- Hidden row and hidden column skip options.
- Corrupted XLSX structure checks.
- Save-time XLSX integrity validation to prevent handing users broken workbooks.
- Formula cached value or formula text read mode.
- Rich text run metadata for shared strings, inline strings, and comments.
- Comments/notes and hyperlinks exposed through `sheetMetadata()` and structured output `metadata`.
- Programmatic XLSX hyperlink writing through `hyperlink()` / `hyperlinks()`.
- Programmatic XLSX legacy note/comment writing through `comment()` / `comments()`.
- Environment capability diagnostics through `MnbExcel::environmentCheck()`.
- XLSX compatibility verification harness through `MnbExcel::verifyXlsxCompatibility()`.
- Advanced object preservation from an existing XLSX template through `preserveAdvancedObjectsFrom()`.
- Public XLSX validation API through `MnbExcel::validateXlsx()` and `MnbExcel::assertValidXlsx()`.
- Upload validation with actual-size verification, MIME inspection, ZIP path-traversal detection, entry-count/uncompressed-size limits, compression-ratio checks, encrypted-content rejection, and optional macro/external-link rejection.

### CSV support

- CSV read/write.
- Dialect presets: `excel`, `semicolon`, `excel_tab`, `unix`, `pipe`.
- Custom delimiter, enclosure, escape, and line ending.
- UTF-8 BOM option.
- Advanced encoding auto-detection.
- UTF-8, UTF-8 BOM, UTF-16LE, UTF-16BE, UTF-32LE, UTF-32BE, Windows-1252, and ISO-8859-1 support.
- Locale-aware number parsing.
- Locale-aware date parsing.
- CSV injection policy modes: `escape`, `tab_escape`, `strip`, `block`, `none`.

### JSON and XML conversion

- Convert XLSX/CSV selected sheet rows to JSON.
- Convert XLSX/CSV selected sheet rows to XML.
- Convert JSON files to XLSX/CSV/XML/SQL through the array-first engine.
- Export array workbooks directly as JSON or XML.
- Preserves associative source keys as JSON object/XML field nodes by default (`preserve_associative_rows => false` restores indexed output).
- Supports list rows as array/cell nodes.
- Safe XML tag-name normalization.
- XML escaping for special characters such as `&`, `<`, `>`, quotes, and apostrophes.
- JSON options for pretty output, Unicode preservation, slash preservation, and zero-fraction preservation.
- JSON import supports list-of-objects, single object, `{"rows": [...]}`, sheet maps, and `{"sheets": ...}` workbook formats.
- Nested JSON objects can be flattened into columns such as `address.city`.
- Large JSON integers are decoded as strings to avoid precision loss.



### JSON to Excel

Simple list-of-objects JSON:

```json
[
  {"student_id":"000123", "name":"Ravi", "address":{"city":"Hyderabad"}},
  {"student_id":"000124", "name":"Sita", "address":{"city":"Vijayawada"}}
]
```

Convert JSON to XLSX:

```php
MnbExcel::fromJson('students.json')
    ->textColumns(['student_id'])
    ->autoWidth()
    ->freezeHeader()
    ->autoFilter()
    ->save('students.xlsx');
```

Workbook-style JSON becomes multiple sheets:

```json
{
  "Students": [
    {"name":"Ravi", "email":"ravi@example.com"}
  ],
  "Teachers": [
    {"name":"Kumar", "subject":"Maths"}
  ]
}
```

```php
MnbExcel::fromJson('school.json')
    ->save('school.xlsx');
```

Read JSON to array:

```php
$rows = MnbExcel::readJson('students.json')
    ->toArray(['header_row' => true]);
```

JSON to CSV/XML:

```php
MnbExcel::fromJson('students.json')->save('students.csv');
MnbExcel::fromJson('students.json')->saveXml('students.xml', [
    'root' => 'students',
    'row' => 'student',
]);
```


### Improved workbook output structure

For normal imports, `toArray()` returns only row data. For API/debug/import workflows, use `toStructuredArray()` to get workbook-level output with all sheets.

```php
$structured = MnbExcel::read('Heilberufe.xlsx')
    ->toStructuredArray([
        'header_row' => true,
        'skip_empty_rows' => true,
        'include_hidden_rows' => false,
        'include_hidden_columns' => false,
        'rename_headers' => [
            'affilication' => 'affiliation',
        ],
    ]);
```

Workbook output shape:

```php
[
    'status' => 'ok',
    'source' => [
        'file' => 'Heilberufe.xlsx',
        'format' => 'xlsx',
        'size_bytes' => 9150,
    ],
    'sheets' => [
        0 => [
            'index' => 1,
            'sheetname' => 'Worksheet',
            'name' => 'Worksheet',
            'state' => 'visible',
            'dimension' => 'A1:F27',
            'headers' => ['author_name', 'email', 'title', 'affiliation', 'article_type', 'link'],
            'columns' => [
                [
                    'index' => 1,
                    'letter' => 'A',
                    'header' => 'author_name',
                    'original_header' => 'author_name',
                    'key' => 'author_name',
                    'renamed' => false,
                ],
            ],
            'rows' => [
                [
                    'index' => 0,
                    'row_number' => 2,
                    'values' => [
                        'author_name' => 'Peter Heinz',
                        'email' => 'heinz@augeninfo.de',
                    ],
                ],
            ],
            'summary' => [
                'source_rows' => 27,
                'processed_rows' => 27,
                'data_rows' => 26,
                'columns' => 6,
                'header_row_number' => 1,
            ],
            'warnings' => [],
            'errors' => [],
        ],
    ],
    'summary' => [
        'sheet_count' => 1,
        'source_rows' => 27,
        'processed_rows' => 27,
        'data_rows' => 26,
        'columns_total' => 6,
        'skipped_empty_rows' => 0,
    ],
    'warnings' => [],
    'errors' => [],
]
```

For selected-sheet-only structure, use `toStructuredSheetArray()` or pass `['structure' => 'sheet']` to `toStructuredArray()`.

```php
$sheetOnly = MnbExcel::read('Heilberufe.xlsx')
    ->sheet('Worksheet')
    ->toStructuredSheetArray([
        'header_row' => true,
    ]);
```

Return workbook structured output as a JSON string without saving:

```php
$json = MnbExcel::read('Heilberufe.xlsx')
    ->toStructuredJson([
        'header_row' => true,
        'rename_headers' => ['affilication' => 'affiliation'],
    ]);

// assign to variable, print, return from controller, or send as API response
echo $json;
```

Return compact JSON for API responses:

```php
header('Content-Type: application/json; charset=UTF-8');

echo MnbExcel::read('Heilberufe.xlsx')
    ->toStructuredJson([
        'header_row' => true,
        'skip_empty_rows' => true,
    ], [
        'pretty' => false,
        'trailing_newline' => false,
    ]);

exit;
```

Save workbook structured output directly as JSON:

```php
MnbExcel::read('Heilberufe.xlsx')
    ->saveStructuredJson('heilberufe-structured.json', [
        'header_row' => true,
        'rename_headers' => ['affilication' => 'affiliation'],
    ]);
```

Header options:

```php
'header_case' => 'snake',      // default: Author Name -> author_name
'header_case' => 'preserve',   // keep original header text
'header_case' => 'camel',      // Author Name -> authorName
'rename_headers' => [
    'affilication' => 'affiliation',
]
```

### Report builder

- `MnbExcel::report()`.
- Title rows.
- Subtitle/custom title rows.
- Summary rows.
- Footer rows.
- Built-in templates: `simple`, `business`, `finance`.
- Named styles.
- Header styling.
- Column, row, cell, and range styles.
- Currency, percentage, date, datetime, integer, decimal, and text formats.
- Auto-width estimation.
- Freeze panes.
- Auto filter.
- Conditional row styling.
- Workbook metadata.
- Safe file and sheet names.

### Import quality tools

- Import preview.
- Required column detection.
- Unexpected column detection.
- Strict column mode.
- Column mapping suggestions.
- Duplicate row detection.
- Original row number preservation.
- Failed rows report.
- Import summary workbook/sheet.
- SQL dry-run before insert.

### Validation

Supported rules:

```text
required
nullable
required_if
email
url
numeric
integer
string
boolean
date
alpha
alpha_num
phone_basic
min
max
length
in
regex
starts_with
ends_with
unique_in_file
```

Validation results include simple error messages and detailed metadata:

```php
[
    'row' => 7,
    'column' => 'email',
    'value' => 'wrong-email',
    'rule' => 'email',
    'message' => 'email must be a valid email.',
]
```

### Formula and cell safety

- Explicit typed cell helpers:
  - `MnbExcel::text()`
  - `MnbExcel::number()`
  - `MnbExcel::bool()`
  - `MnbExcel::date()`
  - `MnbExcel::formula()`
  - `MnbExcel::blank()`
- Formula policies:
  - `safe`
  - `block`
  - `allow`
- Dangerous formula detection for external links, URLs, and risky functions.
- Formula-like text escaping by default.
- Excel 32,767-character cell limit protection.
- Invalid XML/control character handling.
- Long numeric text detection to avoid precision loss.
- Cell safety scanner.

## Examples

### Array to XLSX

```php
use Mnb\PHPExcel\MnbExcel;

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321'],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678'],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->autoWidth()
    ->freezeHeader()
    ->autoFilter()
    ->save('students.xlsx');
```

### XLSX to array

```php
$rows = MnbExcel::read('students.xlsx')
    ->sheet('Students')
    ->toArray([
        'header_row' => true,
        'skip_empty_rows' => true,
        'preserve_original_row_numbers' => true,
        'include_hidden_rows' => false,
        'include_hidden_columns' => false,
    ]);
```

### XLSX metadata: rich text, comments, and hyperlinks

```php
$metadata = MnbExcel::read('source.xlsx')
    ->sheet('Students')
    ->sheetMetadata();

$richTextRuns = $metadata['rich_text'];
$comments = $metadata['comments'];
$hyperlinks = $metadata['hyperlinks'];
```

Structured JSON/XML also includes this metadata by default for XLSX reads:

```php
$json = MnbExcel::read('source.xlsx')->toStructuredJson([
    'header_row' => true,
]);
```

Disable metadata collection in structured output when you only need row/column payloads:

```php
$structured = MnbExcel::read('source.xlsx')->toStructuredArray([
    'header_row' => true,
    'include_cell_metadata' => false,
]);
```


### XLSX comments and hyperlinks writer

You can now create clickable hyperlinks and classic Excel notes/comments directly while generating XLSX files:

```php
MnbExcel::fromArray($rows)
    ->withHeader()
    ->hyperlink('B2', 'https://example.com', 'Open website', [
        'tooltip' => 'Open external reference',
    ])
    ->comment('C2', 'MNB PHPExcel', 'Please verify this value before sharing.')
    ->save('annotated-report.xlsx');
```

For multiple sheets, pass the target sheet name:

```php
MnbExcel::fromWorkbookArray([
    'Students' => $studentRows,
    'Summary' => $summaryRows,
])->withHeader()
    ->hyperlink('B2', 'https://example.com/students', 'Students source', ['sheet' => 'Students'])
    ->comment('A2', 'Reviewer', 'Summary checked.', ['sheet' => 'Summary'])
    ->save('school-report.xlsx');
```

The writer creates the required worksheet relationships, comments XML part, VML note drawing, and content-type declarations. The saved workbook is still checked by `XlsxIntegrityValidator` before the path is returned.

### Environment diagnostics and XLSX compatibility verification

Check whether the current PHP runtime supports the library features needed by your application:

```php
$environment = MnbExcel::environmentCheck();
$alert = MnbExcel::environmentAlert();

if (!$alert['ready']) {
    echo $alert['message'];
    // Example missing requirements: ext-zip, ext-xmlreader.
}

// CLI/admin friendly text:
echo MnbExcel::environmentAlertMessage();
```

The alert helper gives a direct warning when `ext-zip`, `ext-xmlreader`, or `pdo_sqlite` are missing. `pdo_sqlite` is not required for ordinary exports, but it is strongly recommended for very large XLSX imports with huge shared string tables because it enables disk-backed shared-string caching.

Run the compatibility verification harness before release or after changing XLSX internals:

```php
$report = MnbExcel::verifyXlsxCompatibility([
    // Optional real-world fixtures exported from Excel, LibreOffice, Google Sheets, or WPS Office.
    __DIR__ . '/fixtures/excel-comments-hyperlinks.xlsx',
    __DIR__ . '/fixtures/libreoffice-styles-formulas.xlsx',
    __DIR__ . '/fixtures/google-sheets-export.xlsx',
]);
```

Generated fixture checks currently cover:

```text
basic workbook structure
formulas
styles
merged cells
comments/notes
hyperlinks
advanced-object template preservation flow
XLSX integrity validation
```

Compatibility matrix:

| Source / workflow | Current verification |
| --- | --- |
| MNB-generated XLSX | Generated fixtures + integrity validator |
| Microsoft Excel files | Supported through optional external fixture paths |
| LibreOffice Calc files | Supported through optional external fixture paths |
| Google Sheets XLSX export | Supported through optional external fixture paths |
| WPS Office XLSX files | Supported through optional external fixture paths |
| UI-level “opens without repair warning” | Must be confirmed manually in the target desktop app |


### XLSX integrity validation / corruption prevention

XLSX exports are validated after writing and before the save call returns. If the generated package fails integrity checks, the partial file is deleted and `MnbExcelException` is thrown instead of returning a corrupted workbook path.

```php
MnbExcel::fromArray($rows)
    ->withHeader()
    ->save('safe-report.xlsx'); // validated automatically for XLSX output
```

Manual validation is also available for uploaded/generated XLSX files:

```php
$result = MnbExcel::validateXlsx('safe-report.xlsx');

if (!$result['valid']) {
    print_r($result['errors']);
}

MnbExcel::assertValidXlsx('safe-report.xlsx'); // throws on failure
```

The validator checks required XLSX package parts, relationship targets, content type coverage, XML structure, workbook relationship IDs, and worksheet `r:id` references.

For strict XML parsing during save validation, require `ext-xmlreader`:

```php
MnbExcel::fromArray($rows)
    ->withHeader()
    ->strictXlsxIntegrityValidation()
    ->save('strict-safe-report.xlsx');
```

For special/debug workflows only, save-time validation can be disabled:

```php
MnbExcel::fromArray($rows)
    ->skipXlsxIntegrityValidation()
    ->save('debug-output.xlsx');
```

### Preserve advanced XLSX objects while rewriting rows

Use a source workbook as a same-sheet-order template when you need to rewrite row data but keep advanced package parts such as comments, hyperlinks, drawings, legacy notes, tables, VML, embedded media, charts, and pivot/table package parts. Macros are never executed.

```php
MnbExcel::fromArray($newRows)
    ->withHeader()
    ->preserveAdvancedObjectsFrom('source-template.xlsx')
    ->save('updated.xlsx');
```

This is intended for small/normal files. Large-file streaming and full Excel object editing remain future engine layers.


### Numeric accuracy and reader options

Safe numeric conversion is the default. Ordinary integers and decimals become PHP numbers, while values that would overflow an integer or lose significant decimal precision remain exact strings.

```php
$rows = MnbExcel::readCsv('payments.csv', [
    'integer_columns' => ['account_id', 'quantity'],
    'number_columns' => ['amount'],
    'big_integer_mode' => 'string',       // string|float|error
    'numeric_precision' => 'safe',        // safe|native|string
    'max_safe_decimal_digits' => 15,
    'invalid_cast' => 'error',            // preserve|null|error
])->withHeaderRow()->toArray();
```

For IDs, phone numbers, ISBNs, account numbers, and high-precision decimal text, you can also preserve every XLSX numeric cell as text:

```php
$rows = MnbExcel::read('input.xlsx')->toArray([
    'preserve_numeric_strings' => true,
]);

MnbExcel::largeRead('large.xlsx')
    ->withHeader()
    ->preserveNumericStrings()
    ->chunk(1000, function (array $rows): void {
        // Numeric cells are delivered as strings.
    });
```

For XLSX writing, numeric cells are serialized with compact high-precision invariant numbers. Long numeric identifiers should still be written as text because Excel itself has numeric precision limits.

### Reader auto-detection and fluent row operations

`read()` now selects the reader from the file extension and signature. An explicit `format` option can override detection.

```php
$rows = MnbExcel::read('students.csv', [
    'delimiter' => 'auto',
])
    ->withHeaderRow()
    ->skip(10)
    ->limit(100)
    ->selectColumns(['student_id', 'name', 'email'])
    ->toArray();
```

Process rows without creating an intermediate worksheet array for CSV and normal XLSX files:

```php
MnbExcel::read('students.xlsx')
    ->sheet('Students')
    ->withHeaderRow()
    ->chunk(500, function (array $rows, array $state): void {
        // Save or validate this chunk.
    });
```

Other convenience operations:

```php
$first = MnbExcel::read('students.csv')->withHeaderRow()->first();
$count = MnbExcel::read('students.csv')->withHeaderRow()->countRows();

MnbExcel::read('students.csv')->withHeaderRow()->eachRow(
    function (array $row, int $index): bool {
        return true; // return false to stop
    }
);
```

Header positions are explicit:

```php
// Second normalized non-empty data row. withHeaderRow(2) is the same.
$rows = MnbExcel::readCsv('report.csv')->headerAtDataRow(2)->toArray();

// Exact third physical row in the source file, including leading blank/title rows.
$rows = MnbExcel::readCsv('report.csv')->headerAtPhysicalRow(3)->toArray();

// Ignore leading blank rows and use the first non-empty row.
$rows = MnbExcel::readCsv('report.csv')->firstNonEmptyRowAsHeader()->toArray();
```

Duplicate and uneven row widths can be controlled instead of silently losing values:

```php
$rows = MnbExcel::readCsv('users.csv', [
    'duplicate_headers' => 'rename', // rename|error
    'extra_columns' => 'collect',    // ignore|error|generate|collect
    'extra_columns_key' => '_extra',
    'missing_columns' => 'error',    // null|error
])->withHeaderRow()->toArray();
```

### CSV with encoding and delimiter auto-detection

```php
$rows = MnbExcel::readCsv('legacy-users.csv', [
    'encoding' => 'auto',
    'delimiter' => 'auto',
    'trim_values' => true,
    'strict_column_count' => true,
])->withHeaderRow()->toArray();
```

### JSON Lines / NDJSON reading

JSON Lines is detected from content as well as `.jsonl` / `.ndjson` suffixes, so renamed upload files still work.

```php
$rows = MnbExcel::read('events.data')
    ->withHeaderRow()
    ->toArray();

// Force a document mode when required:
$rows = MnbExcel::readJson('events.data', [
    'json_document_mode' => 'ndjson', // auto|json|ndjson
])->withHeaderRow()->toArray();
```

### XML reading

XML created by `saveXml()` can be read back directly:

```php
$rows = MnbExcel::readXml('students.xml')
    ->withHeaderRow()
    ->toArray();
```

For custom XML tags, pass matching reader options:

```php
$rows = MnbExcel::readXml('students.xml', [
    'row_tag' => 'student',
    'sheet_tag' => 'worksheet',
])->withHeaderRow()->toArray();
```

### Excel to JSON

```php
MnbExcel::read('students.xlsx')
    ->sheet('Students')
    ->saveJson('students.json', [
        'header_row' => true,
        'skip_empty_rows' => true,
    ]);
```

Array-first JSON export:

```php
MnbExcel::fromArray($rows)
    ->saveJson('students.json', ['mode' => 'rows']);

// Presentation rows can be omitted explicitly:
MnbExcel::fromArray($rows)
    ->withHeader()
    ->saveJson('students.json', ['data_only' => true]);
```

### Excel to XML

```php
MnbExcel::read('students.xlsx')
    ->sheet('Students')
    ->saveXml('students.xml', [
        'header_row' => true,
        'skip_empty_rows' => true,
    ], [
        'root' => 'students',
        'row' => 'student',
    ]);
```

Array-first XML export:

```php
MnbExcel::fromArray($rows)
    ->withHeader()
    ->saveXml('students.xml', [
        'root' => 'students',
        'row' => 'student',
    ]);
```

### JSON to Excel

Simple list-of-objects JSON:

```json
[
  {"student_id":"000123", "name":"Ravi", "address":{"city":"Hyderabad"}},
  {"student_id":"000124", "name":"Sita", "address":{"city":"Vijayawada"}}
]
```

Convert JSON to XLSX:

```php
MnbExcel::fromJson('students.json')
    ->textColumns(['student_id'])
    ->autoWidth()
    ->freezeHeader()
    ->autoFilter()
    ->save('students.xlsx');
```

Workbook-style JSON becomes multiple sheets:

```json
{
  "Students": [
    {"name":"Ravi", "email":"ravi@example.com"}
  ],
  "Teachers": [
    {"name":"Kumar", "subject":"Maths"}
  ]
}
```

```php
MnbExcel::fromJson('school.json')
    ->save('school.xlsx');
```

Read JSON to array:

```php
$rows = MnbExcel::readJson('students.json')
    ->toArray(['header_row' => true]);
```

JSON to CSV/XML:

```php
MnbExcel::fromJson('students.json')->save('students.csv');
MnbExcel::fromJson('students.json')->saveXml('students.xml', [
    'root' => 'students',
    'row' => 'student',
]);
```


### Report builder

```php
MnbExcel::report($rows, 'business')
    ->title('Student Export Report')
    ->columns([
        'name' => 'Student Name',
        'email' => 'Email Address',
        'marks' => 'Marks',
        'status' => 'Status',
        'amount' => 'Amount',
    ])
    ->integerColumns(['marks'])
    ->currencyColumns(['amount'], '₹')
    ->conditionalRowStyle(
        static fn (array $row): bool => ($row['status'] ?? '') === 'fail',
        'mnb.row.danger'
    )
    ->metadata(['creator' => 'MNB PHPExcel'])
    ->autoWidth()
    ->save('student-report.xlsx');
```

### Validation and failed rows report

```php
$result = MnbExcel::validateImport($rows, [
    'name' => 'required|alpha|max:100',
    'email' => 'required|email|unique_in_file',
    'phone' => 'required|phone_basic',
    'website' => 'nullable|url',
], [
    'row_number_key' => '_mnb_original_row_number',
]);

MnbExcel::fromFailedRows($result['failed'])
    ->withImportSummarySheet(['total_rows' => count($rows)], $result)
    ->save('failed-rows.xlsx');
```


### Database connection from `.env`, config file, constants, or PDO

Real applications usually keep database credentials outside the library, for example in `.env`, `config/database.php`, framework config, or constants. MNB PHPExcel now accepts all of these without requiring Laravel/Symfony or `vlucas/phpdotenv`.

Supported `.env` keys:

```env
DB_CONNECTION=mysql
DB_HOST=127.0.0.1
DB_PORT=3306
DB_DATABASE=app
DB_USERNAME=root
DB_PASSWORD=secret
DB_CHARSET=utf8mb4

# Or direct DSN:
# DB_DSN=mysql:host=127.0.0.1;dbname=app;charset=utf8mb4
```

You can also use library-specific keys when you do not want to reuse the app default database:

```env
MNB_EXCEL_DB_CONNECTION=mysql
MNB_EXCEL_DB_HOST=127.0.0.1
MNB_EXCEL_DB_DATABASE=excel_imports
MNB_EXCEL_DB_USERNAME=root
MNB_EXCEL_DB_PASSWORD=secret
```

Create a PDO connection from a real application config source:

```php
$pdo = MnbExcel::pdo(__DIR__ . '/.env');

$pdo = MnbExcel::pdo(__DIR__ . '/config/database.php'); // PHP file returning array

$pdo = MnbExcel::pdo([
    'driver' => 'mysql',
    'host' => '127.0.0.1',
    'database' => 'app',
    'username' => 'root',
    'password' => 'secret',
]);
```

Constants are also supported, useful for custom apps with `db.php` bootstrap files:

```php
define('MNB_EXCEL_DB_DSN', 'mysql:host=127.0.0.1;dbname=app;charset=utf8mb4');
define('MNB_EXCEL_DB_USERNAME', 'root');
define('MNB_EXCEL_DB_PASSWORD', 'secret');

$pdo = MnbExcel::pdo();
```

All SQL helpers still accept an existing `PDO`, but now they can also accept `.env`, config path, DSN, or array directly:

```php
MnbExcel::fromSql(__DIR__ . '/.env', 'SELECT id, name, email FROM students')
    ->withHeader()
    ->save('students.xlsx');

MnbExcel::largeImportToSql('students-large.xlsx', __DIR__ . '/.env', 'students', [
    'with_header' => true,
    'chunk_size' => 500,
    'batch_size' => 250,
    'resume' => true,
]);
```

For debugging/admin screens:

```php
$config = MnbExcel::dbConfig(__DIR__ . '/.env');
$summary = MnbExcel::databaseConfigSummary(['base_path' => dirname(__DIR__)]);
```

### SQL dry-run

```php
$plan = MnbExcel::read('students.xlsx')
    ->sheet('Students')
    ->dryRunImportToSql($pdo, 'students', [
        'header_row' => true,
        'batch_size' => 500,
    ]);
```

### Release readiness check

```php
$report = MnbExcel::releaseReadiness(__DIR__);
```

## Local testing

Run syntax checks:

```bash
composer run test:syntax
```

Run smoke tests:

```bash
composer run test:smoke
```

Run all configured tests:

```bash
composer test
```

Build and verify isolated modular packages:

```bash
composer build:modules
composer test:modules
```

The Composer test scripts are cross-platform and work on Windows PowerShell/CMD as well as Unix shells. They use PHP-based runners instead of `find`, `xargs`, or shell `for` loops.

In this environment, XLSX runtime tests require `ext-zip` and `ext-xmlreader`. If they are missing, the XLSX smoke test exits safely with a skip message.

## Public release checklist

Before Packagist release:

```bash
git add .
git commit -m "Release MNB PHPExcel"
git tag v1.2.0
git push origin main
git push origin v1.2.0
```

Do not include:

```text
vendor/
storage/
.tmp/
.cache/
generated XLSX/CSV outputs
```

## Scope note

MNB PHPExcel has two intentionally separate paths:

- **Rich small/normal mode**: full workbook arrays, styles, comments, hyperlinks, metadata, reports, validation, JSON/XML/CSV conversion, and XLSX integrity protection.
- **Large import mode**: preflight/advisor, streaming row reads, chunk callbacks, PDO batch import, failed-row CSV, progress manifests, and checkpoint/resume.

Do not use the normal reader for very large XLSX files. Run `MnbExcel::analyzeXlsxForImport()` or `MnbExcel::recommendImportMethod()` first, then use `largeRead()` / `largeImportToSql()` when advised.

## Structured JSON/XML output without saving

Use `toStructuredArray()` when you want a PHP array, `toStructuredJson()` / `toStructuredXml()` when you want a string for an API response or variable, and `saveStructuredJson()` / `saveStructuredXml()` only when you want to save a file.

```php
$structured = MnbExcel::read($file)->toStructuredArray([
    'header_row' => true,
    'skip_empty_rows' => true,
]);

$json = MnbExcel::read($file)->toStructuredJson([
    'header_row' => true,
], [
    'pretty' => false,
    'trailing_newline' => false,
]);

$xml = MnbExcel::read($file)->toStructuredXml([
    'header_row' => true,
], [
    'root' => 'structured_workbook',
]);

MnbExcel::read($file)->saveStructuredJson('output.json', [
    'header_row' => true,
]);

MnbExcel::read($file)->saveStructuredXml('output.xml', [
    'header_row' => true,
], [
    'root' => 'structured_workbook',
]);
```


## Error handling and atomic save guard

All writer save paths now use a temp-file-first workflow for safer exports:

```text
create temp file in the same directory
write full contents
validate when needed, such as XLSX integrity validation
replace final file only after success
delete temp file on failure
keep existing final file unchanged when a write fails before replacement
```

Package exceptions expose stable error codes while preserving the original developer message and previous exception internally.

```php
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

try {
    MnbExcel::fromArray($rows)
        ->withHeader()
        ->save('report.xlsx');
} catch (MnbExcelException $e) {
    // Safe for frontend/API responses. Does not expose local paths or stack details.
    return MnbExcel::safeError($e);
}
```

For logs or admin/debug screens:

```php
$error = MnbExcel::errorReport($e, true);
```

The error response includes:

```text
status
code
category
message
safe_message
recoverable
```

Common codes include:

```text
MNB_FILE_NOT_FOUND
MNB_FILE_WRITE_FAILED
MNB_FILE_REPLACE_FAILED
MNB_UNSUPPORTED_FORMAT
MNB_EXTENSION_MISSING
MNB_JSON_INVALID
MNB_JSON_ENCODE_FAILED
MNB_JSON_WRITE_FAILED
MNB_XML_WRITE_FAILED
MNB_CSV_WRITE_FAILED
MNB_XLSX_WRITE_FAILED
MNB_XLSX_ZIP_ENTRY_FAILED
MNB_XLSX_ZIP_CLOSE_FAILED
MNB_XLSX_INTEGRITY_FAILED
MNB_SQL_IMPORT_FAILED
MNB_SQL_EXPORT_FAILED
MNB_SECURITY_BLOCKED
```

## Large Excel Preflight and Method Advisor

MNB PHPExcel keeps the existing rich reader/writer for small and normal Excel files. For large imports, use the preflight/advisor layer first so the application does not load a huge workbook into PHP arrays.

```php
$plan = MnbExcel::autoImportPlan('students.xlsx', [
    'server' => 'shared',
    'memory_limit' => '256M',
    'max_execution_time' => 30,
]);

print_r($plan['advice']);
```

The analyzer inspects XLSX package structure with ZIP/XML streaming and reports:

- file size and uncompressed worksheet XML size
- sheet count, row count, column count, and estimated cell count
- shared string table size
- formulas, comments, hyperlinks, merged cells, drawings, charts, pivots, external links, macros, and tables
- Excel sheet-limit risk
- recommended import method and chunk size

### Method decision matrix

| Level | Rows / cells | Recommended method |
|---|---:|---|
| Tiny | 1-2,000 rows / up to 50k cells | `normal_read` |
| Small | 2,001-10,000 rows / up to 150k cells | `normal_read` |
| Normal | 10,001-50,000 rows / up to 500k cells | `normal_read_with_warning` or stream on low memory |
| Medium | 50,001-150,000 rows / up to 1.5M cells | `streaming_chunk_import` |
| Large | 150,001-500,000 rows / up to 5M cells | `streaming_chunk_import` |
| Very large | 500,001-1,048,576 rows / 5M+ cells | `cli_queue_recommended` |
| Beyond Excel sheet | Above 1,048,576 rows or 16,384 columns | split sheets / CSV ZIP / DB import |

### Streaming import for large XLSX

```php
$advice = MnbExcel::recommendImportMethod('students.xlsx', [
    'server' => 'vps',
    'memory_limit' => '512M',
    'max_execution_time' => 300,
]);

MnbExcel::largeRead('students.xlsx')
    ->withHeader()
    ->chunk((int) $advice['recommended_chunk_size'], function (array $rows, array $state): void {
        // validate and insert this chunk into a database.
        // Never collect all chunks into one giant array.
    });
```

Useful large-read controls:

```php
MnbExcel::largeRead('students.xlsx')
    ->sheet('Worksheet')
    ->withHeader()
    ->skipRows(0)
    ->limitRows(100000)
    ->onlyColumns(['A', 'B', 'C'])
    ->timeBudgetSeconds(25)
    ->memoryGuardRatio(0.85)
    ->chunk(500, $callback);
```

For shared hosting, prefer small chunks and CLI/queue for very large files. The streaming reader includes memory and time-budget guards, but PHP HTTP timeouts are controlled by the server, so very large imports should run in a CLI worker whenever possible.

## Large Excel Database Import Engine

For very large XLSX imports, do not load the workbook into one PHP array. Use the database import engine so rows are streamed, validated in chunks, inserted with PDO batches, and checkpointed for resume.

The `$pdo` argument may be an existing `PDO`, `.env` path, PHP config file path, DSN string, or config array.

```php
$result = MnbExcel::largeImportToSql('students.xlsx', $pdo, 'students', [
    'with_header' => true,
    'chunk_size' => 500,
    'batch_size' => 250,
    'manifest_path' => __DIR__ . '/storage/students-import.json',
    'failed_rows_csv' => __DIR__ . '/storage/students-failed-rows.csv',
    'resume' => true,
    'time_budget_seconds' => 25,
    'rules' => [
        'email' => 'nullable|email',
        'amount' => 'nullable|numeric',
    ],
]);
```

The engine provides:

- streaming chunk validation with existing validation rules
- PDO batch insert without collecting all rows
- failed-row CSV export with Excel row number, errors, and row JSON
- JSON progress manifest for admin dashboards and CLI/queue jobs
- checkpoint/resume from the last completed Excel row
- transaction-per-chunk safety for database inserts
- soft time-budget guard so shared hosting jobs can pause before timeout

Useful helper:

```php
$manifestPath = MnbExcel::largeImportManifestPath('students.xlsx', 'students');
```


### Large import stability options

For database imports that may be restarted or repeated, enable idempotent duplicate handling with a database unique key:

```php
$result = MnbExcel::largeImportToSql('students.xlsx', $pdo, 'students', [
    'with_header' => true,
    'chunk_size' => 500,
    'batch_size' => 250,
    'resume' => true,
    'idempotent' => true,
    'duplicate_strategy' => 'skip', // fail | skip | update
    'unique_by' => ['id'],
    'failed_rows_format' => 'human',
]);
```

Duplicate strategies:

| Strategy | Behavior |
| --- | --- |
| `fail` | Normal insert; database duplicate errors stop the import. |
| `skip` | Uses driver-specific ignore/do-nothing insert where supported. Best for resume-safe idempotent imports. |
| `update` | Uses upsert where supported. Requires `unique_by`. |

Streaming reads can preserve numeric values as strings and convert Excel date serials using `styles.xml`:

```php
MnbExcel::largeRead('students.xlsx')
    ->withHeader()
    ->preserveNumericStrings()
    ->convertDates(true, 'Y-m-d')
    ->chunk(1000, $callback);
```

Use `preserveNumericStrings()` for IDs, phone numbers, account numbers, ISBNs, and other values where PHP integer/float conversion can damage precision or leading-zero semantics.

### All-sheet large workbook import

Import every sheet without loading the workbook:

```php
$result = MnbExcel::largeImportWorkbookToSql('school.xlsx', $pdo, [
    'Students' => 'students',
    'Scores' => 'scores',
], [
    'chunk_size' => 500,
    'batch_size' => 250,
    'resume' => true,
]);
```

You can also pass a table prefix string. For multi-sheet files, table names are generated as `prefix_sheetname`:

```php
MnbExcel::largeImportWorkbookToSql('school.xlsx', $pdo, 'import');
```

Failed rows are exported in human-readable CSV by default:

```text
excel_row,error_count,errors,id,email,amount
3,1,email must be a valid email.,2,bad-email,15
```

Set `failed_rows_format => 'json'` to use the older compact machine format.

Recommended flow for shared hosting:

1. Run `autoImportPlan()` first.
2. Use the advised chunk size.
3. Set `time_budget_seconds` below the server timeout.
4. Re-run the same import command with `resume => true` until status becomes `completed`.
5. Review `failed_rows_csv` for validation failures.


## Real application and plugin integration layer

MNB PHPExcel can now be used as an application-level import system, not only as a one-off Excel helper. The core package remains framework-neutral: it does not require Laravel, Symfony, CodeIgniter, or PSR packages.

### Application storage paths

Keep import manifests, failed rows, uploads, reports, and temp files outside `vendor`:

```php
use Mnb\PHPExcel\MnbExcel;

MnbExcel::storage([
    'base_path' => __DIR__ . '/storage/mnb-phpexcel',
    'temp_path' => __DIR__ . '/storage/mnb-phpexcel/temp',
    'upload_path' => __DIR__ . '/storage/mnb-phpexcel/uploads',
    'manifest_path' => __DIR__ . '/storage/mnb-phpexcel/manifests',
    'failed_rows_path' => __DIR__ . '/storage/mnb-phpexcel/failed-rows',
    'reports_path' => __DIR__ . '/storage/mnb-phpexcel/reports',
]);
```

### Import profiles

Profiles avoid passing long option arrays in every controller/job:

```php
MnbExcel::registerImportProfile('student_import', [
    'table' => 'students',
    'with_header' => true,
    'chunk_size' => 500,
    'batch_size' => 250,
    'resume' => true,
    'idempotent' => true,
    'duplicate_strategy' => 'update',
    'unique_by' => ['student_id'],
    'rules' => [
        'email' => 'nullable|email',
        'amount' => 'nullable|numeric',
    ],
    'column_map' => [
        'Student ID' => 'student_id',
        'Email Address' => 'email',
    ],
]);

$result = MnbExcel::profile('student_import')
    ->source($_FILES['excel']['tmp_name'])
    ->run(__DIR__ . '/.env');
```

Profiles may also be loaded from a PHP file:

```php
MnbExcel::loadImportProfiles(__DIR__ . '/config/excel-imports.php');
```

### Import status and resume

Admin dashboards can read manifest state without importing again:

```php
$status = MnbExcel::importStatus(__DIR__ . '/storage/imports/students.json');
```

Resume a CLI/queue/cron job from a manifest:

```php
$result = MnbExcel::resumeImport(
    __DIR__ . '/storage/imports/students.json',
    __DIR__ . '/.env',
    ['time_budget_seconds' => 25]
);
```

### Upload safety validation

Before accepting a user-uploaded spreadsheet:

```php
$upload = MnbExcel::validateUpload($_FILES['excel'], [
    'allowed_extensions' => ['xlsx', 'csv'],
    'max_size_mb' => 100,
    'strict_mime' => true,
    'max_zip_entries' => 10000,
    'max_uncompressed_size_mb' => 1024,
    'max_zip_entry_size_mb' => 256,
    'max_compression_ratio' => 200,
    'reject_macros' => true,
    'reject_external_links' => true,
]);

if (!$upload['valid']) {
    // show $upload['errors'] to admin/developer UI
}
```

For XLSX files, the validator checks required package parts, unsafe ZIP paths, archive entry count, expanded size, per-entry size, suspicious compression ratios, encrypted content, macros, and external links when `ext-zip` is enabled. Limits are configurable; unsafe paths, suspicious archives, and encrypted content are rejected by default.

### Event hooks

```php
MnbExcel::on('after_chunk', function (array $event): void {
    // update progress table, cache, websocket, or admin dashboard
});

MnbExcel::on('on_import_failed', function (array $event): void {
    // log/import alert
});
```

Supported application events include `before_import`, `before_chunk`, `after_chunk`, `on_failed_row`, `on_import_paused`, `on_import_completed`, and `on_import_failed`.

### Row transformers

```php
MnbExcel::transformer('clean_student_row', function (array $row): array {
    $row['email'] = strtolower(trim((string) ($row['email'] ?? '')));
    $row['amount'] = str_replace(',', '', (string) ($row['amount'] ?? ''));
    return $row;
});

MnbExcel::largeRead('students.xlsx')
    ->withHeader()
    ->transform('clean_student_row')
    ->chunk(1000, function (array $rows): void {
        // rows are already cleaned
    });
```

Large SQL imports also accept `transformers` in options/profile config.

### Custom validators

```php
MnbExcel::validator('valid_student_id', function (mixed $value): bool|string {
    return preg_match('/^STU-[0-9]{5}$/', (string) $value) === 1
        ? true
        : 'student_id must match STU-00000 format.';
});

$result = MnbExcel::validateImport($rows, [
    'student_id' => 'required|valid_student_id',
]);
```

### Plugins

Framework or application packages can register profiles, validators, transformers, and event listeners:

```php
use Mnb\PHPExcel\Plugin\PluginRegistry;

MnbExcel::plugin(function (PluginRegistry $registry): void {
    $registry->importProfile('student_import', [
        'table' => 'students',
        'with_header' => true,
    ]);

    $registry->validator('valid_status', fn ($value) => in_array($value, ['active', 'inactive'], true));

    $registry->transformer('trim_strings', function (array $row): array {
        return array_map(fn ($value) => is_string($value) ? trim($value) : $value, $row);
    });

    $registry->event('after_chunk', function (array $event): void {
        // progress notification
    });
});
```

### Optional logger

```php
MnbExcel::setLogger(function (string $level, string $message, array $context): void {
    error_log(strtoupper($level) . ': ' . $message . ' ' . json_encode($context));
});
```

PSR-3-like logger objects are supported without requiring `psr/log` as a dependency.

## Release-quality proof suite

MNB PHPExcel now includes proof-oriented helpers for benchmarking, compatibility fixtures, integration readiness, and advanced workbook scope transparency.

### Benchmark plan and local performance table

```php
$plan = MnbExcel::benchmarkPlan([
    'rows' => [100000, 500000, 1000000],
    'columns' => 10,
]);
```

Run internal large writer benchmarks:

```bash
php tools/benchmark-large-writer.php --rows=100000 --cols=10 --json
php tools/benchmark-large-writer.php --rows=500000 --cols=10 --json
php tools/benchmark-large-writer.php --rows=1000000 --cols=10 --json
```

Use the table template in `docs/benchmarks/README.md` to publish measured values for MNB PHPExcel, PhpSpreadsheet, OpenSpout, PHP_XLSXWriter, Laravel Excel/FastExcel. Benchmark numbers should be generated on the same machine, same PHP version, same dataset, and same storage disk.

### Real Excel fixture compatibility

```php
$requirements = MnbExcel::compatibilityFixtureRequirements();
$report = MnbExcel::verifyCompatibilityFixtures(__DIR__ . '/docs/fixtures');
```

Recommended real fixture sources:

```text
Microsoft Excel
LibreOffice Calc
Google Sheets export
WPS Office
```

Run:

```bash
php tools/compatibility-fixture-check.php docs/fixtures
```

Generated XLSX integrity tests are automated. Real application fixtures should still be opened manually in Excel/LibreOffice to confirm there is no repair warning.

### MySQL/PostgreSQL integration readiness

```php
$plan = MnbExcel::databaseIntegrationPlan();
$check = MnbExcel::databaseIntegrationCheck();
```

Run:

```bash
php tools/database-integration-check.php
```

Live MySQL/PostgreSQL checks are skipped unless `MNB_EXCEL_TEST_MYSQL_*` or `MNB_EXCEL_TEST_PGSQL_*` environment variables are configured. See `docs/DATABASE_INTEGRATION_TESTS.md`.

### Packagist-ready comparison

See `docs/PACKAGIST_COMPARISON.md` for the recommended positioning:

```text
MNB PHPExcel is a framework-neutral PHP Excel application toolkit:
upload → preflight → method decision → streaming import/export → validation → database → dashboard/status → recovery.
```

Use MNB PHPExcel for real application import/export workflows. Use specialist spreadsheet engines when you need deep workbook editing, a formula calculation engine, or a Laravel-only integration style.

### Advanced workbook capability matrix

```php
$matrix = MnbExcel::advancedWorkbookCapabilities();
```

The matrix covers:

```text
deep cell-level spreadsheet editing
many spreadsheet formats
advanced style manipulation
formula calculation engine
complex workbook manipulation
```

See `docs/ADVANCED_WORKBOOK_CAPABILITIES.md`.

### Framework examples

Examples are included for:

```text
Slim routes: examples/frameworks/slim_import_routes.php
CodeIgniter-style controller: examples/frameworks/codeigniter_import_controller.php
Import dashboard UI: examples/ui/import_dashboard_example.php
```

Laravel should be released as a separate adapter package, not inside the core library. See `docs/ADAPTER_PACKAGES.md`.
