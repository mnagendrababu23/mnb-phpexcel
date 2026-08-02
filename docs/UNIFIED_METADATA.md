# Unified Metadata API

The unified metadata module provides one versioned response schema for every reader session while allowing format-specific collectors to expose richer information.

## Read metadata

```php
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Format\Xlsx;

$report = MnbExcel::metaInfo('report.xlsx', [
    'profile' => 'standard',
]);

$report = Xlsx::metaInfo('report.xlsx', [
    'profile' => 'full',
    'max_items' => 5000,
]);

$report = Xlsx::read('report.xlsx')->metaInfo([
    'profile' => 'forensic',
    'include_hash' => true,
]);
```

The top-level schema is versioned independently from the package version:

```php
[
    'schema_version' => '1.0',
    'status' => 'ok',
    'profile' => 'full',
    'format' => 'xlsx',
    'format_variant' => 'xlsx',
    'mime_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

    'file' => [],
    'format_details' => [],
    'document' => [],
    'revision' => [],
    'application' => [],
    'workbook' => [],
    'custom_properties' => [],
    'security' => [],
    'macros' => [],
    'named_objects' => [],
    'links' => [],
    'hidden_content' => [],
    'comments_notes' => [],
    'tracked_changes' => [],
    'embedded_objects' => [],
    'calculation' => [],
    'print_settings' => [],
    'validation' => [],
    'pivot_metadata' => [],
    'xml_metadata' => [],
    'statistics' => [],

    'capabilities' => [],
    'warnings' => [],
    'errors' => [],
]
```

Every section contains a `state`, `count`, `items`, `truncated`, and `warnings` envelope. A section may also expose category-specific summary fields.

## Section states

| State | Meaning |
|---|---|
| `available` | The section was read successfully. |
| `partial` | Useful information was returned, but some fields or object types are not fully decoded. |
| `not_supported` | The format can contain the concept, but this reader does not implement it yet. |
| `not_applicable` | The concept does not apply to this format. |
| `not_scanned` | The selected profile deliberately skipped the scan. |
| `encrypted` | Encryption prevented access to the section. |
| `password_required` | A password is required to continue. |
| `error` | The section failed to read. |

Do not infer “no objects” from an empty `items` array without checking `state` and `count`.

## Profiles

| Profile | Intended use |
|---|---|
| `quick` | File properties, document properties, workbook sheets, active sheet, hidden-sheet state, encryption, macros and package-level indicators. It avoids full worksheet scans. |
| `standard` | Adds structural worksheet scans, hidden rows/columns, formulas counts, links, print settings, validations, protection and object counts. |
| `full` | Adds item-level inventories such as formulas, comments, hyperlinks, images, validations and package objects. Accurate row/cell counts are enabled by default. |
| `forensic` | Adds hashes, package-part listings and relationship-oriented package details. It still never executes workbook content. |

Common read options:

```php
[
    'profile' => 'standard',
    'password' => '',
    'max_items' => 1000,
    'include_hash' => false,
    'include_package_parts' => false,
    'include_relationships' => false,
    'accurate_sheet_counts' => false,
]
```

`max_items` is capped at 100,000. Counts continue to describe the complete scanned set when the returned item list is truncated.

## XLSX coverage

The native XLSX collector currently reports:

- filesystem data, SHA-256 on request, ZIP package sizes and OOXML variant;
- title, subject, creator, manager, company, category, keywords, description, language, identifier and document version;
- last saved by, revision number, total editing time, created/modified/printed timestamps;
- application name/version and extended application properties;
- typed custom properties;
- sheet names, IDs, visibility, active sheet, dimensions and optional accurate row/cell counts;
- file encryption, workbook protection, worksheet protection and digital-signature part presence;
- VBA project presence and macro-related package parts;
- defined names, tables, charts, pivots and slicers;
- external links, hyperlinks, data connections and connection-related parts;
- hidden worksheets, rows and columns;
- legacy comments, threaded-comment package parts and reviewer/person information when available;
- images, drawings, OLE objects, embedded packages and ActiveX package parts;
- formulas, calculation mode and recalculation settings;
- print areas/titles, margins, page setup, print options, headers and footers;
- data validations and conditional formatting;
- pivot table/cache package inventory;
- content types, custom XML parts, package parts and structural statistics.

Some complex object categories are intentionally marked `partial`: the module inventories all relevant package parts, but it does not yet decode every chart, PivotTable, slicer, ActiveX, digital-signature, tracked-change, Power Query or custom XML semantic field.

Metadata inspection never executes formulas, macros, ActiveX controls, connections, Power Query, external links or embedded files.

## Encrypted XLSX files

Without a password, safe file/security information is returned and protected sections are marked `encrypted`:

```php
$report = Xlsx::metaInfo('protected.xlsx');

if ($report['status'] === 'password_required') {
    // Ask the application user for the password.
}
```

With a password:

```php
$report = Xlsx::metaInfo('protected.xlsx', [
    'password' => $password,
    'profile' => 'full',
]);
```

The decrypted package exists only in a temporary file and is removed in a `finally` block.

## Update metadata

XLSX metadata updates are atomic and preserve unknown package parts byte-for-byte:

```php
MnbExcel::updateMetaInfo('source.xlsx', 'updated.xlsx', [
    'document' => [
        'title' => 'Annual Report',
        'subject' => 'FY 2026',
        'creator' => 'MNB',
        'keywords' => ['annual', 'finance'],
        'description' => 'Approved annual report',
        'category' => 'Finance',
        'content_status' => 'Final',
        'identifier' => 'REPORT-2026',
        'language' => 'en-IN',
        'document_version' => '1.0',
    ],
    'revision' => [
        'last_saved_by' => 'Release Bot',
        'revision_number' => '7',
        'total_editing_time_seconds' => 3600,
        'last_printed_at' => '2026-08-02T10:00:00Z',
        'document_created_at' => '2026-08-01T10:00:00Z',
        'document_modified_at' => '2026-08-02T10:00:00Z',
    ],
    'application' => [
        'application' => 'MNB PHPExcel',
        'application_version' => '2.0',
        'manager' => 'Finance Manager',
        'company' => 'MNB',
        'operating_system_hint' => 'Linux',
        'document_security' => 0,
        'scale_crop' => false,
        'links_up_to_date' => true,
        'shared_document' => false,
        'hyperlinks_changed' => false,
    ],
    'custom_properties' => [
        'Project ID' => 'PRJ-1001',
        'Department' => ['type' => 'string', 'value' => 'Finance'],
        'Approved' => ['type' => 'boolean', 'value' => true],
        'Budget' => ['type' => 'float', 'value' => 1250.50],
        'Revision Date' => ['type' => 'datetime', 'value' => '2026-08-02T10:00:00Z'],
        'Remove This Property' => null,
    ],
    'workbook' => [
        'active_sheet' => 'Summary',
        'sheet_visibility' => [
            'Raw Data' => 'hidden',
            'Internal' => 'veryHidden',
        ],
        'date1904' => false,
        'code_name' => 'AnnualReport',
        'read_only_recommended' => true,
    ],
    'calculation' => [
        'mode' => 'auto', // auto, manual, autoNoTable
        'calc_id' => 20260802,
        'full_calc_on_load' => true,
        'force_full_calc' => true,
        'calc_on_save' => true,
        'iterate' => false,
        'iterate_count' => 100,
        'iterate_delta' => 0.001,
    ],
]);
```

Use the same source and destination path for an atomic in-place replacement.

Writer options:

```php
[
    'strict' => true,
    'validate_package' => true,
    'replace_custom_properties' => false,
    'password' => '',
    'encryption_options' => [],
]
```

For encrypted input, provide `password`. The output remains encrypted and preserves Agile or Standard mode.

Unknown change sections are rejected in strict mode. At least one worksheet must remain visible, and the active worksheet must be visible.

## Remove personal information

```php
MnbExcel::removePersonalInfo('source.xlsx', 'clean.xlsx', [
    'remove_custom_properties' => true,
    'remove_descriptive_properties' => false,
    'anonymized_author_name' => 'Author',
]);
```

This removes creator, last-saved-by, manager, company and custom properties, and anonymizes legacy comment authors and threaded-comment persons. Set `remove_descriptive_properties` to also remove title, subject, keywords, description and category.

## Format status

- **XLSX:** rich native reader and atomic writer are implemented.
- **CSV:** receives the shared schema envelope through `ReadSession::metaInfo()`. File and synthetic single-sheet information are available; format-specific encoding/dialect statistics are the next collector milestone.
- **XLS:** receives the same shared schema envelope through the core session. Native BIFF/OLE property and feature collectors are the next collector milestone.
- **Other formats:** receive the shared envelope and explicit `not_supported` states until their own collectors implement `MetadataReaderInterface`.

The shared schema prevents format modules from inventing incompatible return structures while allowing each format to progress independently.
