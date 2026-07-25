# Advanced Workbook Capabilities

MNB PHPExcel separates application workflows from specialist spreadsheet manipulation.

## Current strengths

- Small/normal XLSX writer with styles, comments, hyperlinks, merged cells, reports, formulas with cached values, and integrity validation
- Large XLSX streaming reader/writer with low-memory import/export
- Database import/export workflows, progress manifests, failed rows, and dashboard responses
- Plugin-level validators, transformers, events, profiles, and framework-neutral adapters

## Current/planned advanced areas

| Area | Current status | Recommended approach |
|---|---|---|
| Deep cell-level editing | Partial | Use normal writer for generated reports; use future adapter for existing complex workbook editing |
| Many spreadsheet formats | Partial | Core supports XLSX/CSV/JSON/XML/CSV ZIP; add ODS/HTML/NDJSON through plugins |
| Advanced style manipulation | Partial | Normal mode can be rich; large mode should remain style-limited for safety |
| Formula calculation engine | Not in core | Write formulas/cached values; use adapter/specialist library for calculation |
| Complex workbook manipulation | Partial preservation | Preserve advanced parts safely, validate integrity, do not claim full pivot/chart/macro editing in core |

Run in code:

```php
$matrix = MnbExcel::advancedWorkbookCapabilities();
```

