# Advanced Workbook Capabilities

MNB PHPExcel keeps the native XLSX path lightweight while making heavyweight compatibility engines optional.

## v1.3 capability matrix

| Area | Status | Coverage |
|---|---|---|
| Direct cell/range access | Supported | Single cells, multiple cells, ranges, typed snapshots, styles, rich text, images |
| Spreadsheet formats | Modular | XLSX, CSV/TSV, JSON/NDJSON, XML, ODS read, optional XLS read |
| Advanced styles | Supported | Fonts, fills, borders, alignment, protection, number formats, conditional formatting |
| Formula calculation | Optional adapter | Formula/cached values natively; true recalculation through the PhpSpreadsheet adapter |
| Charts | Supported | Column, bar, line, area, pie, doughnut, and scatter chart generation |
| Pivot tables | Template workflow | Preserve pivot parts, rebind source range, and refresh the cache on open |
| Complex templates | Supported for trusted templates | Preserve comments, hyperlinks, drawings, images, charts, pivots, tables, and related parts |

## Deliberate boundaries

- The native core does not contain a heavyweight formula engine.
- Legacy binary XLS reading is isolated in the optional XLS module.
- ODS writing and legacy XLS writing are not currently included.
- Pivot tables are produced through a trusted template workflow; a from-scratch pivot-layout designer is not advertised.
- Macros are never executed.
- Large streaming output intentionally has a smaller formatting surface than normal workbook output.

Run the same matrix in code:

```php
$matrix = MnbExcel::advancedWorkbookCapabilities();
```
