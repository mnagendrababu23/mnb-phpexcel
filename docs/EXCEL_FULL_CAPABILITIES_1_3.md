# Excel Full Capabilities — v1.3.0

Version 1.3.0 promotes the remaining basic XLSX capabilities into explicit, tested public APIs while preserving the v1.2 reader and builder behavior.

## Direct cell and range access

```php
use Mnb\PHPExcel\Format\Xlsx;

$session = Xlsx::read('report.xlsx')->sheet('Sales');
$value = $session->cell('B7');
$values = $session->cells(['A1', 'B7', 'F20']);
$range = $session->rangeValues('A1:F20');
$details = $session->cellDetails('B7');
```

`CellSnapshot` contains the value, formula, cached value, optional recalculated value, rich text, full style, comments, hyperlinks, and images anchored at the cell.

## Formula results

Native XLSX reading exposes formula expressions and workbook-cached values without extra dependencies. True recalculation is available through the optional PhpSpreadsheet adapter:

```bash
composer require phpoffice/phpspreadsheet
```

```php
$result = Xlsx::read('finance.xlsx')->sheet('Model')->calculatedCell('D24');
$results = Xlsx::read('finance.xlsx')->sheet('Model')->calculatedRange('D2:D24');
```

The calculation adapter is intentionally optional so lightweight XLSX users do not download a calculation engine.

## Complete cell formatting metadata

```php
$style = $session->cellStyle('C5');
$styles = $session->rangeStyles('A1:F20');
```

Returned style data includes:

- font family, size, bold, italic, underline, strike, outline, shadow, color, scheme, charset, and vertical alignment;
- pattern and gradient fills;
- per-side borders and colors;
- built-in and custom number formats;
- horizontal/vertical alignment, wrapping, indentation, rotation, shrink-to-fit, and reading order;
- locked/hidden protection flags and style application metadata.

## Rich text and images

```php
$rich = $session->richText('A1');
$images = $session->images();
$written = $session->extractImages(__DIR__ . '/extracted-images');
```

Image metadata includes the media part, relationship, anchor cells and offsets, MIME type, extension, byte size, dimensions, name, and description. `extractImages()` writes safely named files and avoids collisions unless overwrite is explicitly enabled.

## Semantic header detection

```php
$rows = Xlsx::read('customer-upload.xlsx')
    ->autoDetectHeader(sampleRows: 30, minimumConfidence: 0.4)
    ->toArray();

$detection = Xlsx::read('customer-upload.xlsx')->detectHeader();
```

The detector scores text-label density, uniqueness, identifier-like labels, column coverage, following data-like rows, and row position. A strict confidence threshold can be enabled through reader options.

## Import templates

```php
use Mnb\PHPExcel\Format\Xlsx;

Xlsx::writeImportTemplate([
    ['header' => 'SKU', 'required' => true, 'example' => 'SKU-100'],
    ['header' => 'Status', 'list' => ['active', 'inactive']],
    ['header' => 'Price', 'validation' => [
        'type' => 'decimal',
        'operator' => 'greaterThan',
        'formula1' => 0,
    ]],
], 'product-import-template.xlsx', [
    'instructions' => 'Complete all required columns.',
    'validation_rows' => 25000,
]);
```

Templates support styled headers, instructions, examples, comments, auto-width, frozen headers, filters, list validation, numeric/date/custom validation, prompts, and validation errors.

## Freeze panes and filters

```php
MnbExcel::fromArray($rows)
    ->withHeader()
    ->freezeAt('C3')
    ->autoFilterRange('A2:H5000')
    ->filterValues('D', ['Open', 'Pending'])
    ->filterColumn('F', [
        'type' => 'custom',
        'and' => true,
        'filters' => [
            ['operator' => 'greaterThanOrEqual', 'value' => 100],
            ['operator' => 'lessThan', 'value' => 1000],
        ],
    ]);
```

Supported filters include selected values/blanks, custom comparisons, top/bottom values or percentages, dynamic date filters, and color-filter definitions.

## Native conditional formatting

```php
$builder
    ->conditionalCellIs('D2:D500', 'lessThan', 0, 'mnb.error')
    ->conditionalExpression('A2:H500', '$D2="Overdue"', 'mnb.warning')
    ->conditionalColorScale('F2:F500')
    ->conditionalDataBar('G2:G500')
    ->conditionalIconSet('H2:H500');
```

The writer emits native OOXML conditional-format rules and differential styles. Supported rule families include cell comparisons, expressions, color scales, data bars, icon sets, top/bottom rules, duplicate/unique values, text containment, and time periods.

## Native charts

```php
$builder->addChart('column', 'Sales by Month', [[
    'name' => 'Sales',
    'categories' => 'A2:A13',
    'values' => 'B2:B13',
]], [
    'from' => 'D2',
    'to' => 'L20',
    'legend' => 'bottom',
]);
```

Supported chart types are column, bar, line, area, pie, doughnut, and scatter. The writer creates chart, drawing, relationship, and content-type parts directly without requiring PhpSpreadsheet.

## Pivot tables

Pivot support is template-driven because a pivot table is a graph of interdependent OOXML parts rather than a single worksheet element.

1. Create and format the pivot table once in Excel.
2. Regenerate the source rows with MNB PHPExcel.
3. Preserve and rebind the template pivot cache:

```php
MnbExcel::fromWorkbookArray($sheets)
    ->preservePivotTablesFrom(
        'sales-pivot-template.xlsx',
        sourceSheet: 'Raw Data',
        sourceRange: 'A1:H50000'
    )
    ->save('sales-report.xlsx');
```

The writer preserves pivot tables, pivot caches, cache records, worksheet relationships, content types, slicer-related package parts, and drawings. It rebinds the worksheet source range and sets the cache to refresh when Excel opens the generated workbook.

When a preserved worksheet contains relationship-backed template objects, add new charts, drawings, hyperlinks, or comments on a separate worksheet. The writer rejects ambiguous relationship merging rather than silently corrupting the package.

This is full support for production template-based pivot workflows. A from-scratch pivot-layout designer is intentionally not represented as supported.

## Backward compatibility

All v1.2 behavior remains available. New data fields were appended with defaults, new fluent methods are additive, and old array options continue to work.
