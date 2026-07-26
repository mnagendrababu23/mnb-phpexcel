# Modular package distribution

MNB PHPExcel 1.7 can be released as a package family so applications download only the formats they use.

## Install examples

### XLSX only

```bash
composer require mnb/mnb-phpexcel-xlsx
```

Composer installs `mnb/mnb-phpexcel-core` automatically. It does not install CSV, JSON, XML, ODS, legacy XLS, database, or application-facade code.

```php
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;

$options = ReaderOptions::defaults()
    ->withRange(startRow: 2, endRow: 10000, startColumn: 'A', endColumn: 'F')
    ->withMode('streaming');

foreach (Xlsx::read('orders.xlsx', $options)->rows() as $row) {
    // Forward-only rows.
}

Xlsx::write($rows, 'orders.xlsx', [
    'with_header' => true,
    'sheet_name' => 'Orders',
]);
```

Required for XLSX security: `ext-openssl` and `ext-iconv`. Native `ext-zip` and `ext-xmlreader` are recommended for maximum streaming performance; secure pure-PHP fallbacks are included.

### XML only

```bash
composer require mnb/mnb-phpexcel-xml
```

```php
use Mnb\PHPExcel\Format\Xml;

foreach (Xml::read('catalog.xml', [
    'xml_schema' => [
        'row_tag' => 'product',
        'columns' => [
            'id' => ['source' => '@id', 'type' => 'integer'],
            'name' => 'details/name',
            'price' => ['source' => 'details/price', 'type' => 'decimal'],
        ],
    ],
])->withHeaderRow()->rows() as $product) {
    // $product is schema-mapped without loading the full XML document.
}
```

Native `ext-xmlreader` is recommended for true forward-only streaming; a secure pure-PHP fallback is included.

### CSV and JSON only

```bash
composer require mnb/mnb-phpexcel-csv mnb/mnb-phpexcel-json
```

### ODS only

```bash
composer require mnb/mnb-phpexcel-ods
```

The native ODS module provides streaming reads. Native `ext-zip` and `ext-xmlreader` are recommended for the lowest-memory path; pure-PHP fallbacks are included.

### Legacy XLS only

```bash
composer require mnb/mnb-phpexcel-xls
```

The `.xls` adapter is deliberately separate because it depends on `phpoffice/phpspreadsheet`. Native XLSX support does not depend on PhpSpreadsheet.

### Database workflows

```bash
composer require mnb/mnb-phpexcel-database
```

This installs core and XLSX streaming dependencies and requires `ext-pdo`.

### Complete split package family

```bash
composer require mnb/mnb-phpexcel-all
```

### Existing full package

```bash
composer require mnb/mnb-phpexcel
```

The existing package remains the backward-compatible monolith. It declares `replace` entries for the split package names, preventing duplicate class installations.

## Package map

| Package | Purpose | Automatically requires |
|---|---|---|
| `mnb/mnb-phpexcel-core` | Data model, read session, typed options, projection, progress, error policies, registry | PHP, JSON |
| `mnb/mnb-phpexcel-csv` | CSV reader/writer and encoding/dialect handling | core |
| `mnb/mnb-phpexcel-json` | JSON, streaming JSON-array, NDJSON reader/writer | core |
| `mnb/mnb-phpexcel-xml` | Streaming XML, schema mapping, XML writer | core; XMLReader recommended |
| `mnb/mnb-phpexcel-xlsx` | Native XLSX read/write, encryption, protection, formulas, pivots, large streaming | core, OpenSSL, iconv; ZIP/XMLReader recommended |
| `mnb/mnb-phpexcel-ods` | Native ODS streaming reader | core; ZIP/XMLReader recommended |
| `mnb/mnb-phpexcel-xls` | Optional binary XLS adapter | core, PhpSpreadsheet |
| `mnb/mnb-phpexcel-database` | PDO imports, manifests, failed rows, database streaming | core, XLSX, PDO |
| `mnb/mnb-phpexcel-application` | Legacy facade, workbook builder, profiles, storage, plugins | all functional modules |
| `mnb/mnb-phpexcel-all` | Metapackage | application |

## Development and release model

There is one source repository. Module source is not duplicated by hand.

```bash
composer build:modules
composer test:modules
```

`tools/build-modular-packages.php`:

1. Reads `packages/modules.php`.
2. Rejects duplicate or unassigned PHP source files.
3. Creates one isolated package directory per module.
4. Writes module-specific `composer.json` files.
5. Syntax-checks every generated PHP file.
6. Produces ZIP archives and `dist/module-manifest.json` with SHA-256 hashes.

Packagist treats a VCS repository root as one Composer package. Therefore public distribution should use one of these release patterns:

1. **Split repositories:** a release workflow pushes each generated `dist/modules/<module>` tree to its matching read-only repository and tag.
2. **Private Composer repository:** publish the generated archives and package metadata through Satis, Private Packagist, or another Composer repository.
3. **Direct artifact release:** attach the generated ZIP files to a release for manual/custom-repository installs.

Split repositories are the most familiar public Packagist experience. The monorepo remains the source of truth; generated package repositories should not accept direct source edits.

## Backward compatibility rules

- Existing `MnbExcel` calls continue through the full or application package.
- Modular users should prefer `SpreadsheetManager` or `Format\Csv`, `Format\Json`, `Format\Xml`, `Format\Xlsx`, `Format\Ods`, and `Format\Xls`.
- All packages use the same `Mnb\PHPExcel\` namespace. Composer merges their PSR-4 directories.
- A format package must never copy classes owned by another module.
- Module versions should be tagged together during the 1.x line.
