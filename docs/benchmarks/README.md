# MNB PHPExcel Benchmark Proof Suite

This directory documents the reproducible benchmark plan for release testing.
Do not hard-code marketing numbers copied from another machine. Run the same commands locally and publish the machine/PHP/extensions used.

## Recommended benchmark matrix

| Dataset | Operation | Recommended runtime |
|---:|---|---|
| 100,000 rows x 10 columns | large XLSX writer, large reader, DB import | CLI or worker |
| 500,000 rows x 10 columns | large XLSX writer, large reader, DB import | CLI/queue |
| 1,000,000 rows x 10 columns | large XLSX writer, large reader, DB import | CLI/queue only |

## Internal MNB benchmark

```bash
php tools/benchmark-large-writer.php --rows=100000 --cols=10 --json
php tools/benchmark-large-writer.php --rows=500000 --cols=10 --json
php tools/benchmark-large-writer.php --rows=1000000 --cols=10 --json
```

The output includes:

```text
elapsed_seconds
peak_memory_mb
output_size_mb
rows_per_second
```

## Optional comparison libraries

Run third-party comparison benchmarks in a separate benchmark workspace, not in the clean release ZIP.

```bash
composer require --dev phpoffice/phpspreadsheet
composer require --dev openspout/openspout
composer require --dev mk-j/php_xlsxwriter
composer require --dev rap2hpoutre/fast-excel
```

Compare only results generated on the same machine, same PHP version, same dataset, and same storage disk.

## Publish-ready performance table template

| Library | 100k rows time | 100k peak MB | 500k rows time | 500k peak MB | 1M rows time | 1M peak MB | Notes |
|---|---:|---:|---:|---:|---:|---:|---|
| MNB PHPExcel large writer | measured locally | measured locally | measured locally | measured locally | measured locally | measured locally | generator + inline strings |
| PhpSpreadsheet | measured locally | measured locally | measured locally | measured locally | measured locally | measured locally | optional comparison |
| OpenSpout | measured locally | measured locally | measured locally | measured locally | measured locally | measured locally | optional comparison |
| PHP_XLSXWriter | measured locally | measured locally | measured locally | measured locally | measured locally | measured locally | optional comparison |
| FastExcel | measured locally | measured locally | measured locally | measured locally | measured locally | measured locally | optional comparison |

