<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/structured_array.php <workbook.xlsx>\n");
    exit(1);
}

$workbook = MnbExcel::read($path)->toStructuredArray([
    'header_row' => true,
    'header_case' => 'snake',
    'skip_empty_rows' => true,
    'preserve_original_row_numbers' => true,
]);

printf(
    "Status: %s; sheets: %d; data rows: %d\n",
    (string) ($workbook['status'] ?? 'unknown'),
    (int) ($workbook['summary']['sheet_count'] ?? 0),
    (int) ($workbook['summary']['data_rows'] ?? 0),
);

print_r($workbook);
