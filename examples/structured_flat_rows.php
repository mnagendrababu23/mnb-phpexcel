<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/structured_flat_rows.php <workbook.xlsx>\n");
    exit(1);
}

$sheet = MnbExcel::read($path)
    ->sheet(1)
    ->toStructuredArray([
        'structure' => 'sheet',
        'header_row' => true,
        'header_case' => 'snake',
        'row_format' => 'flat',
        'preserve_original_row_numbers' => true,
    ]);

foreach ($sheet['rows'] as $row) {
    print_r($row);
}
