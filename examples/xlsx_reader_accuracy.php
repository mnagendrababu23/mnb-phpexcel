<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$file = __DIR__ . '/students.xlsx';

$inspection = MnbExcel::inspect($file);
print_r($inspection);

$sheetNames = MnbExcel::sheetNames($file);
print_r($sheetNames);

$rows = MnbExcel::read($file)
    ->sheet($sheetNames[0] ?? 1) // supports sheet index or sheet name
    ->toArray([
        'header_row' => true,
        'skip_empty_rows' => true,
        'include_hidden_rows' => false,
        'format_dates' => true,
        'date_format' => 'Y-m-d',
        'datetime_format' => 'Y-m-d H:i:s',
        'formula_cells' => 'formula', // formula or cached_value
    ]);

print_r($rows);
