<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$rows = MnbExcel::read(__DIR__ . '/students.xlsx')
    ->sheet(1)
    ->toArray([
        'header_row' => true,
        'skip_empty_rows' => true,
        'max_rows' => 5000,
    ]);

print_r($rows);
