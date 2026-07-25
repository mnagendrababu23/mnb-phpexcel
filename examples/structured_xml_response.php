<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$file = __DIR__ . '/students.xlsx';

// Use this in an API/controller when you want XML output without saving a file.
header('Content-Type: application/xml; charset=UTF-8');

echo MnbExcel::read($file)
    ->toStructuredXml([
        'header_row' => true,
        'skip_empty_rows' => true,
        'include_hidden_rows' => false,
        'include_hidden_columns' => false,
        'header_case' => 'snake',
    ], [
        'root' => 'structured_workbook',
        'pretty' => true,
    ]);

exit;
