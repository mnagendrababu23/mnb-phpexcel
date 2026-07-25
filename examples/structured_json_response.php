<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$file = __DIR__ . '/students.xlsx';

// Use this in an API/controller when you want JSON output without saving a file.
header('Content-Type: application/json; charset=UTF-8');

echo MnbExcel::read($file)
    ->toStructuredJson([
        'header_row' => true,
        'skip_empty_rows' => true,
        'include_hidden_rows' => false,
        'include_hidden_columns' => false,
        'header_case' => 'snake',
    ], [
        'pretty' => false,
        'trailing_newline' => false,
    ]);

exit;
