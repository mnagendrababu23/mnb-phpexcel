<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');

// Fast mode reads worksheet names, states, dimensions, and XML package sizes.
// It does not scan cell values or build row arrays.
$fast = MnbExcel::sheetsInfo($path);

// Accurate mode streams worksheet XML tags to count physical rows, filled rows,
// cells, and the last used row/column without hydrating workbook rows.
$accurate = MnbExcel::sheetsInfo($path, [
    'accurate_row_count' => true,
]);

print_r([
    'fast' => $fast,
    'accurate' => $accurate,
]);
