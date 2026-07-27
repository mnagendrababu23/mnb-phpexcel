<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');
$sheet = $argv[2] ?? 1;

// Default mode counts rows containing at least one cell by streaming XML.
$filledRows = MnbExcel::rowCount($path, $sheet);

// Other modes are useful when formatting or sparse row indexes matter.
$physicalRows = MnbExcel::rowCount($path, $sheet, ['mode' => 'physical']);
$lastUsedRow = MnbExcel::rowCount($path, $sheet, ['mode' => 'last_row']);
$declaredLastRow = MnbExcel::rowCount($path, $sheet, ['mode' => 'declared']);
$allSheetCounts = MnbExcel::rowCounts($path);

print_r([
    'selected_sheet' => $sheet,
    'filled_rows' => $filledRows,
    'physical_rows' => $physicalRows,
    'last_used_row' => $lastUsedRow,
    'declared_last_row' => $declaredLastRow,
    'all_sheets' => $allSheetCounts,
]);
