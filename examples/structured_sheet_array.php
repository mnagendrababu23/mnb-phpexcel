<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
$sheet = $argv[2] ?? 1;

if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/structured_sheet_array.php <workbook.xlsx> [sheet-name-or-number]\n");
    exit(1);
}

$selectedSheet = is_numeric($sheet) ? (int) $sheet : $sheet;

$output = MnbExcel::read($path)
    ->sheet($selectedSheet)
    ->toStructuredSheetArray([
        'header_row' => true,
        'header_case' => 'snake',
        'skip_empty_rows' => true,
        'include_cell_metadata' => true,
    ]);

print_r([
    'sheet' => $output['sheet'],
    'headers' => $output['headers'],
    'columns' => $output['columns'],
    'summary' => $output['summary'],
    'warnings' => $output['warnings'],
]);
