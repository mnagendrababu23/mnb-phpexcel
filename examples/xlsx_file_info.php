<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');

// Reads file stat, ZIP package totals, document properties, sheet count, and
// sheet names. Worksheet cell rows are not loaded into PHP arrays.
$info = MnbExcel::fileInfo($path);

print_r([
    'name' => $info['name'],
    'size_bytes' => $info['size_bytes'],
    'encrypted' => $info['encrypted'],
    'sheet_count' => $info['sheet_count'],
    'sheet_names' => $info['sheet_names'],
    'zip_entries' => $info['zip_entries'],
    'has_macros' => $info['has_macros'],
    'properties' => $info['properties'],
]);
