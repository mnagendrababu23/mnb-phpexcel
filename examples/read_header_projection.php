<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/read_header_projection.php <workbook.xlsx>\n");
    exit(1);
}

$session = MnbExcel::read($path)->sheet(1);
$detection = $session->detectHeader(['header_detection_rows' => 30]);

printf("Detected header at row %d with confidence %.2f\n", $detection->row, $detection->confidence);

$rows = $session
    ->range(startRow: 1, endRow: 5000, startColumn: 'A', endColumn: 'F')
    ->projectColumns(['A', 'B', 'F'])
    ->autoDetectHeader(sampleRows: 30, minimumConfidence: 0.45)
    ->toArray([
        'strict_header_detection' => true,
        'header_case' => 'snake',
    ]);

print_r(array_slice($rows, 0, 5));
