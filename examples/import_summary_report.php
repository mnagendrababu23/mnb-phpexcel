<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'marks' => 95],
    ['name' => '', 'email' => 'bad-email', 'marks' => 'A+'],
];

$preview = MnbExcel::previewImport($rows, [
    'required_columns' => ['name', 'email', 'marks'],
    'allowed_columns' => ['name', 'email', 'marks'],
    'strict_columns' => true,
]);

$result = MnbExcel::validateImport($rows, [
    'name' => 'required|string',
    'email' => 'required|email',
    'marks' => 'required|numeric|min:0|max:100',
]);

$output = MnbExcel::fromFailedRows($result['failed'])
    ->withImportSummarySheet($preview, $result)
    ->metadata([
        'title' => 'Import Error Report',
        'subject' => 'Failed rows with import summary sheet',
        'creator' => 'MNB PHPExcel',
    ])
    ->autoWidth(['min' => 12, 'max' => 50])
    ->saveSafe(__DIR__ . '/../storage/examples', 'Import Error Report', 'xlsx');

echo "Exported: {$output}\n";
