<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'marks' => 95],
    ['name' => '', 'email' => 'bad-email', 'marks' => 'A+'],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'marks' => 101],
];

$result = MnbExcel::validateArray($rows, [
    'name' => 'required|string|max:100',
    'email' => 'required|email',
    'marks' => 'required|numeric|min:0|max:100',
]);

MnbExcel::fromFailedRows($result['failed'])
    ->columnWidths([
        'A' => 12,
        'B' => 45,
        'C' => 20,
        'D' => 30,
        'E' => 12,
    ])
    ->freezeHeader()
    ->autoFilter()
    ->save(__DIR__ . '/output/failed-rows-report.xlsx');

echo "Valid rows: " . count($result['valid']) . PHP_EOL;
echo "Failed rows: " . count($result['failed']) . PHP_EOL;
