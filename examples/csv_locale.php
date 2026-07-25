<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

$rows = [
    ['name' => 'Ravi', 'amount' => '1.234,50', 'started' => '31.12.2025', 'note' => '=cmd'],
    ['name' => 'Sita', 'amount' => '2.500,75', 'started' => '01.01.2026', 'note' => 'safe'],
];

$csv = __DIR__ . '/exports/locale-import.csv';

MnbExcel::fromArray($rows)
    ->withHeader()
    ->csvDialect('semicolon')
    ->csvBom(true)
    ->csvInjectionPolicy('escape')
    ->save($csv);

$cleanRows = MnbExcel::readCsv($csv, [
    'dialect' => 'semicolon',
    'encoding' => 'UTF-8',
    'trim_values' => true,
])->toArray([
    'header_row' => true,
    'locale' => 'de_DE',
    'number_columns' => ['amount'],
    'date_columns' => ['started' => 'Y-m-d'],
]);

print_r($cleanRows);
