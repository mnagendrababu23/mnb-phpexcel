<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

$path = __DIR__ . '/output/legacy-users.csv';

$encoding = MnbExcel::detectEncoding($path);
print_r($encoding);

$rows = MnbExcel::readCsv($path, [
    'encoding' => 'auto',
    'trim_values' => true,
])->toArray([
    'header_row' => true,
    'skip_empty_rows' => true,
]);

print_r(array_slice($rows, 0, 3));

MnbExcel::fromArray($rows)
    ->withHeader()
    ->csvEncoding('UTF-8')
    ->csvBom(true)
    ->save(__DIR__ . '/output/legacy-users-normalized.csv');
