<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

require __DIR__ . '/bootstrap.php';

$tmp = sys_get_temp_dir() . '/mnb-phpexcel-locale.csv';

$rows = [
    ['name' => 'Ravi', 'amount' => '1.234,50', 'started' => '31.12.2025', 'note' => '=cmd'],
    ['name' => 'Sita', 'amount' => '2.500,75', 'started' => '01.01.2026', 'note' => 'safe'],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->csvDialect('semicolon')
    ->csvBom(false)
    ->csvInjectionPolicy('strip')
    ->save($tmp);

$raw = file_get_contents($tmp);
assert(is_string($raw));
assert(!str_starts_with($raw, "\xEF\xBB\xBF"));
assert(str_contains($raw, ';'));
assert(!str_contains($raw, '=cmd'));

$read = MnbExcel::readCsv($tmp, [
    'dialect' => 'semicolon',
    'trim_values' => true,
])->toArray([
    'header_row' => true,
    'locale' => 'de_DE',
    'number_columns' => ['amount'],
    'date_columns' => ['started' => 'Y-m-d'],
]);

assert($read[0]['amount'] === 1234.5);
assert($read[1]['amount'] === 2500.75);
assert($read[0]['started'] === '2025-12-31');
assert($read[0]['note'] === 'cmd');

$blocked = false;
try {
    MnbExcel::fromArray([['danger' => '=HYPERLINK("http://evil.test","Click")']])
        ->withHeader()
        ->csvInjectionPolicy('block')
        ->save(sys_get_temp_dir() . '/mnb-phpexcel-block.csv');
} catch (MnbExcelException) {
    $blocked = true;
}
assert($blocked === true);

unlink($tmp);

echo "CSV locale smoke test passed.\n";
