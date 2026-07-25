<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$utf16 = sys_get_temp_dir() . '/mnb-phpexcel-utf16.csv';
$cp1252 = sys_get_temp_dir() . '/mnb-phpexcel-cp1252.csv';

$rows = [
    ['name' => 'Nagendra', 'city' => 'Hyderabad', 'note' => '₹ symbol'],
    ['name' => 'Sita', 'city' => 'Bengaluru', 'note' => 'safe'],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->csvEncoding('UTF-16LE')
    ->csvBom(true)
    ->save($utf16);

$detectedUtf16 = MnbExcel::detectEncoding($utf16);
assert($detectedUtf16['encoding'] === 'UTF-16LE');
assert($detectedUtf16['bom'] === true);

$readUtf16 = MnbExcel::readCsv($utf16, [
    'encoding' => 'auto',
    'trim_values' => true,
])->toArray(['header_row' => true]);

assert($readUtf16[0]['name'] === 'Nagendra');
assert($readUtf16[0]['city'] === 'Hyderabad');

$cp1252Content = "name,note\r\n" . "Ravi,\x93smart quotes\x94 and \x80 euro\r\n";
file_put_contents($cp1252, $cp1252Content);

$detectedCp = MnbExcel::detectEncoding($cp1252);
assert(in_array($detectedCp['encoding'], ['Windows-1252', 'ISO-8859-1'], true));

$readCp = MnbExcel::readCsv($cp1252, [
    'encoding' => 'auto',
])->toArray(['header_row' => true]);

assert($readCp[0]['name'] === 'Ravi');
assert(str_contains((string) $readCp[0]['note'], 'smart quotes'));

@unlink($utf16);
@unlink($cp1252);

echo "Encoding auto-detection smoke test passed.\n";
