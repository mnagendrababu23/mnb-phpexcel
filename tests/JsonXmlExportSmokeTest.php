<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$tmpDir = sys_get_temp_dir();
$jsonPath = $tmpDir . '/mnb-phpexcel-json-export.json';
$xmlPath = $tmpDir . '/mnb-phpexcel-xml-export.xml';
$csvPath = $tmpDir . '/mnb-phpexcel-json-xml-source.csv';

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => MnbExcel::text('0987654321'), 'active' => true],
    ['name' => 'Sita & Co', 'email' => 'sita@example.com', 'phone' => MnbExcel::text('0912345678'), 'active' => false],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->saveJson($jsonPath, ['mode' => 'rows', 'preserve_associative_rows' => false]);

$json = file_get_contents($jsonPath);
assert(is_string($json));
$decoded = json_decode($json, true, flags: JSON_THROW_ON_ERROR);
assert($decoded[1][2] === '0987654321');
assert($decoded[2][0] === 'Sita & Co');

MnbExcel::fromArray($rows)
    ->withHeader()
    ->saveXml($xmlPath, ['root' => 'students', 'row' => 'student']);

$xml = file_get_contents($xmlPath);
assert(is_string($xml));
assert(str_contains($xml, '<students>'));
assert(str_contains($xml, 'Sita &amp; Co'));
assert(str_contains($xml, '0987654321'));

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->save($csvPath);

$jsonFromCsv = MnbExcel::readCsv($csvPath)->toJson(['header_row' => true]);
$decodedCsv = json_decode($jsonFromCsv, true, flags: JSON_THROW_ON_ERROR);
assert($decodedCsv[0]['phone'] === '0987654321');

$xmlFromCsv = MnbExcel::readCsv($csvPath)->toXml(['header_row' => true], ['root' => 'students', 'row' => 'student']);
assert(str_contains($xmlFromCsv, '<name>Ravi</name>'));
assert(str_contains($xmlFromCsv, '<phone>0987654321</phone>'));

@unlink($jsonPath);
@unlink($xmlPath);
@unlink($csvPath);

echo "JSON/XML export smoke test passed.\n";
