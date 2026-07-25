<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$tmpDir = sys_get_temp_dir();
$jsonPath = $tmpDir . '/mnb-phpexcel-json-import.json';
$csvPath = $tmpDir . '/mnb-phpexcel-json-import.csv';
$xmlPath = $tmpDir . '/mnb-phpexcel-json-import.xml';
$jsonOutPath = $tmpDir . '/mnb-phpexcel-json-import-roundtrip.json';

$data = [
    [
        'student_id' => '000123456789012345',
        'name' => 'Ravi',
        'address' => ['city' => 'Hyderabad', 'state' => 'Telangana'],
        'phone' => '0987654321',
        'marks' => 95,
    ],
    [
        'student_id' => '000999999999999999',
        'name' => 'Sita',
        'address' => ['city' => 'Vijayawada', 'state' => 'Andhra Pradesh'],
        'phone' => '0912345678',
        'marks' => 88,
        'extra' => 'mixed key should stay',
    ],
];

file_put_contents($jsonPath, json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR));

MnbExcel::fromJson($jsonPath)
    ->textColumns(['student_id', 'phone'])
    ->autoWidth()
    ->save($csvPath);

$rows = MnbExcel::readCsv($csvPath)->toArray(['header_row' => true]);
assert($rows[0]['student_id'] === '000123456789012345');
assert($rows[0]['address_city'] === 'Hyderabad');
assert($rows[1]['extra'] === 'mixed key should stay');

$readJsonRows = MnbExcel::readJson($jsonPath)->toArray(['header_row' => true]);
assert($readJsonRows[0]['address_city'] === 'Hyderabad');
assert($readJsonRows[1]['student_id'] === '000999999999999999');

MnbExcel::fromJson($jsonPath)
    ->saveXml($xmlPath, ['root' => 'students', 'row' => 'student']);
$xml = file_get_contents($xmlPath);
assert(is_string($xml));
assert(str_contains($xml, '<address.city>Hyderabad</address.city>')); // dots are valid XML name characters
assert(str_contains($xml, 'Hyderabad'));

$workbookJsonPath = $tmpDir . '/mnb-phpexcel-json-import-workbook.json';
file_put_contents($workbookJsonPath, json_encode([
    'Students' => $data,
    'Teachers' => [
        ['name' => 'Kumar', 'subject' => 'Maths'],
    ],
], JSON_PRETTY_PRINT | JSON_THROW_ON_ERROR));

$names = MnbExcel::readJson($workbookJsonPath)->sheetNames();
assert($names === ['Students', 'Teachers']);
$teachers = MnbExcel::readJson($workbookJsonPath)->sheet('Teachers')->toArray(['header_row' => true]);
assert($teachers[0]['subject'] === 'Maths');

MnbExcel::fromJson($workbookJsonPath)->saveJson($jsonOutPath, ['include_sheet_names' => false]);
$roundtrip = json_decode((string) file_get_contents($jsonOutPath), true, flags: JSON_THROW_ON_ERROR);
assert(isset($roundtrip['sheets']['Students']));

@unlink($jsonPath);
@unlink($csvPath);
@unlink($xmlPath);
@unlink($jsonOutPath);
@unlink($workbookJsonPath);

echo "JSON import smoke test passed.\n";
