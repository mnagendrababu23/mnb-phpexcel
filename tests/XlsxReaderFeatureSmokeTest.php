<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class)) {
    echo "XLSX reader feature smoke test skipped: ext-zip/ext-xmlreader unavailable.\n";
    exit(0);
}

$file = sys_get_temp_dir() . '/mnb-phpexcel-reader-feature-smoke.xlsx';

MnbExcel::fromWorkbookArray([
    'Students' => [
        ['name' => 'Ravi', 'dob' => '2020-01-15', 'marks' => 95],
        ['name' => 'Sita', 'dob' => '2020-01-16', 'marks' => 88],
    ],
    'Teachers' => [
        ['name' => 'Kumar', 'subject' => 'Maths'],
    ],
])
    ->withHeader()
    ->save($file);

$names = MnbExcel::sheetNames($file);
assert($names === ['Students', 'Teachers']);

$rows = MnbExcel::read($file)->sheet('Students')->toArray(['header_row' => true]);
assert(count($rows) === 2);
assert($rows[0]['name'] === 'Ravi');

$inspection = MnbExcel::inspect($file);
assert(in_array($inspection['status'], ['ok', 'warning'], true));
assert(count($inspection['sheets']) === 2);

@unlink($file);

echo "XLSX reader feature smoke test passed.\n";
