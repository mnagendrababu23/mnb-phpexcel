<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "ArrayToXlsxSmokeTest\n";

if (!extension_loaded('zip') || !extension_loaded('xmlreader')) {
    echo "ArrayToXlsxSmokeTest: SKIP ext-zip/ext-xmlreader unavailable\n";
    exit(0);
}

$dir = smoke_temp_dir('array_to_xlsx');
$path = $dir . DIRECTORY_SEPARATOR . 'students.xlsx';

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'marks' => 95],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'marks' => 88],
];

smoke_run('writes an XLSX file from an array', function () use ($rows, $path): void {
    $saved = MnbExcel::fromArray($rows)
        ->withHeader()
        ->columns([
            'name' => 'Student Name',
            'email' => 'Email Address',
            'marks' => 'Marks',
        ])
        ->numberColumns(['marks'])
        ->freezeHeader()
        ->autoFilter()
        ->save($path);

    smoke_assert_equals($path, $saved, 'save() should return the path it was given');
    smoke_assert_true(is_file($path), 'XLSX file should exist after save()');
    smoke_assert_true(filesize($path) > 0, 'XLSX file should not be empty');
});

smoke_run('reads the file back into the same shape of data', function () use ($rows, $path): void {
    $read = MnbExcel::read($path)->toArray();

    smoke_assert_equals(count($rows), count($read), 'row count should round-trip');
    smoke_assert_equals('Ravi', $read[0]['Student Name'] ?? $read[0]['name'] ?? null, 'first row name should round-trip');
    smoke_assert_equals(95, (int) ($read[0]['Marks'] ?? $read[0]['marks'] ?? 0), 'first row marks should round-trip as a number');
});

smoke_run('passes save-time XLSX integrity validation', function () use ($path): void {
    $result = MnbExcel::validateXlsx($path);
    smoke_assert_true(($result['valid'] ?? false) === true, 'generated XLSX should pass integrity validation: ' . json_encode($result));
});

echo "ArrayToXlsxSmokeTest: all assertions passed.\n";
