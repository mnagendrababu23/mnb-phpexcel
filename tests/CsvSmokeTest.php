<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "CsvSmokeTest\n";

$dir = smoke_temp_dir('csv');
$path = $dir . DIRECTORY_SEPARATOR . 'contacts.csv';

$rows = [
    ['name' => 'Ravi', 'phone' => '0987654321'],
    ['name' => 'Sita', 'phone' => '0912345678'],
];

smoke_run('writes a CSV file from an array', function () use ($rows, $path): void {
    $saved = MnbExcel::fromArray($rows)
        ->withHeader()
        ->textColumns(['phone'])
        ->save($path);

    smoke_assert_equals($path, $saved, 'save() should return the path it was given');
    smoke_assert_true(is_file($path), 'CSV file should exist after save()');
});

smoke_run('preserves leading zeros in a text column', function () use ($path): void {
    $contents = (string) file_get_contents($path);
    smoke_assert_contains('0987654321', $contents, 'phone number text should not be mangled into a number');
});

smoke_run('reads the CSV back into an array with the same row count', function () use ($rows, $path): void {
    $read = MnbExcel::readCsv($path)->withHeaderRow()->toArray();
    smoke_assert_equals(count($rows), count($read), 'row count should round-trip through CSV');
    smoke_assert_equals('0987654321', $read[0]['phone'] ?? null, 'header-aware CSV reads should return associative rows');
});

smoke_run('blocks formula-injection-style CSV content by default', function (): void {
    $unsafeRows = [
        ['name' => '=cmd|\'/c calc\'!A1', 'phone' => '123'],
    ];

    $issues = MnbExcel::scanCells($unsafeRows);
    smoke_assert_true(($issues['status'] ?? null) === 'warning', 'scanCells should flag formula-like text in plain string cells');
});

echo "CsvSmokeTest: all assertions passed.\n";
