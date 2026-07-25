<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "LargeWriterSmokeTest\n";

if (!extension_loaded('zip') || !extension_loaded('xmlreader')) {
    echo "LargeWriterSmokeTest: SKIP ext-zip/ext-xmlreader unavailable\n";
    exit(0);
}

/**
 * Deliberately small (not a real "large file") — this is a smoke test for
 * the streaming code path itself, not a performance benchmark. See
 * tools/benchmark-large-writer.php and docs/benchmarks/README.md for
 * actual large-scale performance runs.
 */
function smoke_row_generator(int $count): \Generator
{
    for ($i = 1; $i <= $count; $i++) {
        yield ['id' => $i, 'name' => 'Row ' . $i, 'amount' => $i * 1.5];
    }
}

$dir = smoke_temp_dir('large_writer');
$path = $dir . DIRECTORY_SEPARATOR . 'streamed.xlsx';
$rowCount = 250;

smoke_run('streams a generator into an XLSX file without buffering all rows', function () use ($path, $rowCount): void {
    $result = MnbExcel::largeExport(smoke_row_generator($rowCount))
        ->withHeader()
        ->sheetName('Rows')
        ->formatColumn('amount', 'decimal')
        ->save($path);

    smoke_assert_true(is_file($path), 'streaming writer should produce an XLSX file');
    smoke_assert_true(is_array($result), 'save() should return a result summary array');
});

smoke_run('reads the streamed file back with the large streaming reader', function () use ($path, $rowCount): void {
    $rows = [];
    MnbExcel::largeRead($path)
        ->withHeader()
        ->eachRow(function (array $row) use (&$rows): void {
            $rows[] = $row;
        });

    smoke_assert_equals($rowCount, count($rows), 'streamed row count should match what was written');
});

smoke_run('the large writer also passes save-time XLSX integrity validation', function () use ($path): void {
    $result = MnbExcel::validateXlsx($path);
    smoke_assert_true(($result['valid'] ?? false) === true, 'large-writer output should pass integrity validation: ' . json_encode($result));
});

echo "LargeWriterSmokeTest: all assertions passed.\n";
