<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

echo "ReaderEnhancementSmokeTest\n";

smoke_run('auto-detects CSV and its delimiter', function (): void {
    $dir = smoke_temp_dir('reader-auto-csv');
    $path = $dir . DIRECTORY_SEPARATOR . 'people.data';
    file_put_contents($path, "id;name;city\n1;Ravi;Hyderabad\n2;Sita;Pune\n3;Asha;Delhi\n");

    smoke_assert_equals('csv', MnbExcel::detectFormat($path), 'content detection should identify CSV');
    $rows = MnbExcel::read($path, ['delimiter' => 'auto'])
        ->withHeaderRow()
        ->skip(1)
        ->limit(1)
        ->selectColumns(['name', 'city'])
        ->toArray();

    smoke_assert_equals([['name' => 'Sita', 'city' => 'Pune']], $rows, 'fluent reader transforms should apply after the header');
});

smoke_run('supports first, countRows, eachRow, and chunk', function (): void {
    $dir = smoke_temp_dir('reader-iteration');
    $path = $dir . DIRECTORY_SEPARATOR . 'people.csv';
    file_put_contents($path, "id,name\n1,Ravi\n2,Sita\n3,Asha\n");

    $session = MnbExcel::read($path)->withHeaderRow();
    smoke_assert_equals(['id' => '1', 'name' => 'Ravi'], $session->first(), 'first() should return the first normalized data row');
    smoke_assert_equals(3, $session->countRows(), 'countRows() should count normalized data rows');

    $seen = [];
    $state = $session->eachRow(function (array $row) use (&$seen): bool {
        $seen[] = $row['name'];
        return count($seen) < 2;
    });
    smoke_assert_equals(['Ravi', 'Sita'], $seen, 'eachRow() should stop when callback returns false');
    smoke_assert_true($state['stopped'], 'eachRow() should report callback stop');

    $chunks = [];
    $chunkState = $session->chunk(2, function (array $rows) use (&$chunks): void {
        $chunks[] = array_column($rows, 'name');
    });
    smoke_assert_equals([['Ravi', 'Sita'], ['Asha']], $chunks, 'chunk() should deliver the final partial chunk');
    smoke_assert_equals(2, $chunkState['chunks'], 'chunk() should report delivered chunks');
});

smoke_run('reads JSON Lines and preserves large integer text', function (): void {
    $dir = smoke_temp_dir('reader-ndjson');
    $path = $dir . DIRECTORY_SEPARATOR . 'events.ndjson';
    file_put_contents($path, "{\"id\":12345678901234567890,\"event\":\"created\"}\n{\"id\":2,\"event\":\"updated\"}\n");

    smoke_assert_equals('json', MnbExcel::detectFormat($path), 'NDJSON extension should map to JSON reader');
    $rows = MnbExcel::read($path)->withHeaderRow()->toArray();
    smoke_assert_equals('12345678901234567890', $rows[0]['id'] ?? null, 'large JSON integers should remain strings');
    smoke_assert_equals('updated', $rows[1]['event'] ?? null, 'second JSON Lines record should be read');
});

smoke_run('rejects inconsistent CSV rows in strict mode', function (): void {
    $dir = smoke_temp_dir('reader-strict-csv');
    $path = $dir . DIRECTORY_SEPARATOR . 'broken.csv';
    file_put_contents($path, "id,name\n1,Ravi\n2,Sita,extra\n");

    $thrown = false;
    try {
        MnbExcel::readCsv($path, ['strict_column_count' => true])->toArray();
    } catch (MnbExcelException $e) {
        $thrown = str_contains($e->getMessage(), 'row 3');
    }
    smoke_assert_true($thrown, 'strict CSV mode should report the inconsistent physical row');
});

smoke_run('round-trips library XML when ext-xmlreader is available', function (): void {
    if (!class_exists(XMLReader::class)) {
        echo ' SKIP ext-xmlreader unavailable';
        return;
    }

    $dir = smoke_temp_dir('reader-xml');
    $path = $dir . DIRECTORY_SEPARATOR . 'people.xml';
    MnbExcel::fromArray([
        ['id' => '001', 'name' => 'Ravi'],
        ['id' => '002', 'name' => 'Sita'],
    ])->saveXml($path);

    smoke_assert_equals('xml', MnbExcel::detectFormat($path), 'XML signature should be detected');
    $rows = MnbExcel::read($path)->withHeaderRow()->toArray();
    smoke_assert_equals('001', $rows[0]['id'] ?? null, 'XML associative keys should become headers');
    smoke_assert_equals('Sita', $rows[1]['name'] ?? null, 'XML rows should round-trip');
});

echo "ReaderEnhancementSmokeTest: all assertions passed.\n";
