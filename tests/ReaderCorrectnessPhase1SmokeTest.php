<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Reader\XlsxStyleMap;
use Mnb\PHPExcel\Reader\XmlReader;
use Mnb\PHPExcel\Support\MnbExcelException;

echo "ReaderCorrectnessPhase1SmokeTest\n";

smoke_run('allocates collision-free duplicate header names', function (): void {
    $dir = smoke_temp_dir('reader-duplicate-headers');
    $path = $dir . DIRECTORY_SEPARATOR . 'duplicates.csv';
    file_put_contents($path, "a,a,a_2\n1,2,3\n");

    $rows = MnbExcel::readCsv($path)->withHeaderRow()->toArray();
    smoke_assert_equals(
        [['a' => '1', 'a_2' => '2', 'a_2_2' => '3']],
        $rows,
        'renamed duplicate headers must never collide with a real header'
    );

    $thrown = false;
    try {
        MnbExcel::readCsv($path, ['duplicate_headers' => 'error'])->withHeaderRow()->toArray();
    } catch (MnbExcelException $e) {
        $thrown = str_contains($e->getMessage(), 'Duplicate header');
    }
    smoke_assert_true($thrown, 'duplicate_headers=error should reject duplicate headers');
});

smoke_run('preserves large integers and precision-sensitive decimals', function (): void {
    $dir = smoke_temp_dir('reader-safe-numbers');
    $path = $dir . DIRECTORY_SEPARATOR . 'numbers.csv';
    file_put_contents($path, "id,small,amount\n12345678901234567890,42,123456789012345.67\n");

    $rows = MnbExcel::readCsv($path, [
        'integer_columns' => ['id', 'small'],
        'number_columns' => ['amount'],
    ])->withHeaderRow()->toArray();

    smoke_assert_equals('12345678901234567890', $rows[0]['id'] ?? null, 'oversized integer must remain exact text');
    smoke_assert_equals(42, $rows[0]['small'] ?? null, 'safe integer should still become an integer');
    smoke_assert_equals('123456789012345.67', $rows[0]['amount'] ?? null, 'precision-sensitive decimal must remain exact text');

    $thrown = false;
    try {
        MnbExcel::readCsv($path, [
            'integer_columns' => ['id'],
            'big_integer_mode' => 'error',
        ])->withHeaderRow()->toArray();
    } catch (MnbExcelException $e) {
        $thrown = str_contains($e->getMessage(), 'exceeds the PHP integer range');
    }
    smoke_assert_true($thrown, 'big_integer_mode=error should reject oversized integers');

    $xlsxMethod = (new ReflectionClass(XlsxReader::class))->getMethod('cellValueFromXml');
    $xlsxMethod->setAccessible(true);
    $xlsxValue = $xlsxMethod->invoke(
        new XlsxReader(),
        '<c><v>12345678901234567890</v></c>',
        '',
        null,
        [],
        XlsxStyleMap::fromXml(null),
        false,
        []
    );
    smoke_assert_equals('12345678901234567890', $xlsxValue, 'XLSX numeric decoding should use the safe integer path');

    $xmlMethod = (new ReflectionClass(XmlReader::class))->getMethod('inferValue');
    $xmlMethod->setAccessible(true);
    $xmlValue = $xmlMethod->invoke(new XmlReader(), '12345678901234567890', ['infer_types' => true]);
    smoke_assert_equals('12345678901234567890', $xmlValue, 'XML type inference should use the safe integer path');
    $leadingZero = $xmlMethod->invoke(new XmlReader(), '00123', ['infer_types' => true]);
    smoke_assert_equals('00123', $leadingZero, 'XML inference should preserve leading-zero identifiers by default');
});

smoke_run('applies extra and missing column policies', function (): void {
    $dir = smoke_temp_dir('reader-column-policies');
    $extraPath = $dir . DIRECTORY_SEPARATOR . 'extra.csv';
    file_put_contents($extraPath, "id,name\n1,Ravi,admin\n2,Sita\n");

    $generated = MnbExcel::readCsv($extraPath, ['extra_columns' => 'generate'])
        ->withHeaderRow()
        ->toArray();
    smoke_assert_equals(
        [
            ['id' => '1', 'name' => 'Ravi', 'column_3' => 'admin'],
            ['id' => '2', 'name' => 'Sita', 'column_3' => null],
        ],
        $generated,
        'extra_columns=generate should create stable generated fields and backfill prior rows'
    );

    $generatedWithFluentRange = MnbExcel::readCsv($extraPath, ['extra_columns' => 'generate'])
        ->withHeaderRow()
        ->skip(0)
        ->limit(2)
        ->toArray();
    smoke_assert_equals($generated, $generatedWithFluentRange, 'generated columns should stay stable through fluent range operations');

    $collected = MnbExcel::readCsv($extraPath, ['extra_columns' => 'collect'])
        ->withHeaderRow()
        ->toArray();
    smoke_assert_equals(['admin'], $collected[0]['_extra'] ?? null, 'extra_columns=collect should preserve surplus values');
    smoke_assert_equals([], $collected[1]['_extra'] ?? null, 'extra_columns=collect should keep a stable empty collection on normal rows');

    $extraThrown = false;
    try {
        MnbExcel::readCsv($extraPath, ['extra_columns' => 'error'])->withHeaderRow()->toArray();
    } catch (MnbExcelException $e) {
        $extraThrown = str_contains($e->getMessage(), 'header defines 2');
    }
    smoke_assert_true($extraThrown, 'extra_columns=error should reject surplus values');

    $missingPath = $dir . DIRECTORY_SEPARATOR . 'missing.csv';
    file_put_contents($missingPath, "id,name,role\n1,Ravi\n");
    $missingThrown = false;
    try {
        MnbExcel::readCsv($missingPath, ['missing_columns' => 'error'])->withHeaderRow()->toArray();
    } catch (MnbExcelException $e) {
        $missingThrown = str_contains($e->getMessage(), 'expected at least 3');
    }
    smoke_assert_true($missingThrown, 'missing_columns=error should reject short rows');
});

smoke_run('detects NDJSON content without relying on filename suffix', function (): void {
    $dir = smoke_temp_dir('reader-ndjson-content');
    $path = $dir . DIRECTORY_SEPARATOR . 'events.data';
    file_put_contents($path, "{\"id\":1,\"event\":\"created\"}\n{\"id\":2,\"event\":\"updated\"}\n");

    smoke_assert_equals('json', MnbExcel::detectFormat($path), 'NDJSON object content should select the JSON reader');
    $rows = MnbExcel::read($path)->withHeaderRow()->toArray();
    smoke_assert_equals(2, count($rows), 'content-detected NDJSON should return both records');
    smoke_assert_equals('updated', $rows[1]['event'] ?? null, 'second NDJSON record should be decoded');

    $scalarPath = $dir . DIRECTORY_SEPARATOR . 'values.data';
    file_put_contents($scalarPath, "1\n2\n");
    smoke_assert_equals('json', MnbExcel::detectFormat($scalarPath), 'multiple standalone JSON scalar records should be detected as JSON Lines');
});

smoke_run('separates physical and normalized data header positions', function (): void {
    $dir = smoke_temp_dir('reader-header-semantics');
    $path = $dir . DIRECTORY_SEPARATOR . 'report.csv';
    file_put_contents($path, "\nReport title\nid,name\n1,Ravi\n");

    $physical = MnbExcel::readCsv($path)->headerAtPhysicalRow(3)->toArray();
    smoke_assert_equals([['id' => '1', 'name' => 'Ravi']], $physical, 'physical header row should use the exact source row');

    $data = MnbExcel::readCsv($path)->headerAtDataRow(2)->toArray();
    smoke_assert_equals([['id' => '1', 'name' => 'Ravi']], $data, 'data header row should count normalized non-empty rows');

    $alias = MnbExcel::readCsv($path)->withHeaderRow(2)->toArray();
    smoke_assert_equals($data, $alias, 'withHeaderRow should have explicit normalized-data semantics');

    $legacy = MnbExcel::readCsv($path, ['header_row' => 2])->toArray();
    smoke_assert_true(array_key_exists('report_title', $legacy[0] ?? []), 'raw header_row options should retain legacy physical-first behavior');

    $simplePath = $dir . DIRECTORY_SEPARATOR . 'simple.csv';
    file_put_contents($simplePath, "\n\nid,name\n1,Ravi\n");
    $first = MnbExcel::readCsv($simplePath)->firstNonEmptyRowAsHeader()->toArray();
    smoke_assert_equals([['id' => '1', 'name' => 'Ravi']], $first, 'firstNonEmptyRowAsHeader should ignore leading blank rows');
});

echo "ReaderCorrectnessPhase1SmokeTest: all assertions passed.\n";
