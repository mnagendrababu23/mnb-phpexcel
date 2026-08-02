<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Csv;
use Mnb\PHPExcel\Format\Xls;
use Mnb\PHPExcel\MnbExcel;

echo "CsvXlsMetadataSmokeTest\n";
$dir = smoke_temp_dir('csv-xls-metadata');

smoke_run('returns normalized full CSV metadata', function () use ($dir): void {
    $path = $dir . DIRECTORY_SEPARATOR . 'metadata.csv';
    file_put_contents($path, "Name,Age,Formula\r\nAlice,30,=1+1\r\nBob,25,ok\r\nBad,row\r\n");
    $meta = Csv::metaInfo($path, ['profile' => 'full']);
    smoke_assert_equals('csv', $meta['format'] ?? null, 'CSV format should be normalized');
    smoke_assert_equals(4, $meta['statistics']['row_count'] ?? null, 'CSV row count should be complete');
    smoke_assert_equals(1, $meta['statistics']['ragged_row_count'] ?? null, 'CSV ragged row should be detected');
    smoke_assert_equals(1, $meta['security']['formula_like_cell_count'] ?? null, 'CSV formula-like cell should be reported');
    smoke_assert_equals('not_applicable', $meta['document']['state'] ?? null, 'CSV document properties should be not applicable');
    $generic = MnbExcel::metaInfo($path, ['profile' => 'standard']);
    smoke_assert_equals(4, $generic['statistics']['row_count'] ?? null, 'Generic facade should route CSV metadata');
});

smoke_run('writes reads and inspects native BIFF8 metadata', function () use ($dir): void {
    $path = $dir . DIRECTORY_SEPARATOR . 'metadata.xls';
    Xls::write(new WorkbookData([
        new WorksheetData('Data', [
            ['Name', 'Amount', 'Formula'],
            ['Alice', 10, CellValue::formula('B2*2', 20)],
        ], mergeCells: ['A4:B4']),
        new WorksheetData('Second', [['x']]),
    ], [
        'title' => 'Native XLS metadata',
        'creator' => 'MNB Author',
        'company' => 'MNB Company',
        'custom_properties' => ['Project ID' => 'P-1'],
    ]), $path);
    $meta = Xls::metaInfo($path, ['profile' => 'full']);
    smoke_assert_equals('xls', $meta['format'] ?? null, 'XLS format should be normalized');
    smoke_assert_equals(2, $meta['workbook']['sheet_count'] ?? null, 'XLS sheet count should be complete');
    smoke_assert_equals(1, $meta['statistics']['formula_count'] ?? null, 'XLS formula count should be reported');
    smoke_assert_equals(1, $meta['statistics']['merged_range_count'] ?? null, 'XLS merged ranges should be reported');
    smoke_assert_equals(false, $meta['security']['encrypted'] ?? null, 'Native XLS should report encryption state');
    $generic = MnbExcel::metaInfo($path, ['profile' => 'quick']);
    smoke_assert_equals(2, $generic['workbook']['sheet_count'] ?? null, 'Generic facade should route XLS metadata');
    $rows = MnbExcel::readXls($path)->sheet('Data')->toArray();
    smoke_assert_equals('Alice', $rows[1][0] ?? null, 'Mono should use the native XLS reader');

    $updated = $dir . DIRECTORY_SEPARATOR . 'metadata-updated.xls';
    MnbExcel::updateMetaInfo($path, $updated, [
        'document' => ['title' => 'Updated native XLS', 'creator' => 'Updated Author'],
        'application' => ['company' => 'Updated Company'],
        'custom_properties' => ['Project ID' => 'P-2', 'Approved' => true],
        'workbook' => ['active_sheet' => 'Second', 'sheet_visibility' => ['Data' => 'hidden']],
        'calculation' => ['mode' => 'manual', 'iterate_count' => 12],
    ]);
    $updatedMeta = MnbExcel::metaInfo($updated, ['profile' => 'full']);
    smoke_assert_equals('Updated native XLS', $updatedMeta['document']['title'] ?? null, 'Generic facade should update XLS document metadata');
    smoke_assert_equals('Updated Company', $updatedMeta['application']['company'] ?? null, 'Generic facade should update XLS application metadata');
    smoke_assert_equals('Second', $updatedMeta['workbook']['active_sheet_name'] ?? null, 'Generic facade should update active XLS sheet');
    smoke_assert_equals(1, $updatedMeta['hidden_content']['hidden_sheet_count'] ?? null, 'Generic facade should update XLS visibility');
    smoke_assert_equals('manual', $updatedMeta['calculation']['mode'] ?? null, 'Generic facade should update XLS calculation mode');

    $clean = $dir . DIRECTORY_SEPARATOR . 'metadata-clean.xls';
    MnbExcel::removePersonalInfo($updated, $clean);
    $cleanMeta = MnbExcel::metaInfo($clean, ['profile' => 'full']);
    smoke_assert_true(!array_key_exists('creator', $cleanMeta['document']), 'XLS personal-information removal should clear creator');
    smoke_assert_true(!array_key_exists('company', $cleanMeta['application']), 'XLS personal-information removal should clear company');
    smoke_assert_equals(0, $cleanMeta['custom_properties']['count'] ?? null, 'XLS personal-information removal should clear custom properties');
});

echo "CsvXlsMetadataSmokeTest: all assertions passed.\n";
