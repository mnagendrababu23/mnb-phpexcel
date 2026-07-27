<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\MnbExcel;

smoke_run('reads XLSX file, sheet, and row information without hydrating workbook rows', static function (): void {
    $dir = smoke_temp_dir('xlsx_quick_info');
    $path = $dir . '/quick-info.xlsx';

    WorkbookBuilder::fromWorkbookArray([
        'Orders' => [
            ['Order ID' => 1001, 'Customer' => 'Alice', 'Total' => 125.50],
            ['Order ID' => 1002, 'Customer' => 'Bob', 'Total' => 80.00],
        ],
        'Notes' => [
            ['Message' => 'Generated for quick-info testing'],
        ],
    ])->withHeader()->save($path);

    $file = MnbExcel::fileInfo($path);
    smoke_assert_equals('ok', $file['status'], 'fileInfo should inspect a valid package');
    smoke_assert_equals(2, $file['sheet_count'], 'fileInfo should return the worksheet count');
    smoke_assert_equals(['Orders', 'Notes'], $file['sheet_names'], 'fileInfo should return worksheet names');
    smoke_assert_true($file['size_bytes'] > 0, 'fileInfo should return file size');
    smoke_assert_true($file['zip_entries'] > 0, 'fileInfo should return ZIP package entry count');

    $fastSheets = MnbExcel::sheetsInfo($path);
    smoke_assert_equals(2, count($fastSheets), 'sheetsInfo should return all worksheets');
    smoke_assert_equals('Orders', $fastSheets[0]['name'], 'sheetsInfo should preserve worksheet names');
    smoke_assert_equals(3, $fastSheets[0]['declared_last_row'], 'fast sheet info should read the worksheet dimension');
    smoke_assert_equals(false, $fastSheets[0]['accurate_row_count'], 'fast sheet info should not scan worksheet rows');

    $orders = MnbExcel::sheetInfo($path, 'Orders', ['accurate_row_count' => true]);
    smoke_assert_equals(3, $orders['filled_rows'], 'accurate sheet info should count header plus data rows');
    smoke_assert_equals(3, $orders['physical_rows'], 'accurate sheet info should count physical row elements');
    smoke_assert_equals(3, $orders['columns'], 'accurate sheet info should detect columns');
    smoke_assert_equals(9, $orders['cells'], 'accurate sheet info should count cells');

    smoke_assert_equals(3, MnbExcel::rowCount($path, 'Orders'), 'default rowCount should count filled rows');
    smoke_assert_equals(3, MnbExcel::rowCount($path, 'Orders', ['mode' => 'physical']), 'physical rowCount should count row tags');
    smoke_assert_equals(3, MnbExcel::rowCount($path, 'Orders', ['mode' => 'last_row']), 'last_row mode should return the highest row index');
    smoke_assert_equals(3, MnbExcel::rowCount($path, 'Orders', ['mode' => 'declared']), 'declared mode should use worksheet dimension only');

    $counts = MnbExcel::rowCounts($path);
    smoke_assert_equals(3, $counts['Orders'], 'rowCounts should return Orders count');
    smoke_assert_equals(2, $counts['Notes'], 'rowCounts should return Notes count');

    smoke_assert_equals(3, Xlsx::rowCount($path, 'Orders'), 'Xlsx facade should expose lightweight rowCount');
    smoke_assert_equals(2, count(Xlsx::sheetsInfo($path)), 'Xlsx facade should expose sheetsInfo');
});

echo "XlsxQuickInfoSmokeTest passed\n";
