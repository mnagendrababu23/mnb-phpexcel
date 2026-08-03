<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Snapshot\VisualSnapshot;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

smoke_run('Unified visual snapshot and style round-trip', static function (): void {
    $dir = smoke_temp_dir('visual_snapshot');
    $xlsx = $dir . '/styled.xlsx';
    $xlsxRoundTrip = $dir . '/styled-roundtrip.xlsx';
    $xls = $dir . '/styled.xls';
    $csv = $dir . '/styled.csv';
    $encryptedXlsx = $dir . '/styled-encrypted.xlsx';

    $rows = [
        [
            'Order ID' => 'ORD-1001',
            'Order Date' => '2026-07-01',
            'Customer' => 'Aarav Retail',
            'Product' => 'Wireless Keyboard',
            'Quantity' => 4,
            'Unit Price' => 2499.00,
            'Total' => 9996.00,
            'Status' => 'Completed',
        ],
        [
            'Order ID' => 'ORD-1002',
            'Order Date' => '2026-07-03',
            'Customer' => 'Bright Stores',
            'Product' => 'Ergonomic Mouse',
            'Quantity' => 8,
            'Unit Price' => 1299.00,
            'Total' => 10392.00,
            'Status' => 'Completed',
        ],
    ];

    MnbExcel::report($rows)
        ->withHeader()
        ->styleHeader([
            'font' => ['bold' => true, 'color' => '#FFFFFF'],
            'fill' => ['color' => '#1F4E78'],
            'alignment' => ['horizontal' => 'center'],
        ])
        ->dateStyleColumns(['Order Date'], 'yyyy-mm-dd')
        ->integerColumns(['Quantity'])
        ->currencyColumns(['Unit Price', 'Total'], '₹')
        ->columnWidths(['A' => 15, 'B' => 14, 'C' => 22, 'D' => 28, 'E' => 12, 'F' => 15, 'G' => 16, 'H' => 14])
        ->rowHeight(1, 30)
        ->freezeHeader()
        ->autoFilter()
        ->rangeStyle('A2:H3', [
            'borders' => [
                'all' => ['style' => 'thin', 'color' => '#808080'],
            ],
        ])
        ->conditionalColorScale('G2:G3', ['#F8696B', '#FFEB84', '#63BE7B'])
        ->dataValidation('H2:H3', 'list', ['values' => ['Completed', 'Pending']])
        ->comment('A2', 'QA', 'Verified order')
        ->hyperlink('D2', 'https://example.com/products/keyboard', 'Product page')
        ->save($xlsx);

    $snapshot = MnbExcel::visualSnapshot($xlsx);
    smoke_assert_equals('mnb-phpexcel.visual-snapshot', $snapshot['schema'] ?? null, 'Snapshot schema mismatch');
    smoke_assert_equals('A1:H3', $snapshot['sheets'][0]['dimension'] ?? null, 'Snapshot lost Total or Status columns');
    smoke_assert_true(isset($snapshot['sheets'][0]['cells']['H3']), 'Snapshot must retain sparse H3 cell');
    smoke_assert_equals('date', $snapshot['sheets'][0]['cells']['B2']['type'] ?? null, 'Date column must be a typed date');
    smoke_assert_equals(14.0, $snapshot['sheets'][0]['layout']['column_widths']['H'] ?? null, 'Column width was not captured');
    smoke_assert_equals(30.0, $snapshot['sheets'][0]['layout']['row_heights'][1] ?? null, 'Row height was not captured');
    smoke_assert_equals(1, $snapshot['sheets'][0]['layout']['freeze_panes']['rows'] ?? null, 'Frozen header was not captured');
    smoke_assert_equals('A1:H3', $snapshot['sheets'][0]['layout']['auto_filter_range'] ?? null, 'Auto-filter was not captured');
    smoke_assert_equals('color_scale', $snapshot['sheets'][0]['conditional_formats'][0]['type'] ?? null, 'Color scale was not captured');
    smoke_assert_equals('list', $snapshot['sheets'][0]['data_validations'][0]['type'] ?? null, 'Data validation was not captured');
    smoke_assert_equals('Verified order', $snapshot['sheets'][0]['comments'][0]['text'] ?? null, 'Comment was not captured');
    smoke_assert_equals('https://example.com/products/keyboard', $snapshot['sheets'][0]['hyperlinks'][0]['url'] ?? null, 'Hyperlink was not captured');

    $headerStyleId = $snapshot['sheets'][0]['cells']['A1']['style_id'] ?? null;
    smoke_assert_true((is_int($headerStyleId) || is_string($headerStyleId)) && array_key_exists($headerStyleId, $snapshot['styles']), 'Header style ID is missing');
    $headerStyle = $snapshot['styles'][$headerStyleId] ?? [];
    smoke_assert_true(($headerStyle['font']['bold'] ?? false) === true, 'Header bold style was not read');
    smoke_assert_equals('FFFFFFFF', $headerStyle['font']['color']['rgb'] ?? null, 'Header font color was not read');
    smoke_assert_equals('FF1F4E78', $headerStyle['fill']['foreground']['rgb'] ?? null, 'Header fill was not read');

    $dataStyleId = $snapshot['sheets'][0]['cells']['A2']['style_id'] ?? null;
    smoke_assert_true((is_int($dataStyleId) || is_string($dataStyleId)) && array_key_exists($dataStyleId, $snapshot['styles']), 'Data style ID is missing');
    $dataStyle = $snapshot['styles'][$dataStyleId] ?? [];
    smoke_assert_equals('thin', $dataStyle['border']['left']['style'] ?? null, 'borders.all was not written/read');

    $zip = new ZipArchive();
    smoke_assert_true($zip->open($xlsx) === true, 'Unable to open generated XLSX package');
    $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    $zip->close();
    smoke_assert_true(is_string($sheetXml), 'Worksheet XML missing');
    smoke_assert_true(preg_match('/<c r="B2"[^>]*><v>\d+(?:\.\d+)?<\/v><\/c>/', $sheetXml) === 1, 'Date must be stored as a numeric Excel serial');

    $snapshotJson = VisualSnapshot::toJson($snapshot);
    MnbExcel::createFromVisualSnapshot($snapshotJson, $xlsxRoundTrip);
    $roundTrip = MnbExcel::visualSnapshot($xlsxRoundTrip);
    smoke_assert_equals('A1:H3', $roundTrip['sheets'][0]['dimension'] ?? null, 'XLSX round-trip lost cells');
    smoke_assert_equals('color_scale', $roundTrip['sheets'][0]['conditional_formats'][0]['type'] ?? null, 'XLSX round-trip lost color scale');
    smoke_assert_equals('list', $roundTrip['sheets'][0]['data_validations'][0]['type'] ?? null, 'XLSX round-trip lost data validation');
    smoke_assert_equals('Verified order', $roundTrip['sheets'][0]['comments'][0]['text'] ?? null, 'XLSX round-trip lost comment');
    smoke_assert_equals('https://example.com/products/keyboard', $roundTrip['sheets'][0]['hyperlinks'][0]['url'] ?? null, 'XLSX round-trip lost hyperlink');

    MnbExcel::createFromVisualSnapshot($snapshot, $xls);
    $xlsSnapshot = MnbExcel::visualSnapshot($xls);
    smoke_assert_equals('xls', $xlsSnapshot['format'] ?? null, 'Generic facade did not route XLS');
    smoke_assert_equals('A1:H3', $xlsSnapshot['sheets'][0]['dimension'] ?? null, 'XLS round-trip lost cells');
    smoke_assert_equals('date', $xlsSnapshot['sheets'][0]['cells']['B2']['type'] ?? null, 'XLS round-trip lost typed date');

    MnbExcel::createFromVisualSnapshot($snapshot, $csv);
    $csvSnapshot = MnbExcel::visualSnapshot($csv);
    smoke_assert_equals('csv', $csvSnapshot['format'] ?? null, 'Generic facade did not route CSV');
    smoke_assert_equals('not_applicable', $csvSnapshot['capabilities']['styles'] ?? null, 'CSV must report styles as not applicable');
    smoke_assert_true(isset($csvSnapshot['sheets'][0]['cells']['H3']), 'CSV round-trip lost Status column');

    $password = 'Visual-Snapshot-2026!';
    MnbExcel::createFromVisualSnapshot($snapshot, $encryptedXlsx, [
        'password' => $password,
        'encryption_mode' => 'standard',
    ]);
    smoke_assert_true(MnbExcel::isEncryptedXlsx($encryptedXlsx), 'Snapshot destination encryption was ignored');
    $encryptedSnapshot = MnbExcel::visualSnapshot($encryptedXlsx, ['password' => $password]);
    smoke_assert_equals('A1:H3', $encryptedSnapshot['sheets'][0]['dimension'] ?? null, 'Encrypted visual snapshot lost cells');
});
