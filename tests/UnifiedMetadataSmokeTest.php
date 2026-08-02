<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Writer\XlsxWriter;

smoke_run('provides one metadata schema through Mono and the XLSX facade', static function (): void {
    $dir = smoke_temp_dir('unified_metadata');
    $source = $dir . '/metadata.xlsx';
    $updated = $dir . '/updated.xlsx';
    $clean = $dir . '/clean.xlsx';

    $data = new WorksheetData(
        'Data',
        [
            ['Name', 'Amount', 'Result'],
            ['Alice', 10, CellValue::formula('B2*2', 20)],
        ],
        hasHeader: true,
        hyperlinks: [
            ['cell' => 'A2', 'url' => 'https://example.com', 'display' => 'Alice'],
        ],
        comments: [
            ['cell' => 'B2', 'author' => 'Reviewer', 'text' => 'Check this amount.'],
        ],
        dataValidations: [
            ['type' => 'list', 'range' => 'A2', 'values' => ['Alice', 'Bob']],
        ],
    );
    $archive = new WorksheetData('Archive', [['Value'], [1]]);
    (new XlsxWriter())->write(new WorkbookData([$data, $archive], [
        'title' => 'Mono Metadata',
        'creator' => 'MNB',
        'last_modified_by' => 'QA',
        'company' => 'MNB Technologies',
    ]), $source);

    MnbExcel::updateMetaInfo($source, $updated, [
        'document' => ['title' => 'Mono Metadata Updated'],
        'custom_properties' => [
            'Project ID' => 'PRJ-MONO-1',
            'Approved' => true,
        ],
        'workbook' => [
            'sheet_visibility' => ['Data' => 'hidden', 'Archive' => 'visible'],
            'active_sheet' => 'Archive',
        ],
    ]);

    $quick = MnbExcel::metaInfo($updated, ['profile' => 'quick']);
    smoke_assert_equals('1.0', $quick['schema_version'], 'Mono should return metadata schema 1.0');
    smoke_assert_equals('xlsx', $quick['format'], 'Mono should detect XLSX metadata format');
    smoke_assert_equals('Mono Metadata Updated', $quick['document']['title'], 'Mono should read the updated title');
    smoke_assert_equals('Archive', $quick['workbook']['active_sheet']['name'], 'Mono should read the active sheet');
    smoke_assert_equals(1, $quick['hidden_content']['hidden_sheet_count'], 'Mono should report hidden worksheets');

    $full = Xlsx::read($updated)->metaInfo(['profile' => 'full']);
    smoke_assert_equals(1, $full['comments_notes']['comment_count'], 'Full profile should inventory comments');
    smoke_assert_equals(1, $full['links']['hyperlink_count'], 'Full profile should inventory hyperlinks');
    smoke_assert_equals(1, $full['calculation']['formula_count'], 'Full profile should inventory formulas');
    smoke_assert_equals(2, $full['custom_properties']['count'], 'Full profile should read custom properties');
    smoke_assert_true($full['capabilities']['document']['write'], 'XLSX should advertise document metadata write support');

    MnbExcel::removePersonalInfo($updated, $clean);
    $cleanInfo = MnbExcel::metaInfo($clean, ['profile' => 'full']);
    smoke_assert_equals(null, $cleanInfo['document']['creator'], 'Personal cleanup should remove creator');
    smoke_assert_equals(null, $cleanInfo['revision']['last_saved_by'], 'Personal cleanup should remove last saved by');
    smoke_assert_equals(null, $cleanInfo['application']['company'], 'Personal cleanup should remove company');
    smoke_assert_equals(0, $cleanInfo['custom_properties']['count'], 'Personal cleanup should remove custom properties');
    smoke_assert_equals('Author', $cleanInfo['comments_notes']['items'][0]['author'], 'Personal cleanup should anonymize comment authors');
});

smoke_run('routes encrypted XLSX metadata through the generic Mono facade', static function (): void {
    smoke_assert_true(extension_loaded('openssl'), 'ext-openssl is required for encrypted metadata tests');
    smoke_assert_true(function_exists('iconv'), 'ext-iconv is required for encrypted metadata tests');

    $dir = smoke_temp_dir('encrypted_unified_metadata');
    $plain = $dir . '/plain.xlsx';
    $encrypted = $dir . '/encrypted.xlsx';
    $updated = $dir . '/updated.xlsx';
    $password = 'Metadata-Encrypted-Test!';

    Xlsx::write([['name' => 'Alice']], $plain, ['with_header' => true]);
    MnbExcel::encryptXlsx($plain, $encrypted, $password, ['mode' => 'agile', 'spin_count' => 1000]);

    $locked = MnbExcel::metaInfo($encrypted, ['profile' => 'quick']);
    smoke_assert_equals('password_required', $locked['status'], 'Encrypted metadata without a password should return password_required.');
    smoke_assert_equals('xlsx', $locked['format'], 'Encrypted OOXML should not be misrouted as legacy XLS.');

    $opened = MnbExcel::metaInfo($encrypted, ['profile' => 'quick', 'password' => $password]);
    smoke_assert_equals('xlsx', $opened['format'], 'Encrypted XLSX should use the XLSX metadata reader.');

    MnbExcel::updateMetaInfo(
        $encrypted,
        $updated,
        ['document' => ['title' => 'Encrypted Mono Metadata']],
        ['password' => $password],
    );
    smoke_assert_true(MnbExcel::isEncryptedXlsx($updated), 'Generic metadata update should preserve XLSX encryption.');
    $updatedInfo = MnbExcel::metaInfo($updated, ['profile' => 'quick', 'password' => $password]);
    smoke_assert_equals('Encrypted Mono Metadata', $updatedInfo['document']['title'], 'Generic encrypted metadata update failed.');
});

smoke_run('returns the shared metadata envelope for CSV', static function (): void {
    $dir = smoke_temp_dir('csv_metadata_envelope');
    $path = $dir . '/data.csv';
    file_put_contents($path, "name,amount\nAlice,10\n");

    $metadata = MnbExcel::metaInfo($path, ['profile' => 'quick', 'include_hash' => true]);
    smoke_assert_equals('1.0', $metadata['schema_version'], 'CSV should use metadata schema 1.0');
    smoke_assert_equals('csv', $metadata['format'], 'CSV format should be reported');
    smoke_assert_equals('available', $metadata['file']['state'], 'CSV file metadata should be available');
    smoke_assert_equals('partial', $metadata['workbook']['state'], 'CSV should expose a synthetic workbook section');
    smoke_assert_equals(1, $metadata['workbook']['sheet_count'], 'CSV should expose one synthetic worksheet');
    smoke_assert_equals('not_supported', $metadata['document']['state'], 'CSV document properties should not be fabricated');
    smoke_assert_true(is_string($metadata['file']['sha256']) && strlen($metadata['file']['sha256']) === 64, 'CSV hash should be available on request');
});

echo "UnifiedMetadataSmokeTest passed\n";
