<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "UploadSafetyHardeningSmokeTest\n";

smoke_run('validates actual size, MIME metadata, and safe CSV names', function (): void {
    $dir = smoke_temp_dir('upload_safety');
    $path = $dir . DIRECTORY_SEPARATOR . 'students.csv';
    file_put_contents($path, "id,name\n1,Ravi\n");

    $result = MnbExcel::validateUpload([
        'tmp_name' => $path,
        'name' => 'students.csv',
        'size' => filesize($path),
        'error' => UPLOAD_ERR_OK,
    ], [
        'allowed_extensions' => ['csv'],
        'max_size_mb' => 1,
    ]);

    smoke_assert_true($result['valid'] === true, 'safe CSV upload should pass validation');
    smoke_assert_equals(filesize($path), $result['size_bytes'], 'validator should use actual file size');
    if (class_exists(finfo::class)) {
        smoke_assert_true(is_string($result['mime_type']), 'fileinfo-enabled validation should return a MIME type');
    }
});

smoke_run('rejects path-like upload names and PHP upload errors', function (): void {
    $dir = smoke_temp_dir('upload_name');
    $path = $dir . DIRECTORY_SEPARATOR . 'safe.csv';
    file_put_contents($path, "id\n1\n");

    $unsafe = MnbExcel::validateUpload([
        'tmp_name' => $path,
        'name' => '../unsafe.csv',
        'size' => filesize($path),
        'error' => UPLOAD_ERR_PARTIAL,
    ], ['allowed_extensions' => ['csv']]);

    smoke_assert_true($unsafe['valid'] === false, 'unsafe filename/upload error should fail validation');
    smoke_assert_true(count($unsafe['errors']) >= 2, 'validator should report both independent failures');
});

smoke_run('rejects ZIP path traversal when ext-zip is available', function (): void {
    if (!class_exists(ZipArchive::class)) {
        echo "SKIP ext-zip unavailable ";
        return;
    }

    $dir = smoke_temp_dir('upload_zip');
    $path = $dir . DIRECTORY_SEPARATOR . 'unsafe.xlsx';
    $zip = new ZipArchive();
    smoke_assert_true($zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true, 'test ZIP should open');
    $zip->addFromString('[Content_Types].xml', '<Types/>');
    $zip->addFromString('_rels/.rels', '<Relationships/>');
    $zip->addFromString('xl/workbook.xml', '<workbook/>');
    $zip->addFromString('../escape.txt', 'unsafe');
    $zip->close();

    $result = MnbExcel::validateUpload($path, ['allowed_extensions' => ['xlsx']]);
    smoke_assert_true($result['valid'] === false, 'ZIP path traversal should fail validation');
});

echo "UploadSafetyHardeningSmokeTest: all assertions passed.\n";
