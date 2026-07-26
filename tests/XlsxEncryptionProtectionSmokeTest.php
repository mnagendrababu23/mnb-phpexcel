<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;

smoke_run('encrypted XLSX read/write and document protection', static function (): void {
    smoke_assert_true(extension_loaded('openssl'), 'ext-openssl is required for encryption tests');
    smoke_assert_true(function_exists('iconv'), 'ext-iconv is required for encryption tests');

    $dir = smoke_temp_dir('xlsx_security');
    $plain = $dir . '/plain.xlsx';
    $standard = $dir . '/standard.xlsx';
    $agile = $dir . '/agile.xlsx';
    $decrypted = $dir . '/decrypted.xlsx';
    $combined = $dir . '/combined.xlsx';
    $password = 'S3cure-Workbook!';

    Xlsx::write([
        ['name' => 'Alice', 'amount' => 125.50],
        ['name' => 'Bob', 'amount' => 75],
    ], $plain, ['with_header' => true]);

    MnbExcel::encryptXlsx($plain, $standard, $password, ['mode' => 'standard']);
    smoke_assert_true(MnbExcel::isEncryptedXlsx($standard), 'Standard output should be detected as encrypted');
    smoke_assert_equals('standard', MnbExcel::xlsxEncryptionMode($standard), 'Standard mode should be reported');
    smoke_assert_equals('Alice', Xlsx::cell($standard, 'A2', 1, ReaderOptions::defaults()->withPassword($password)), 'Standard encrypted cell should be readable');
    $largeRows = [];
    $largeState = MnbExcel::largeRead($standard, ['password' => $password])
        ->withHeader()
        ->chunk(1, static function (array $rows, array $state) use (&$largeRows): void {
            $largeRows = array_merge($largeRows, $rows);
        });
    smoke_assert_equals('Alice', $largeRows[0]['name'] ?? null, 'Large encrypted read should stream decrypted rows');
    smoke_assert_true((bool) ($largeState['encrypted_source'] ?? false), 'Large read state should report encrypted source');

    $integrity = MnbExcel::validateXlsx($standard, ['password' => $password]);
    smoke_assert_true((bool) ($integrity['valid'] ?? false), 'Encrypted workbook should pass integrity validation with its password');

    $upload = MnbExcel::validateUpload([
        'tmp_name' => $standard,
        'name' => 'secure.xlsx',
        'size' => filesize($standard),
        'error' => UPLOAD_ERR_OK,
    ], ['password' => $password]);
    smoke_assert_true((bool) ($upload['valid'] ?? false), 'Encrypted upload should be validated after decryption');
    smoke_assert_true((bool) ($upload['features']['encrypted'] ?? false), 'Upload metadata should retain encrypted=true');
    smoke_assert_true((bool) ($upload['features']['decrypted_for_validation'] ?? false), 'Encrypted upload should be inspected internally');

    $passwordRequired = MnbExcel::validateUpload([
        'tmp_name' => $standard,
        'name' => 'secure.xlsx',
        'size' => filesize($standard),
        'error' => UPLOAD_ERR_OK,
    ]);
    smoke_assert_true(!(bool) ($passwordRequired['valid'] ?? true), 'Encrypted upload without a password should fail safe validation');

    $wrongPasswordRejected = false;
    try {
        Xlsx::cell($standard, 'A2', 1, ['password' => 'wrong-password']);
    } catch (Throwable) {
        $wrongPasswordRejected = true;
    }
    smoke_assert_true($wrongPasswordRejected, 'Wrong password should be rejected');

    MnbExcel::decryptXlsx($standard, $decrypted, $password);
    smoke_assert_equals('Bob', Xlsx::cell($decrypted, 'A3'), 'Decrypted workbook should be a normal readable XLSX');

    MnbExcel::encryptXlsx($plain, $agile, $password, ['mode' => 'agile', 'spin_count' => 1000]);
    smoke_assert_equals('agile', MnbExcel::xlsxEncryptionMode($agile), 'Agile mode should be reported');
    smoke_assert_equals(125.5, Xlsx::cell($agile, 'B2', 1, ['password' => $password]), 'Agile encrypted cell should be readable');

    Xlsx::write([
        ['name' => 'Protected', 'amount' => 10],
    ], $combined, [
        'with_header' => true,
        'password' => $password,
        'encryption_mode' => 'standard',
        'protection_password' => $password,
        'protect_workbook' => true,
        'protect_sheets' => true,
    ]);

    $protection = MnbExcel::xlsxProtection($combined, 1, ['password' => $password]);
    smoke_assert_true((bool) ($protection['file_encrypted'] ?? false), 'Combined output should be file encrypted');
    smoke_assert_true((bool) ($protection['workbook_protected'] ?? false), 'Workbook structure should be protected');
    smoke_assert_true((bool) ($protection['worksheet_protected'] ?? false), 'Worksheet should be protected');
    smoke_assert_equals('[present]', $protection['workbook']['workbookHashValue'] ?? null, 'Workbook verifier should be redacted');
    smoke_assert_equals('[present]', $protection['worksheet']['hashValue'] ?? null, 'Worksheet verifier should be redacted');
});
