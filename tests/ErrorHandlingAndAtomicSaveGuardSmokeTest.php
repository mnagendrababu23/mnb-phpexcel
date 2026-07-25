<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

require __DIR__ . '/bootstrap.php';

$tmpDir = sys_get_temp_dir() . '/mnb_phpexcel_error_guard_' . bin2hex(random_bytes(4));
if (!mkdir($tmpDir, 0775, true) && !is_dir($tmpDir)) {
    throw new RuntimeException('Unable to create temp directory.');
}

$csvPath = $tmpDir . '/atomic.csv';
file_put_contents($csvPath, "original\n");

try {
    MnbExcel::fromArray([
        ['name' => '=cmd|bad'],
    ])
        ->withHeader()
        ->csvInjectionPolicy('block')
        ->save($csvPath);
    assert(false, 'Expected CSV injection block exception.');
} catch (MnbExcelException $e) {
    assert($e->getErrorCode() === ErrorCode::SECURITY_BLOCKED);
    $safe = MnbExcel::safeError($e);
    assert($safe['status'] === 'error');
    assert($safe['code'] === ErrorCode::SECURITY_BLOCKED);
    assert(!isset($safe['developer_message']));
}

assert(file_get_contents($csvPath) === "original\n");
assert(glob($tmpDir . '/.atomic.csv.tmp.*') === []);

$jsonPath = $tmpDir . '/ok.json';
MnbExcel::fromArray([['name' => 'Ravi']])->withHeader()->saveJson($jsonPath);
assert(is_file($jsonPath));
assert(str_contains((string) file_get_contents($jsonPath), 'Ravi'));

$directoryAsTarget = $tmpDir . '/target-dir';
if (!mkdir($directoryAsTarget) && !is_dir($directoryAsTarget)) {
    throw new RuntimeException('Unable to create target directory.');
}
try {
    MnbExcel::fromArray([['name' => 'Should fail']])->withHeader()->saveJson($directoryAsTarget);
    assert(false, 'Expected atomic replace failure when target is a directory.');
} catch (MnbExcelException $e) {
    assert(in_array($e->getErrorCode(), [ErrorCode::FILE_REPLACE_FAILED, ErrorCode::FILE_WRITE_FAILED], true));
    $debug = MnbExcel::errorReport($e, true);
    assert(isset($debug['developer_message']));
    assert($debug['status'] === 'error');
}

assert(is_dir($directoryAsTarget));
assert(glob($tmpDir . '/.target-dir.tmp.*') === []);

$xmlPath = $tmpDir . '/ok.xml';
MnbExcel::fromArray([['name' => 'Sita']])->withHeader()->saveXml($xmlPath);
assert(is_file($xmlPath));
assert(str_contains((string) file_get_contents($xmlPath), 'Sita'));

foreach (glob($tmpDir . '/*') ?: [] as $path) {
    if (is_file($path)) {
        unlink($path);
    } elseif (is_dir($path)) {
        rmdir($path);
    }
}
rmdir($tmpDir);

echo "Error handling and atomic save guard smoke test passed.\n";
