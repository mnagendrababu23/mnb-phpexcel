<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;

$missing = MnbExcel::validateXlsx(__DIR__ . '/missing-file.xlsx');
assert($missing['status'] === 'fail');
assert($missing['valid'] === false);
assert($missing['errors'] !== []);

$validator = new XlsxIntegrityValidator();
$missingDirect = $validator->validate(__DIR__ . '/missing-direct.xlsx');
assert($missingDirect['status'] === 'fail');

if (!class_exists(ZipArchive::class)) {
    echo "XlsxIntegrityValidatorSmokeTest skipped XLSX runtime checks: ext-zip is not available.\n";
    return;
}

$outputDir = __DIR__ . '/tmp';
if (!is_dir($outputDir)) {
    mkdir($outputDir, 0775, true);
}

$validPath = $outputDir . '/integrity-valid.xlsx';
MnbExcel::fromArray([
    ['id' => 1, 'name' => 'Asha'],
    ['id' => 2, 'name' => 'Ravi'],
])->withHeader()->save($validPath);

$result = MnbExcel::validateXlsx($validPath);
assert($result['valid'] === true);
assert(in_array($result['status'], ['pass', 'warning'], true));
MnbExcel::assertValidXlsx($validPath);

$brokenPath = $outputDir . '/integrity-broken.xlsx';
copy($validPath, $brokenPath);
$zip = new ZipArchive();
$zip->open($brokenPath);
$zip->deleteName('xl/workbook.xml');
$zip->close();

$broken = MnbExcel::validateXlsx($brokenPath);
assert($broken['status'] === 'fail');
assert($broken['valid'] === false);

@unlink($validPath);
@unlink($brokenPath);

echo "XlsxIntegrityValidatorSmokeTest passed.\n";
