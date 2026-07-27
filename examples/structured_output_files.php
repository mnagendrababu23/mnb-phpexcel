<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
$outputDirectory = $argv[2] ?? (__DIR__ . '/output');

if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/structured_output_files.php <workbook.xlsx> [output-directory]\n");
    exit(1);
}

if (!is_dir($outputDirectory) && !mkdir($outputDirectory, 0775, true) && !is_dir($outputDirectory)) {
    throw new RuntimeException('Unable to create output directory: ' . $outputDirectory);
}

$reader = MnbExcel::read($path);
$readOptions = [
    'header_row' => true,
    'header_case' => 'snake',
    'skip_empty_rows' => true,
];

$jsonPath = $reader->saveStructuredJson(
    $outputDirectory . '/workbook.json',
    $readOptions,
    ['pretty' => true, 'trailing_newline' => true],
);

$xmlPath = $reader->saveStructuredXml(
    $outputDirectory . '/workbook.xml',
    $readOptions,
    ['pretty' => true],
);

printf("Saved structured JSON: %s\nSaved structured XML: %s\n", $jsonPath, $xmlPath);
