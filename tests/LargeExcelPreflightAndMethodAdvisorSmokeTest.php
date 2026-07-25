<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Large\ImportMethodAdvisor;

$sampleLarge = [
    'file_size_mb' => 88.18,
    'sheet_count' => 1,
    'total_rows' => 934262,
    'total_cells' => 22422288,
    'max_columns' => 24,
    'total_uncompressed_sheet_xml_mb' => 840.06,
    'features' => [
        'formulas' => 0,
        'comments' => 0,
        'hyperlinks' => 0,
        'merged_cells' => 0,
        'drawings' => 0,
        'charts' => 0,
        'pivot_tables' => 0,
        'external_links' => 0,
        'macros' => false,
    ],
    'sheets' => [[
        'name' => 'Worksheet',
        'rows' => 934262,
        'columns' => 24,
        'cells' => 22422288,
    ]],
];

$sharedAdvice = MnbExcel::recommendImportMethodFromProfile($sampleLarge, [
    'server' => 'shared',
    'memory_limit' => '256M',
    'max_execution_time' => 30,
]);

assert($sharedAdvice['method'] === ImportMethodAdvisor::METHOD_CLI_QUEUE_RECOMMENDED);
assert($sharedAdvice['normal_reader_allowed'] === false);
assert($sharedAdvice['recommended_chunk_size'] === 250);
assert($sharedAdvice['level'] === 'very_large');

$vpsAdvice = MnbExcel::recommendImportMethodFromProfile([
    'file_size_mb' => 28,
    'sheet_count' => 1,
    'total_rows' => 120000,
    'total_cells' => 1200000,
    'max_columns' => 10,
    'total_uncompressed_sheet_xml_mb' => 220,
    'features' => [],
], [
    'server' => 'vps',
    'memory_limit' => '512M',
    'max_execution_time' => 300,
]);

assert($vpsAdvice['method'] === ImportMethodAdvisor::METHOD_STREAMING_CHUNK_IMPORT);
assert($vpsAdvice['recommended_chunk_size'] === 1000);

$smallAdvice = MnbExcel::recommendImportMethodFromProfile([
    'file_size_mb' => 2,
    'sheet_count' => 1,
    'total_rows' => 1500,
    'total_cells' => 12000,
    'max_columns' => 8,
    'total_uncompressed_sheet_xml_mb' => 5,
    'features' => [],
], ['memory_limit' => '128M']);

assert($smallAdvice['method'] === ImportMethodAdvisor::METHOD_NORMAL_READ);
assert($smallAdvice['normal_reader_allowed'] === true);

$matrix = MnbExcel::importMethodMatrix();
assert(isset($matrix['very_large']));

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class)) {
    echo "LargeExcelPreflightAndMethodAdvisorSmokeTest skipped XLSX runtime preflight because ext-zip/ext-xmlreader is unavailable.\n";
    echo "LargeExcelPreflightAndMethodAdvisorSmokeTest passed\n";
    return;
}

$tmp = sys_get_temp_dir() . '/mnb-large-preflight-' . uniqid('', true) . '.xlsx';
MnbExcel::fromArray([
    ['id' => 1, 'name' => 'One'],
    ['id' => 2, 'name' => 'Two'],
])->withHeader()->save($tmp);

$profile = MnbExcel::analyzeXlsxForImport($tmp, ['accurate_row_count' => true]);
assert($profile['status'] === 'ok');
assert($profile['sheet_count'] === 1);
assert($profile['total_rows'] >= 3);
assert($profile['max_columns'] >= 2);
assert(isset($profile['method_advice']['method']));
@unlink($tmp);

echo "LargeExcelPreflightAndMethodAdvisorSmokeTest passed\n";
