<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$source = $argv[1] ?? '';
if ($source === '' || !is_file($source)) {
    fwrite(STDERR, "Usage: php examples/unified_metadata.php spreadsheet.(xlsx|xls|csv) [updated file]\n");
    exit(1);
}

$report = MnbExcel::metaInfo($source, [
    'profile' => 'standard',
]);

echo json_encode($report, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE), PHP_EOL;

$destination = $argv[2] ?? '';
if ($destination !== '') {
    $extension = strtolower((string) pathinfo($source, PATHINFO_EXTENSION));
    if ($extension === 'csv' || $extension === 'tsv') {
        fwrite(STDERR, "CSV/TSV files do not contain an embedded Office metadata store; inspection is supported, metadata updates are not.\n");
        exit(2);
    }

    MnbExcel::updateMetaInfo($source, $destination, [
        'document' => [
            'title' => 'Updated with MNB PHPExcel',
            'creator' => 'MNB PHPExcel',
        ],
        'custom_properties' => [
            'Metadata Schema' => '1.0',
        ],
    ]);
    echo 'Updated workbook: ' . $destination . PHP_EOL;
}
