<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$environment = MnbExcel::environmentCheck();
print_r($environment['capabilities']);

$report = MnbExcel::verifyXlsxCompatibility([
    // Add optional real-world fixtures created/exported by Excel, LibreOffice, or Google Sheets:
    // __DIR__ . '/fixtures/excel-comments-hyperlinks.xlsx',
]);

print_r([
    'status' => $report['status'],
    'summary' => $report['summary'],
    'warnings' => $report['warnings'] ?? [],
]);
