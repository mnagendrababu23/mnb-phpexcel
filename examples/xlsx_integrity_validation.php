<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$rows = [
    ['invoice' => 'INV-001', 'customer' => 'Ravi', 'total' => 1250.50],
    ['invoice' => 'INV-002', 'customer' => 'Sita', 'total' => 980.00],
];

$output = __DIR__ . '/output/invoices-safe.xlsx';

MnbExcel::fromArray($rows)
    ->withHeader()
    ->autoWidth()
    ->strictXlsxIntegrityValidation()
    ->save($output);

$result = MnbExcel::validateXlsx($output);

if (!$result['valid']) {
    echo "XLSX integrity failed:\n";
    print_r($result['errors']);
    exit(1);
}

echo "XLSX integrity status: {$result['status']}\n";
echo "Saved: {$output}\n";
