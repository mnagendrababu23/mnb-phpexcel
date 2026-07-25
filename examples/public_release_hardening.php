<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

$rows = [
    ['name' => 'Ravi Kumar', 'email' => 'ravi@example.com', 'phone' => '0987654321', 'website' => 'https://example.com'],
    ['name' => 'Sita Devi', 'email' => 'sita@example.com', 'phone' => '0912345678', 'website' => 'https://example.org'],
];

$result = MnbExcel::validateImport($rows, [
    'name' => 'required|alpha|max:100',
    'email' => 'required|email|unique_in_file',
    'phone' => 'required|phone_basic',
    'website' => 'nullable|url',
]);

if ($result['failed'] !== []) {
    MnbExcel::fromFailedRows($result['failed'])->save(__DIR__ . '/failed-rows.xlsx');
}

$readiness = MnbExcel::releaseReadiness(dirname(__DIR__));
print_r($readiness);
