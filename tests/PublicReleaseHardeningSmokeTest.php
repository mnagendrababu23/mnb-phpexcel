<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\ReleaseReadiness;

require __DIR__ . '/bootstrap.php';

$rows = [
    ['_mnb_original_row_number' => 2, 'name' => 'Ravi Kumar', 'email' => 'ravi@example.com', 'phone' => '+91 98765 43210', 'website' => 'https://example.com', 'status' => 'active', 'code' => 'STU-001'],
    ['_mnb_original_row_number' => 3, 'name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678', 'website' => 'https://example.org', 'status' => 'inactive', 'code' => 'STU-002'],
    ['_mnb_original_row_number' => 4, 'name' => 'Bad 123', 'email' => 'ravi@example.com', 'phone' => 'bad', 'website' => 'not-url', 'status' => 'draft', 'code' => 'BAD'],
];

$result = MnbExcel::validateImport($rows, [
    'name' => 'required|alpha|max:100',
    'email' => 'required|email|unique_in_file',
    'phone' => 'required|phone_basic',
    'website' => 'nullable|url',
    'status' => 'required|in:active,inactive',
    'code' => 'required|starts_with:STU-|length:7',
], [
    'row_number_key' => '_mnb_original_row_number',
    'strict_columns' => true,
    'allowed_columns' => ['_mnb_original_row_number', 'name', 'email', 'phone', 'website', 'status', 'code'],
]);

assert(count($result['valid']) === 2);
assert(count($result['failed']) === 1);
assert($result['failed'][0]['row'] === 4);
assert(isset($result['failed'][0]['error_details']));
assert(count($result['failed'][0]['error_details']) >= 5);

$messages = implode(' ', $result['failed'][0]['errors']);
assert(str_contains($messages, 'valid URL'));
assert(str_contains($messages, 'unique in file'));
assert(str_contains($messages, 'phone number'));
assert(str_contains($messages, 'must start with'));

$workbook = MnbExcel::fromArray([
    ['phone' => '0987654321', 'amount' => 1200.5, 'started' => '2026-07-01'],
])
    ->withHeader()
    ->textColumns(['phone'])
    ->numberColumns(['amount'])
    ->dateColumns(['started' => 'Y-m-d'])
    ->toWorkbookData();

$sheet = $workbook->sheets[0];
assert(($sheet->columnStyles[1]['format'] ?? null) === 'text');
assert(($sheet->columnStyles[2]['format'] ?? null) === 'number');
assert(($sheet->columnStyles[3]['format'] ?? null) === 'date');

$readiness = ReleaseReadiness::check(dirname(__DIR__));
assert(in_array($readiness['status'], ['ready', 'warning'], true));
assert($readiness['summary']['failed'] === 0);

$apiReadiness = MnbExcel::releaseReadiness(dirname(__DIR__));
assert($apiReadiness['summary']['failed'] === 0);

echo "Public release hardening smoke test passed.\n";
