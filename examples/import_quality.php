<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../tests/bootstrap.php';

$rows = [
    ['_mnb_original_row_number' => 2, 'student_name' => 'Ravi', 'email_address' => 'ravi@example.com', 'phone' => '0987654321', 'marks' => 95],
    ['_mnb_original_row_number' => 3, 'student_name' => 'Sita', 'email_address' => 'sita@example.com', 'phone' => '0912345678', 'marks' => 88],
    ['_mnb_original_row_number' => 4, 'student_name' => '', 'email_address' => 'bad-email', 'phone' => '0987654321', 'marks' => 'A+'],
];

$preview = MnbExcel::previewImport($rows, [
    'required_columns' => ['student_name', 'email_address', 'phone', 'marks'],
    'allowed_columns' => ['_mnb_original_row_number', 'student_name', 'email_address', 'phone', 'marks'],
    'strict_columns' => true,
    'duplicate_by' => ['phone'],
]);

print_r($preview['warnings']);

$map = MnbExcel::suggestColumnMap(
    ['Student Name', 'Email Address', 'Mobile No', 'Total Marks'],
    ['name', 'email', 'phone', 'marks'],
    [
        'phone' => ['mobile', 'mobile no', 'phone number'],
        'marks' => ['total marks', 'score'],
    ]
);

print_r($map);

$result = MnbExcel::validateImport($rows, [
    'student_name' => 'required|string|max:100',
    'email_address' => 'required|email',
    'marks' => 'required|numeric|min:0|max:100',
], [
    'row_number_key' => '_mnb_original_row_number',
    'strict_columns' => true,
    'allowed_columns' => ['_mnb_original_row_number', 'student_name', 'email_address', 'phone', 'marks'],
    'duplicate_by' => ['phone'],
]);

MnbExcel::fromFailedRows($result['failed'])
    ->freezeHeader()
    ->autoFilter()
    ->save(__DIR__ . '/import-failed-rows.csv');

echo "Import quality example complete. Failed report written to examples/import-failed-rows.csv\n";
