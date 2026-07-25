<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$rows = [
    ['_mnb_original_row_number' => 2, 'name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321', 'marks' => 95],
    ['_mnb_original_row_number' => 3, 'name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678', 'marks' => 88],
    ['_mnb_original_row_number' => 4, 'name' => '', 'email' => 'bad-email', 'phone' => '0987654321', 'marks' => 'A+'],
];

$preview = MnbExcel::previewImport($rows, [
    'required_columns' => ['name', 'email', 'phone', 'marks'],
    'allowed_columns' => ['_mnb_original_row_number', 'name', 'email', 'phone', 'marks'],
    'strict_columns' => true,
    'duplicate_by' => ['phone'],
]);

assert($preview['status'] === 'warning');
assert($preview['total_rows'] === 3);
assert(count($preview['duplicate_groups']) === 1);
assert($preview['duplicate_groups'][0]['rows'] === [2, 4]);

$map = MnbExcel::suggestColumnMap(
    ['Student Name', 'Email Address', 'Mobile No', 'Total Marks'],
    ['name', 'email', 'phone', 'marks'],
    ['phone' => ['mobile no'], 'marks' => ['total marks']]
);

assert($map['Mobile No']['target'] === 'phone');
assert($map['Total Marks']['target'] === 'marks');

$result = MnbExcel::validateImport($rows, [
    'name' => 'required|string|max:100',
    'email' => 'required|email',
    'marks' => 'required|numeric|min:0|max:100',
], [
    'row_number_key' => '_mnb_original_row_number',
    'strict_columns' => true,
    'allowed_columns' => ['_mnb_original_row_number', 'name', 'email', 'phone', 'marks'],
    'duplicate_by' => ['phone'],
]);

assert(count($result['valid']) === 2);
assert(count($result['failed']) === 1);
assert($result['failed'][0]['row'] === 4);
assert(str_contains(implode(' ', $result['failed'][0]['errors']), 'Duplicate row'));

$report = MnbExcel::fromFailedRows($result['failed'])->toArray();
assert($report[0][0] === 'row_number');
assert($report[1][0] === 4);

$pdo = new class extends PDO {
    public function __construct()
    {
    }
};

$dryRun = MnbExcel::fromArray(array_map(static function (array $row): array {
    unset($row['_mnb_original_row_number']);
    return $row;
}, $result['valid']))->importToSql($pdo, 'students', ['dry_run' => true]);

assert($dryRun['status'] === 'dry_run');
assert($dryRun['planned_rows'] === 2);
assert($dryRun['inserted_rows'] === 0);
assert($dryRun['planned_batches'] === 1);

echo "Import quality smoke test passed.\n";
