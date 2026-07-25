<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "ValidationSmokeTest\n";

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'age' => 30],
    ['name' => '', 'email' => 'not-an-email', 'age' => 'thirty'],
];

$rules = [
    'name' => 'required',
    'email' => 'required|email',
    'age' => 'required|numeric',
];

smoke_run('passes valid rows and rejects invalid ones', function () use ($rows, $rules): void {
    $result = MnbExcel::validateArray($rows, $rules);

    smoke_assert_equals(1, count($result['valid'] ?? []), 'exactly one row should pass validation');
    smoke_assert_equals(1, count($result['failed'] ?? []), 'exactly one row should fail validation');

    $failedRow = $result['failed'][0] ?? [];
    smoke_assert_true(count($failedRow['errors'] ?? []) >= 2, 'the invalid row should report at least two errors (name + email/age)');
});

smoke_run('produces a failed-rows report that can be re-exported to XLSX', function () use ($rows, $rules): void {
    $result = MnbExcel::validateArray($rows, $rules);
    $failedRows = $result['failed'] ?? [];

    smoke_assert_true($failedRows !== [], 'expected at least one failed row to build a report from');

    $builder = MnbExcel::fromFailedRows($failedRows);
    smoke_assert_true(is_object($builder), 'fromFailedRows() should return a workbook builder');
});

echo "ValidationSmokeTest: all assertions passed.\n";
