<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$pdo = new PDO('mysql:host=localhost;dbname=app;charset=utf8mb4', 'user', 'password', [
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
]);

$plan = MnbExcel::autoImportPlan(__DIR__ . '/large.xlsx', [
    'server' => 'shared',
    'memory_limit' => '256M',
    'max_execution_time' => 30,
]);

$result = MnbExcel::largeImportToSql(__DIR__ . '/large.xlsx', $pdo, 'imports', [
    'chunk_size' => $plan['chunk_size'],
    'batch_size' => 250,
    'with_header' => true,
    'manifest_path' => __DIR__ . '/storage/large-import-manifest.json',
    'failed_rows_csv' => __DIR__ . '/storage/large-import-failed-rows.csv',
    'resume' => true,
    'time_budget_seconds' => 25,
    'rules' => [
        'email' => 'nullable|email',
        'amount' => 'nullable|numeric',
    ],
    'progress' => static function (array $state): void {
        echo 'Rows scanned: ' . ($state['rows_scanned'] ?? 0) . PHP_EOL;
    },
]);

print_r($result);
