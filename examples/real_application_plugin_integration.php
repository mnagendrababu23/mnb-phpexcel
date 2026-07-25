<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Plugin\PluginRegistry;

MnbExcel::storage([
    'base_path' => __DIR__ . '/../storage/mnb-phpexcel',
]);

MnbExcel::setLogger(function (string $level, string $message, array $context): void {
    echo '[' . strtoupper($level) . '] ' . $message . PHP_EOL;
});

MnbExcel::plugin(function (PluginRegistry $registry): void {
    $registry->transformer('clean_rows', function (array $row): array {
        foreach ($row as $key => $value) {
            if (is_string($value)) {
                $row[$key] = trim($value);
            }
        }
        if (isset($row['email'])) {
            $row['email'] = strtolower((string) $row['email']);
        }
        return $row;
    });

    $registry->validator('active_status', function (mixed $value): bool|string {
        return in_array((string) $value, ['active', 'inactive'], true)
            ? true
            : 'status must be active or inactive.';
    });

    $registry->importProfile('student_import', [
        'table' => 'students',
        'with_header' => true,
        'chunk_size' => 500,
        'batch_size' => 250,
        'resume' => true,
        'idempotent' => true,
        'duplicate_strategy' => 'skip',
        'unique_by' => ['student_id'],
        'transformers' => ['clean_rows'],
        'rules' => [
            'student_id' => 'required',
            'email' => 'nullable|email',
            'status' => 'nullable|active_status',
        ],
    ]);

    $registry->event('after_chunk', function (array $event): void {
        echo 'Imported rows scanned: ' . ($event['rows_scanned'] ?? 0) . PHP_EOL;
    });
});

// Validate upload before storing/importing.
// $upload = MnbExcel::validateUpload($_FILES['excel'], ['allowed_extensions' => ['xlsx'], 'max_size_mb' => 100]);
// if (!$upload['valid']) { print_r($upload['errors']); exit; }

// Run profile from a real .env/config/PDO source:
// $result = MnbExcel::profile('student_import')->source($_FILES['excel']['tmp_name'])->run(__DIR__ . '/../.env');
// print_r($result);

// Dashboard progress:
// $status = MnbExcel::importStatus(__DIR__ . '/../storage/mnb-phpexcel/manifests/students.json');
// print_r($status);

echo "Real application/plugin integration example configured.\n";
