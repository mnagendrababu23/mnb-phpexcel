<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-dashboard-' . uniqid('', true);
mkdir($dir, 0775, true);
$failed = $dir . DIRECTORY_SEPARATOR . 'failed.csv';
file_put_contents($failed, "excel_row,error\n2,Invalid email\n");
$manifest = $dir . DIRECTORY_SEPARATOR . 'manifest.json';
file_put_contents($manifest, json_encode([
    'status' => 'paused',
    'rows_scanned' => 50,
    'inserted_rows' => 45,
    'failed_rows' => 5,
    'valid_rows' => 45,
    'chunks_completed' => 2,
    'total_rows' => 100,
    'last_row_number' => 51,
    'failed_rows_csv' => $failed,
    'created_at' => gmdate('c', time() - 100),
    'updated_at' => gmdate('c'),
]));

$response = MnbExcel::importDashboard($manifest, ['download_base_url' => '/downloads']);
assert($response['status'] === 'paused');
assert($response['progress']['percent'] === 50.0);
assert($response['resume']['ready'] === true);
assert($response['resume']['recommended_action'] === 'resume_now');
assert($response['files']['failed_rows_exists'] === true);
assert($response['files']['failed_rows_download_url'] === '/downloads/failed.csv');

@unlink($manifest);
@unlink($failed);
@rmdir($dir);

echo "ImportDashboardHelperSmokeTest passed\n";
