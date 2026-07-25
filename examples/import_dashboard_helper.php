<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$response = MnbExcel::importDashboard(__DIR__ . '/storage/import-manifest.json', [
    'download_base_url' => '/admin/imports/downloads',
]);

header('Content-Type: application/json');
echo json_encode($response, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
