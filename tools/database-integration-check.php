<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$result = MnbExcel::databaseIntegrationCheck();
echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
exit(($result['status'] ?? 'failed') === 'failed' ? 1 : 0);
