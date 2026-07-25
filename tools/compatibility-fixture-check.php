<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$dir = $argv[1] ?? (__DIR__ . '/../docs/fixtures');
$result = MnbExcel::verifyCompatibilityFixtures($dir);
echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
exit(($result['status'] ?? 'failed') === 'failed' ? 1 : 0);
