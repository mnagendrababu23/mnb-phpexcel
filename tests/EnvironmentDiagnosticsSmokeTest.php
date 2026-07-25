<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$report = MnbExcel::environmentCheck();

assert(isset($report['status'], $report['extensions'], $report['capabilities'], $report['checks']));
assert($report['extensions']['json'] === extension_loaded('json'));
assert($report['capabilities']['csv_ready'] === true);
assert(array_key_exists('xlsx_write_ready', $report['capabilities']));

echo "EnvironmentDiagnosticsSmokeTest: PASS\n";
