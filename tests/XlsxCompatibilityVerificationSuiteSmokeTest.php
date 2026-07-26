<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$report = MnbExcel::verifyXlsxCompatibility();

assert(isset($report['status'], $report['environment'], $report['cases'], $report['summary']));
assert(in_array($report['status'], ['pass', 'warning'], true), 'Compatibility verification should pass generated fixtures.');
assert(($report['summary']['failed'] ?? 1) === 0);
assert(count($report['cases']) >= 3);

echo "XlsxCompatibilityVerificationSuiteSmokeTest: PASS\n";
