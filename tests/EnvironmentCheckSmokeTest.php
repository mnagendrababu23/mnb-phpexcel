<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "EnvironmentCheckSmokeTest\n";

smoke_run('environmentCheck() runs without throwing and reports PHP version status', function (): void {
    $report = MnbExcel::environmentCheck();

    smoke_assert_true(is_array($report), 'environmentCheck() should return an array');
    smoke_assert_true(array_key_exists('php_version', $report) || array_key_exists('checks', $report), 'environmentCheck() report should describe at least PHP version or a checks list');
});

smoke_run('releaseReadiness() runs against this checkout without throwing', function (): void {
    $root = dirname(__DIR__);
    $report = MnbExcel::releaseReadiness($root);

    smoke_assert_true(is_array($report), 'releaseReadiness() should return an array');
});

smoke_run('version() returns a non-empty string', function (): void {
    $version = MnbExcel::version();
    smoke_assert_true(is_string($version) && $version !== '', 'version() should return a non-empty string');
});

echo "EnvironmentCheckSmokeTest: all assertions passed.\n";
