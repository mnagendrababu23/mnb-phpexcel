<?php

declare(strict_types=1);

/**
 * Cross-platform smoke test runner for Composer scripts.
 * Works on Windows PowerShell/CMD and Unix shells without shell for-loops.
 */

$root = dirname(__DIR__);
$testsDir = $root . DIRECTORY_SEPARATOR . 'tests';
$files = glob($testsDir . DIRECTORY_SEPARATOR . '*SmokeTest.php') ?: [];
sort($files);

if ($files === []) {
    fwrite(STDERR, "No smoke tests found.\n");
    exit(1);
}

$failed = [];
foreach ($files as $file) {
    echo 'Running ' . basename($file) . "...\n";
    $command = escapeshellarg(PHP_BINARY) . ' -d zend.assertions=1 -d assert.exception=1 ' . escapeshellarg($file);
    passthru($command, $exitCode);
    if ($exitCode !== 0) {
        $failed[] = basename($file);
    }
}

if ($failed !== []) {
    fwrite(STDERR, 'Smoke tests failed: ' . implode(', ', $failed) . "\n");
    exit(1);
}

echo 'All smoke tests passed (' . count($files) . ").\n";
