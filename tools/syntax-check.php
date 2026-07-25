<?php

declare(strict_types=1);

/**
 * Cross-platform PHP syntax checker for Composer scripts.
 * Works on Windows PowerShell/CMD and Unix shells without find/xargs.
 */

$root = dirname(__DIR__);
$dirs = array_slice($argv, 1);
if ($dirs === []) {
    $dirs = ['src', 'tests', 'examples', 'tools'];
}

$files = [];
foreach ($dirs as $dir) {
    $path = $root . DIRECTORY_SEPARATOR . $dir;
    if (!is_dir($path)) {
        continue;
    }
    $iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($path, FilesystemIterator::SKIP_DOTS));
    foreach ($iterator as $file) {
        /** @var SplFileInfo $file */
        if ($file->isFile() && strtolower($file->getExtension()) === 'php') {
            $files[] = $file->getPathname();
        }
    }
}

sort($files);
if ($files === []) {
    fwrite(STDERR, "No PHP files found for syntax check.\n");
    exit(1);
}

$failed = [];
foreach ($files as $file) {
    $command = escapeshellarg(PHP_BINARY) . ' -l ' . escapeshellarg($file);
    $output = [];
    $exitCode = 0;
    exec($command, $output, $exitCode);
    if ($exitCode !== 0) {
        $failed[$file] = implode(PHP_EOL, $output);
    }
}

if ($failed !== []) {
    foreach ($failed as $file => $message) {
        fwrite(STDERR, "Syntax error in {$file}:\n{$message}\n");
    }
    exit(1);
}

echo 'PHP syntax check passed for ' . count($files) . " files.\n";
