<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$definitions = require $root . '/packages/modules.php';
$out = $root . '/dist/modules';
$zipOut = $root . '/dist';

if (!is_array($definitions)) {
    fwrite(STDERR, "Invalid package definitions.\n");
    exit(1);
}

$removeTree = static function (string $path) use (&$removeTree): void {
    if (!is_dir($path)) {
        return;
    }
    $items = scandir($path);
    if ($items === false) {
        throw new RuntimeException('Unable to list ' . $path);
    }
    foreach ($items as $item) {
        if ($item === '.' || $item === '..') {
            continue;
        }
        $target = $path . DIRECTORY_SEPARATOR . $item;
        is_dir($target) ? $removeTree($target) : unlink($target);
    }
    rmdir($path);
};

$copyFile = static function (string $source, string $target): void {
    $dir = dirname($target);
    if (!is_dir($dir) && !mkdir($dir, 0777, true) && !is_dir($dir)) {
        throw new RuntimeException('Unable to create ' . $dir);
    }
    if (!copy($source, $target)) {
        throw new RuntimeException('Unable to copy ' . $source);
    }
};

$removeTree($out);
if (!is_dir($out) && !mkdir($out, 0777, true) && !is_dir($out)) {
    throw new RuntimeException('Unable to create output directory.');
}
if (!is_dir($zipOut) && !mkdir($zipOut, 0777, true) && !is_dir($zipOut)) {
    throw new RuntimeException('Unable to create dist directory.');
}

$owners = [];
$built = [];
foreach ($definitions as $slug => $package) {
    $packageDir = $out . '/' . $slug;
    if (!mkdir($packageDir, 0777, true) && !is_dir($packageDir)) {
        throw new RuntimeException('Unable to create ' . $packageDir);
    }

    $resolved = [];
    foreach ((array) ($package['files'] ?? []) as $pattern) {
        $matches = glob($root . '/' . $pattern, GLOB_NOSORT);
        if ($matches === false || $matches === []) {
            throw new RuntimeException("Package {$slug} pattern matched no files: {$pattern}");
        }
        foreach ($matches as $source) {
            if (!is_file($source)) {
                continue;
            }
            $relative = str_replace('\\', '/', substr($source, strlen($root) + 1));
            if (isset($owners[$relative])) {
                throw new RuntimeException("Source file {$relative} is assigned to both {$owners[$relative]} and {$slug}");
            }
            $owners[$relative] = $slug;
            $resolved[$relative] = $source;
        }
    }
    ksort($resolved);

    foreach ($resolved as $relative => $source) {
        $copyFile($source, $packageDir . '/' . $relative);
    }
    if (($package['type'] ?? 'library') !== 'metapackage') {
        $copyFile($root . '/LICENSE', $packageDir . '/LICENSE');
    }

    $composer = [
        'name' => $package['name'],
        'description' => $package['description'],
        'type' => $package['type'] ?? 'library',
        'license' => 'MIT',
        'require' => (object) ($package['require'] ?? []),
        'minimum-stability' => 'stable',
        'prefer-stable' => true,
    ];
    if (($package['type'] ?? 'library') !== 'metapackage') {
        $composer['autoload'] = ['psr-4' => ['Mnb\\PHPExcel\\' => 'src/']];
        $composer['conflict'] = ['mnb/mnb-phpexcel' => '*'];
    }
    file_put_contents(
        $packageDir . '/composer.json',
        json_encode($composer, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . "\n"
    );

    $install = 'composer require ' . $package['name'];
    $readme = '# ' . $package['name'] . "\n\n" . $package['description'] . "\n\n"
        . "This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.\n\n"
        . "## Install\n\n```bash\n{$install}\n```\n\n"
        . "See the main project documentation for typed options, streaming reads, and compatibility notes.\n";
    file_put_contents($packageDir . '/README.md', $readme);

    $built[$slug] = [
        'name' => $package['name'],
        'type' => $package['type'] ?? 'library',
        'files' => count($resolved),
        'require' => $package['require'] ?? [],
    ];
}

$allSource = [];
$iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($root . '/src', FilesystemIterator::SKIP_DOTS));
foreach ($iterator as $file) {
    if ($file->isFile() && strtolower($file->getExtension()) === 'php') {
        $relative = str_replace('\\', '/', substr($file->getPathname(), strlen($root) + 1));
        $allSource[] = $relative;
    }
}
sort($allSource);
$unassigned = array_values(array_diff($allSource, array_keys($owners)));
if ($unassigned !== []) {
    fwrite(STDERR, "Unassigned source files:\n - " . implode("\n - ", $unassigned) . "\n");
    exit(1);
}

foreach ($definitions as $slug => $_package) {
    $packageDir = $out . '/' . $slug;
    foreach (glob($packageDir . '/src/**/*.php') ?: [] as $_unused) {
        // Recursive syntax validation is done below with an iterator.
    }
    if (is_dir($packageDir . '/src')) {
        $it = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($packageDir . '/src', FilesystemIterator::SKIP_DOTS));
        foreach ($it as $file) {
            if (!$file->isFile() || strtolower($file->getExtension()) !== 'php') {
                continue;
            }
            $command = escapeshellarg(PHP_BINARY) . ' -l ' . escapeshellarg($file->getPathname());
            exec($command, $lines, $status);
            if ($status !== 0) {
                throw new RuntimeException('Syntax validation failed for ' . $file->getPathname() . "\n" . implode("\n", $lines));
            }
        }
    }

    $archive = $zipOut . '/mnb-phpexcel-' . $slug . '.zip';
    if (is_file($archive)) {
        unlink($archive);
    }
    $command = 'cd ' . escapeshellarg($out) . ' && zip -qr ' . escapeshellarg($archive) . ' ' . escapeshellarg($slug);
    exec($command, $lines, $status);
    if ($status !== 0) {
        throw new RuntimeException('Unable to create archive for ' . $slug);
    }
    $built[$slug]['archive'] = basename($archive);
    $built[$slug]['sha256'] = hash_file('sha256', $archive);
}

file_put_contents(
    $zipOut . '/module-manifest.json',
    json_encode(['generated_at' => gmdate(DATE_ATOM), 'packages' => $built], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . "\n"
);

printf("Built %d modular packages; assigned %d source files.\n", count($built), count($owners));
foreach ($built as $slug => $meta) {
    printf("- %-12s %3d source files  %s\n", $slug, $meta['files'], $meta['archive']);
}
