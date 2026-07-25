<?php

declare(strict_types=1);

$root = dirname(__DIR__);
require $root . '/tools/build-modular-packages.php';

$definitions = require $root . '/packages/modules.php';
$moduleRoot = $root . '/dist/modules';

$fail = static function (string $message): never {
    fwrite(STDERR, "Modular package test failed: {$message}\n");
    exit(1);
};
$assert = static function (bool $condition, string $message) use ($fail): void {
    if (!$condition) {
        $fail($message);
    }
};

foreach ($definitions as $slug => $definition) {
    $composerPath = $moduleRoot . '/' . $slug . '/composer.json';
    $assert(is_file($composerPath), "Missing composer.json for {$slug}");
    $composer = json_decode((string) file_get_contents($composerPath), true);
    $assert(is_array($composer), "Invalid composer.json for {$slug}");
    $assert(($composer['name'] ?? null) === $definition['name'], "Wrong package name for {$slug}");
    foreach ((array) ($composer['require'] ?? []) as $dependency => $_constraint) {
        if (!str_starts_with((string) $dependency, 'mnb/mnb-phpexcel-')) {
            continue;
        }
        $found = false;
        foreach ($definitions as $candidate) {
            if ($candidate['name'] === $dependency) {
                $found = true;
                break;
            }
        }
        $assert($found, "Unknown internal dependency {$dependency} in {$slug}");
    }
}

$runIsolated = static function (string $name, array $modules, string $body) use ($moduleRoot, $fail): void {
    $dirs = [];
    foreach ($modules as $module) {
        $dir = $moduleRoot . '/' . $module . '/src';
        if (is_dir($dir)) {
            $dirs[] = $dir;
        }
    }
    $script = <<<'BOOT'
<?php
declare(strict_types=1);
$dirs = __DIRS__;
spl_autoload_register(static function (string $class) use ($dirs): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) return;
    $relative = str_replace('\\', DIRECTORY_SEPARATOR, substr($class, strlen($prefix))) . '.php';
    foreach ($dirs as $dir) {
        $path = $dir . DIRECTORY_SEPARATOR . $relative;
        if (is_file($path)) { require $path; return; }
    }
});
function check(bool $condition, string $message): void { if (!$condition) throw new RuntimeException($message); }
__BODY__
BOOT;
    $script = str_replace('__DIRS__', var_export($dirs, true), $script);
    $script = str_replace('__BODY__', $body, $script);
    $path = sys_get_temp_dir() . '/mnb_module_' . preg_replace('/[^a-z0-9]+/i', '_', $name) . '_' . bin2hex(random_bytes(4)) . '.php';
    file_put_contents($path, $script);
    $command = escapeshellarg(PHP_BINARY) . ' ' . escapeshellarg($path) . ' 2>&1';
    exec($command, $output, $status);
    @unlink($path);
    if ($status !== 0) {
        $fail($name . "\n" . implode("\n", $output));
    }
};

$runIsolated('core + JSON', ['core', 'json'], <<<'PHP'
$temp = tempnam(sys_get_temp_dir(), 'mnb_json_');
file_put_contents($temp, '[{"id":"001","name":"A"},{"id":"002","name":"B"}]');
try {
    $rows = \Mnb\PHPExcel\Format\Json::read($temp)
        ->withHeaderRow()
        ->projectColumns(['id'])
        ->streaming()
        ->toArray();
    check($rows === [['id' => '001'], ['id' => '002']], 'JSON isolated read failed');
    check(\Mnb\PHPExcel\SpreadsheetManager::create()->formats() === ['json'], 'Only JSON should be registered');
} finally { @unlink($temp); }
PHP);

$runIsolated('core + CSV', ['core', 'csv'], <<<'PHP'
$temp = tempnam(sys_get_temp_dir(), 'mnb_csv_');
try {
    \Mnb\PHPExcel\Format\Csv::write([['id' => '001', 'name' => 'A']], $temp, ['with_header' => true]);
    $rows = \Mnb\PHPExcel\Format\Csv::read($temp)->withHeaderRow()->toArray();
    check($rows === [['id' => '001', 'name' => 'A']], 'CSV isolated round-trip failed');
    check(\Mnb\PHPExcel\SpreadsheetManager::create()->formats() === ['csv'], 'Only CSV should be registered');
} finally { @unlink($temp); }
PHP);

foreach (['xml', 'xlsx', 'ods', 'xls'] as $format) {
    $runIsolated('core + ' . strtoupper($format), ['core', $format], "\n"
        . '$formats = \\Mnb\\PHPExcel\\SpreadsheetManager::create()->formats();' . "\n"
        . "check(\$formats === ['{$format}'], 'Only {$format} should be registered');\n");
}

$runIsolated('complete application family', ['core', 'csv', 'json', 'xml', 'xlsx', 'ods', 'xls', 'database', 'application'], <<<'PHP'
check(\Mnb\PHPExcel\MnbExcel::version() === '1.2.0', 'Legacy facade version mismatch');
$formats = \Mnb\PHPExcel\SpreadsheetManager::create()->formats();
foreach (['csv', 'json', 'xml', 'xlsx', 'ods', 'xls'] as $format) {
    check(in_array($format, $formats, true), 'Missing installed format: ' . $format);
}
PHP);

echo "Modular package tests passed for " . count($definitions) . " packages and isolated install combinations.\n";
