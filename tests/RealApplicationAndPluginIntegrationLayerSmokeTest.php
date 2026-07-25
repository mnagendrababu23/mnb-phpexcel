<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Plugin\PluginRegistry;

$base = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-real-app-' . uniqid('', true);
$paths = MnbExcel::storage(['base_path' => $base]);
assert(is_dir($paths['temp_path']));
assert(is_dir($paths['manifest_path']));

$csv = $paths['upload_path'] . DIRECTORY_SEPARATOR . 'students.csv';
file_put_contents($csv, "id,email\n1,a@example.com\n");
$upload = MnbExcel::validateUpload(['tmp_name' => $csv, 'name' => 'students.csv', 'size' => filesize($csv)], [
    'allowed_extensions' => ['csv'],
    'max_size_mb' => 1,
]);
assert($upload['valid'] === true);

$events = [];
MnbExcel::on('demo_event', function (array $payload) use (&$events): void {
    $events[] = $payload['value'] ?? null;
});
MnbExcel::dispatch('demo_event', ['value' => 42]);
assert($events === [42]);

MnbExcel::transformer('trim_strings', function (array $row): array {
    foreach ($row as $key => $value) {
        if (is_string($value)) {
            $row[$key] = trim($value);
        }
    }
    return $row;
});

MnbExcel::validator('even_number', function (mixed $value, array $row, array $context): bool|string {
    return ((int) $value) % 2 === 0 ? true : ($context['column'] . ' must be even.');
});

$result = MnbExcel::validateArray([
    ['code' => 2],
    ['code' => 3],
], ['code' => 'required|even_number']);
assert(count($result['valid']) === 1);
assert(count($result['failed']) === 1);
assert($result['failed'][0]['errors'][0] === 'code must be even.');

MnbExcel::plugin(function (PluginRegistry $registry): void {
    $registry->importProfile('student_import', [
        'table' => 'students',
        'with_header' => true,
        'chunk_size' => 250,
        'batch_size' => 100,
        'rules' => ['email' => 'nullable|email'],
        'transformers' => ['trim_strings'],
    ]);
    $registry->event('plugin_event', function (array $payload): void {
        $GLOBALS['mnb_plugin_event_seen'] = $payload['ok'] ?? false;
    });
});
assert(in_array('closure@0', MnbExcel::plugins(), true));
assert(MnbExcel::profile('student_import')->config()['table'] === 'students');
MnbExcel::dispatch('plugin_event', ['ok' => true]);
assert(($GLOBALS['mnb_plugin_event_seen'] ?? false) === true);

$manifest = MnbExcel::storagePath('manifest_path', 'status.json');
file_put_contents($manifest, json_encode([
    'status' => 'paused',
    'source_path' => 'large.xlsx',
    'table' => 'students',
    'rows_scanned' => 50,
    'total_rows' => 100,
    'inserted_rows' => 49,
    'failed_rows' => 1,
], JSON_PRETTY_PRINT));
$status = MnbExcel::importStatus($manifest);
assert($status['status'] === 'paused');
assert($status['percent'] === 50.0);
assert($status['resume_ready'] === true);

$logs = [];
MnbExcel::setLogger(function (string $level, string $message, array $context) use (&$logs): void {
    $logs[] = [$level, $message, $context];
});
\Mnb\PHPExcel\Application\LoggerBridge::info('hello', ['x' => 1]);
assert($logs[0][0] === 'info');

$it = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($base, FilesystemIterator::SKIP_DOTS), RecursiveIteratorIterator::CHILD_FIRST);
foreach ($it as $file) {
    $file->isDir() ? rmdir($file->getPathname()) : unlink($file->getPathname());
}
rmdir($base);

echo "RealApplicationAndPluginIntegrationLayerSmokeTest passed\n";
