<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$options = getopt('', ['rows::', 'cols::', 'path::', 'progress::', 'json']);
$rows = max(1, (int) ($options['rows'] ?? 10000));
$cols = max(1, (int) ($options['cols'] ?? 10));
$path = (string) ($options['path'] ?? sys_get_temp_dir() . DIRECTORY_SEPARATOR . "mnb-large-writer-{$rows}x{$cols}.xlsx");
$progress = max(0, (int) ($options['progress'] ?? 0));

$result = MnbExcel::benchmarkLargeWriter($rows, $cols, [
    'path' => $path,
    'progress_every' => $progress,
]);

if (isset($options['json'])) {
    echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;
    exit(0);
}

echo 'MNB PHPExcel large writer benchmark' . PHP_EOL;
echo 'Rows: ' . $result['rows'] . PHP_EOL;
echo 'Columns: ' . $result['columns'] . PHP_EOL;
echo 'Elapsed seconds: ' . $result['elapsed_seconds'] . PHP_EOL;
echo 'Peak memory MB: ' . $result['peak_memory_mb'] . PHP_EOL;
echo 'Output size MB: ' . $result['output_size_mb'] . PHP_EOL;
echo 'Rows/sec: ' . $result['rows_per_second'] . PHP_EOL;
echo 'Output: ' . $result['path'] . PHP_EOL;
