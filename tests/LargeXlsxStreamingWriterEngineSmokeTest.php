<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$rows = (static function (): Generator {
    yield ['id' => 1, 'name' => 'Alpha', 'amount' => 12.3456];
    yield ['id' => 2, 'name' => 'Beta', 'amount' => 99.01];
})();

$session = MnbExcel::largeExport($rows)
    ->withHeader()
    ->sheetName('Large Export')
    ->autoSplitSheets(true, 1000)
    ->formatColumn('amount', 'decimal')
    ->progress(static function (array $state): void {
        assert(isset($state['rows_exported']));
    }, 1);

$options = $session->options();
assert(($options['with_header'] ?? false) === true);
assert(($options['sheet_name'] ?? '') === 'Large Export');
assert(($options['max_rows_per_sheet'] ?? 0) === 1000);
assert(($options['column_formats']['amount'] ?? '') === 'decimal');

if (!class_exists(ZipArchive::class)) {
    echo "LargeXlsxStreamingWriterEngineSmokeTest skipped runtime save because ext-zip is unavailable.\n";
    exit(0);
}

$out = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-large-writer-' . uniqid('', true) . '.xlsx';
$result = $session->save($out);
assert(is_file($out));
assert(($result['rows_exported'] ?? 0) === 2);
assert(($result['sheets_created'] ?? 0) === 1);
@unlink($out);

echo "LargeXlsxStreamingWriterEngineSmokeTest passed\n";
