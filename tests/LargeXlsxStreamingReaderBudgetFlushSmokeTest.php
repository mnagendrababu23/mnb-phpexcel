<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class)) {
    echo "LargeXlsxStreamingReaderBudgetFlushSmokeTest skipped because ext-zip/ext-xmlreader is unavailable.\n";
    return;
}

$tmp = sys_get_temp_dir() . '/mnb-large-budget-' . uniqid('', true) . '.xlsx';
MnbExcel::fromArray([
    ['id' => 1, 'name' => 'One'],
    ['id' => 2, 'name' => 'Two'],
    ['id' => 3, 'name' => 'Three'],
])->withHeader()->save($tmp);

$chunks = [];
$summary = MnbExcel::largeRead($tmp)
    ->withHeader()
    ->limitRows(1)
    ->chunk(10, function (array $rows) use (&$chunks): void {
        $chunks[] = $rows;
    });

assert($summary['stopped'] === true);
assert($summary['stop_reason'] === 'row_limit_reached');
assert($summary['rows_delivered'] === 1);
assert(count($chunks) === 1);
assert(count($chunks[0]) === 1);
assert($chunks[0][0]['id'] === 1);

@unlink($tmp);

echo "LargeXlsxStreamingReaderBudgetFlushSmokeTest passed\n";
