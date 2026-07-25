<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class)) {
    echo "LargeXlsxStreamingReaderSmokeTest skipped because ext-zip/ext-xmlreader is unavailable.\n";
    return;
}

$tmp = sys_get_temp_dir() . '/mnb-large-stream-' . uniqid('', true) . '.xlsx';
MnbExcel::fromArray([
    ['id' => 1, 'name' => 'One'],
    ['id' => 2, 'name' => 'Two'],
    ['id' => 3, 'name' => 'Three'],
])->withHeader()->save($tmp);

$chunks = [];
$summary = MnbExcel::largeRead($tmp)
    ->withHeader()
    ->chunk(2, function (array $rows) use (&$chunks): void {
        $chunks[] = $rows;
    });

assert($summary['rows_delivered'] === 3);
assert(count($chunks) === 2);
assert($chunks[0][0]['id'] === 1);
assert($chunks[0][1]['name'] === 'Two');
assert($chunks[1][0]['id'] === 3);
@unlink($tmp);

echo "LargeXlsxStreamingReaderSmokeTest passed\n";
