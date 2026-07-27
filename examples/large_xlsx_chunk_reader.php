<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = $argv[1] ?? null;
if ($path === null || !is_file($path)) {
    fwrite(STDERR, "Usage: php examples/large_xlsx_chunk_reader.php <large-workbook.xlsx>\n");
    exit(1);
}

$plan = MnbExcel::autoImportPlan($path, [
    'memory_limit' => ini_get('memory_limit') ?: '128M',
]);

print_r($plan);

MnbExcel::largeRead($path)
    ->sheet(1)
    ->withHeader()
    ->onlyColumns(['A', 'B', 'F'])
    ->chunk(1000, static function (array $rows, array $state): void {
        foreach ($rows as $row) {
            // Replace this line with a batch insert, queue dispatch, or domain handler.
            print_r($row);
        }

        printf("Chunks delivered: %d\n", (int) ($state['chunks_delivered'] ?? 0));
    });
