<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$rows = (static function (): Generator {
    for ($i = 1; $i <= 100000; $i++) {
        yield [
            'id' => $i,
            'name' => 'Customer ' . $i,
            'amount' => $i / 10,
            'created_at' => date('Y-m-d', strtotime('2026-01-01 +' . $i . ' days')),
        ];
    }
})();

$result = MnbExcel::largeExport($rows)
    ->withHeader()
    ->sheetName('Customers')
    ->formatColumn('id', 'integer')
    ->formatColumn('amount', 'decimal')
    ->formatColumn('created_at', 'date')
    ->autoSplitSheets(true)
    ->progress(static function (array $state): void {
        echo 'Exported rows: ' . $state['rows_exported'] . PHP_EOL;
    }, 10000)
    ->save(__DIR__ . '/large-customers.xlsx');

print_r($result);
