<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

try {
    $path = MnbExcel::fromArray([
        ['name' => 'Ravi', 'amount' => 1200],
        ['name' => 'Sita', 'amount' => 1800],
    ])
        ->withHeader()
        ->save(__DIR__ . '/safe-report.xlsx');

    echo "Saved safely: {$path}\n";
} catch (MnbExcelException $e) {
    // Safe for API/frontend output. Does not expose local file paths or internals.
    $publicError = MnbExcel::safeError($e);
    echo json_encode($publicError, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES) . PHP_EOL;

    // Use this only for local logs/admin/debug screens.
    error_log(json_encode(MnbExcel::errorReport($e, true), JSON_UNESCAPED_SLASHES));
}
