<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$path = __DIR__ . '/large-import.xlsx';

$plan = MnbExcel::autoImportPlan($path, [
    'server' => 'shared',
    'memory_limit' => '256M',
    'max_execution_time' => 30,
], [
    'accurate_row_count' => true,
    'scan_features' => true,
    'time_budget_seconds' => 20,
]);

print_r([
    'method' => $plan['selected_method'],
    'chunk_size' => $plan['chunk_size'],
    'route' => $plan['route'],
    'rows' => $plan['profile']['total_rows'],
    'columns' => $plan['profile']['max_columns'],
    'risk' => $plan['profile']['risk'],
]);

if ($plan['selected_method'] === 'streaming_chunk_import' || $plan['selected_method'] === 'cli_queue_recommended') {
    MnbExcel::largeRead($path)
        ->withHeader()
        ->timeBudgetSeconds(25) // HTTP-safe soft stop; use 0 in CLI workers.
        ->chunk((int) $plan['chunk_size'], function (array $rows, array $state): void {
            // validate and insert this chunk into DB here.
            // Do not collect chunks into one giant array.
            echo 'Imported rows: ' . $state['rows_delivered'] . PHP_EOL;
        });
}
