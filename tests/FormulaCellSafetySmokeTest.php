<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

$rows = [
    [
        'name' => 'Ravi',
        'phone' => MnbExcel::text('0987654321'),
        'total' => MnbExcel::formula('SUM(C2:D2)', 150),
        'note' => '=HYPERLINK("http://example.com","bad")',
    ],
];

$scan = MnbExcel::scanCells($rows);
if ($scan['total_issues'] < 1) {
    throw new RuntimeException('Expected formula-like text issue.');
}

$builder = MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->cellSafety(['max_text_length' => 32767])
    ->formulaPolicy('safe');

$csv = tempnam(sys_get_temp_dir(), 'mnb_formula_safety_') . '.csv';
$builder->save($csv);
$content = file_get_contents($csv);
@unlink($csv);

if (!is_string($content) || !str_contains($content, "'=HYPERLINK")) {
    throw new RuntimeException('Expected dangerous formula-like text to be escaped in CSV.');
}

try {
    MnbExcel::fromArray([
        ['bad' => MnbExcel::formula('HYPERLINK("http://evil.test","bad")')],
    ])->withHeader()->toArray();
    throw new RuntimeException('Expected unsafe explicit formula to be blocked.');
} catch (MnbExcelException $e) {
    if (!str_contains($e->getMessage(), 'Unsafe formula')) {
        throw $e;
    }
}

try {
    MnbExcel::fromArray([
        ['long' => str_repeat('A', 20)],
    ])->cellSafety(['max_text_length' => 10, 'long_text_policy' => 'error'])->toArray();
    throw new RuntimeException('Expected long text to be blocked.');
} catch (MnbExcelException $e) {
    if (!str_contains($e->getMessage(), 'exceeds maximum length')) {
        throw $e;
    }
}

echo "Formula/cell safety smoke test passed.\n";
