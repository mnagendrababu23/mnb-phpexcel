<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$rows = [
    [
        'student' => 'Ravi',
        'student_id' => MnbExcel::text('000123456789012345'),
        'maths' => 80,
        'science' => 70,
        'total' => MnbExcel::formula('SUM(C2:D2)', 150),
        'comment' => '=HYPERLINK("http://example.com", "Unsafe text")',
    ],
    [
        'student' => 'Sita',
        'student_id' => MnbExcel::text('000987654321098765'),
        'maths' => 90,
        'science' => 85,
        'total' => MnbExcel::formula('SUM(C3:D3)', 175),
        'comment' => 'Safe text value',
    ],
];

$safety = MnbExcel::scanCells($rows);
print_r($safety);

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['student_id'])
    ->formulaPolicy('safe')
    ->cellSafety([
        'max_text_length' => 32767,
        'long_text_policy' => 'truncate',
        'control_char_policy' => 'strip',
    ])
    ->autoWidth()
    ->freezeHeader()
    ->autoFilter()
    ->save(__DIR__ . '/output/formula-cell-safety.xlsx');

MnbExcel::fromArray($rows)
    ->withHeader()
    ->csvInjectionPolicy('escape')
    ->save(__DIR__ . '/output/formula-cell-safety.csv');
