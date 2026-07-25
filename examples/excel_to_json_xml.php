<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

// XLSX example:
// MnbExcel::read(__DIR__ . '/students.xlsx')
//     ->sheet('Students')
//     ->saveJson(__DIR__ . '/exports/students.json', [
//         'header_row' => true,
//         'skip_empty_rows' => true,
//     ]);
//
// MnbExcel::read(__DIR__ . '/students.xlsx')
//     ->sheet('Students')
//     ->saveXml(__DIR__ . '/exports/students.xml', [
//         'header_row' => true,
//         'skip_empty_rows' => true,
//     ], [
//         'root' => 'students',
//         'row' => 'student',
//     ]);

// Array-first example, useful when data already comes from SQL/API/PHP arrays.
$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => MnbExcel::text('0987654321')],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => MnbExcel::text('0912345678')],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->saveJson(__DIR__ . '/students.json', ['mode' => 'rows']);

MnbExcel::fromArray($rows)
    ->withHeader()
    ->textColumns(['phone'])
    ->saveXml(__DIR__ . '/students.xml', [
        'root' => 'students',
        'row' => 'student',
    ]);

echo "Created students.json and students.xml\n";
