<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

// Example JSON file shape:
// [
//   {"student_id":"000123", "name":"Ravi", "address":{"city":"Hyderabad"}},
//   {"student_id":"000124", "name":"Sita", "address":{"city":"Vijayawada"}}
// ]

MnbExcel::fromJson(__DIR__ . '/students.json')
    ->textColumns(['student_id', 'phone'])
    ->autoWidth()
    ->freezeHeader()
    ->autoFilter()
    ->metadata([
        'title' => 'Students from JSON',
        'creator' => 'MNB PHPExcel',
    ])
    ->save(__DIR__ . '/students-from-json.xlsx');

// Multi-sheet JSON is also supported:
// {
//   "Students": [{"name":"Ravi"}],
//   "Teachers": [{"name":"Kumar", "subject":"Maths"}]
// }

MnbExcel::fromJson(__DIR__ . '/school.json')
    ->autoWidth()
    ->save(__DIR__ . '/school-from-json.xlsx');
