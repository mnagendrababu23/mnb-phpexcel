<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$data = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321', 'marks' => 95],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678', 'marks' => 88],
];

MnbExcel::fromArray($data)
    ->withHeader()
    ->columns([
        'name' => 'Student Name',
        'email' => 'Email Address',
        'phone' => 'Phone Number',
        'marks' => 'Marks',
    ])
    ->textColumns(['phone'])
    ->numberColumns(['marks'])
    ->styleHeader(['fill' => '#EEF4FF'])
    ->freezeHeader()
    ->autoFilter()
    ->save(__DIR__ . '/students.xlsx');

echo "Created examples/students.xlsx\n";
