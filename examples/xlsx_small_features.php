<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$rows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321', 'marks' => 95],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678', 'marks' => 88],
    ['name' => 'Kiran', 'email' => 'kiran@example.com', 'phone' => '0900000001', 'marks' => 76],
];

MnbExcel::fromArray($rows)
    ->withHeader()
    ->columns([
        'name' => 'Student Name',
        'email' => 'Email Address',
        'phone' => 'Phone Number',
        'marks' => 'Marks',
    ])
    ->textColumns(['phone'])
    ->numberColumns(['marks'])
    ->styleHeader([
        'font' => ['bold' => true, 'color' => '#FFFFFF', 'size' => 12],
        'fill' => '#1F6FEB',
        'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
        'border' => ['color' => '#D0D7DE'],
    ])
    ->columnWidths([
        'A' => 22,
        'B' => 30,
        'C' => 18,
        'D' => 12,
    ])
    ->rowHeight(1, 24)
    ->freezeHeader()
    ->autoFilter()
    ->mergeCells('A5:D5')
    ->addImage(__DIR__ . '/assets/mnb-logo.png', 'F1', ['width' => 80, 'height' => 40, 'name' => 'MNB Logo'])
    ->save(__DIR__ . '/output/students-small-features.xlsx');

echo "Created examples/output/students-small-features.xlsx\n";
