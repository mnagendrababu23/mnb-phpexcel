<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$data = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321'],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678'],
];

MnbExcel::fromArray($data)
    ->withHeader()
    ->textColumns(['phone'])
    ->save(__DIR__ . '/students.csv');

echo "Created examples/students.csv\n";
