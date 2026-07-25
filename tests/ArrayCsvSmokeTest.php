<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$tmp = sys_get_temp_dir() . '/mnb-phpexcel-smoke.csv';

$data = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'phone' => '0987654321'],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'phone' => '0912345678'],
];

MnbExcel::fromArray($data)
    ->withHeader()
    ->textColumns(['phone'])
    ->save($tmp);

$rows = MnbExcel::readCsv($tmp)->toArray(['header_row' => true]);

assert($rows[0]['phone'] === '0987654321');
assert($rows[1]['email'] === 'sita@example.com');

unlink($tmp);

echo "Array/CSV smoke test passed.\n";
