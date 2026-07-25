<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

$pdo = new PDO('sqlite::memory:');
$pdo->exec('CREATE TABLE students (name TEXT, email TEXT, phone TEXT)');
$pdo->exec("INSERT INTO students VALUES ('Ravi', 'ravi@example.com', '0987654321')");

MnbExcel::fromSql($pdo, 'SELECT name, email, phone FROM students')
    ->withHeader()
    ->textColumns(['phone'])
    ->save(__DIR__ . '/sql-students.xlsx');

echo "Created examples/sql-students.xlsx\n";
