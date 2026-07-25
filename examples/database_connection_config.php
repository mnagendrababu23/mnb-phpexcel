<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\MnbExcel;

// Option 1: use your real application .env file.
$pdo = MnbExcel::pdo(__DIR__ . '/../.env');

// Option 2: use constants from your app bootstrap/config file.
// define('MNB_EXCEL_DB_DSN', 'mysql:host=127.0.0.1;dbname=app;charset=utf8mb4');
// define('MNB_EXCEL_DB_USERNAME', 'root');
// define('MNB_EXCEL_DB_PASSWORD', 'secret');
// $pdo = MnbExcel::pdo();

// Option 3: use a PHP config file that returns an array.
// $pdo = MnbExcel::pdo(__DIR__ . '/../config/database.php');

// Option 4: pass an array directly.
// $pdo = MnbExcel::pdo([
//     'driver' => 'mysql',
//     'host' => '127.0.0.1',
//     'database' => 'app',
//     'username' => 'root',
//     'password' => 'secret',
// ]);

// Use the same connection for normal SQL export.
MnbExcel::fromSql($pdo, 'SELECT id, name, email FROM students')
    ->withHeader()
    ->save('students.xlsx');

// Or let the SQL import helper create the PDO from .env/config automatically.
MnbExcel::largeImportToSql('students-large.xlsx', __DIR__ . '/../.env', 'students', [
    'with_header' => true,
    'chunk_size' => 500,
    'batch_size' => 250,
    'resume' => true,
    'idempotent' => true,
    'duplicate_strategy' => 'skip',
    'unique_by' => ['id'],
]);
