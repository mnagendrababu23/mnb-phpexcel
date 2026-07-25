<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class) || !in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    echo "LargeExcelDatabaseImportEngineSmokeTest skipped because ext-zip/ext-xmlreader/pdo_sqlite is unavailable.\n";
    return;
}

$tmpBase = sys_get_temp_dir() . '/mnb-large-db-' . uniqid('', true);
$xlsx = $tmpBase . '.xlsx';
$manifest = $tmpBase . '.manifest.json';
$failed = $tmpBase . '.failed.csv';
$db = $tmpBase . '.sqlite';

MnbExcel::fromArray([
    ['id' => 1, 'email' => 'one@example.com', 'amount' => 10],
    ['id' => 2, 'email' => 'bad-email', 'amount' => 15],
    ['id' => 3, 'email' => 'three@example.com', 'amount' => 20],
])->withHeader()->save($xlsx);

$pdo = new PDO('sqlite:' . $db);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE imports (id INTEGER, email TEXT, amount INTEGER)');

$result = MnbExcel::largeImportToSql($xlsx, $pdo, 'imports', [
    'chunk_size' => 2,
    'batch_size' => 2,
    'manifest_path' => $manifest,
    'failed_rows_csv' => $failed,
    'rules' => [
        'id' => 'required|integer',
        'email' => 'required|email',
        'amount' => 'required|numeric',
    ],
]);

assert($result['status'] === 'completed');
assert($result['rows_scanned'] === 3);
assert($result['valid_rows'] === 2);
assert($result['failed_rows'] === 1);
assert($result['inserted_rows'] === 2);
assert(is_file($manifest));
assert(is_file($failed));
assert((int) $pdo->query('SELECT COUNT(*) FROM imports')->fetchColumn() === 2);

@unlink($xlsx);
@unlink($manifest);
@unlink($failed);
@unlink($db);

echo "LargeExcelDatabaseImportEngineSmokeTest passed\n";
