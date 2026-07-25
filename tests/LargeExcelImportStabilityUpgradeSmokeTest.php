<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class) || !class_exists(PDO::class) || !in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    echo "LargeExcelImportStabilityUpgradeSmokeTest skipped because ext-zip/ext-xmlreader/pdo_sqlite is unavailable.\n";
    return;
}

$tmpBase = sys_get_temp_dir() . '/mnb-large-stability-db-' . uniqid('', true);
$xlsx = $tmpBase . '.xlsx';
$db = $tmpBase . '.sqlite';
$manifest = $tmpBase . '.manifest.json';
$failed = $tmpBase . '.failed.csv';

MnbExcel::fromArray([
    ['id' => 1, 'email' => 'one@example.com', 'amount' => 10],
    ['id' => 2, 'email' => 'bad-email', 'amount' => 15],
    ['id' => 3, 'email' => 'three@example.com', 'amount' => 20],
])->withHeader()->save($xlsx);

$pdo = new PDO('sqlite:' . $db);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$pdo->exec('CREATE TABLE imports (id INTEGER PRIMARY KEY, email TEXT, amount INTEGER)');

$result = MnbExcel::largeImportToSql($xlsx, $pdo, 'imports', [
    'chunk_size' => 2,
    'batch_size' => 2,
    'manifest_path' => $manifest,
    'failed_rows_csv' => $failed,
    'failed_rows_format' => 'human',
    'idempotent' => true,
    'unique_by' => ['id'],
    'rules' => [
        'id' => 'required|integer',
        'email' => 'required|email',
        'amount' => 'required|numeric',
    ],
]);

assert($result['status'] === 'completed');
assert($result['inserted_rows'] === 2);
assert((int) $pdo->query('SELECT COUNT(*) FROM imports')->fetchColumn() === 2);
$failedContent = file_get_contents($failed);
assert(is_string($failedContent) && str_contains($failedContent, 'excel_row,error_count,errors,id,email,amount'));
assert(str_contains($failedContent, 'bad-email'));

// Re-run from the start with idempotent duplicate skip. Existing IDs must not fail.
$secondManifest = $tmpBase . '.second.manifest.json';
$second = MnbExcel::largeImportToSql($xlsx, $pdo, 'imports', [
    'chunk_size' => 3,
    'batch_size' => 3,
    'manifest_path' => $secondManifest,
    'failed_rows_csv' => $tmpBase . '.second.failed.csv',
    'resume' => false,
    'idempotent' => true,
    'unique_by' => ['id'],
    'rules' => [
        'email' => 'nullable|email',
    ],
]);
assert($second['status'] === 'completed');
assert((int) $pdo->query('SELECT COUNT(*) FROM imports')->fetchColumn() === 2);

// All-sheet orchestrator imports selected sheets without loading the workbook.
$multiXlsx = $tmpBase . '.multi.xlsx';
MnbExcel::fromWorkbookArray([
    'Students' => [
        ['id' => 1, 'name' => 'A'],
        ['id' => 2, 'name' => 'B'],
    ],
    'Scores' => [
        ['id' => 1, 'score' => 90],
        ['id' => 2, 'score' => 95],
    ],
])->withHeader()->save($multiXlsx);
$pdo->exec('CREATE TABLE students (id INTEGER PRIMARY KEY, name TEXT)');
$pdo->exec('CREATE TABLE scores (id INTEGER PRIMARY KEY, score INTEGER)');
$all = MnbExcel::largeImportWorkbookToSql($multiXlsx, $pdo, [
    'Students' => 'students',
    'Scores' => 'scores',
], [
    'chunk_size' => 1,
    'batch_size' => 1,
    'manifest_paths' => [
        'Students' => $tmpBase . '.students.manifest.json',
        'Scores' => $tmpBase . '.scores.manifest.json',
    ],
    'failed_rows_csvs' => [
        'Students' => $tmpBase . '.students.failed.csv',
        'Scores' => $tmpBase . '.scores.failed.csv',
    ],
]);
assert($all['status'] === 'completed');
assert($all['sheets_imported'] === 2);
assert((int) $pdo->query('SELECT COUNT(*) FROM students')->fetchColumn() === 2);
assert((int) $pdo->query('SELECT COUNT(*) FROM scores')->fetchColumn() === 2);

foreach (glob($tmpBase . '*') ?: [] as $file) {
    @unlink($file);
}

echo "LargeExcelImportStabilityUpgradeSmokeTest passed\n";
