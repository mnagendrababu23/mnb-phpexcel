<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\DatabaseConfigResolver;
use Mnb\PHPExcel\Support\DatabaseConnectionFactory;

$env = tempnam(sys_get_temp_dir(), 'mnb-db-env-');
file_put_contents($env, "DB_CONNECTION=mysql\nDB_HOST=localhost\nDB_PORT=3307\nDB_DATABASE=mnb_test\nDB_USERNAME=mnb_user\nDB_PASSWORD=secret\nDB_CHARSET=utf8mb4\n");
$config = MnbExcel::dbConfig($env);
assert($config['driver'] === 'mysql');
assert($config['dsn'] === 'mysql:host=localhost;port=3307;dbname=mnb_test;charset=utf8mb4');
assert($config['username'] === 'mnb_user');
assert($config['password'] === 'secret');
@unlink($env);

$phpConfig = tempnam(sys_get_temp_dir(), 'mnb-db-php-') . '.php';
file_put_contents($phpConfig, "<?php\nreturn ['driver' => 'pgsql', 'host' => 'db.local', 'port' => 5433, 'database' => 'app', 'username' => 'u', 'password' => 'p'];\n");
$config = MnbExcel::dbConfig($phpConfig);
assert($config['driver'] === 'pgsql');
assert($config['dsn'] === 'pgsql:host=db.local;port=5433;dbname=app');
@unlink($phpConfig);

$config = MnbExcel::dbConfig(['dsn' => 'sqlite::memory:']);
assert($config['driver'] === 'sqlite');
assert($config['dsn'] === 'sqlite::memory:');

$config = MnbExcel::dbConfig('mysql:host=127.0.0.1;dbname=app;charset=utf8mb4', ['username' => 'root']);
assert($config['driver'] === 'mysql');
assert($config['username'] === 'root');

$summary = MnbExcel::databaseConfigSummary(['base_path' => __DIR__]);
assert($summary['status'] === 'ok');
assert(is_array($summary['checked_paths']));

if (class_exists(PDO::class) && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = MnbExcel::pdo(['dsn' => 'sqlite::memory:']);
    assert($pdo instanceof PDO);
    $pdo2 = DatabaseConnectionFactory::make($pdo);
    assert($pdo2 === $pdo);
}

echo "DatabaseConnectionConfigSmokeTest passed\n";
