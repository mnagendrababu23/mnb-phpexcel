<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$alert = MnbExcel::environmentAlert();
$message = MnbExcel::environmentAlertMessage();

assert(isset($alert['status'], $alert['alerts'], $alert['missing'], $alert['message']));
assert(is_array($alert['alerts']));
assert(is_string($message));
assert($message !== '');
assert(array_key_exists('ext_zip', $alert['missing']));
assert(array_key_exists('ext_xmlreader', $alert['missing']));
assert(array_key_exists('pdo_sqlite', $alert['missing']));

if (!class_exists(ZipArchive::class)) {
    assert(str_contains($message, 'ext-zip'));
}
if (!class_exists(XMLReader::class)) {
    assert(str_contains($message, 'ext-xmlreader'));
}
if (!class_exists(PDO::class) || !in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    assert(str_contains($message, 'pdo_sqlite'));
}

echo "EnvironmentAlertSmokeTest passed\n";
