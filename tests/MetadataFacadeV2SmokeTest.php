<?php
declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Metadata\MetadataDiff;
use Mnb\PHPExcel\Metadata\MetadataResult;

$tmp = tempnam(sys_get_temp_dir(), 'mnb_meta_v2_');
$csv = $tmp . '.csv';
@unlink($tmp);
file_put_contents($csv, "name,amount\nA,10\nB,20\n");

try {
    $query = MnbExcel::meta($csv)->quick()->only(['file','workbook','statistics']);
    $result = $query->read();

    if (!$result instanceof MetadataResult) throw new RuntimeException('Expected MetadataResult.');
    if ($result->format() !== 'csv') throw new RuntimeException('Expected CSV format.');
    if ($result->sheetCount() !== 1) throw new RuntimeException('Expected one synthetic CSV sheet.');
    if (!isset($result['file'])) throw new RuntimeException('ArrayAccess failed.');
    if (isset($result->toArray()['security'])) throw new RuntimeException('only() did not filter sections.');

    $diff = MetadataDiff::between(['a'=>1], ['a'=>2]);
    if (($diff['count'] ?? 0) !== 1) throw new RuntimeException('Metadata diff failed.');

    $a = MnbExcel::meta($csv)->quick()->read()->toArray();
    $b = MnbExcel::meta($csv)->quick()->read()->toArray();
    if ($a !== $b) throw new RuntimeException('Cached metadata mismatch.');

    echo "Metadata facade V2 smoke test passed.\n";
} finally {
    @unlink($csv);
}
