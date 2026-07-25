<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

echo "JsonXmlSmokeTest\n";

$rows = [
    ['sku' => 'A100', 'qty' => 5],
    ['sku' => 'B200', 'qty' => 12],
];

smoke_run('converts an array to a JSON string', function () use ($rows): void {
    $json = MnbExcel::fromArray($rows)->toJson();
    $decoded = json_decode($json, true);

    smoke_assert_true(is_array($decoded), 'toJson() output should decode back into an array');
    smoke_assert_equals(2, count($decoded), 'decoded JSON should have the same row count');
    smoke_assert_equals('A100', $decoded[0]['sku'] ?? null, 'first row sku should round-trip through JSON');

    $indexed = json_decode(MnbExcel::fromArray($rows)->toJson(['preserve_associative_rows' => false]), true);
    smoke_assert_equals('A100', $indexed[0][0] ?? null, 'indexed structured output should remain available as an opt-out');
});

smoke_run('converts an array to an XML string', function () use ($rows): void {
    $xml = MnbExcel::fromArray($rows)->toXml();

    smoke_assert_true($xml !== '', 'toXml() should not return an empty string');
    smoke_assert_contains('<sku>A100</sku>', $xml, 'XML output should preserve associative field names');

    if (function_exists('simplexml_load_string')) {
        $loaded = simplexml_load_string($xml);
        smoke_assert_true($loaded !== false, 'generated XML should be well-formed');
    }
});

smoke_run('round-trips a saved JSON file back through readJson()', function () use ($rows): void {
    $dir = smoke_temp_dir('json');
    $path = $dir . DIRECTORY_SEPARATOR . 'inventory.json';

    MnbExcel::fromArray($rows)->saveJson($path);
    smoke_assert_true(is_file($path), 'saveJson() should write a file');

    $read = MnbExcel::readJson($path)->withHeaderRow()->toArray();
    smoke_assert_equals(count($rows), count($read), 'row count should round-trip through a saved JSON file');
});

echo "JsonXmlSmokeTest: all assertions passed.\n";
