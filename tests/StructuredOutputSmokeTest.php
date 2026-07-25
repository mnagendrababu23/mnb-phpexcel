<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

$tmp = sys_get_temp_dir() . '/mnb_structured_output_' . uniqid('', true) . '.csv';
file_put_contents($tmp, "Author Name,Email,Affilication\nPeter Heinz,heinz@example.com,Old header\n");

$session = MnbExcel::readCsv($tmp);
$structured = $session->toStructuredArray([
    'header_row' => true,
    'skip_empty_rows' => true,
    'rename_headers' => [
        'affilication' => 'affiliation',
    ],
]);

if (($structured['status'] ?? null) !== 'ok') {
    throw new RuntimeException('Workbook structured output status failed.');
}
if (!isset($structured['sheets'][0])) {
    throw new RuntimeException('Workbook sheets structure missing.');
}
$sheet = $structured['sheets'][0];
if (($sheet['sheetname'] ?? null) !== 'Sheet1') {
    throw new RuntimeException('Sheet name metadata failed.');
}
if (($sheet['headers'][0] ?? null) !== 'author_name') {
    throw new RuntimeException('Header normalization failed.');
}
if (($sheet['headers'][2] ?? null) !== 'affiliation') {
    throw new RuntimeException('Header rename failed.');
}
if (($sheet['rows'][0]['row_number'] ?? null) !== 2) {
    throw new RuntimeException('Original row number failed.');
}
if (($sheet['rows'][0]['values']['affiliation'] ?? null) !== 'Old header') {
    throw new RuntimeException('Structured row values failed.');
}
if (($sheet['columns'][0]['letter'] ?? null) !== 'A') {
    throw new RuntimeException('Column metadata failed.');
}
if (($sheet['summary']['data_rows'] ?? null) !== 1) {
    throw new RuntimeException('Sheet structured summary failed.');
}
if (($structured['summary']['data_rows'] ?? null) !== 1) {
    throw new RuntimeException('Workbook structured summary failed.');
}

$selectedSheet = $session->toStructuredSheetArray([
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
]);
if (($selectedSheet['headers'][2] ?? null) !== 'affiliation') {
    throw new RuntimeException('Selected-sheet structured output failed.');
}

$json = $session->toStructuredJson([
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
]);
if (!str_contains($json, '"sheets"') || !str_contains($json, '"affiliation"')) {
    throw new RuntimeException('Structured JSON failed.');
}

$compactJson = $session->toStructuredJson([
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
], [
    'pretty' => false,
    'trailing_newline' => false,
]);
if (str_contains($compactJson, "\n") || !str_starts_with($compactJson, '{')) {
    throw new RuntimeException('Structured JSON options failed.');
}

$structuredJsonFile = sys_get_temp_dir() . '/mnb_structured_output_' . uniqid('', true) . '.json';
$savedPath = $session->saveStructuredJson($structuredJsonFile, [
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
], [
    'pretty' => false,
]);
if ($savedPath !== $structuredJsonFile || !is_file($structuredJsonFile) || !str_contains((string) file_get_contents($structuredJsonFile), '"sheets"')) {
    throw new RuntimeException('Structured JSON save failed.');
}

unlink($structuredJsonFile);
unlink($tmp);

echo "Structured output smoke test passed.\n";

// Structured XML should also support direct string return and file save.
$tmpXml = sys_get_temp_dir() . '/mnb_structured_output_' . uniqid('', true) . '.csv';
file_put_contents($tmpXml, "Author Name,Email,Affilication\nPeter Heinz,heinz@example.com,Old header\n");
$xmlSession = MnbExcel::readCsv($tmpXml);
$xml = $xmlSession->toStructuredXml([
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
], [
    'root' => 'structured_workbook',
]);
if (!str_contains($xml, '<structured_workbook>') || !str_contains($xml, '<sheets>') || !str_contains($xml, 'affiliation')) {
    throw new RuntimeException('Structured XML failed.');
}
$structuredXmlFile = sys_get_temp_dir() . '/mnb_structured_output_' . uniqid('', true) . '.xml';
$savedXmlPath = $xmlSession->saveStructuredXml($structuredXmlFile, [
    'header_row' => true,
    'rename_headers' => ['affilication' => 'affiliation'],
], [
    'root' => 'structured_workbook',
]);
if ($savedXmlPath !== $structuredXmlFile || !is_file($structuredXmlFile) || !str_contains((string) file_get_contents($structuredXmlFile), '<structured_workbook>')) {
    throw new RuntimeException('Structured XML save failed.');
}
unlink($structuredXmlFile);
unlink($tmpXml);
