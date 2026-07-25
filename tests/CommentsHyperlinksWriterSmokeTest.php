<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

if (!class_exists(ZipArchive::class)) {
    echo "CommentsHyperlinksWriterSmokeTest: SKIP ext-zip unavailable\n";
    return;
}

$path = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-comments-hyperlinks-' . uniqid('', true) . '.xlsx';

MnbExcel::fromArray([
    ['Name' => 'MNB PHPExcel', 'URL' => 'Repository', 'Status' => 'Ready'],
])->withHeader()
    ->hyperlink('B2', 'https://github.com/mnagendrababu23/mnb-phpexcel', 'Open repository', ['tooltip' => 'GitHub repository'])
    ->comment('C2', 'QA', 'Generated note/comment metadata must be valid.')
    ->save($path);

$result = MnbExcel::validateXlsx($path);
assert(($result['valid'] ?? false) === true, 'Generated XLSX with comments/hyperlinks must pass integrity validation.');

$zip = new ZipArchive();
assert($zip->open($path) === true);

$sheetXml = (string) $zip->getFromName('xl/worksheets/sheet1.xml');
$relsXml = (string) $zip->getFromName('xl/worksheets/_rels/sheet1.xml.rels');
$commentsXml = (string) $zip->getFromName('xl/comments1.xml');
$vmlXml = (string) $zip->getFromName('xl/drawings/vmlDrawing1.vml');
$contentTypes = (string) $zip->getFromName('[Content_Types].xml');
$zip->close();

assert(str_contains($sheetXml, '<hyperlink ref="B2"'));
assert(str_contains($sheetXml, '<legacyDrawing r:id='));
assert(str_contains($relsXml, '/relationships/hyperlink'));
assert(str_contains($relsXml, '/relationships/comments'));
assert(str_contains($relsXml, '/relationships/vmlDrawing'));
assert(str_contains($commentsXml, '<author>QA</author>'));
assert(str_contains($commentsXml, 'Generated note/comment metadata must be valid.'));
assert(str_contains($vmlXml, 'ObjectType="Note"'));
assert(str_contains($contentTypes, '/xl/comments1.xml'));

if (class_exists(XMLReader::class)) {
    $metadata = MnbExcel::read($path)->sheet(1)->sheetMetadata();
    assert(($metadata['summary']['comments'] ?? 0) === 1);
    assert(($metadata['summary']['hyperlinks'] ?? 0) === 1);
}

@unlink($path);

echo "CommentsHyperlinksWriterSmokeTest: PASS\n";
