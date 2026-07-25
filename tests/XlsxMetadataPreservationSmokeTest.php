<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/bootstrap.php';

if (!class_exists(ZipArchive::class) || !class_exists(XMLReader::class)) {
    echo "XLSX metadata preservation smoke test skipped: ext-zip/ext-xmlreader unavailable.\n";
    exit(0);
}

$source = sys_get_temp_dir() . '/mnb-phpexcel-metadata-source-' . uniqid('', true) . '.xlsx';
$output = sys_get_temp_dir() . '/mnb-phpexcel-metadata-output-' . uniqid('', true) . '.xlsx';

$zip = new ZipArchive();
if ($zip->open($source, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    throw new RuntimeException('Unable to create test XLSX.');
}

$zip->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    . '<Default Extension="xml" ContentType="application/xml"/>'
    . '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
    . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    . '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    . '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    . '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'
    . '</Types>');
$zip->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    . '</Relationships>');
$zip->addFromString('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    . '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
$zip->addFromString('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    . '</Relationships>');
$zip->addFromString('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    . '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    . '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    . '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    . '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    . '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
$zip->addFromString('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">'
    . '<si><r><t>Hello </t></r><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>World</t></r></si>'
    . '<si><t>Website</t></si>'
    . '</sst>');
$zip->addFromString('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    . '<dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>'
    . '<hyperlinks><hyperlink ref="B1" r:id="rId1" tooltip="Open website"/></hyperlinks>'
    . '<legacyDrawing r:id="rId2"/>'
    . '</worksheet>');
$zip->addFromString('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
    . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
    . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>'
    . '</Relationships>');
$zip->addFromString('xl/comments1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    . '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    . '<authors><author>Nagendra</author></authors>'
    . '<commentList><comment ref="A1" authorId="0"><text><r><t>Check this value</t></r></text></comment></commentList>'
    . '</comments>');
$zip->addFromString('xl/drawings/vmlDrawing1.vml', '<xml></xml>');
$zip->close();

$metadata = MnbExcel::read($source)->sheet(1)->sheetMetadata();
if (($metadata['summary']['rich_text_cells'] ?? 0) !== 1) {
    throw new RuntimeException('Rich text metadata was not detected.');
}
if (($metadata['rich_text'][0]['runs'][1]['bold'] ?? false) !== true) {
    throw new RuntimeException('Rich text run formatting was not detected.');
}
if (($metadata['summary']['comments'] ?? 0) !== 1 || ($metadata['comments'][0]['author'] ?? null) !== 'Nagendra') {
    throw new RuntimeException('Comment metadata was not detected.');
}
if (($metadata['summary']['hyperlinks'] ?? 0) !== 1 || ($metadata['hyperlinks'][0]['target'] ?? null) !== 'https://example.com') {
    throw new RuntimeException('Hyperlink metadata was not detected.');
}

$structured = MnbExcel::read($source)->toStructuredArray(['include_cell_metadata' => true]);
if (($structured['sheets'][0]['metadata']['summary']['comments'] ?? 0) !== 1) {
    throw new RuntimeException('Structured metadata was not included.');
}

MnbExcel::fromArray([
    ['Title' => 'Updated', 'Link' => 'Website'],
])
    ->withHeader()
    ->preserveAdvancedObjectsFrom($source)
    ->save($output);

$check = new ZipArchive();
$check->open($output);
if ($check->locateName('xl/comments1.xml') === false || $check->locateName('xl/worksheets/_rels/sheet1.xml.rels') === false) {
    $check->close();
    throw new RuntimeException('Advanced object parts were not preserved.');
}
$sheetXml = (string) $check->getFromName('xl/worksheets/sheet1.xml');
$check->close();
if (!str_contains($sheetXml, '<hyperlinks>') || !str_contains($sheetXml, '<legacyDrawing')) {
    throw new RuntimeException('Advanced worksheet elements were not preserved.');
}

@unlink($source);
@unlink($output);

echo "XLSX metadata preservation smoke test passed.\n";
