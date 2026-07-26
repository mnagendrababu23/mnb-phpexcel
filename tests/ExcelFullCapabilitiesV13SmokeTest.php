<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Core\RichTextRun;
use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Format\Csv;
use Mnb\PHPExcel\Reader\Formula\PhpSpreadsheetFormulaEvaluator;
use Mnb\PHPExcel\Reader\HeaderDetector;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\XlsxImageExtractor;
use Mnb\PHPExcel\Reader\XlsxMetadataExtractor;
use Mnb\PHPExcel\Reader\XlsxStyleMap;
use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Writer\XlsxWriter;

function v13_private(object $object, string $method, mixed ...$args): mixed
{
    $reflection = new ReflectionMethod($object, $method);
    $reflection->setAccessible(true);
    return $reflection->invoke($object, ...$args);
}

echo "ExcelFullCapabilitiesV13SmokeTest\n";

smoke_run('detects semantic header rows without requiring mbstring', function (): void {
    $detection = (new HeaderDetector())->detect([
        0 => ['Quarterly sales upload'],
        1 => [],
        2 => ['Order ID', 'Customer Name', 'Amount', 'Order Date'],
        3 => ['A-100', 'Ravi', 125.50, '2026-07-01'],
        4 => ['A-101', 'Maya', 99.00, '2026-07-02'],
    ]);
    smoke_assert_equals(3, $detection->row, 'The semantic header row should be detected after a title and blank row.');
    smoke_assert_true($detection->confidence > 0.35, 'Header confidence should be usable.');

    $options = ReaderOptions::defaults()->withAutoHeader(30, 0.4)->toArray();
    smoke_assert_equals('auto', $options['header_row'], 'Typed options should enable automatic header detection.');
});

smoke_run('integrates automatic headers into read sessions', function (): void {
    $dir = smoke_temp_dir('auto_header');
    $path = $dir . '/upload.csv';
    file_put_contents($path, "Quarterly upload,,,\n,,,\nOrder ID,Customer,Amount,Date\nA-1,Ravi,10.50,2026-07-01\n");
    $rows = Csv::read($path)->autoDetectHeader()->toArray();
    smoke_assert_equals('A-1', $rows[0]['order_id'], 'ReadSession should use the detected row as headers.');
    smoke_assert_equals('Ravi', $rows[0]['customer'], 'Detected headers should map data rows normally.');
});

smoke_run('exposes typed rich-text runs', function (): void {
    $rich = new RichText([
        new RichTextRun('Hello ', ['bold' => true]),
        new RichTextRun('World', ['color' => ['rgb' => 'FFFF0000']]),
    ]);
    smoke_assert_equals('Hello World', $rich->text(), 'Rich text should preserve plain text.');
    smoke_assert_equals(true, $rich->jsonSerialize()['runs'][0]['bold'], 'Rich text should preserve run formatting.');
});

smoke_run('builds full native presentation features', function (): void {
    $builder = WorkbookBuilder::fromArray([
        ['Region' => 'North', 'Sales' => 120, 'Target' => 100],
        ['Region' => 'South', 'Sales' => 80, 'Target' => 100],
    ])
        ->withHeader()
        ->freezeAt('B2')
        ->autoFilterRange('A1:C3')
        ->filterValues('A', ['North'])
        ->conditionalCellIs('B2:B3', 'greaterThan', 100, ['font' => ['bold' => true, 'color' => '#008000']])
        ->conditionalColorScale('B2:B3')
        ->validationList('A2:A100', ['North', 'South', 'East', 'West'])
        ->addChart('column', 'Sales by Region', [[
            'name' => 'Sales',
            'categories' => 'A2:A3',
            'values' => 'B2:B3',
        ]], ['from' => 'E2', 'to' => 'L18']);

    $workbook = $builder->toWorkbookData();
    $sheet = $workbook->sheets[0];
    smoke_assert_equals(1, $sheet->freezeRows, 'freezeAt should freeze the row above the selected cell.');
    smoke_assert_equals(1, $sheet->freezeColumns, 'freezeAt should freeze the column left of the selected cell.');
    smoke_assert_equals('A1:C3', $sheet->autoFilterRange, 'Explicit filter ranges should be preserved.');
    smoke_assert_equals(1, count($sheet->filterColumns), 'Native filter criteria should be stored.');
    smoke_assert_equals(2, count($sheet->conditionalFormats), 'Native conditional-format rules should be stored.');
    smoke_assert_equals(1, count($sheet->dataValidations), 'Data validations should be stored.');
    smoke_assert_equals(1, count($sheet->charts), 'Native chart definitions should be stored.');

    $writer = new XlsxWriter();
    v13_private($writer, 'buildStyleRegistry', $workbook);
    $sheetXml = v13_private($writer, 'worksheetXml', $sheet, true, null);
    smoke_assert_contains('xSplit="1"', $sheetXml, 'Worksheet XML should contain a horizontal frozen pane split.');
    smoke_assert_contains('ySplit="1"', $sheetXml, 'Worksheet XML should contain a vertical frozen pane split.');
    smoke_assert_contains('<autoFilter ref="A1:C3">', $sheetXml, 'Worksheet XML should contain the selected filter range.');
    smoke_assert_contains('<filterColumn colId="0">', $sheetXml, 'Worksheet XML should contain native filter criteria.');
    smoke_assert_contains('<conditionalFormatting sqref="B2:B3">', $sheetXml, 'Worksheet XML should contain native conditional formatting.');
    smoke_assert_contains('<dataValidations count="1">', $sheetXml, 'Worksheet XML should contain data validation.');
    smoke_assert_contains('<drawing r:id="rId1"/>', $sheetXml, 'Worksheet XML should link its chart drawing.');

    $stylesXml = v13_private($writer, 'stylesXml');
    smoke_assert_contains('<dxfs count="1">', $stylesXml, 'Conditional styles should be written as differential formats.');

    $chartPlan = v13_private($writer, 'buildChartPlan', $workbook);
    $chartXml = v13_private($writer, 'chartXml', $chartPlan[1][0], $sheet->name);
    smoke_assert_contains('<c:barChart>', $chartXml, 'Column charts should emit native chart parts.');
    smoke_assert_contains('&apos;Sheet1&apos;!A2:A3', $chartXml, 'Chart category ranges should point at worksheet cells.');

    $scatterBuilder = WorkbookBuilder::fromArray([[1, 2], [2, 4]])->addChart('scatter', 'Scatter', [[
        'name' => 'Series', 'categories' => 'A1:A2', 'values' => 'B1:B2',
    ]]);
    $scatterWorkbook = $scatterBuilder->toWorkbookData();
    $scatterPlan = v13_private($writer, 'buildChartPlan', $scatterWorkbook);
    $scatterXml = v13_private($writer, 'chartXml', $scatterPlan[1][0], $scatterWorkbook->sheets[0]->name);
    smoke_assert_equals(2, substr_count($scatterXml, '<c:valAx>'), 'Scatter charts should use two numeric axes.');
    smoke_assert_equals(0, substr_count($scatterXml, '<c:catAx>'), 'Scatter charts should not use a category axis.');
});

smoke_run('creates reusable import templates with validation and instructions', function (): void {
    $workbook = WorkbookBuilder::importTemplate([
        ['header' => 'SKU', 'required' => true, 'example' => 'SKU-100'],
        ['header' => 'Status', 'list' => ['active', 'inactive'], 'description' => 'Choose a status'],
        ['header' => 'Price', 'validation' => ['type' => 'decimal', 'operator' => 'greaterThan', 'formula1' => 0]],
    ], ['instructions' => 'Complete every required column.'])->toWorkbookData();

    $sheet = $workbook->sheets[0];
    smoke_assert_true($sheet->hasHeader, 'Import templates should include a real header row.');
    smoke_assert_equals(2, count($sheet->dataValidations), 'Template definitions should produce list and numeric validations.');
    smoke_assert_equals(1, count($sheet->comments), 'Required fields should receive explanatory comments.');
});

smoke_run('parses complete XLSX formatting metadata', function (): void {
    $xml = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        . '<numFmts count="1"><numFmt numFmtId="165" formatCode="$#,##0.00"/></numFmts>'
        . '<fonts count="1"><font><b/><i/><u val="double"/><strike/><sz val="14"/><name val="Aptos"/><color rgb="FF123456"/></font></fonts>'
        . '<fills count="1"><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>'
        . '<borders count="1"><border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom style="double"><color theme="1"/></bottom><diagonal/></border></borders>'
        . '<cellXfs count="1"><xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1" textRotation="45"/><protection locked="0" hidden="1"/></xf></cellXfs>'
        . '</styleSheet>';
    $style = XlsxStyleMap::fromXml($xml)->styleForIndex(0);
    smoke_assert_equals('Aptos', $style['font']['name'], 'Font family should be parsed.');
    smoke_assert_equals(true, $style['font']['bold'], 'Bold should be parsed.');
    smoke_assert_equals('solid', $style['fill']['pattern'], 'Fill pattern should be parsed.');
    smoke_assert_equals('thin', $style['border']['left']['style'], 'Border sides should be parsed.');
    smoke_assert_equals('$#,##0.00', $style['number_format'], 'Custom number formats should be parsed.');
    smoke_assert_equals('right', $style['alignment']['horizontal'], 'Alignment should be parsed.');
    smoke_assert_equals(false, $style['protection']['locked'], 'Protection flags should be parsed.');
});

smoke_run('rebinds template pivot caches for refresh-on-open', function (): void {
    $writer = new XlsxWriter();
    $xml = '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" saveData="1"><cacheSource type="worksheet"><worksheetSource ref="A1:C20" sheet="Old Data"/></cacheSource></pivotCacheDefinition>';
    $updated = v13_private($writer, 'rebindPivotCacheDefinition', $xml, [
        'pivot_source_sheet' => 'New Data',
        'pivot_source_range' => 'A1:F500',
        'pivot_refresh_on_load' => true,
    ]);
    smoke_assert_contains('sheet="New Data"', $updated, 'Pivot cache should target the new source sheet.');
    smoke_assert_contains('ref="A1:F500"', $updated, 'Pivot cache should target the new source range.');
    smoke_assert_contains('refreshOnLoad="1"', $updated, 'Pivot cache should refresh when opened.');

    $sheetElements = v13_private($writer, 'extractPreservedSheetElements',
        '<worksheet xmlns="urn:test" xmlns:r="urn:rels"><sheetData/><ignoredErrors/><drawing r:id="rId2"/><tableParts count="1"/><pivotTableParts count="1"><pivotTablePart r:id="rId9"/></pivotTableParts><extLst/></worksheet>'
    );
    smoke_assert_equals(5, count($sheetElements), 'Pivot and related worksheet elements should all be preserved.');
    smoke_assert_contains('<ignoredErrors', $sheetElements[0], 'Preserved elements should retain source schema order.');
    smoke_assert_contains('<pivotTableParts', $sheetElements[3], 'Worksheet pivot-table references should be preserved.');
    smoke_assert_equals(true, v13_private($writer, 'elementsRequireRelationships', $sheetElements), 'Pivot and drawing elements should retain their relationships.');
});

smoke_run('parses namespace-prefixed cell formula and value XML', function (): void {
    $reader = new \Mnb\PHPExcel\Reader\XlsxReader();
    $formula = v13_private($reader, 'formulaFromXml', '<x:c xmlns:x="urn:test"><x:f t="shared" si="2">SUM(A1:A3)</x:f><x:v>6</x:v></x:c>');
    smoke_assert_equals('SUM(A1:A3)', $formula['expression'], 'Formula parsing should be namespace-prefix independent.');
    smoke_assert_equals('2', $formula['metadata']['si'], 'Shared-formula metadata should be preserved.');
    smoke_assert_equals('6', v13_private($reader, 'readV', '<x:c xmlns:x="urn:test"><x:v>6</x:v></x:c>'), 'Cached values should be namespace-prefix independent.');
    smoke_assert_equals('Hello World', v13_private($reader, 'textFromRichXml', '<x:is xmlns:x="urn:test"><x:r><x:t>Hello </x:t></x:r><x:r><x:t>World</x:t></x:r></x:is>'), 'Rich text extraction should be namespace-prefix independent.');
});

smoke_run('parses arbitrary OOXML namespace prefixes across styles, workbook metadata, rich text, and images', function (): void {
    $styleXml = '<?xml version="1.0"?><s:styleSheet xmlns:s="urn:test">'
        . '<s:numFmts count="1"><s:numFmt numFmtId="165" formatCode="0.000"/></s:numFmts>'
        . '<s:fonts count="1"><s:font><s:b/><s:name val="Aptos"/><s:color rgb="FF112233"/></s:font></s:fonts>'
        . '<s:fills count="1"><s:fill><s:patternFill patternType="solid"><s:fgColor rgb="FFFFFF00"/></s:patternFill></s:fill></s:fills>'
        . '<s:borders count="1"><s:border><s:left style="thin"><s:color rgb="FF000000"/></s:left></s:border></s:borders>'
        . '<s:cellXfs count="1"><s:xf numFmtId="165" fontId="0" fillId="0" borderId="0"><s:alignment horizontal="center"/></s:xf></s:cellXfs>'
        . '</s:styleSheet>';
    $style = XlsxStyleMap::fromXml($styleXml)->styleForIndex(0);
    smoke_assert_equals('Aptos', $style['font']['name'], 'Prefixed font elements should be parsed.');
    smoke_assert_equals('solid', $style['fill']['pattern'], 'Prefixed fill elements should be parsed.');
    smoke_assert_equals('thin', $style['border']['left']['style'], 'Prefixed border elements should be parsed.');
    smoke_assert_equals('center', $style['alignment']['horizontal'], 'Prefixed alignment elements should be parsed.');

    $resolver = new XlsxWorkbookResolver();
    $sheets = v13_private($resolver, 'parseWorkbookSheets', '<w:workbook xmlns:w="urn:test" xmlns:rel="urn:rels"><w:sheets><w:sheet name="Data" sheetId="4" rel:id="rId7"/></w:sheets></w:workbook>');
    smoke_assert_equals('rId7', $sheets[0]['relationship_id'], 'Workbook relationship IDs should use local-name matching.');

    $rich = v13_private(new XlsxMetadataExtractor(), 'richTextFromXml', '<x:si xmlns:x="urn:test"><x:r><x:rPr><x:b/><x:rFont val="Aptos"/><x:color rgb="FFFF0000"/></x:rPr><x:t>Hello</x:t></x:r></x:si>');
    smoke_assert_equals(true, $rich['runs'][0]['bold'], 'Prefixed rich-text formatting should be parsed.');
    smoke_assert_equals('Aptos', $rich['runs'][0]['font'], 'Prefixed rich-text fonts should be parsed.');

    $marker = v13_private(new XlsxImageExtractor(), 'anchorMarker', '<d:from xmlns:d="urn:test"><d:col>1</d:col><d:colOff>10</d:colOff><d:row>2</d:row><d:rowOff>20</d:rowOff></d:from>', 'from');
    smoke_assert_equals('B3', $marker['cell'], 'Prefixed drawing anchors should resolve to cell coordinates.');
});

smoke_run('provides explicit optional formula calculation support', function (): void {
    $evaluator = new PhpSpreadsheetFormulaEvaluator();
    if (!$evaluator->available()) {
        try {
            $evaluator->calculateCell('/missing.xlsx', 1, 'A1');
            throw new SmokeTestFailure('Missing optional calculator should fail clearly.');
        } catch (Throwable $e) {
            smoke_assert_contains('phpoffice/phpspreadsheet', $e->getMessage(), 'Formula calculation error should contain installation guidance.');
        }
    }
});

echo "ExcelFullCapabilitiesV13SmokeTest passed\n";
