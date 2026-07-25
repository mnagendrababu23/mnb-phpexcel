<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Reader\CsvReader;

smoke_run('CSV requested boundaries stop cleanly', static function (): void {
    $dir = smoke_temp_dir('csv_boundaries');
    $path = $dir . '/rows.csv';
    file_put_contents($path, "a,b\n1,2\n3,4\n5,6\n");

    $reader = new CsvReader();
    $endRows = $reader->readSheet($path, 1, ['end_row' => 2]);
    smoke_assert_equals([['a', 'b'], ['1', '2']], $endRows, 'end_row should return the requested physical rows');

    $limited = $reader->readSheet($path, 1, ['source_limit_rows' => 2]);
    smoke_assert_equals([['a', 'b'], ['1', '2']], $limited, 'source_limit_rows should return the requested number of rows');

    $empty = $reader->readSheet($path, 1, ['source_limit_rows' => 0]);
    smoke_assert_equals([], $empty, 'source_limit_rows=0 should return no rows');
});

smoke_run('worksheet names are Unicode-safe and unique', static function (): void {
    $longUnicode = str_repeat('é', 40);
    $builder = WorkbookBuilder::fromWorkbookArray([
        'Same/Name' => [[1]],
        'Same\\Name' => [[2]],
        $longUnicode => [[3]],
    ]);

    $workbook = $builder->toWorkbookData();
    $names = array_map(static fn($sheet): string => $sheet->name, $workbook->sheets);

    smoke_assert_equals(count($names), count(array_unique(array_map('strtolower', $names))), 'sanitized worksheet names should be case-insensitively unique');
    smoke_assert_true($names[0] !== $names[1], 'colliding source names should receive unique suffixes');
    smoke_assert_true(preg_match('//u', $names[2]) === 1, 'truncated Unicode worksheet name should remain valid UTF-8');

    preg_match_all('/./u', $names[2], $characters);
    smoke_assert_true(count($characters[0]) <= 31, 'worksheet name should contain no more than 31 Unicode characters');

    $safe = MnbExcel::safeSheetName($longUnicode);
    smoke_assert_true(preg_match('//u', $safe) === 1, 'safeSheetName should preserve valid UTF-8');

    $caseFolded = WorkbookBuilder::fromWorkbookArray([
        'École' => [[1]],
        'école' => [[2]],
    ])->toWorkbookData();
    smoke_assert_true($caseFolded->sheets[0]->name !== $caseFolded->sheets[1]->name, 'Unicode case variants should not create duplicate worksheet names');
});

smoke_run('images remain scoped to their selected worksheet', static function (): void {
    $dir = smoke_temp_dir('image_scope');
    $imagePath = $dir . '/marker.png';
    file_put_contents($imagePath, 'not-rendered-in-this-test');

    $workbook = WorkbookBuilder::fromWorkbookArray([
        'First' => [[1]],
        'Second' => [[2]],
    ])
        ->addImage($imagePath, 'A1', ['sheet' => 'Second'])
        ->toWorkbookData();

    smoke_assert_equals(0, count($workbook->sheets[0]->images), 'first worksheet should not inherit an image assigned to another sheet');
    smoke_assert_equals(1, count($workbook->sheets[1]->images), 'selected worksheet should contain its assigned image');
});

smoke_run('headerless structured columns report generated metadata', static function (): void {
    $dir = smoke_temp_dir('generated_columns');
    $path = $dir . '/rows.csv';
    file_put_contents($path, "one,two\nthree,four\n");

    $structured = MnbExcel::readCsv($path)->toStructuredSheetArray([
        'header_row' => false,
        'include_cell_metadata' => false,
    ]);

    smoke_assert_true($structured['columns'][0]['generated'] === true, 'column_1 should be marked generated without a header row');
    smoke_assert_true($structured['columns'][1]['generated'] === true, 'column_2 should be marked generated without a header row');
});
