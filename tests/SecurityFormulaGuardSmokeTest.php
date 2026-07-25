<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Security\FormulaGuard;

echo "SecurityFormulaGuardSmokeTest\n";

smoke_run('flags known-dangerous formula functions as risky', function (): void {
    $dangerous = ['WEBSERVICE("http://evil.example/x")', 'IMPORTXML("http://evil.example","//a")', 'DDE("cmd";"/c calc";"")'];

    foreach ($dangerous as $formula) {
        $risk = FormulaGuard::risk($formula);
        smoke_assert_equals('high', $risk['risk'], 'formula should be flagged high risk: ' . $formula);
    }
});

smoke_run('allows an ordinary safe formula', function (): void {
    $risk = FormulaGuard::risk('SUM(A1:A10)');
    smoke_assert_equals('none', $risk['risk'], 'a plain SUM formula should not be flagged');
});

smoke_run('assertSafe() throws for a dangerous formula under the default "safe" policy', function (): void {
    $threw = false;
    try {
        FormulaGuard::assertSafe('HYPERLINK("http://evil.example","click me")');
    } catch (\Throwable $e) {
        $threw = true;
    }
    smoke_assert_true($threw, 'assertSafe() should throw for a dangerous formula');
});

smoke_run('writing a row with a blocked formula fails safely instead of producing a file', function (): void {
    $dir = smoke_temp_dir('formula_guard');
    $path = $dir . DIRECTORY_SEPARATOR . 'unsafe.xlsx';

    $rows = [
        ['label' => 'bad', 'calc' => MnbExcel::formula('HYPERLINK("http://evil.example","click me")')],
    ];

    $threw = false;
    try {
        MnbExcel::fromArray($rows)->save($path);
    } catch (\Throwable $e) {
        $threw = true;
    }

    smoke_assert_true($threw, 'saving a workbook containing a dangerous formula should throw under the default policy');
    smoke_assert_true(!is_file($path), 'no partial/corrupt file should be left behind after a blocked save');
});

echo "SecurityFormulaGuardSmokeTest: all assertions passed.\n";
