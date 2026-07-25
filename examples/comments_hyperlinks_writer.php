<?php

declare(strict_types=1);

require __DIR__ . '/../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;

$path = __DIR__ . '/comments-hyperlinks-report.xlsx';

try {
    MnbExcel::fromArray([
        ['Resource' => 'Repository', 'Link' => 'Open GitHub', 'Status' => 'Review'],
        ['Resource' => 'Documentation', 'Link' => 'Open Docs', 'Status' => 'Ready'],
    ])->withHeader()
        ->hyperlink('B2', 'https://github.com/mnagendrababu23/mnb-phpexcel', 'Open repository', ['tooltip' => 'Project source code'])
        ->comment('C2', 'MNB PHPExcel', 'Check this row before sharing the report.')
        ->save($path);

    print_r(MnbExcel::validateXlsx($path));
} catch (MnbExcelException $e) {
    print_r(MnbExcel::safeError($e));
}
