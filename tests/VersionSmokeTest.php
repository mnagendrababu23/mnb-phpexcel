<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

smoke_assert_equals('1.4.0', MnbExcel::VERSION, 'VERSION constant should match this release');
smoke_assert_equals('1.4.0', MnbExcel::version(), 'version() should match this release');

echo "VersionSmokeTest passed\n";
