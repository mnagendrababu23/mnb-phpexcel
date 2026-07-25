<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Writer\XlsxWriter;

$writer = new XlsxWriter();
$method = new ReflectionMethod($writer, 'number');
$method->setAccessible(true);

assert($method->invoke($writer, 123.123456789) === '123.123456789');
assert($method->invoke($writer, '123.123456789123') === '123.123456789123');
assert($method->invoke($writer, 100.0) === '100');
assert($method->invoke($writer, -0.0) === '0');
assert($method->invoke($writer, '1.230000') === '1.23');

$cell = CellValue::number('123.123456789123');
assert($cell->value() === '123.123456789123');

$formula = CellValue::formula('A1/3', '0.333333333333');
assert($formula->cachedValue() === '0.333333333333');

echo "NumericPrecisionSmokeTest passed\n";
