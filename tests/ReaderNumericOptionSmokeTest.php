<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Reader\XlsxStyleMap;

$reader = new XlsxReader();
$method = new ReflectionMethod($reader, 'cellValueFromXml');
$method->setAccessible(true);
$styleMap = XlsxStyleMap::fromXml(null);

$xml = '<c r="A1"><v>12345678901234567890</v></c>';
$value = $method->invoke($reader, $xml, '', null, [], $styleMap, false, ['preserve_numeric_strings' => true]);
assert($value === '12345678901234567890');

$value = $method->invoke($reader, '<c r="A1"><v>1.23456789123</v></c>', '', null, [], $styleMap, false, ['preserve_numeric_strings' => true]);
assert($value === '1.23456789123');

echo "ReaderNumericOptionSmokeTest passed\n";
