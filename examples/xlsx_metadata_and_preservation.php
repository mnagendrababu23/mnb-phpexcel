<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

require __DIR__ . '/../vendor/autoload.php';

$source = __DIR__ . '/source-template.xlsx';
$output = __DIR__ . '/updated-with-preserved-objects.xlsx';

// Read non-tabular XLSX metadata separately from plain row values.
$metadata = MnbExcel::read($source)->sheet(1)->sheetMetadata();

print_r([
    'rich_text_cells' => $metadata['summary']['rich_text_cells'] ?? 0,
    'comments' => $metadata['summary']['comments'] ?? 0,
    'hyperlinks' => $metadata['summary']['hyperlinks'] ?? 0,
    'advanced_object_parts' => $metadata['summary']['advanced_object_parts'] ?? 0,
]);

// Rewrite rows while preserving supported advanced package objects from the source template.
$newRows = [
    ['name' => 'Ravi', 'email' => 'ravi@example.com', 'website' => 'https://example.com/ravi'],
    ['name' => 'Sita', 'email' => 'sita@example.com', 'website' => 'https://example.com/sita'],
];

MnbExcel::fromArray($newRows)
    ->withHeader()
    ->preserveAdvancedObjectsFrom($source)
    ->save($output);

echo "Saved: {$output}\n";
