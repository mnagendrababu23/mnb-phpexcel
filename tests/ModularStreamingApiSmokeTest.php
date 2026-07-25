<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Contracts\ReaderPluginInterface;
use Mnb\PHPExcel\Format\Csv;
use Mnb\PHPExcel\Format\Json;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\Options\RowErrorPolicy;
use Mnb\PHPExcel\Reader\ReaderRegistry;
use Mnb\PHPExcel\Reader\SharedStrings\InMemorySharedStringProvider;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Reader\XmlSchemaMapping;
use Mnb\PHPExcel\Reader\State\ReadProgress;
use Mnb\PHPExcel\Reader\State\RowState;
use Mnb\PHPExcel\Reader\StreamingJsonArrayParser;
use Mnb\PHPExcel\SpreadsheetManager;

$dir = smoke_temp_dir('modular_streaming');

echo "ModularStreamingApiSmokeTest\n";

smoke_run('streams NDJSON and applies source projection', function () use ($dir): void {
    $path = $dir . '/rows.ndjson';
    file_put_contents($path, "{\"id\":1,\"name\":\"A\",\"drop\":true}\n{\"id\":2,\"name\":\"B\",\"drop\":false}\n");

    $rows = Json::read($path, ReaderOptions::defaults()->withColumns(['name']))
        ->withHeaderRow()
        ->streaming()
        ->toArray();

    smoke_assert_equals([['name' => 'A'], ['name' => 'B']], $rows, 'NDJSON should stream through the shared ReadSession API.');
});

smoke_run('streams top-level JSON arrays without workbook materialization', function () use ($dir): void {
    $path = $dir . '/rows.json';
    file_put_contents($path, '[{"id":"0001","name":"A"},{"id":"0002","name":"B"}]');

    $states = iterator_to_array(
        Json::read($path)
            ->withHeaderRow()
            ->range(1, 3, 1, 2)
            ->rowStates(),
        false
    );

    smoke_assert_equals(2, count($states), 'Two data-row states should be emitted.');
    smoke_assert_true($states[0] instanceof RowState, 'rowStates() should emit typed RowState objects.');
    smoke_assert_equals(1, $states[0]->outputRow, 'Output row numbers should be one-based.');
    smoke_assert_equals('0001', $states[0]->values['id'], 'Numeric-looking identifiers should remain strings.');
});

smoke_run('parses nested and scalar JSON array items across tiny chunks', function () use ($dir): void {
    $path = $dir . '/complex.json';
    $text = str_repeat('x', 5000) . 'a,]\"x';
    file_put_contents($path, json_encode([['text' => $text, 'nested' => [1, ['a' => 2]]], null, true, 12.50], JSON_THROW_ON_ERROR));
    $items = iterator_to_array((new StreamingJsonArrayParser())->parse($path, ['json_chunk_bytes' => 4096]), false);
    smoke_assert_equals(4, count($items), 'Streaming parser should decode every top-level item.');
    smoke_assert_equals($text, $items[0]['text'], 'String delimiters and escapes across stream chunks must not end an item.');
    smoke_assert_equals(2, $items[0]['nested'][1]['a'], 'Nested arrays and objects should be preserved.');
    smoke_assert_equals(true, $items[2], 'Scalar values should be supported.');
});

smoke_run('reports typed progress and row errors', function () use ($dir): void {
    $path = $dir . '/invalid.csv';
    file_put_contents($path, "id,name\n1,A\n2\n3,C\n");
    $progress = [];

    $session = Csv::read($path)
        ->withHeaderRow()
        ->onRowError(RowErrorPolicy::Collect)
        ->onProgress(static function (ReadProgress $state) use (&$progress): void {
            $progress[] = $state;
        }, 1);

    $rows = $session->toArray(['missing_columns' => 'error']);
    smoke_assert_equals([['id' => '1', 'name' => 'A'], ['id' => '3', 'name' => 'C']], $rows, 'Collect policy should skip malformed rows and continue.');
    smoke_assert_equals(1, count($session->rowErrors()), 'One row error should be collected.');
    smoke_assert_true($progress !== [] && end($progress)->completed, 'Final typed progress event should be marked completed.');
    smoke_assert_equals(4, end($progress)->sourceRows, 'Progress should include the physical header row.');
});

smoke_run('maps XML schemas and exposes richer formula/shared-string values', function (): void {
    $mapping = XmlSchemaMapping::from([
        'columns' => [
            'id' => ['source' => '@id', 'type' => 'integer'],
            'name' => 'customer/name',
            'active' => ['source' => 'active', 'type' => 'boolean'],
            'missing' => ['source' => 'unknown', 'default' => 'n/a'],
        ],
    ]);
    smoke_assert_true($mapping instanceof XmlSchemaMapping, 'Schema mapping should be created.');
    $mapped = $mapping->map(['customer' => ['name' => 'Ravi'], 'active' => 'true'], ['id' => '42']);
    smoke_assert_equals(['id' => 42, 'name' => 'Ravi', 'active' => true, 'missing' => 'n/a'], $mapped, 'XML schema mapping should support paths, attributes, types, and defaults.');

    $strings = new InMemorySharedStringProvider(['Alpha', 'Beta']);
    smoke_assert_equals('Beta', $strings->get(1), 'Shared-string provider should resolve indexed strings.');
    smoke_assert_equals('memory', $strings->mode(), 'In-memory provider should report its mode.');

    $formula = new FormulaResult('=SUM(A1:A2)', 3, 'number', ['t' => 'shared', 'si' => '0']);
    smoke_assert_equals(3, $formula->jsonSerialize()['cached_value'], 'FormulaResult should preserve cached results.');
});

smoke_run('supports instance-scoped reader plugins', function () use ($dir): void {
    $path = $dir . '/custom.foo';
    file_put_contents($path, "custom\n");

    $plugin = new class implements ReaderPluginInterface {
        public function supports(string $path, array $options = []): bool
        {
            return str_ends_with(strtolower($path), '.foo');
        }

        public function read(string $path, array $options = []): iterable
        {
            yield ['id' => 7, 'source' => basename($path)];
        }
    };

    $manager = SpreadsheetManager::create(new ReaderRegistry());
    $manager->registerReaderPlugin($plugin);
    $rows = $manager->read($path)->toArray();

    smoke_assert_equals([[7, 'custom.foo']], $rows, 'Plugin rows should use the common read session.');
    smoke_assert_equals([], SpreadsheetManager::create(new ReaderRegistry())->formats(), 'A separate manager should not inherit plugin state.');
});

smoke_run('format-specific CSV facade writes without the legacy facade', function () use ($dir): void {
    $path = $dir . '/facade.csv';
    Csv::write([['id' => '001', 'name' => 'A', 'status' => 'open']], $path, ['with_header' => true]);
    smoke_assert_true(is_file($path), 'Csv::write() should create a file.');
    $rows = Csv::read($path)
        ->withHeaderRow()
        ->projectColumns(['name', 'A'])
        ->toArray();
    smoke_assert_equals([['name' => 'A', 'id' => '001']], $rows, 'Named and positional projection should be resolved after CSV headers are mapped.');
});

echo "ModularStreamingApiSmokeTest passed\n";
