# Stable reading API — 1.2

## One session for normal and streaming reads

```php
use Mnb\PHPExcel\SpreadsheetManager;

$session = SpreadsheetManager::create()->read('orders.xlsx');

$small = $session->normal()->toArray();

foreach ($session->streaming()->rows() as $row) {
    // Forward-only processing.
}

foreach ($session->streaming()->chunks(1000) as $chunk) {
    // Same session vocabulary for large files.
}
```

`normal()` and `streaming()` are policies, not separate application APIs. Legacy `largeRead()` remains available for compatibility and specialized checkpoint/budget behavior.

## Typed options

```php
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\Options\ReadMode;
use Mnb\PHPExcel\Reader\Options\RowErrorPolicy;

$options = ReaderOptions::defaults()
    ->withRange(2, 50000, 'A', 'H')
    ->withColumns(['order_id', 'amount'])
    ->withMode(ReadMode::Streaming)
    ->withRowErrorPolicy(RowErrorPolicy::Collect)
    ->withProgress($progressCallback, everyRows: 1000);
```

Legacy arrays remain supported. Stable keys are normalized to old aliases internally.

## Unified ranges and source projection

```php
$session
    ->range(startRow: 2, endRow: 1000, startColumn: 'B', endColumn: 'F')
    ->projectColumns(['customer_id', 'total']);
```

Rules:

- Rows and positional columns are one-based.
- Uppercase `A` through `XFD` are positional column selectors.
- Other strings are associative source keys.
- Projection is performed inside each source reader before header mapping and general row normalization.
- `compact_selected_columns` defaults to `true` for positional rows.

## Row state

```php
foreach ($session->rowStates() as $state) {
    echo $state->sourceRow;
    echo $state->outputRow;
    var_dump($state->values, $state->errors, $state->skipped);
}
```

`RowState` is JSON serializable and contains one-based source/output row numbers, sheet, values, errors, and skipped status.

## Progress

```php
$session->onProgress(function (ReadProgress $progress): void {
    printf(
        "%d source / %d output / %d errors\n",
        $progress->sourceRows,
        $progress->outputRows,
        $progress->errorRows,
    );
}, everyRows: 1000);
```

A final callback is always emitted with `completed === true`, including when iteration stops early or throws.

## Row error policies

- `throw`: stop immediately; default.
- `skip`: skip invalid rows without retaining their details in memory.
- `collect`: skip and retain details in `rowErrors()`.
- `callback`: invoke a handler; returning an array replaces the invalid row, returning `null` skips it; callback failures are retained for inspection.

```php
$session = $session->onRowError('callback', function (Throwable $error, RowState $state): ?array {
    if (str_contains($error->getMessage(), 'missing')) {
        return $state->values + ['missing_value' => null];
    }
    return null;
});
```

## Reader plugins

```php
$manager = SpreadsheetManager::create()
    ->registerReaderPlugin($myReaderPlugin, priority: 100);

$rows = $manager->read('data.custom')->toArray();
```

The registry belongs to the manager instance, so plugins do not leak between long-running worker requests or tenants. The legacy static facade remains available but is not recommended for dynamically changing per-request registrations.
