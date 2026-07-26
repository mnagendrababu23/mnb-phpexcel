# MNB PHPExcel v1.5 Full Workflow APIs

v1.5 promotes the remaining application integrations to first-class, framework-neutral library services.

## Multiple-file imports

```php
$result = MnbExcel::importDomainFiles('products', $files, $pdo, 'products', [
    'stop_on_error' => false,
    'import_options' => [
        'duplicate_strategy' => 'update',
        'unique_by' => ['sku'],
    ],
    'progress' => function (array $state): void {
        echo $state['files_completed'] . '/' . $state['files_total'];
    },
]);
```

## Built-in queue and worker

```php
$queue = MnbExcel::queue(__DIR__ . '/storage/queue');
$queue->enqueueDomainImport('products', 'products.xlsx', $dbConfig, 'products');
$queue->enqueueExport($rows, __DIR__ . '/reports/report.xlsx');

$result = $queue->work([
    'max_jobs' => 50,
    'max_attempts' => 3,
    'retry_delay_seconds' => 30,
]);
```

Jobs are stored atomically in `pending`, `processing`, `completed`, and `failed` directories. The worker can be launched from CLI, cron, a process manager, or wrapped by a framework queue command.

## AJAX uploads

```php
$response = MnbExcel::handleAjaxUpload($_FILES['spreadsheet'], [
    'directory' => __DIR__ . '/storage/uploads',
    'allowed_extensions' => ['xlsx', 'xls', 'ods', 'csv', 'tsv'],
    'max_size_mb' => 250,
]);
```

The response includes an HTTP-oriented status, safe stored filename, SHA-256 digest, size, warnings, and validation errors.

## API dispatcher

```php
$response = MnbExcel::api('import', [
    'domain' => 'users',
    'path' => $storedPath,
    'table' => 'users',
    'options' => ['duplicate_strategy' => 'update'],
], $pdo);
```

Supported actions are `upload`, `preview`, `import`, `import-many`, `status`, and `export`.

## Email generated spreadsheets

```php
MnbExcel::emailGeneratedExcel(
    MnbExcel::report($rows),
    'manager@example.com',
    'Daily report',
    'The report is attached.',
    ['filename' => 'daily-report.xlsx']
);
```

The default transport uses PHP `mail()` with a correct multipart MIME attachment. A callback or `MailerInterface` implementation can be supplied for SMTP, SES, Postmark, Laravel Mail, Symfony Mailer, or another provider.

## Scheduled imports and exports

```php
$scheduler = MnbExcel::scheduler(__DIR__ . '/storage/schedule.json');
$scheduler->scheduleImport('erp-sync', '*/15 * * * *', 'erp.xlsx', $dbConfig, 'products');
$scheduler->scheduleExport('daily-report', '0 6 * * *', $rows, 'daily.xlsx');

// Invoke once per minute from cron or a service manager.
$scheduler->runDue();
```

The scheduler supports standard five-field minute, hour, day, month, and weekday expressions with lists, ranges, and step values.

## Native formula evaluation

```php
$value = Xlsx::read('finance.xlsx')->sheet('Summary')->calculatedCell('D12');
```

The built-in evaluator handles arithmetic, comparisons, ranges, `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, logical functions, rounding, text functions, and common date functions. When PhpSpreadsheet is installed, the evaluator automatically uses it for broader Excel-function compatibility.

## Native pivot tables

```php
MnbExcel::fromWorkbookArray([
    'Data' => $rows,
    'Summary' => [],
])
    ->withHeader()
    ->addPivotTable('SalesPivot', 'Data', 'A1:F5000', 'A1', [
        'sheet' => 'Summary',
        'rows' => ['Region'],
        'columns' => ['Month'],
        'filters' => ['Status'],
        'values' => [
            ['field' => 'Amount', 'function' => 'sum', 'name' => 'Total Sales'],
        ],
    ])
    ->save('sales-pivot.xlsx');
```

Native pivots are generated with refresh-on-open cache definitions. Existing template preservation remains the best route for slicers, OLAP/data-model pivots, and deeply customized layouts.
