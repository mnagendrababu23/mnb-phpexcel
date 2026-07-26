# MNB PHPExcel v1.6 — Full Runtime Compatibility

v1.6 closes the runtime-dependent and scoped workflow gaps identified in the v1.5 capability audit while preserving the modular package model.

## Portable package and XML runtime

The Core module now provides `Support\Zip\ZipArchive` and `Support\Xml\XmlReader` adapters.

- When native `ext-zip` or `ext-xmlreader` is installed, the adapters delegate automatically to the native implementations.
- Without those extensions, secure pure-PHP readers and writers provide functional XLSX, XML, and ODS operation.
- The XML fallback rejects DTD/entity input.
- Native XMLReader remains the preferred huge-file path because it is genuinely forward-only and bounded-memory; the fallback prioritizes deployment portability.

```php
use Mnb\PHPExcel\Format\Xlsx;

$rows = Xlsx::read('orders.xlsx')->sheet('Orders')->toArray();
Xlsx::write($rows, 'copy.xlsx');
```

No conditional extension check is required in application code.

## Legacy XLS read and write

The dedicated `mnb/mnb-phpexcel-xls` package declares its BIFF8 compatibility dependency and now exposes both operations:

```php
use Mnb\PHPExcel\Format\Xls;

$session = Xls::read('legacy.xls');
Xls::write($session->sheet(1)->toArray(), 'export.xls');
```

The complete `mnb/mnb-phpexcel-all` install includes this package automatically. Lightweight XLSX-only users do not download the legacy engine.

## Expanded formula calculation

The native evaluator now supports arithmetic, comparisons, ranges, conditional aggregates, lookup/index functions, text functions, dates, times, rounding, logarithms, trigonometry, information functions, and custom functions.

```php
use Mnb\PHPExcel\Reader\Formula\FormulaFunctionRegistry;
use Mnb\PHPExcel\Reader\Formula\NativeFormulaEvaluator;

$functions = (new FormulaFunctionRegistry())
    ->register('MARGIN', static fn (array $args): float =>
        ((float) $args[0] - (float) $args[1]) / (float) $args[0]
    );

$engine = new NativeFormulaEvaluator(null, $functions);
```

The XLS/all install also supplies the compatibility evaluator for specialist Excel functions. Vendor-specific functions can be registered without process-global state.

## Advanced standard pivots

`WorkbookBuilder::addPivotTable()` supports:

- compact, outline, and tabular layouts;
- row, column, page/filter, and data fields;
- subtotals and sorting;
- repeated labels and blank-row controls;
- `show_data_as` calculations and base fields/items;
- field captions and item visibility;
- style stripes, headers, formatting preservation, drill-down, field-list, empty-row/column, and cache controls.

OLAP/data-model pivots and slicers are preserved from trusted templates because those are external data-model features rather than ordinary worksheet pivots.

## Production queues

Queue consumers program against `QueueBackendInterface`.

```php
$queue = MnbExcel::pdoQueue($pdo, ['table' => 'spreadsheet_jobs']);
$queue->enqueueDomainImport('products', 'products.xlsx', $db, 'products');

$queue->work([
    'worker_id' => gethostname() . '-' . getmypid(),
    'visibility_timeout_seconds' => 900,
    'max_attempts' => 5,
]);
```

The PDO backend provides transactional claims, visibility timeouts, delayed retries, reservation recovery, multi-host workers, and persistent results. The original filesystem backend remains useful for one-host deployments.

## HTTP and AJAX

`SpreadsheetHttpEndpoint` is a complete backend endpoint helper rather than a loose dispatcher. It includes:

- route/action dispatch;
- JSON, form, query, and upload parsing;
- method enforcement;
- bearer token and HMAC authentication;
- CSRF validation;
- CORS and preflight handling;
- file or PDO rate limiting;
- consistent JSON responses and errors.

`AjaxUploader::html()` provides a dependency-free browser form and XMLHttpRequest client with upload progress and completion/error events.

Framework routing, user identity, and business authorization remain host-application policy, but the spreadsheet transport workflow itself is implemented by the library.

## Mail delivery

`SmtpMailer` supports SMTP, implicit TLS, STARTTLS, AUTH LOGIN/PLAIN, multiple recipients, multipart MIME, HTML/text bodies, and spreadsheet attachments.

```php
MnbExcel::emailGeneratedExcel($builder, 'reports@example.com', 'Daily report', 'Attached.', [
    'filename' => 'daily-report.xlsx',
    'smtp' => [
        'host' => 'smtp.example.com',
        'port' => 587,
        'encryption' => 'starttls',
        'username' => getenv('SMTP_USER'),
        'password' => getenv('SMTP_PASSWORD'),
        'from' => 'noreply@example.com',
    ],
]);
```

Callback/provider adapters and PHP `mail()` remain available.

## Scheduling

Schedules can use `FileScheduler` or transactional `PdoScheduler`. Cron syntax supports five-field expressions, lists, ranges, steps, named months/weekdays, and common macros.

```php
$scheduler = MnbExcel::pdoScheduler($pdo);
$scheduler->scheduleExport('daily', '@daily', $rows, '/reports/daily.xlsx');

MnbExcel::runScheduler($scheduler, [
    'interval_seconds' => 30,
    'lock_path' => '/var/run/mnb-excel-scheduler.lock',
]);
```

The runner includes process locking, bounded test cycles, retries through dispatched jobs, and graceful SIGTERM/SIGINT/SIGHUP handling. As with any scheduler, an operating-system service must start or supervise the process.

## Package release impact

v1.6 source changes affect these split repositories:

1. `mnb-phpexcel-core`
2. `mnb-phpexcel-xml`
3. `mnb-phpexcel-xlsx`
4. `mnb-phpexcel-ods`
5. `mnb-phpexcel-xls`
6. `mnb-phpexcel-application`
7. `mnb-phpexcel-all`

CSV, JSON, and database packages can keep their existing releases because their source files did not change and their constraints remain compatible.
