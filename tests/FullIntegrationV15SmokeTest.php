<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Application\AjaxUploadHandler;
use Mnb\PHPExcel\Application\Mail\CallbackMailer;
use Mnb\PHPExcel\Application\Mail\SpreadsheetMailer;
use Mnb\PHPExcel\Application\MultiFileImportManager;
use Mnb\PHPExcel\Application\Queue\FileQueue;
use Mnb\PHPExcel\Application\Queue\QueueWorker;
use Mnb\PHPExcel\Application\Queue\SpreadsheetQueue;
use Mnb\PHPExcel\Application\Schedule\CronExpression;
use Mnb\PHPExcel\Application\Schedule\FileScheduler;
use Mnb\PHPExcel\Application\Schedule\SpreadsheetScheduler;
use Mnb\PHPExcel\Application\SpreadsheetApi;
use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Reader\Formula\FormulaEvaluatorFactory;
use Mnb\PHPExcel\Reader\Formula\NativeFormulaEvaluator;
use Mnb\PHPExcel\Writer\XlsxWriter;

function v15_private(object $object, string $method, mixed ...$args): mixed
{
    $reflection = new ReflectionMethod($object, $method);
    $reflection->setAccessible(true);
    return $reflection->invoke($object, ...$args);
}

echo "FullIntegrationV15SmokeTest\n";

smoke_run('evaluates common formulas without external dependencies', function (): void {
    $engine = new NativeFormulaEvaluator();
    $cells = ['A1' => 10, 'A2' => 20, 'B1' => 5, 'C1' => 'North'];
    $resolver = static fn (int|string $sheet, string $cell): mixed => $cells[$cell] ?? 0;
    smoke_assert_equals(40.0, $engine->evaluate('=SUM(A1:A2)+B1*2', $resolver), 'Native formula sums and arithmetic should work.');
    smoke_assert_equals('NORTH-30', $engine->evaluate('=UPPER(C1)&"-"&SUM(A1:A2)', $resolver), 'Text and range functions should work.');
    smoke_assert_equals('large', $engine->evaluate('=IF(A2>=20,"large","small")', $resolver), 'Conditional formulas should work.');
    smoke_assert_true(FormulaEvaluatorFactory::create(true) instanceof NativeFormulaEvaluator, 'Formula factory should expose a dependency-free native engine.');
});

smoke_run('builds pivot tables from scratch', function (): void {
    $builder = WorkbookBuilder::fromWorkbookArray([
        'Data' => [
            ['Region' => 'North', 'Product' => 'A', 'Amount' => 100],
            ['Region' => 'South', 'Product' => 'A', 'Amount' => 150],
        ],
        'Summary' => [],
    ])->withHeader()->addPivotTable('SalesPivot', 'Data', 'A1:C3', 'A1', [
        'sheet' => 'Summary',
        'rows' => ['Region'],
        'columns' => ['Product'],
        'values' => [['field' => 'Amount', 'function' => 'sum', 'name' => 'Total Sales']],
    ]);
    $workbook = $builder->toWorkbookData();
    smoke_assert_equals(1, count($workbook->sheets[1]->pivotTables), 'Target worksheet should contain a pivot definition.');
    $writer = new XlsxWriter();
    $plan = v15_private($writer, 'buildPivotPlan', $workbook);
    smoke_assert_equals(1, count($plan[2]), 'Pivot plan should target the second worksheet.');
    $cache = v15_private($writer, 'pivotCacheDefinitionXml', $plan[2][0]);
    $pivot = v15_private($writer, 'pivotTableXml', $plan[2][0]);
    smoke_assert_contains('<worksheetSource ref="A1:C3" sheet="Data"/>', $cache, 'Pivot cache should point at source data.');
    smoke_assert_contains('<dataField name="Total Sales" fld="2" subtotal="sum"/>', $pivot, 'Pivot table should include the configured aggregate.');
    $sheetXml = v15_private($writer, 'worksheetXml', $workbook->sheets[1], false, $plan[2], null);
    smoke_assert_contains('<pivotTableParts count="1">', $sheetXml, 'Worksheet should link generated pivot parts.');
});

smoke_run('coordinates multiple files with aggregate progress and errors', function (): void {
    $manager = new MultiFileImportManager();
    $progressCalls = 0;
    $result = $manager->run(['one.csv', 'two.csv', 'bad.csv'], static function (string|array $file): array {
        if ($file === 'bad.csv') {
            throw new RuntimeException('bad input');
        }
        return ['rows_scanned' => 2, 'rows_inserted' => 2];
    }, ['validate_uploads' => false, 'progress' => static function () use (&$progressCalls): void { $progressCalls++; }]);
    smoke_assert_equals('completed_with_errors', $result['status'], 'Batch import should retain successful files when one fails.');
    smoke_assert_equals(4, $result['summary']['rows_inserted'], 'Batch totals should aggregate row counts.');
    smoke_assert_equals(3, $progressCalls, 'Progress should fire after every file.');
});

smoke_run('runs durable filesystem queue jobs', function (): void {
    $dir = smoke_temp_dir('queue_v15');
    $queue = new FileQueue($dir);
    $queue->enqueue('double', ['value' => 21]);
    $worker = (new QueueWorker($queue))->register('double', static fn (array $payload): array => ['value' => $payload['value'] * 2]);
    $result = $worker->work(['max_jobs' => 1]);
    smoke_assert_equals(1, $result['completed'], 'Queue worker should complete jobs.');
    smoke_assert_equals(1, $queue->stats()['completed'], 'Completed job should be persisted.');

    $csv = $dir . '/queued.csv';
    $spreadsheetQueue = new SpreadsheetQueue($queue);
    $spreadsheetQueue->enqueueExport([['name' => 'A'], ['name' => 'B']], $csv);
    $spreadsheetQueue->work(['max_jobs' => 1]);
    smoke_assert_true(is_file($csv), 'Built-in export jobs should generate files.');
});

smoke_run('handles AJAX uploads and API exports', function (): void {
    $dir = smoke_temp_dir('api_v15');
    $source = $dir . '/incoming.csv';
    file_put_contents($source, "name\nAlice\n");
    $stored = (new AjaxUploadHandler())->handle(['tmp_name' => $source, 'name' => 'users.csv', 'size' => filesize($source), 'error' => UPLOAD_ERR_OK], [
        'directory' => $dir . '/uploads',
        'strict_mime' => false,
    ]);
    smoke_assert_equals(true, $stored['ok'], 'AJAX upload should validate and store a file.');
    smoke_assert_true(is_file($stored['file']['path']), 'Stored upload should exist.');

    $output = $dir . '/api.csv';
    $response = (new SpreadsheetApi())->handle('export', ['path' => $output, 'rows' => [['name' => 'Alice']]]);
    smoke_assert_equals(true, $response['ok'], 'API export should return a successful response.');
    smoke_assert_true(is_file($output), 'API export should create the requested file.');
});

smoke_run('emails generated spreadsheets through a first-class mailer', function (): void {
    $dir = smoke_temp_dir('mail_v15');
    $path = $dir . '/report.csv';
    WorkbookBuilder::fromArray([['name' => 'Alice']])->withHeader()->save($path);
    $captured = null;
    $mailer = new CallbackMailer(static function ($message) use (&$captured): bool { $captured = $message; return true; });
    $sent = (new SpreadsheetMailer($mailer))->send($path, 'report@example.com', 'Daily report', 'Attached.');
    smoke_assert_equals(true, $sent, 'Spreadsheet mailer should delegate delivery.');
    smoke_assert_equals('report.csv', basename($captured->attachments[0]['path']), 'Generated spreadsheet should be attached.');

    $sentFacade = MnbExcel::emailGeneratedExcel($path, 'report@example.com', 'Facade report', '', [
        'mailer' => static fn (): bool => true,
    ]);
    smoke_assert_equals(true, $sentFacade, 'Facade should support callback mail transports.');
});

smoke_run('schedules imports and exports with cron expressions', function (): void {
    $dir = smoke_temp_dir('schedule_v15');
    $now = new DateTimeImmutable('2026-07-26 10:15:00');
    smoke_assert_equals(true, (new CronExpression('15 10 * * *'))->isDue($now), 'Cron matcher should detect due tasks.');
    $store = $dir . '/schedule.json';
    $scheduler = new FileScheduler($store);
    $scheduler->add('test', '15 10 * * *', 'custom', ['value' => 7]);
    $run = $scheduler->runDue(static fn (string $type, array $payload): array => ['value' => $payload['value'] * 2], $now);
    smoke_assert_equals(1, $run['ran'], 'Due scheduled task should run once.');
    smoke_assert_equals(0, $scheduler->runDue(static fn (): array => [], $now)['ran'], 'Same task should not run twice in one minute.');

    $csv = $dir . '/scheduled.csv';
    $spreadsheet = new SpreadsheetScheduler(new FileScheduler($dir . '/spreadsheet-schedule.json'));
    $spreadsheet->scheduleExport('daily-export', '15 10 * * *', [['name' => 'Alice']], $csv);
    $spreadsheet->runDue($now);
    smoke_assert_true(is_file($csv), 'Scheduled exports should generate files.');
});

echo "FullIntegrationV15SmokeTest passed\n";
