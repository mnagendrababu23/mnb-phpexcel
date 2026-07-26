<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Application\AjaxUploader;
use Mnb\PHPExcel\Application\Http\FileRateLimiter;
use Mnb\PHPExcel\Application\Http\SpreadsheetHttpEndpoint;
use Mnb\PHPExcel\Application\Mail\MailMessage;
use Mnb\PHPExcel\Application\Mail\SmtpMailer;
use Mnb\PHPExcel\Application\Queue\FileQueue;
use Mnb\PHPExcel\Application\Queue\QueueWorker;
use Mnb\PHPExcel\Application\Schedule\CronExpression;
use Mnb\PHPExcel\Application\Schedule\FileScheduler;
use Mnb\PHPExcel\Application\Schedule\SpreadsheetScheduler;
use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Format\Ods;
use Mnb\PHPExcel\Format\Xml;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Reader\Formula\FormulaFunctionRegistry;
use Mnb\PHPExcel\Reader\Formula\NativeFormulaEvaluator;
use Mnb\PHPExcel\Support\EnvironmentDiagnostics;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Writer\XlsxWriter;

function v16_private(object $object, string $method, mixed ...$args): mixed
{
    $reflection = new ReflectionMethod($object, $method);
    $reflection->setAccessible(true);
    return $reflection->invoke($object, ...$args);
}

echo "FullCompatibilityV16SmokeTest\n";

smoke_run('reads and writes XLSX without ext-zip or ext-xmlreader', function (): void {
    $dir = smoke_temp_dir('pure_xlsx_v16');
    $path = $dir . '/pure.xlsx';
    WorkbookBuilder::fromArray([
        ['name' => 'Alice', 'amount' => 10],
        ['name' => 'Bob', 'amount' => 20],
    ])->withHeader()->save($path);
    smoke_assert_true(is_file($path), 'Pure-PHP ZIP writer should create XLSX.');
    $rows = Xlsx::read($path)->sheet(1)->toArray();
    smoke_assert_equals('Alice', $rows[1][0], 'Pure-PHP XML/ZIP reader should read XLSX values.');
    smoke_assert_equals(20, $rows[2][1], 'Numeric cells should round-trip.');
    $zip = new ZipArchive();
    smoke_assert_equals(true, $zip->open($path), 'Fallback ZIP reader should open generated XLSX.');
    smoke_assert_true($zip->locateName('xl/workbook.xml') !== false, 'Workbook part should exist.');
    $zip->close();
    $environment = EnvironmentDiagnostics::check();
    smoke_assert_equals(true, $environment['capabilities']['xlsx_read_ready'], 'XLSX reads should be ready through native or fallback runtime.');
    smoke_assert_equals(true, $environment['capabilities']['xlsx_write_ready'], 'XLSX writes should be ready through native or fallback runtime.');
});

smoke_run('reads XML and ODS through pure-PHP fallbacks', function (): void {
    $dir = smoke_temp_dir('pure_xml_ods_v16');
    $xmlPath = $dir . '/rows.xml';
    Xml::write([
        ['name' => 'Alice', 'amount' => 10],
        ['name' => 'Bob', 'amount' => 20],
    ], $xmlPath);
    $xmlRows = Xml::read($xmlPath)->toArray();
    smoke_assert_equals('Alice', $xmlRows[1][0], 'Pure-PHP XML reader should read generated XML.');
    smoke_assert_equals('20', (string) $xmlRows[2][1], 'Pure-PHP XML reader should preserve XML values.');

    $odsPath = $dir . '/fallback.ods';
    $zip = new ZipArchive();
    smoke_assert_equals(true, $zip->open($odsPath, ZipArchive::CREATE | ZipArchive::OVERWRITE), 'ODS package should open for writing.');
    $zip->addFromString('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    $zip->addFromString('content.xml', '<?xml version="1.0" encoding="UTF-8"?>'
        . '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
        . '<office:body><office:spreadsheet><table:table table:name="Sheet1">'
        . '<table:table-row><table:table-cell office:value-type="string"><text:p>Name</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Amount</text:p></table:table-cell></table:table-row>'
        . '<table:table-row><table:table-cell office:value-type="string"><text:p>Alice</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell></table:table-row>'
        . '</table:table></office:spreadsheet></office:body></office:document-content>');
    $zip->close();
    $odsRows = Ods::read($odsPath)->sheet(1)->toArray();
    smoke_assert_equals('Alice', $odsRows[1][0], 'Pure-PHP ODS reader should read cell text.');
    smoke_assert_equals(10, $odsRows[1][1], 'Pure-PHP ODS reader should read numeric values.');
});

smoke_run('supports expanded formulas and custom function registration', function (): void {
    $cells = ['A1' => 'A', 'B1' => 10, 'A2' => 'B', 'B2' => 20, 'A3' => 'B', 'B3' => 30];
    $registry = (new FormulaFunctionRegistry())->register('DOUBLE', static fn (array $args): float => (float) ($args[0] ?? 0) * 2);
    $engine = new NativeFormulaEvaluator(null, $registry);
    $resolver = static fn (int|string $sheet, string $cell): mixed => $cells[$cell] ?? null;
    smoke_assert_equals(50.0, $engine->evaluate('=SUMIF(A1:A3,"B",B1:B3)', $resolver), 'SUMIF should aggregate matching rows.');
    smoke_assert_equals(20, $engine->evaluate('=VLOOKUP("B",A1:B3,2,FALSE)', $resolver), 'VLOOKUP should support two-dimensional ranges.');
    smoke_assert_equals('A/B/B', $engine->evaluate('=TEXTJOIN("/",TRUE,A1:A3)', $resolver), 'TEXTJOIN should flatten ranges.');
    smoke_assert_equals(42.0, $engine->evaluate('=DOUBLE(21)', $resolver), 'Custom functions should be callable.');
});

smoke_run('generates advanced pivot configuration', function (): void {
    $builder = WorkbookBuilder::fromWorkbookArray([
        'Data' => [
            ['Region' => 'North', 'Month' => 'Jan', 'Amount' => 100],
            ['Region' => 'South', 'Month' => 'Jan', 'Amount' => 150],
        ],
        'Summary' => [],
    ])->withHeader()->addPivotTable('AdvancedPivot', 'Data', 'A1:C3', 'A1', [
        'sheet' => 'Summary',
        'rows' => [['field' => 'Region', 'subtotals' => ['sum'], 'sort' => 'ascending', 'repeat_labels' => true]],
        'columns' => ['Month'],
        'values' => [['field' => 'Amount', 'function' => 'sum', 'name' => 'Share', 'show_data_as' => 'percent_total']],
        'layout' => 'tabular',
        'show_row_stripes' => true,
        'show_empty_rows' => true,
    ]);
    $writer = new XlsxWriter();
    $plan = v16_private($writer, 'buildPivotPlan', $builder->toWorkbookData());
    $xml = v16_private($writer, 'pivotTableXml', $plan[2][0]);
    smoke_assert_contains('sortType="ascending"', $xml, 'Pivot row sorting should be serialized.');
    smoke_assert_contains('sumSubtotal="1"', $xml, 'Pivot subtotals should be serialized.');
    smoke_assert_contains('showDataAs="percentOfTotal"', $xml, 'Show-values-as should be serialized.');
    smoke_assert_contains('showRowStripes="1"', $xml, 'Pivot style options should be serialized.');
});

smoke_run('recovers expired queue reservations and runs scheduler loops', function (): void {
    $dir = smoke_temp_dir('workers_v16');
    $queue = new FileQueue($dir . '/queue');
    $queue->enqueue('noop', []);
    $job = $queue->reserve(300, 'worker-a');
    smoke_assert_true($job !== null, 'Job should reserve.');
    $processing = $dir . '/queue/processing/' . $job->id . '.json';
    $data = json_decode((string) file_get_contents($processing), true, 512, JSON_THROW_ON_ERROR);
    $data['reserved_at'] = time() - 1000;
    file_put_contents($processing, json_encode($data, JSON_THROW_ON_ERROR));
    smoke_assert_equals(1, $queue->releaseExpired(60), 'Expired reservation should return to pending.');
    $worker = (new QueueWorker($queue))->register('noop', static fn (): array => ['ok' => true]);
    smoke_assert_equals(1, $worker->work(['max_jobs' => 1])['completed'], 'Released job should be processed.');

    smoke_assert_equals(true, (new CronExpression('@hourly'))->isDue(new DateTimeImmutable('2026-07-26 10:00:00')), 'Cron macros should work.');
    smoke_assert_equals(true, (new CronExpression('0 10 * JUL SUN'))->isDue(new DateTimeImmutable('2026-07-26 10:00:00')), 'Named cron fields should work.');
    $scheduler = new SpreadsheetScheduler(new FileScheduler($dir . '/schedule.json'));
    $output = $dir . '/scheduled.csv';
    $scheduler->scheduleExport('every-minute', '* * * * *', [['name' => 'Alice']], $output);
    $result = $scheduler->runForever(['max_cycles' => 1, 'interval_seconds' => 1, 'lock_path' => $dir . '/runner.lock']);
    smoke_assert_equals(1, $result['cycles'], 'Built-in scheduler runner should execute bounded cycles.');
    smoke_assert_true(is_file($output), 'Scheduler runner should execute due exports.');
});

smoke_run('provides authenticated rate-limited HTTP and complete AJAX client', function (): void {
    $dir = smoke_temp_dir('http_v16');
    $endpoint = new SpreadsheetHttpEndpoint(options: [
        'base_path' => '/spreadsheet',
        'bearer_tokens' => ['secret-token'],
        'allowed_origins' => ['https://app.example'],
        'rate_limiter' => new FileRateLimiter($dir . '/rate'),
        'rate_limit' => 1,
        'rate_window_seconds' => 60,
    ]);
    $path = $dir . '/api.csv';
    $server = [
        'REQUEST_METHOD' => 'POST',
        'REQUEST_URI' => '/spreadsheet/export',
        'CONTENT_TYPE' => 'application/json',
        'HTTP_AUTHORIZATION' => 'Bearer secret-token',
        'HTTP_ORIGIN' => 'https://app.example',
        'REMOTE_ADDR' => '127.0.0.1',
    ];
    $body = json_encode(['path' => $path, 'rows' => [['name' => 'Alice']]], JSON_THROW_ON_ERROR);
    $response = $endpoint->handle($server, [], [], [], $body);
    smoke_assert_equals(200, $response->status, 'Authenticated API request should succeed.');
    smoke_assert_true(is_file($path), 'HTTP API should export a spreadsheet.');
    smoke_assert_equals(429, $endpoint->handle($server, [], [], [], $body)->status, 'Rate limiter should reject excess requests.');
    $html = AjaxUploader::html('/spreadsheet/upload');
    smoke_assert_contains('XMLHttpRequest', $html, 'AJAX helper should include a complete browser client.');
    smoke_assert_contains('mnb-upload-complete', $html, 'AJAX helper should dispatch completion events.');
});

smoke_run('delivers attachments over native SMTP', function (): void {
    if (!function_exists('pcntl_fork')) {
        echo 'SKIP pcntl unavailable ';
        return;
    }
    $server = stream_socket_server('tcp://127.0.0.1:0', $errno, $error);
    smoke_assert_true(is_resource($server), 'Mock SMTP server should start.');
    $address = stream_socket_get_name($server, false);
    $port = (int) substr((string) strrchr((string) $address, ':'), 1);
    $pid = pcntl_fork();
    smoke_assert_true($pid >= 0, 'SMTP mock should fork.');
    if ($pid === 0) {
        $client = @stream_socket_accept($server, 10);
        if (!is_resource($client)) { exit(2); }
        fwrite($client, "220 localhost ESMTP\r\n");
        $inData = false;
        while (($line = fgets($client)) !== false) {
            $trim = rtrim($line, "\r\n");
            if ($inData) {
                if ($trim === '.') { fwrite($client, "250 queued\r\n"); $inData = false; }
                continue;
            }
            if (str_starts_with($trim, 'EHLO')) fwrite($client, "250-localhost\r\n250 SIZE 10000000\r\n");
            elseif (str_starts_with($trim, 'MAIL FROM')) fwrite($client, "250 ok\r\n");
            elseif (str_starts_with($trim, 'RCPT TO')) fwrite($client, "250 ok\r\n");
            elseif ($trim === 'DATA') { fwrite($client, "354 end with dot\r\n"); $inData = true; }
            elseif ($trim === 'QUIT') { fwrite($client, "221 bye\r\n"); break; }
            else fwrite($client, "250 ok\r\n");
        }
        fclose($client); fclose($server); exit(0);
    }
    fclose($server);
    $dir = smoke_temp_dir('smtp_v16');
    $attachment = $dir . '/report.csv';
    file_put_contents($attachment, "name\nAlice\n");
    $mailer = new SmtpMailer(['host' => '127.0.0.1', 'port' => $port, 'encryption' => 'none', 'from' => 'sender@example.com']);
    $sent = $mailer->send(new MailMessage(['receiver@example.com'], 'Report', 'Attached.', false, [['path' => $attachment]]));
    smoke_assert_equals(true, $sent, 'SMTP mailer should complete a transaction.');
    pcntl_waitpid($pid, $status);
    smoke_assert_equals(0, pcntl_wexitstatus($status), 'Mock SMTP server should exit cleanly.');
});

smoke_run('exposes legacy XLS writing as a first-class API', function (): void {
    smoke_assert_true(method_exists(Mnb\PHPExcel\Format\Xls::class, 'write'), 'XLS facade should expose write().');
    $dir = smoke_temp_dir('xls_v16');
    try {
        WorkbookBuilder::fromArray([['name' => 'Alice']])->withHeader()->save($dir . '/legacy.xls');
        smoke_assert_true(class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet'), 'Successful XLS write requires the XLS compatibility dependency.');
    } catch (Throwable $e) {
        smoke_assert_contains('mnb/mnb-phpexcel-xls', $e->getMessage(), 'Missing dependency error should provide installation guidance.');
    }
});

echo "FullCompatibilityV16SmokeTest passed\n";
