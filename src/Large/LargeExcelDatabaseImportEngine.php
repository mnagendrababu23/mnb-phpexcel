<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Application\LoggerBridge;
use Mnb\PHPExcel\Application\RowTransformerPipeline;
use Mnb\PHPExcel\Events\EventDispatcher;
use Mnb\PHPExcel\Import\SqlImporter;
use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Validation\ArrayValidator;
use PDO;
use Throwable;

/**
 * Streams large XLSX rows into a database using chunk validation, PDO batch inserts,
 * failed-row CSV export, and resumable JSON progress manifests.
 */
final class LargeExcelDatabaseImportEngine
{
    private LargeXlsxStreamingReader $reader;
    private ArrayValidator $validator;
    private SqlImporter $sqlImporter;
    private XlsxWorkbookResolver $resolver;

    public function __construct(
        ?LargeXlsxStreamingReader $reader = null,
        ?ArrayValidator $validator = null,
        ?SqlImporter $sqlImporter = null
    ) {
        $this->reader = $reader ?? new LargeXlsxStreamingReader();
        $this->validator = $validator ?? new ArrayValidator();
        $this->sqlImporter = $sqlImporter ?? new SqlImporter();
        $this->resolver = new XlsxWorkbookResolver();
    }

    /**
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function importToSql(string $path, PDO $pdo, string $table, array $options = []): array
    {
        $sheet = $options['sheet'] ?? 1;
        $withHeader = (bool) ($options['with_header'] ?? true);
        $resume = (bool) ($options['resume'] ?? true);
        $dryRun = (bool) ($options['dry_run'] ?? false);
        $chunkSize = max(1, (int) ($options['chunk_size'] ?? 1000));
        $batchSize = max(1, (int) ($options['batch_size'] ?? min(500, $chunkSize)));
        $rules = is_array($options['rules'] ?? null) ? $options['rules'] : (is_array($options['validation_rules'] ?? null) ? $options['validation_rules'] : []);
        $validationOptions = is_array($options['validation_options'] ?? null) ? $options['validation_options'] : [];
        $map = is_array($options['map'] ?? null) ? $options['map'] : (is_array($options['column_map'] ?? null) ? $options['column_map'] : []);
        $strictValidation = (bool) ($options['strict_validation'] ?? false);
        $transactionPerChunk = (bool) ($options['transaction_per_chunk'] ?? true);
        $progress = $options['progress'] ?? null;
        $idempotent = (bool) ($options['idempotent'] ?? false);
        $duplicateStrategy = strtolower((string) ($options['duplicate_strategy'] ?? ($idempotent ? 'skip' : 'fail')));
        $uniqueBy = $this->stringList($options['unique_by'] ?? []);
        $failedRowsFormat = strtolower((string) ($options['failed_rows_format'] ?? 'human'));
        $transformers = is_array($options['transformers'] ?? null) ? $options['transformers'] : [];

        $manifestPath = (string) ($options['manifest_path'] ?? LargeImportManifest::defaultPath($path, $table, $sheet));
        $manifest = new LargeImportManifest($manifestPath);
        $existing = $resume ? $manifest->load() : [];

        $failedRowsCsv = (string) ($options['failed_rows_csv'] ?? $this->defaultFailedRowsPath($path, $manifestPath));

        $headers = [];
        $startAfterRowNumber = null;
        if ($existing !== [] && $resume) {
            $headers = is_array($existing['headers'] ?? null) ? array_values($existing['headers']) : [];
            $startAfterRowNumber = max(0, (int) ($existing['last_row_number'] ?? 0));
        }

        // Fresh imports should not pre-read the header in a separate pass. The streaming reader
        // consumes the header during the real import pass, which keeps row counters accurate and
        // avoids Windows/XAMPP edge cases where a generated inline-string header could be missed.
        // If an older/incomplete manifest is resumed without saved headers, recover them once.
        if ($withHeader && $headers === [] && $resume && $existing !== [] && (int) ($startAfterRowNumber ?? 0) > 0) {
            $headers = $this->readHeader($path, $sheet, $options);
        }

        $failedWriter = new LargeFailedRowsCsvWriter($failedRowsCsv, !$resume || $existing === [], [
            'format' => $failedRowsFormat,
            'columns' => $headers,
        ]);

        if ($existing === [] || !$resume) {
            $manifest->start([
                'source_path' => $path,
                'table' => $table,
                'sheet' => $sheet,
                'status' => 'running',
                'headers' => $headers,
                'with_header' => $withHeader,
                'chunk_size' => $chunkSize,
                'batch_size' => $batchSize,
                'dry_run' => $dryRun,
                'failed_rows_csv' => $failedRowsCsv,
                'failed_rows_format' => $failedRowsFormat,
                'idempotent' => $idempotent,
                'duplicate_strategy' => $duplicateStrategy,
                'unique_by' => $uniqueBy,
                'last_row_number' => (int) ($startAfterRowNumber ?? 0),
            ]);
        } else {
            $manifest->update([
                'status' => 'running',
                'runs' => ((int) ($existing['runs'] ?? 0)) + 1,
                'last_stop_reason' => null,
            ]);
        }

        $totals = $manifest->data();
        $rowsScanned = (int) ($totals['rows_scanned'] ?? 0);
        $resumeBaseRowsScanned = $rowsScanned;
        $validRows = (int) ($totals['valid_rows'] ?? 0);
        $failedRows = (int) ($totals['failed_rows'] ?? 0);
        $insertedRows = (int) ($totals['inserted_rows'] ?? 0);
        $chunksCompleted = (int) ($totals['chunks_completed'] ?? 0);

        $readerOptions = $options;
        $readerOptions['sheet'] = $sheet;
        $readerOptions['with_header'] = $withHeader;
        if ($headers !== []) {
            $readerOptions['headers'] = $headers;
        }
        if ($withHeader && $startAfterRowNumber !== null && $startAfterRowNumber > 0) {
            $readerOptions['header_consumed'] = true;
        }
        if ($startAfterRowNumber !== null && $startAfterRowNumber > 0) {
            $readerOptions['start_after_row_number'] = $startAfterRowNumber;
        }

        EventDispatcher::safeDispatch('before_import', ['path' => $path, 'table' => $table, 'sheet' => $sheet, 'manifest_path' => $manifestPath]);
        LoggerBridge::info('Large Excel import started.', ['path' => $path, 'table' => $table, 'sheet' => $sheet]);

        try {
            $summary = $this->reader->chunk($path, $chunkSize, function (array $rows, array $state) use (
                $pdo,
                $table,
                $rules,
                $validationOptions,
                $map,
                $strictValidation,
                $transactionPerChunk,
                $dryRun,
                $batchSize,
                $duplicateStrategy,
                $uniqueBy,
                $failedWriter,
                $manifest,
                $transformers,
                $progress,
                $withHeader,
                &$headers,
                &$rowsScanned,
                &$validRows,
                &$failedRows,
                &$insertedRows,
                &$chunksCompleted
            ): bool|null {
            EventDispatcher::safeDispatch('before_chunk', ['state' => $state, 'rows' => count($rows)]);
            if ($withHeader && $headers === [] && isset($rows[0]) && is_array($rows[0])) {
                $keys = array_keys($rows[0]);
                if ($keys !== range(0, count($keys) - 1)) {
                    $headers = array_values(array_map(static fn (mixed $key): string => (string) $key, $keys));
                    $manifest->update(['headers' => $headers]);
                }
            }
            if ($transformers !== []) {
                $rows = RowTransformerPipeline::applyRows($rows, $transformers, $state);
            }
            $rowNumbers = is_array($state['chunk_row_numbers'] ?? null) ? array_values($state['chunk_row_numbers']) : [];
            $rowsWithNumbers = $this->attachRowNumbers($rows, $rowNumbers);
            $rowsScanned += count($rowsWithNumbers);

            $valid = $rowsWithNumbers;
            $failed = [];
            if ($rules !== []) {
                $validation = $this->validator->validate($rowsWithNumbers, $rules, array_merge($validationOptions, [
                    'row_number_key' => '_mnb_excel_row',
                ]));
                $valid = $validation['valid'];
                $failed = $validation['failed'];
            }

            if ($failed !== []) {
                $failedRows += $failedWriter->append($failed);
                EventDispatcher::safeDispatch('on_failed_row', ['failed_rows' => $failed, 'count' => count($failed), 'state' => $state]);
                if ($strictValidation) {
                    $manifest->update([
                        'status' => 'failed',
                        'last_stop_reason' => 'validation_failed',
                        'last_row_number' => (int) ($state['chunk_last_row_number'] ?? 0),
                        'rows_scanned' => $rowsScanned,
                        'valid_rows' => $validRows,
                        'failed_rows' => $failedRows,
                        'inserted_rows' => $insertedRows,
                        'chunks_completed' => $chunksCompleted,
                    ]);
                    throw MnbExcelException::withCode('Large Excel import failed validation.', ErrorCode::VALIDATION_FAILED, ['failed_rows' => count($failed)]);
                }
            }

            $validRows += count($valid);
            $insertable = $this->stripInternalColumns($valid);
            if ($insertable !== []) {
                $insertedRows += $this->insertRows($pdo, $table, $insertable, [
                    'batch_size' => $batchSize,
                    'map' => $map,
                    'dry_run' => $dryRun,
                    'transaction_per_chunk' => $transactionPerChunk,
                    'duplicate_strategy' => $duplicateStrategy,
                    'unique_by' => $uniqueBy,
                ]);
            }

            $chunksCompleted++;
            $manifestPatch = [
                'status' => 'running',
                'last_row_number' => (int) ($state['chunk_last_row_number'] ?? $state['current_row_number'] ?? 0),
                'rows_scanned' => $rowsScanned,
                'valid_rows' => $validRows,
                'failed_rows' => $failedRows,
                'inserted_rows' => $insertedRows,
                'chunks_completed' => $chunksCompleted,
                'last_elapsed_seconds' => $state['elapsed_seconds'] ?? null,
                'last_memory_usage_mb' => $state['memory_usage_mb'] ?? null,
            ];
            $manifest->update($manifestPatch);

            EventDispatcher::safeDispatch('after_chunk', array_merge($state, $manifestPatch));
            LoggerBridge::info('Large Excel import chunk completed.', $manifestPatch);

            if (is_callable($progress)) {
                $progress(array_merge($state, $manifestPatch));
            }

                return null;
            }, $readerOptions);
        } catch (Throwable $e) {
            EventDispatcher::safeDispatch('on_import_failed', ['path' => $path, 'table' => $table, 'sheet' => $sheet, 'exception' => $e]);
            LoggerBridge::error('Large Excel import failed.', ['path' => $path, 'table' => $table, 'message' => $e->getMessage()]);
            $manifest->update([
                'status' => 'failed',
                'last_stop_reason' => $e instanceof MnbExcelException ? $e->getErrorCode() : 'exception',
                'last_error_message' => $e->getMessage(),
                'rows_scanned' => $rowsScanned,
                'valid_rows' => $validRows,
                'failed_rows' => $failedRows,
                'inserted_rows' => $insertedRows,
                'chunks_completed' => $chunksCompleted,
            ]);
            throw $e;
        }

        if (isset($summary['rows_delivered'])) {
            $rowsScanned = max($rowsScanned, $resumeBaseRowsScanned + (int) $summary['rows_delivered']);
        }

        $status = ($summary['stopped'] ?? false) ? 'paused' : 'completed';
        EventDispatcher::safeDispatch($status === 'completed' ? 'on_import_completed' : 'on_import_paused', ['path' => $path, 'table' => $table, 'sheet' => $sheet, 'summary' => $summary]);
        LoggerBridge::info('Large Excel import ' . $status . '.', ['path' => $path, 'table' => $table, 'summary' => $summary]);
        $manifest->update([
            'status' => $status,
            'last_stop_reason' => $summary['stop_reason'] ?? null,
            'completed_at' => $status === 'completed' ? gmdate('c') : null,
            'rows_scanned' => $rowsScanned,
            'valid_rows' => $validRows,
            'failed_rows' => $failedRows,
            'inserted_rows' => $insertedRows,
            'chunks_completed' => $chunksCompleted,
        ]);

        return [
            'status' => $status,
            'source_path' => $path,
            'table' => $table,
            'sheet' => $sheet,
            'rows_scanned' => $rowsScanned,
            'valid_rows' => $validRows,
            'failed_rows' => $failedRows,
            'inserted_rows' => $insertedRows,
            'chunks_completed' => $chunksCompleted,
            'chunk_size' => $chunkSize,
            'batch_size' => $batchSize,
            'idempotent' => $idempotent,
            'duplicate_strategy' => $duplicateStrategy,
            'unique_by' => $uniqueBy,
            'failed_rows_format' => $failedRowsFormat,
            'manifest_path' => $manifest->path(),
            'failed_rows_csv' => $failedRowsCsv,
            'reader_summary' => $summary,
            'resume_ready' => $status !== 'completed',
        ];
    }

    /**
     * Import all or selected sheets from a workbook. Each sheet is imported by the same streaming engine.
     *
     * @param string|array<int|string,string> $tableMap
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function importWorkbookToSql(string $path, PDO $pdo, string|array $tableMap, array $options = []): array
    {
        $realPath = realpath($path);
        if ($realPath === false || !is_file($path)) {
            throw MnbExcelException::withCode('XLSX file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $sheets = $this->resolver->sheets($realPath);
        $onlySheets = $this->stringList($options['sheets'] ?? []);
        $results = [];
        $status = 'completed';
        $stopOnPause = (bool) ($options['stop_on_pause'] ?? true);

        foreach ($sheets as $sheet) {
            if (!($sheet['exists'] ?? false)) {
                continue;
            }
            $sheetName = (string) $sheet['name'];
            $sheetIndex = (int) $sheet['index'];
            if ($onlySheets !== [] && !in_array($sheetName, $onlySheets, true) && !in_array((string) $sheetIndex, $onlySheets, true)) {
                continue;
            }

            $table = $this->tableForSheet($tableMap, $sheetName, $sheetIndex, count($sheets));
            $sheetOptions = $options;
            $sheetOptions['sheet'] = $sheetName;
            $manifestPaths = is_array($options['manifest_paths'] ?? null) ? $options['manifest_paths'] : [];
            $failedCsvs = is_array($options['failed_rows_csvs'] ?? null) ? $options['failed_rows_csvs'] : [];
            $sheetOptions['manifest_path'] = (string) ($manifestPaths[$sheetName] ?? $manifestPaths[$sheetIndex] ?? LargeImportManifest::defaultPath($path, $table, $sheetName));
            $sheetOptions['failed_rows_csv'] = (string) ($failedCsvs[$sheetName] ?? $failedCsvs[$sheetIndex] ?? $this->defaultFailedRowsPath($path, $sheetOptions['manifest_path']));

            $result = $this->importToSql($path, $pdo, $table, $sheetOptions);
            $results[$sheetName] = $result;
            if (($result['status'] ?? '') !== 'completed') {
                $status = (string) $result['status'];
                if ($stopOnPause) {
                    break;
                }
            }
        }

        return [
            'status' => $status,
            'source_path' => $path,
            'sheets_total' => count($sheets),
            'sheets_imported' => count($results),
            'results' => $results,
            'rows_scanned' => array_sum(array_map(static fn(array $r): int => (int) ($r['rows_scanned'] ?? 0), $results)),
            'inserted_rows' => array_sum(array_map(static fn(array $r): int => (int) ($r['inserted_rows'] ?? 0), $results)),
            'failed_rows' => array_sum(array_map(static fn(array $r): int => (int) ($r['failed_rows'] ?? 0), $results)),
            'resume_ready' => $status !== 'completed',
        ];
    }

    /** @param array<string,mixed> $options @return list<string> */
    private function readHeader(string $path, int|string $sheet, array $options): array
    {
        $headerRows = [];
        $readerOptions = $options;
        $readerOptions['sheet'] = $sheet;
        $readerOptions['with_header'] = false;
        $readerOptions['limit_rows'] = 1;
        unset($readerOptions['headers'], $readerOptions['start_after_row_number'], $readerOptions['skip_rows']);

        $this->reader->chunk($path, 1, static function (array $rows) use (&$headerRows): bool {
            $headerRows = $rows;
            return false;
        }, $readerOptions);

        $first = $headerRows[0] ?? [];
        $headers = [];
        foreach ($first as $i => $value) {
            $name = trim((string) $value);
            $name = $name !== '' ? $name : 'column_' . ((int) $i + 1);
            $base = $name;
            $suffix = 2;
            while (in_array($name, $headers, true)) {
                $name = $base . '_' . $suffix;
                $suffix++;
            }
            $headers[] = $name;
        }
        return $headers;
    }

    /** @param list<array<int|string,mixed>> $rows @param list<int> $rowNumbers @return list<array<int|string,mixed>> */
    private function attachRowNumbers(array $rows, array $rowNumbers): array
    {
        foreach ($rows as $i => $row) {
            if (is_array($row)) {
                $row['_mnb_excel_row'] = (int) ($rowNumbers[$i] ?? ($i + 1));
                $rows[$i] = $row;
            }
        }
        return $rows;
    }

    /** @param list<array<int|string,mixed>> $rows @return list<array<int|string,mixed>> */
    private function stripInternalColumns(array $rows): array
    {
        foreach ($rows as $i => $row) {
            unset($row['_mnb_excel_row']);
            $rows[$i] = $row;
        }
        return $rows;
    }

    /** @param list<array<int|string,mixed>> $rows @param array<string,mixed> $options */
    private function insertRows(PDO $pdo, string $table, array $rows, array $options): int
    {
        if ($rows === []) {
            return 0;
        }

        $transaction = (bool) ($options['transaction_per_chunk'] ?? true);
        $startedTransaction = false;
        try {
            if ($transaction && !$pdo->inTransaction()) {
                $pdo->beginTransaction();
                $startedTransaction = true;
            }
            $result = $this->sqlImporter->importRows($pdo, $table, $rows, [
                'batch_size' => (int) ($options['batch_size'] ?? 500),
                'map' => is_array($options['map'] ?? null) ? $options['map'] : [],
                'dry_run' => (bool) ($options['dry_run'] ?? false),
                'duplicate_strategy' => (string) ($options['duplicate_strategy'] ?? 'fail'),
                'unique_by' => $this->stringList($options['unique_by'] ?? []),
            ]);
            if ($startedTransaction) {
                $pdo->commit();
            }
            return (int) ($result['inserted_rows'] ?? 0);
        } catch (Throwable $e) {
            if ($startedTransaction && $pdo->inTransaction()) {
                $pdo->rollBack();
            }
            if ($e instanceof MnbExcelException) {
                throw $e;
            }
            throw MnbExcelException::withCode('Large Excel database import failed: ' . $e->getMessage(), ErrorCode::SQL_IMPORT_FAILED, ['table' => $table], $e);
        }
    }

    /** @return list<string> */
    private function stringList(mixed $value): array
    {
        if ($value === null || $value === '') {
            return [];
        }
        if (is_string($value)) {
            $value = array_map('trim', explode(',', $value));
        }
        if (!is_array($value)) {
            return [];
        }
        return array_values(array_filter(array_map(static fn(mixed $item): string => (string) $item, $value), static fn(string $item): bool => trim($item) !== ''));
    }

    /** @param string|array<int|string,string> $tableMap */
    private function tableForSheet(string|array $tableMap, string $sheetName, int $sheetIndex, int $sheetCount): string
    {
        if (is_array($tableMap)) {
            return (string) ($tableMap[$sheetName] ?? $tableMap[$sheetIndex] ?? $tableMap[(string) $sheetIndex] ?? $tableMap['*'] ?? $this->safeTableName($sheetName));
        }

        $base = trim($tableMap);
        if ($base === '') {
            return $this->safeTableName($sheetName);
        }
        if ($sheetCount <= 1) {
            return $base;
        }
        return $base . '_' . $this->safeTableName($sheetName);
    }

    private function safeTableName(string $name): string
    {
        $safe = strtolower((string) preg_replace('/[^A-Za-z0-9_]+/', '_', trim($name)));
        $safe = trim($safe, '_');
        if ($safe === '' || preg_match('/^[0-9]/', $safe) === 1) {
            $safe = 'sheet_' . $safe;
        }
        return $safe;
    }

    private function defaultFailedRowsPath(string $sourcePath, string $manifestPath): string
    {
        $name = pathinfo($sourcePath, PATHINFO_FILENAME) ?: 'large-import';
        return dirname($manifestPath) . DIRECTORY_SEPARATOR . $name . '-failed-rows.csv';
    }
}
