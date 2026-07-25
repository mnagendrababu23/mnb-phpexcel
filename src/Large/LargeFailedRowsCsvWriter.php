<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Writes failed large-import rows without keeping all failures in memory. */
final class LargeFailedRowsCsvWriter
{
    private bool $headerWritten = false;
    /** @var list<string> */
    private array $columns = [];
    private string $format;

    /** @param array<string,mixed> $options */
    public function __construct(private string $path, bool $reset = false, array $options = [])
    {
        if ($reset && is_file($path)) {
            @unlink($path);
        }
        $this->headerWritten = is_file($path) && filesize($path) > 0;
        $this->format = strtolower((string) ($options['format'] ?? $options['failed_rows_format'] ?? 'human'));
        if (!in_array($this->format, ['human', 'json'], true)) {
            $this->format = 'human';
        }
        $this->columns = $this->normalizeColumns($options['columns'] ?? []);
    }

    public function path(): string
    {
        return $this->path;
    }

    /**
     * @param list<array{row:int,errors:list<string>,data:array<string,mixed>}> $failedRows
     */
    public function append(array $failedRows): int
    {
        if ($failedRows === []) {
            return 0;
        }

        $dir = dirname($this->path);
        if (!is_dir($dir) && !@mkdir($dir, 0775, true) && !is_dir($dir)) {
            throw MnbExcelException::withCode('Unable to create failed-row export directory: ' . $dir, ErrorCode::DIRECTORY_CREATE_FAILED, ['path' => $dir]);
        }

        $handle = @fopen($this->path, 'ab');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open failed-row CSV: ' . $this->path, ErrorCode::FILE_OPEN_FAILED, ['path' => $this->path]);
        }

        try {
            if ($this->format === 'json') {
                $this->appendJsonRows($handle, $failedRows);
            } else {
                $this->appendHumanRows($handle, $failedRows);
            }
        } finally {
            if (!@fclose($handle)) {
                throw MnbExcelException::withCode('Unable to close failed-row CSV: ' . $this->path, ErrorCode::CSV_WRITE_FAILED, ['path' => $this->path]);
            }
        }

        return count($failedRows);
    }

    /** @param resource $handle @param list<array{row:int,errors:list<string>,data:array<string,mixed>}> $failedRows */
    private function appendJsonRows($handle, array $failedRows): void
    {
        if (!$this->headerWritten) {
            $this->writeCsv($handle, ['excel_row', 'errors', 'data_json']);
            $this->headerWritten = true;
        }
        foreach ($failedRows as $failed) {
            $data = $this->cleanData($failed['data'] ?? []);
            $dataJson = json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
            if ($dataJson === false) {
                $dataJson = '{}';
            }
            $this->writeCsv($handle, [
                (string) ($failed['row'] ?? ''),
                implode('; ', $failed['errors'] ?? []),
                $dataJson,
            ]);
        }
    }

    /** @param resource $handle @param list<array{row:int,errors:list<string>,data:array<string,mixed>}> $failedRows */
    private function appendHumanRows($handle, array $failedRows): void
    {
        if ($this->columns === []) {
            $this->columns = $this->columnsFromFailedRows($failedRows);
        }
        if (!$this->headerWritten) {
            $this->writeCsv($handle, array_merge(['excel_row', 'error_count', 'errors'], $this->columns));
            $this->headerWritten = true;
        }

        foreach ($failedRows as $failed) {
            $data = $this->cleanData($failed['data'] ?? []);
            $row = [
                (string) ($failed['row'] ?? ''),
                (string) count($failed['errors'] ?? []),
                implode('; ', $failed['errors'] ?? []),
            ];
            foreach ($this->columns as $column) {
                $value = $data[$column] ?? '';
                $encoded = json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
                $row[] = is_scalar($value) || $value === null ? (string) $value : ($encoded !== false ? $encoded : '');
            }
            $this->writeCsv($handle, $row);
        }
    }

    /** @param resource $handle @param list<string> $row */
    private function writeCsv($handle, array $row): void
    {
        if (@fputcsv($handle, $row) === false) {
            throw MnbExcelException::withCode('Unable to write failed-row CSV: ' . $this->path, ErrorCode::CSV_WRITE_FAILED, ['path' => $this->path]);
        }
    }

    /** @param array<string,mixed> $data @return array<string,mixed> */
    private function cleanData(array $data): array
    {
        unset($data['_mnb_excel_row']);
        return $data;
    }

    /** @param mixed $columns @return list<string> */
    private function normalizeColumns(mixed $columns): array
    {
        if (is_string($columns)) {
            $columns = array_map('trim', explode(',', $columns));
        }
        if (!is_array($columns)) {
            return [];
        }
        return array_values(array_filter(array_map(static fn(mixed $column): string => (string) $column, $columns), static fn(string $column): bool => trim($column) !== '' && $column !== '_mnb_excel_row'));
    }

    /** @param list<array{data:array<string,mixed>}> $failedRows @return list<string> */
    private function columnsFromFailedRows(array $failedRows): array
    {
        $columns = [];
        foreach ($failedRows as $failed) {
            foreach (array_keys($this->cleanData($failed['data'] ?? [])) as $column) {
                $column = (string) $column;
                if ($column !== '' && !in_array($column, $columns, true)) {
                    $columns[] = $column;
                }
            }
        }
        return $columns;
    }
}
