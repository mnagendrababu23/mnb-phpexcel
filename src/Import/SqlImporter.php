<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Import;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;

final class SqlImporter
{
    /**
     * @param list<array<string, mixed>|list<mixed>> $rows
     * @return array<string,mixed>
     */
    public function importRows(PDO $pdo, string $table, array $rows, array $options = []): array
    {
        if ($rows === []) {
            return [
                'status' => 'completed',
                'total_rows' => 0,
                'inserted_rows' => 0,
                'failed_rows' => 0,
                'errors' => [],
            ];
        }

        $batchSize = max(1, (int) ($options['batch_size'] ?? 500));
        $map = is_array($options['map'] ?? null) ? $options['map'] : [];
        $skipInvalidRows = (bool) ($options['skip_invalid_rows'] ?? false);
        $dryRun = (bool) ($options['dry_run'] ?? false);
        $duplicateStrategy = strtolower((string) ($options['duplicate_strategy'] ?? 'fail'));
        if (!in_array($duplicateStrategy, ['fail', 'skip', 'update'], true)) {
            throw MnbExcelException::withCode('Unsupported duplicate strategy: ' . $duplicateStrategy, ErrorCode::SQL_IMPORT_FAILED);
        }
        $uniqueBy = $this->stringList($options['unique_by'] ?? []);
        $columns = $this->resolveColumns($rows[0], $map);

        $inserted = 0;
        $planned = 0;
        $errors = [];
        $batch = [];

        foreach ($rows as $index => $row) {
            try {
                $batch[] = $this->normalizeRow($row, $columns, $map);
                $planned++;
            } catch (Throwable $e) {
                if (!$skipInvalidRows) {
                    throw $e;
                }
                $errors[] = ['row' => $index + 1, 'message' => $e->getMessage()];
            }

            if (count($batch) >= $batchSize) {
                if (!$dryRun) {
                    $inserted += $this->insertBatch($pdo, $table, $columns, $batch, $duplicateStrategy, $uniqueBy);
                }
                $batch = [];
            }
        }

        if ($batch !== []) {
            if (!$dryRun) {
                $inserted += $this->insertBatch($pdo, $table, $columns, $batch, $duplicateStrategy, $uniqueBy);
            }
        }

        return [
            'status' => $dryRun ? 'dry_run' : 'completed',
            'total_rows' => count($rows),
            'planned_rows' => $planned,
            'inserted_rows' => $dryRun ? 0 : $inserted,
            'failed_rows' => count($errors),
            'columns' => $columns,
            'batch_size' => $batchSize,
            'planned_batches' => (int) ceil($planned / $batchSize),
            'errors' => $errors,
        ];
    }

    /**
     * @param array<string, mixed>|list<mixed> $firstRow
     * @param array<string, string> $map
     * @return list<string>
     */
    private function resolveColumns(array $firstRow, array $map): array
    {
        if ($map !== []) {
            return array_values($map);
        }

        $keys = array_keys($firstRow);
        if ($keys === range(0, count($keys) - 1)) {
            throw MnbExcelException::withCode('SQL import needs associative rows or a column map.', ErrorCode::SQL_IMPORT_FAILED);
        }

        return array_map(static fn($key): string => (string) $key, $keys);
    }

    /**
     * @param array<string, mixed>|list<mixed> $row
     * @param list<string> $columns
     * @param array<string, string> $map
     * @return array<string, mixed>
     */
    private function normalizeRow(array $row, array $columns, array $map): array
    {
        $normalized = [];

        if ($map !== []) {
            foreach ($map as $source => $target) {
                $normalized[$target] = $row[$source] ?? null;
            }
            return $normalized;
        }

        foreach ($columns as $column) {
            $normalized[$column] = $row[$column] ?? null;
        }

        return $normalized;
    }

    /**
     * @param list<string> $columns
     * @param list<array<string, mixed>> $batch
     */
    private function insertBatch(PDO $pdo, string $table, array $columns, array $batch, string $duplicateStrategy = 'fail', array $uniqueBy = []): int
    {
        if ($batch === []) {
            return 0;
        }

        $driver = strtolower((string) $pdo->getAttribute(PDO::ATTR_DRIVER_NAME));
        $quotedTable = $this->quoteIdentifier($table, $driver);
        $quotedColumns = array_map(fn(string $column): string => $this->quoteIdentifier($column, $driver), $columns);
        $placeholders = [];
        $values = [];

        foreach ($batch as $row) {
            $placeholders[] = '(' . implode(',', array_fill(0, count($columns), '?')) . ')';
            foreach ($columns as $column) {
                $values[] = $row[$column] ?? null;
            }
        }

        $existingBefore = null;
        if ($duplicateStrategy === 'skip' && $uniqueBy !== []) {
            $existingBefore = $this->countExistingByUnique($pdo, $table, $driver, $batch, $uniqueBy);
        }

        $insertVerb = 'INSERT INTO';
        if ($duplicateStrategy === 'skip') {
            if ($driver === 'sqlite') {
                $insertVerb = 'INSERT OR IGNORE INTO';
            } elseif ($driver === 'mysql') {
                $insertVerb = 'INSERT IGNORE INTO';
            } elseif ($driver !== 'pgsql') {
                throw MnbExcelException::withCode('Duplicate skip strategy is not supported for PDO driver: ' . $driver, ErrorCode::SQL_IMPORT_FAILED);
            }
        }

        $sql = $insertVerb . ' ' . $quotedTable . ' (' . implode(',', $quotedColumns) . ') VALUES ' . implode(',', $placeholders);
        $sql .= $this->duplicateClause($driver, $duplicateStrategy, $columns, $uniqueBy);

        try {
            $stmt = $pdo->prepare($sql);
            if ($stmt === false) {
                throw MnbExcelException::withCode('Unable to prepare SQL import statement.', ErrorCode::SQL_IMPORT_FAILED);
            }
            $stmt->execute($values);
            if ($duplicateStrategy === 'fail') {
                return count($batch);
            }
            if ($duplicateStrategy === 'skip' && $existingBefore !== null) {
                return max(0, count($batch) - $existingBefore);
            }
            if ($duplicateStrategy === 'update') {
                return count($batch);
            }
            $affected = $stmt->rowCount();
            return $affected > 0 ? $affected : count($batch);
        } catch (MnbExcelException $e) {
            throw $e;
        } catch (Throwable $e) {
            throw MnbExcelException::withCode(
                'SQL import failed: ' . $e->getMessage(),
                ErrorCode::SQL_IMPORT_FAILED,
                ['table' => $table, 'columns' => $columns, 'rows' => count($batch), 'duplicate_strategy' => $duplicateStrategy, 'unique_by' => $uniqueBy],
                $e
            );
        }
    }


    /** @param list<array<string, mixed>> $batch @param list<string> $uniqueBy */
    private function countExistingByUnique(PDO $pdo, string $table, string $driver, array $batch, array $uniqueBy): int
    {
        if ($batch === [] || $uniqueBy === []) {
            return 0;
        }

        $quotedTable = $this->quoteIdentifier($table, $driver);
        $conditions = [];
        $values = [];
        foreach ($batch as $row) {
            $parts = [];
            foreach ($uniqueBy as $column) {
                $parts[] = $this->quoteIdentifier($column, $driver) . ' = ?';
                $values[] = $row[$column] ?? null;
            }
            $conditions[] = '(' . implode(' AND ', $parts) . ')';
        }

        if ($conditions === []) {
            return 0;
        }

        $sql = 'SELECT COUNT(*) FROM ' . $quotedTable . ' WHERE ' . implode(' OR ', $conditions);
        $stmt = $pdo->prepare($sql);
        if ($stmt === false) {
            return 0;
        }
        $stmt->execute($values);
        return max(0, (int) $stmt->fetchColumn());
    }

    /** @param list<string> $columns @param list<string> $uniqueBy */
    private function duplicateClause(string $driver, string $strategy, array $columns, array $uniqueBy): string
    {
        if ($strategy === 'fail') {
            return '';
        }

        $quotedUnique = array_map(fn(string $column): string => $this->quoteIdentifier($column, $driver), $uniqueBy);
        $updateColumns = array_values(array_diff($columns, $uniqueBy));

        if ($strategy === 'skip') {
            if ($driver === 'pgsql') {
                return $quotedUnique !== [] ? ' ON CONFLICT (' . implode(',', $quotedUnique) . ') DO NOTHING' : ' ON CONFLICT DO NOTHING';
            }
            return '';
        }

        if ($uniqueBy === []) {
            throw MnbExcelException::withCode('Duplicate update strategy requires unique_by columns.', ErrorCode::SQL_IMPORT_FAILED);
        }
        if ($updateColumns === []) {
            return $driver === 'pgsql' || $driver === 'sqlite' ? ' ON CONFLICT (' . implode(',', $quotedUnique) . ') DO NOTHING' : '';
        }

        if ($driver === 'sqlite' || $driver === 'pgsql') {
            $assignments = array_map(fn(string $column): string => $this->quoteIdentifier($column, $driver) . ' = excluded.' . $this->quoteIdentifier($column, $driver), $updateColumns);
            return ' ON CONFLICT (' . implode(',', $quotedUnique) . ') DO UPDATE SET ' . implode(',', $assignments);
        }

        if ($driver === 'mysql') {
            $assignments = array_map(fn(string $column): string => $this->quoteIdentifier($column, $driver) . ' = VALUES(' . $this->quoteIdentifier($column, $driver) . ')', $updateColumns);
            return ' ON DUPLICATE KEY UPDATE ' . implode(',', $assignments);
        }

        throw MnbExcelException::withCode('Duplicate update strategy is not supported for PDO driver: ' . $driver, ErrorCode::SQL_IMPORT_FAILED);
    }

    private function quoteIdentifier(string $identifier, string $driver = 'mysql'): string
    {
        if (!preg_match('/^[A-Za-z_][A-Za-z0-9_]*$/', $identifier)) {
            throw MnbExcelException::withCode('Unsafe SQL identifier: ' . $identifier, ErrorCode::SQL_IMPORT_FAILED, ['identifier' => $identifier]);
        }

        if ($driver === 'pgsql') {
            return '"' . str_replace('"', '""', $identifier) . '"';
        }

        return '`' . str_replace('`', '``', $identifier) . '`';
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
}
