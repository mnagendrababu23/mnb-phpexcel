<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Import;

use Mnb\PHPExcel\Support\MnbExcelException;

final class ImportQualityAnalyzer
{
    /**
     * Build a safe import preview from array rows.
     *
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function preview(array $rows, array $options = []): array
    {
        $sampleSize = max(1, (int) ($options['sample_size'] ?? 5));
        $headers = $this->resolveHeaders($rows, $options);
        $requiredColumns = $this->stringList($options['required_columns'] ?? []);
        $allowedColumns = $this->stringList($options['allowed_columns'] ?? []);
        $strictColumns = (bool) ($options['strict_columns'] ?? false);
        $duplicateBy = $this->stringList($options['duplicate_by'] ?? []);

        $missingRequired = $this->missingColumns($headers, $requiredColumns);
        $unexpectedColumns = $allowedColumns === [] ? [] : array_values(array_diff($headers, $allowedColumns));
        $emptyRows = 0;
        $columnStats = [];

        foreach ($headers as $header) {
            $columnStats[$header] = [
                'filled' => 0,
                'empty' => 0,
                'types' => [],
                'examples' => [],
            ];
        }

        foreach ($rows as $row) {
            if ($this->isEmptyRow($row)) {
                $emptyRows++;
            }

            $assoc = $this->rowToAssoc($row, $headers);
            foreach ($headers as $header) {
                $value = $assoc[$header] ?? null;
                $isEmpty = $this->isEmptyValue($value);
                if ($isEmpty) {
                    $columnStats[$header]['empty']++;
                } else {
                    $columnStats[$header]['filled']++;
                    $type = $this->detectType($value);
                    $columnStats[$header]['types'][$type] = ($columnStats[$header]['types'][$type] ?? 0) + 1;
                    if (count($columnStats[$header]['examples']) < 3) {
                        $columnStats[$header]['examples'][] = $value;
                    }
                }
            }
        }

        $duplicates = $duplicateBy === [] ? [] : $this->findDuplicates($rows, $duplicateBy, $options);
        $warnings = [];

        if ($missingRequired !== []) {
            $warnings[] = 'Missing required columns: ' . implode(', ', $missingRequired) . '.';
        }
        if ($strictColumns && $unexpectedColumns !== []) {
            $warnings[] = 'Unexpected columns found in strict mode: ' . implode(', ', $unexpectedColumns) . '.';
        }
        if ($emptyRows > 0) {
            $warnings[] = 'Empty rows found: ' . $emptyRows . '.';
        }
        if ($duplicates !== []) {
            $warnings[] = 'Duplicate rows found: ' . count($duplicates) . ' duplicate group(s).';
        }

        return [
            'status' => $warnings === [] ? 'ok' : 'warning',
            'total_rows' => count($rows),
            'total_columns' => count($headers),
            'columns' => $headers,
            'required_columns' => $requiredColumns,
            'missing_required_columns' => $missingRequired,
            'unexpected_columns' => $unexpectedColumns,
            'empty_rows' => $emptyRows,
            'duplicate_groups' => $duplicates,
            'sample_rows' => array_slice($rows, 0, $sampleSize),
            'column_stats' => $columnStats,
            'warnings' => $warnings,
        ];
    }

    /**
     * Suggest map from source columns to target columns.
     *
     * @param list<string> $sourceColumns
     * @param list<string> $targetColumns
     * @param array<string,list<string>|string> $aliases
     * @return array<string,array{target:?string,confidence:float,reason:string}>
     */
    public function suggestColumnMap(array $sourceColumns, array $targetColumns, array $aliases = [], float $minConfidence = 0.55): array
    {
        $suggestions = [];
        $normalizedTargets = [];
        foreach ($targetColumns as $target) {
            $normalizedTargets[$target] = $this->normalizeColumnName($target);
        }

        $aliasLookup = [];
        foreach ($aliases as $target => $aliasValues) {
            foreach ((array) $aliasValues as $alias) {
                $aliasLookup[$this->normalizeColumnName((string) $alias)] = (string) $target;
            }
        }

        foreach ($sourceColumns as $source) {
            $sourceNorm = $this->normalizeColumnName($source);

            if (isset($aliasLookup[$sourceNorm]) && in_array($aliasLookup[$sourceNorm], $targetColumns, true)) {
                $suggestions[$source] = [
                    'target' => $aliasLookup[$sourceNorm],
                    'confidence' => 1.0,
                    'reason' => 'alias_match',
                ];
                continue;
            }

            $bestTarget = null;
            $bestScore = 0.0;
            $bestReason = 'no_match';

            foreach ($normalizedTargets as $target => $targetNorm) {
                if ($sourceNorm === $targetNorm) {
                    $bestTarget = $target;
                    $bestScore = 1.0;
                    $bestReason = 'exact_match';
                    break;
                }

                $score = $this->similarity($sourceNorm, $targetNorm);
                if ($score > $bestScore) {
                    $bestTarget = $target;
                    $bestScore = $score;
                    $bestReason = 'similarity_match';
                }
            }

            $suggestions[$source] = [
                'target' => $bestScore >= $minConfidence ? $bestTarget : null,
                'confidence' => round($bestScore, 3),
                'reason' => $bestScore >= 0.55 ? $bestReason : 'low_confidence',
            ];
        }

        return $suggestions;
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param list<string> $columns
     * @param array<string,mixed> $options
     * @return list<array{key:string,count:int,rows:list<int>}>
     */
    public function findDuplicates(array $rows, array $columns, array $options = []): array
    {
        if ($columns === []) {
            return [];
        }

        $headers = $this->resolveHeaders($rows, $options);
        $columns = $this->stringList($columns);
        $rowNumberKey = isset($options['row_number_key']) ? (string) $options['row_number_key'] : (isset($options['original_row_number_key']) ? (string) $options['original_row_number_key'] : '_mnb_original_row_number');
        $startRow = (int) ($options['start_row'] ?? 1);
        $seen = [];

        foreach ($rows as $index => $row) {
            $assoc = $this->rowToAssoc($row, $headers);
            $parts = [];
            foreach ($columns as $column) {
                $parts[] = strtolower(trim((string) ($assoc[$column] ?? '')));
            }

            $key = implode('|', $parts);
            if (trim(str_replace('|', '', $key)) === '') {
                continue;
            }

            $rowNumber = $rowNumberKey !== null && isset($assoc[$rowNumberKey]) && is_numeric($assoc[$rowNumberKey])
                ? (int) $assoc[$rowNumberKey]
                : $startRow + $index;

            $seen[$key][] = $rowNumber;
        }

        $duplicates = [];
        foreach ($seen as $key => $rowNumbers) {
            if (count($rowNumbers) > 1) {
                $duplicates[] = [
                    'key' => $key,
                    'count' => count($rowNumbers),
                    'rows' => $rowNumbers,
                ];
            }
        }

        return $duplicates;
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     * @return list<string>
     */
    private function resolveHeaders(array $rows, array $options): array
    {
        if (isset($options['columns']) && is_array($options['columns'])) {
            return array_values(array_map('strval', $options['columns']));
        }

        if ($rows === []) {
            return [];
        }

        $first = $rows[0];
        if (!is_array($first)) {
            throw new MnbExcelException('Import rows must be arrays.');
        }

        $keys = array_keys($first);
        if ($keys === range(0, count($keys) - 1)) {
            return array_map(static fn(int $index): string => 'column_' . ($index + 1), range(0, max(0, count($first) - 1)));
        }

        return array_map('strval', $keys);
    }

    /** @param array<string,mixed>|list<mixed> $row @param list<string> $headers @return array<string,mixed> */
    private function rowToAssoc(array $row, array $headers): array
    {
        $keys = array_keys($row);
        if ($keys !== range(0, count($keys) - 1)) {
            /** @var array<string,mixed> $row */
            return $row;
        }

        $assoc = [];
        foreach ($headers as $index => $header) {
            $assoc[$header] = $row[$index] ?? null;
        }
        return $assoc;
    }

    /** @param array<string,mixed>|list<mixed> $row */
    private function isEmptyRow(array $row): bool
    {
        foreach ($row as $value) {
            if (!$this->isEmptyValue($value)) {
                return false;
            }
        }
        return true;
    }

    private function isEmptyValue(mixed $value): bool
    {
        return $value === null || trim((string) $value) === '';
    }

    private function detectType(mixed $value): string
    {
        if (is_bool($value)) {
            return 'boolean';
        }
        if (is_int($value)) {
            return 'integer';
        }
        if (is_float($value)) {
            return 'float';
        }
        if ($value instanceof \DateTimeInterface) {
            return 'date';
        }
        if (is_string($value)) {
            if (is_numeric($value)) {
                return 'numeric_string';
            }
            if (strtotime($value) !== false && preg_match('/\d{2,4}[-\/]/', $value)) {
                return 'date_string';
            }
            return 'string';
        }
        return get_debug_type($value);
    }

    /** @param list<string> $headers @param list<string> $required */
    private function missingColumns(array $headers, array $required): array
    {
        $lookup = array_map(fn(string $v): string => $this->normalizeColumnName($v), $headers);
        $missing = [];
        foreach ($required as $column) {
            if (!in_array($this->normalizeColumnName($column), $lookup, true)) {
                $missing[] = $column;
            }
        }
        return $missing;
    }

    /** @return list<string> */
    private function stringList(mixed $value): array
    {
        if ($value === null || $value === '') {
            return [];
        }
        if (is_string($value)) {
            return [$value];
        }
        if (!is_array($value)) {
            return [(string) $value];
        }
        return array_values(array_map('strval', $value));
    }

    private function normalizeColumnName(string $column): string
    {
        $column = strtolower(trim($column));
        $column = preg_replace('/[^a-z0-9]+/', '', $column) ?: '';
        return $column;
    }

    private function similarity(string $a, string $b): float
    {
        if ($a === '' || $b === '') {
            return 0.0;
        }
        if ($a === $b) {
            return 1.0;
        }
        if (str_contains($a, $b) || str_contains($b, $a)) {
            return 0.85;
        }
        similar_text($a, $b, $percent);
        return $percent / 100;
    }
}
