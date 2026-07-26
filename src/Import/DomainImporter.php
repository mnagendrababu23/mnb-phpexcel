<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Import;

use Mnb\PHPExcel\Application\LoggerBridge;
use Mnb\PHPExcel\Application\RowTransformerPipeline;
use Mnb\PHPExcel\Domain\DomainImportPreset;
use Mnb\PHPExcel\Domain\DomainImportRegistry;
use Mnb\PHPExcel\Domain\DomainImportType;
use Mnb\PHPExcel\Events\EventDispatcher;
use Mnb\PHPExcel\Large\LargeFailedRowsCsvWriter;
use Mnb\PHPExcel\Reader\ReaderRegistry;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Support\DatabaseConnectionFactory;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\FileFormatDetector;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Validation\ArrayValidator;
use PDO;
use Throwable;

final class DomainImporter
{
    public function __construct(
        private readonly DomainImportRegistry $domains,
        private readonly ReaderRegistry $readers,
        private readonly ArrayValidator $validator = new ArrayValidator(),
        private readonly SqlImporter $sqlImporter = new SqlImporter(),
        private readonly ImportQualityAnalyzer $quality = new ImportQualityAnalyzer()
    ) {}

    public static function create(?DomainImportRegistry $domains = null, ?ReaderRegistry $readers = null): self
    {
        return new self($domains ?? DomainImportRegistry::withBuiltIns(), $readers ?? ReaderRegistry::withBuiltIns());
    }

    public function registry(): DomainImportRegistry { return $this->domains; }
    /** @return array<string,mixed> */
    public function schema(DomainImportType|string $domain): array { return $this->domains->get($domain)->toArray(); }
    /** @return array<string,array<string,mixed>> */
    public function schemas(): array
    {
        $out = [];
        foreach ($this->domains->all() as $name => $preset) $out[$name] = $preset->toArray();
        return $out;
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function preview(DomainImportType|string $domain, string $path, array $options = []): array
    {
        $options['dry_run'] = true;
        $options['limit'] = max(1, (int) ($options['limit'] ?? $options['preview_rows'] ?? 25));
        $options['collect_rows'] = true;
        $options['max_collected_rows'] = $options['limit'];
        $options['max_errors'] = max(1, (int) ($options['max_errors'] ?? $options['limit']));
        return $this->import($domain, $path, null, (string) ($options['table'] ?? ''), $options);
    }

    /**
     * @param PDO|array<string,mixed>|string|null $pdo
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function import(DomainImportType|string $domain, string $path, PDO|array|string|null $pdo = null, string $table = '', array $options = []): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('File not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $preset = $this->domains->get($domain);
        $table = trim($table) !== '' ? trim($table) : trim((string) ($options['table'] ?? $preset->defaultTable));
        if ($table === '') throw MnbExcelException::withCode('A database table is required for domain import.', ErrorCode::VALIDATION_FAILED);

        $dryRun = (bool) ($options['dry_run'] ?? false);
        $connection = $dryRun ? null : DatabaseConnectionFactory::make($pdo, is_array($options['db'] ?? null) ? $options['db'] : []);
        $batchSize = max(1, (int) ($options['batch_size'] ?? 500));
        $collectRows = (bool) ($options['collect_rows'] ?? false);
        $maxCollectedRows = max(0, (int) ($options['max_collected_rows'] ?? 100));
        $maxErrors = max(0, (int) ($options['max_errors'] ?? 100));
        $strictValidation = (bool) ($options['strict_validation'] ?? false);
        $skipInvalid = (bool) ($options['skip_invalid_rows'] ?? true);
        $fileDuplicatePolicy = strtolower(trim((string) ($options['file_duplicate_policy'] ?? 'error')));
        if (!in_array($fileDuplicatePolicy, ['error','skip','allow'], true)) {
            throw MnbExcelException::withCode('file_duplicate_policy must be error, skip, or allow.', ErrorCode::VALIDATION_FAILED);
        }
        $databaseDuplicateStrategy = strtolower(trim((string) ($options['duplicate_strategy'] ?? 'fail')));
        if (!in_array($databaseDuplicateStrategy, ['fail','skip','update'], true)) {
            throw MnbExcelException::withCode('duplicate_strategy must be fail, skip, or update.', ErrorCode::VALIDATION_FAILED);
        }

        $rules = array_replace($preset->rules(), is_array($options['rules'] ?? null) ? $options['rules'] : []);
        $aliases = $this->mergeAliases($preset->aliases(), is_array($options['aliases'] ?? null) ? $options['aliases'] : []);
        $uniqueBy = $this->stringList($options['unique_by'] ?? $preset->uniqueBy());
        $unknownUnique = array_values(array_diff($uniqueBy, $preset->columns()));
        if ($unknownUnique !== []) {
            throw MnbExcelException::withCode('Unknown unique_by domain columns: ' . implode(', ', $unknownUnique) . '.', ErrorCode::VALIDATION_FAILED);
        }
        if ($databaseDuplicateStrategy === 'update' && $uniqueBy === []) {
            throw MnbExcelException::withCode('duplicate_strategy=update requires unique_by columns.', ErrorCode::VALIDATION_FAILED);
        }
        $defaults = array_replace($preset->defaults(), is_array($options['defaults'] ?? null) ? $options['defaults'] : []);
        $normalizers = is_array($options['normalizers'] ?? null) ? $options['normalizers'] : [];
        $transformers = is_array($options['transformers'] ?? null) ? $options['transformers'] : [];
        $explicitMap = is_array($options['map'] ?? null) ? $options['map'] : (is_array($options['column_map'] ?? null) ? $options['column_map'] : []);
        $mappingConfidence = max(0.0, min(1.0, (float) ($options['mapping_min_confidence'] ?? 0.55)));
        $strictMapping = (bool) ($options['strict_mapping'] ?? true);
        $preserveUnmapped = (bool) ($options['preserve_unmapped'] ?? false);
        $applyDefaultsToEmpty = (bool) ($options['apply_defaults_to_empty'] ?? true);
        $transactionPerBatch = (bool) ($options['transaction_per_batch'] ?? true);
        $progress = $options['progress'] ?? null;
        $rowErrorCallback = $options['on_row_error'] ?? null;
        $format = FileFormatDetector::detect($path, $options);

        $readerOptions = $this->readerOptions($options);
        $reader = $this->readers->resolve($path, $readerOptions);
        $session = (new ReadSession($path, $reader, $readerOptions))->sheet($options['sheet'] ?? 1);

        $failedRowsPath = trim((string) ($options['failed_rows_csv'] ?? ''));
        $failedWriter = $failedRowsPath !== '' ? new LargeFailedRowsCsvWriter($failedRowsPath, true, [
            'format' => (string) ($options['failed_rows_format'] ?? 'human'),
            'columns' => $preset->columns(),
        ]) : null;

        $mapping = null;
        $mappingDetails = [];
        $sourceColumns = [];
        $missingRequired = [];
        $seen = [];
        $pending = [];
        $totals = ['rows_scanned'=>0,'valid_rows'=>0,'failed_rows'=>0,'inserted_rows'=>0,'planned_rows'=>0,'batches'=>0,'skipped_duplicate_rows'=>0];
        $errors = [];
        $sampleRows = [];
        $startedAt = microtime(true);

        EventDispatcher::safeDispatch('before_domain_import', ['domain'=>$preset->type->value,'path'=>$path,'table'=>$table,'dry_run'=>$dryRun]);
        LoggerBridge::info('Domain import started.', ['domain'=>$preset->type->value,'path'=>$path,'table'=>$table,'dry_run'=>$dryRun]);

        $flush = function() use (&$pending,&$totals,&$errors,&$sampleRows,$connection,$table,$dryRun,$batchSize,$rules,$preset,$uniqueBy,$fileDuplicatePolicy,&$seen,$strictValidation,$skipInvalid,$maxErrors,$collectRows,$maxCollectedRows,$failedWriter,$rowErrorCallback,$options,$progress,$startedAt,$transactionPerBatch,$databaseDuplicateStrategy): void {
            if ($pending === []) return;
            $validation = $rules !== [] ? $this->validator->validate($pending, $rules, [
                'row_number_key'=>'_mnb_excel_row',
                'strict_columns'=>(bool)($options['strict_columns'] ?? false),
                'allowed_columns'=>array_merge($preset->columns(), ['_mnb_excel_row']),
            ]) : ['valid'=>$pending,'failed'=>[]];
            $valid = $validation['valid'];
            $failed = $validation['failed'];
            foreach ($valid as $index => $row) {
                $rowNumber = (int) ($row['_mnb_excel_row'] ?? ($index + 1));
                $customErrors = $this->runRowValidators($preset, $row, $rowNumber);
                if ($customErrors !== []) {
                    $failed[] = ['row'=>$rowNumber,'errors'=>$customErrors,'data'=>$row];
                    unset($valid[$index]);
                }
            }
            $valid = array_values($valid);
            [$valid, $duplicateFailures, $skipped] = $this->applyDuplicatePolicy($valid, $uniqueBy, $fileDuplicatePolicy, $seen);
            $totals['skipped_duplicate_rows'] += $skipped;
            if ($duplicateFailures !== []) $failed = array_merge($failed, $duplicateFailures);

            if ($failed !== []) {
                $totals['failed_rows'] += count($failed);
                if ($failedWriter !== null) $failedWriter->append($failed);
                foreach ($failed as $failure) {
                    if (count($errors) < $maxErrors) $errors[] = $failure;
                    if (is_callable($rowErrorCallback)) $rowErrorCallback($failure);
                }
                EventDispatcher::safeDispatch('on_domain_import_failed_rows', ['count'=>count($failed)]);
                if ($strictValidation || !$skipInvalid) {
                    $first = $failed[0];
                    throw MnbExcelException::withCode('Domain import failed validation at source row ' . (int)($first['row'] ?? 0) . ': ' . implode('; ', (array)($first['errors'] ?? [])), ErrorCode::VALIDATION_FAILED);
                }
            }

            $clean = [];
            foreach ($valid as $row) {
                $sourceRow = (int) ($row['_mnb_excel_row'] ?? 0);
                unset($row['_mnb_excel_row']);
                $clean[] = $row;
                if ($collectRows && count($sampleRows) < $maxCollectedRows) $sampleRows[] = ['source_row'=>$sourceRow,'data'=>$row];
            }
            $totals['valid_rows'] += count($clean);
            $totals['planned_rows'] += count($clean);
            if ($clean !== [] && !$dryRun) {
                if (!$connection instanceof PDO) throw MnbExcelException::withCode('A PDO connection is required unless dry_run is enabled.', ErrorCode::DB_CONNECTION_FAILED);
                $startedTransaction = false;
                try {
                    if ($transactionPerBatch && !$connection->inTransaction()) $startedTransaction = $connection->beginTransaction();
                    $result = $this->sqlImporter->importRows($connection, $table, $clean, [
                        'batch_size'=>$batchSize,
                        'duplicate_strategy'=>$databaseDuplicateStrategy,
                        'unique_by'=>$uniqueBy,
                        'skip_invalid_rows'=>false,
                    ]);
                    if ($startedTransaction && $connection->inTransaction()) $connection->commit();
                    $totals['inserted_rows'] += (int) ($result['inserted_rows'] ?? 0);
                } catch (Throwable $e) {
                    if ($startedTransaction && $connection->inTransaction()) $connection->rollBack();
                    throw $e;
                }
            }
            $totals['batches']++;
            $pending = [];
            if (is_callable($progress)) $progress([
                'domain'=>$preset->type->value,'table'=>$table,'rows_scanned'=>$totals['rows_scanned'],'valid_rows'=>$totals['valid_rows'],
                'failed_rows'=>$totals['failed_rows'],'inserted_rows'=>$totals['inserted_rows'],'batches'=>$totals['batches'],
                'skipped_duplicate_rows'=>$totals['skipped_duplicate_rows'],'elapsed_seconds'=>microtime(true)-$startedAt,'completed'=>false,
            ]);
        };

        try {
            foreach ($session->rows($readerOptions) as $index => $row) {
                if (!is_array($row)) continue;
                $totals['rows_scanned']++;
                if ($mapping === null) {
                    $sourceColumns = array_values(array_filter(array_map('strval', array_keys($row)), static fn(string $c): bool => !str_starts_with($c, '_mnb_')));
                    $numericColumns = $sourceColumns === [] ? [] : array_map('strval', range(0, count($sourceColumns)-1));
                    if ($sourceColumns === $numericColumns) {
                        throw MnbExcelException::withCode('Domain imports require a header row. Enable auto header detection or provide header_row.', ErrorCode::VALIDATION_FAILED);
                    }
                    [$mapping,$mappingDetails] = $this->resolveMapping($sourceColumns, $preset, $aliases, $explicitMap, $mappingConfidence);
                    $missingRequired = $this->missingRequiredColumns($preset, $mapping, $defaults);
                    if ($strictMapping && $missingRequired !== []) {
                        throw MnbExcelException::withCode('Required domain columns could not be mapped: ' . implode(', ', $missingRequired) . '.', ErrorCode::VALIDATION_FAILED, ['source_columns'=>$sourceColumns,'mapping'=>$mapping]);
                    }
                }
                $sourceRow = isset($row['_mnb_original_row_number']) && is_numeric($row['_mnb_original_row_number']) ? (int)$row['_mnb_original_row_number'] : $index + 2;
                $normalized = $this->normalizeRow($row, $mapping, $preset, $defaults, $normalizers, $preserveUnmapped, $applyDefaultsToEmpty, $sourceRow);
                if ($transformers !== []) $normalized = RowTransformerPipeline::apply($normalized, $transformers, ['domain'=>$preset->type->value,'source_row'=>$sourceRow,'source_path'=>$path,'table'=>$table]);
                $normalized['_mnb_excel_row'] = $sourceRow;
                $pending[] = $normalized;
                if (count($pending) >= $batchSize) $flush();
            }
            $flush();
        } catch (Throwable $e) {
            EventDispatcher::safeDispatch('on_domain_import_failed', ['domain'=>$preset->type->value,'path'=>$path,'table'=>$table,'exception'=>$e]);
            LoggerBridge::error('Domain import failed.', ['domain'=>$preset->type->value,'path'=>$path,'table'=>$table,'message'=>$e->getMessage()]);
            throw $e;
        }

        $status = $dryRun ? ($totals['failed_rows'] > 0 ? 'dry_run_with_errors' : 'dry_run') : ($totals['failed_rows'] > 0 ? 'completed_with_errors' : 'completed');
        $result = [
            'status'=>$status,'domain'=>$preset->type->value,'description'=>$preset->description,'source_path'=>$path,'source_format'=>$format,
            'sheet'=>$options['sheet'] ?? 1,'table'=>$table,'dry_run'=>$dryRun,'source_columns'=>$sourceColumns,'mapping'=>$mapping ?? [],
            'mapping_details'=>$mappingDetails,'missing_required_columns'=>$missingRequired,'canonical_columns'=>$preset->columns(),
            'required_columns'=>$preset->requiredColumns(),'unique_by'=>$uniqueBy,'rows_scanned'=>$totals['rows_scanned'],
            'valid_rows'=>$totals['valid_rows'],'failed_rows'=>$totals['failed_rows'],'planned_rows'=>$totals['planned_rows'],
            'inserted_rows'=>$totals['inserted_rows'],'batches'=>$totals['batches'],'skipped_duplicate_rows'=>$totals['skipped_duplicate_rows'],
            'batch_size'=>$batchSize,'errors'=>$errors,'errors_truncated'=>$totals['failed_rows'] > count($errors),
            'failed_rows_csv'=>$failedRowsPath !== '' ? $failedRowsPath : null,'sample_rows'=>$sampleRows,
            'elapsed_seconds'=>round(microtime(true)-$startedAt, 6),
        ];
        EventDispatcher::safeDispatch('on_domain_import_completed', $result);
        LoggerBridge::info('Domain import completed.', array_intersect_key($result, array_flip(['status','domain','source_path','source_format','table','rows_scanned','valid_rows','failed_rows','inserted_rows','elapsed_seconds'])));
        if (is_callable($progress)) $progress($result + ['completed'=>true]);
        return $result;
    }

    public function importUsers(string $path, PDO|array|string|null $pdo = null, string $table = 'users', array $options = []): array { return $this->import(DomainImportType::Users,$path,$pdo,$table,$options); }
    public function importProducts(string $path, PDO|array|string|null $pdo = null, string $table = 'products', array $options = []): array { return $this->import(DomainImportType::Products,$path,$pdo,$table,$options); }
    public function importOrders(string $path, PDO|array|string|null $pdo = null, string $table = 'orders', array $options = []): array { return $this->import(DomainImportType::Orders,$path,$pdo,$table,$options); }
    public function importInventory(string $path, PDO|array|string|null $pdo = null, string $table = 'inventory', array $options = []): array { return $this->import(DomainImportType::Inventory,$path,$pdo,$table,$options); }
    public function importStudents(string $path, PDO|array|string|null $pdo = null, string $table = 'students', array $options = []): array { return $this->import(DomainImportType::Students,$path,$pdo,$table,$options); }
    public function importAttendance(string $path, PDO|array|string|null $pdo = null, string $table = 'attendance', array $options = []): array { return $this->import(DomainImportType::Attendance,$path,$pdo,$table,$options); }
    public function importMarks(string $path, PDO|array|string|null $pdo = null, string $table = 'marks', array $options = []): array { return $this->import(DomainImportType::Marks,$path,$pdo,$table,$options); }
    public function importContacts(string $path, PDO|array|string|null $pdo = null, string $table = 'contacts', array $options = []): array { return $this->import(DomainImportType::Contacts,$path,$pdo,$table,$options); }
    public function importLocations(string $path, PDO|array|string|null $pdo = null, string $table = 'locations', array $options = []): array { return $this->import(DomainImportType::Locations,$path,$pdo,$table,$options); }
    public function importBlogPosts(string $path, PDO|array|string|null $pdo = null, string $table = 'blog_posts', array $options = []): array { return $this->import(DomainImportType::BlogPosts,$path,$pdo,$table,$options); }
    public function importImagesWithPaths(string $path, PDO|array|string|null $pdo = null, string $table = 'media', array $options = []): array { return $this->import(DomainImportType::Media,$path,$pdo,$table,$options); }
    public function importMedia(string $path, PDO|array|string|null $pdo = null, string $table = 'media', array $options = []): array { return $this->importImagesWithPaths($path,$pdo,$table,$options); }
    public function importCategories(string $path, PDO|array|string|null $pdo = null, string $table = 'categories', array $options = []): array { return $this->import(DomainImportType::Categories,$path,$pdo,$table,$options); }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    private function readerOptions(array $options): array
    {
        $out = is_array($options['reader_options'] ?? null) ? $options['reader_options'] : [];
        foreach (['format','sheet','start_row','end_row','start_column','end_column','limit','offset','encoding','delimiter','enclosure','escape','xml_schema','row_element','stream_json_array','formula_cells'] as $key) {
            if (array_key_exists($key,$options)) $out[$key] = $options[$key];
        }
        if (!array_key_exists('header_row',$out) && !array_key_exists('header_row_mode',$out)) {
            $out['header_row'] = $options['header_row'] ?? 'auto';
            $out['header_row_mode'] = $options['header_row_mode'] ?? 'auto';
        }
        $out['header_detection_rows'] = max(1,(int)($options['header_detection_rows'] ?? $out['header_detection_rows'] ?? 25));
        $out['header_min_confidence'] = max(0.0,min(1.0,(float)($options['header_min_confidence'] ?? $out['header_min_confidence'] ?? 0.2)));
        $out['strict_header_detection'] = (bool)($options['strict_header_detection'] ?? false);
        $out['preserve_original_row_numbers'] = true;
        return $out;
    }

    /** @param list<string> $sourceColumns @param array<string,list<string>> $aliases @param array<string,mixed> $explicitMap @return array{0:array<string,string>,1:array<string,array<string,mixed>>} */
    private function resolveMapping(array $sourceColumns, DomainImportPreset $preset, array $aliases, array $explicitMap, float $minConfidence): array
    {
        $mapping = [];
        $details = [];
        $canonical = $preset->columns();
        foreach ($explicitMap as $left => $right) {
            $left = (string)$left; $right = (string)$right;
            if (in_array($left,$canonical,true)) { $source = $this->findSourceColumn($sourceColumns,$right); $target = $left; }
            else { $source = $this->findSourceColumn($sourceColumns,$left); $target = $right; }
            if ($source !== null && in_array($target,$canonical,true)) {
                $mapping[$source] = $target;
                $details[$source] = ['target'=>$target,'confidence'=>1.0,'reason'=>'explicit_map'];
            }
        }
        $remaining = array_values(array_filter($sourceColumns, static fn(string $s): bool => !isset($mapping[$s])));
        $suggestions = $this->quality->suggestColumnMap($remaining,$canonical,$aliases,$minConfidence);
        $claimed = array_values($mapping);
        foreach ($suggestions as $source => $suggestion) {
            $target = $suggestion['target'] ?? null;
            if (is_string($target) && !in_array($target,$claimed,true)) {
                $mapping[$source] = $target;
                $claimed[] = $target;
            }
            $details[$source] = $suggestion;
        }
        return [$mapping,$details];
    }

    /** @param array<string,string> $mapping @param array<string,mixed> $defaults @return list<string> */
    private function missingRequiredColumns(DomainImportPreset $preset, array $mapping, array $defaults): array
    {
        $mapped = array_values($mapping); $missing = [];
        foreach ($preset->requiredColumns() as $required) if (!in_array($required,$mapped,true) && !array_key_exists($required,$defaults)) $missing[] = $required;
        return $missing;
    }

    /** @param array<string,mixed> $row @param array<string,string> $mapping @param array<string,mixed> $defaults @param array<string,callable|string> $normalizers @return array<string,mixed> */
    private function normalizeRow(array $row, array $mapping, DomainImportPreset $preset, array $defaults, array $normalizers, bool $preserveUnmapped, bool $applyDefaultsToEmpty, int $sourceRow): array
    {
        $normalized = $defaults;
        foreach ($mapping as $source=>$target) $normalized[$target] = $row[$source] ?? null;
        if ($preserveUnmapped) foreach ($row as $source=>$value) if (!isset($mapping[(string)$source]) && $source !== '_mnb_original_row_number') $normalized[(string)$source] = $value;
        foreach ($preset->fields() as $field=>$definition) {
            if (!array_key_exists($field,$normalized)) $normalized[$field] = null;
            if ($applyDefaultsToEmpty && $this->isEmptyValue($normalized[$field]) && array_key_exists($field,$defaults)) $normalized[$field] = $defaults[$field];
            $normalizer = $normalizers[$field] ?? ($definition['normalizer'] ?? 'string');
            $normalized[$field] = $this->normalizeValue($normalized[$field],$normalizer,$field,$normalized,$sourceRow);
        }
        if (in_array($preset->type,[DomainImportType::Users,DomainImportType::Students,DomainImportType::Contacts],true) && trim((string)($normalized['name'] ?? '')) === '') {
            $normalized['name'] = trim((string)($normalized['first_name'] ?? '') . ' ' . (string)($normalized['last_name'] ?? ''));
        }
        if ($preset->type === DomainImportType::BlogPosts && trim((string)($normalized['slug'] ?? '')) === '') $normalized['slug'] = $this->slug((string)($normalized['title'] ?? ''));
        if ($preset->type === DomainImportType::Categories && trim((string)($normalized['slug'] ?? '')) === '') $normalized['slug'] = $this->slug((string)($normalized['name'] ?? ''));
        return $normalized;
    }

    private function normalizeValue(mixed $value, mixed $normalizer, string $field, array $row, int $sourceRow): mixed
    {
        if (!is_string($normalizer) && is_callable($normalizer)) return $normalizer($value,$row,['field'=>$field,'source_row'=>$sourceRow]);
        $type = strtolower(trim((string)$normalizer));
        if ($value === null) return null;
        if (is_string($value)) { $value = trim($value); if ($value === '') return null; }
        return match($type) {
            'email','lowercase' => $this->lower((string)$value),
            'uppercase' => $this->upper((string)$value),
            'integer' => $this->integer($value),
            'decimal','numeric' => $this->decimal($value),
            'boolean' => $this->boolean($value),
            'date' => $this->dateValue($value,false),
            'datetime' => $this->dateValue($value,true),
            'slug' => $this->slug((string)$value),
            'raw' => $value,
            default => is_scalar($value) ? trim((string)$value) : $value,
        };
    }

    private function integer(mixed $value): mixed
    {
        if (is_int($value)) return $value;
        if (is_float($value) && floor($value) === $value) return (int)$value;
        $text = str_replace([',',' '],'',(string)$value);
        return preg_match('/^[+-]?\d+$/',$text) === 1 ? (int)$text : $value;
    }
    private function decimal(mixed $value): mixed
    {
        if (is_int($value)||is_float($value)) return $value;
        $text = trim((string)$value);
        $text = preg_replace('/^[^\d+\-.]+|[^\d]+$/u','',$text) ?? $text;
        $text = str_replace([',',' '],'',$text);
        return is_numeric($text) ? (float)$text : $value;
    }
    private function boolean(mixed $value): mixed
    {
        if (is_bool($value)) return $value;
        $v = strtolower(trim((string)$value));
        if (in_array($v,['1','true','yes','y','active','enabled'],true)) return true;
        if (in_array($v,['0','false','no','n','inactive','disabled'],true)) return false;
        return $value;
    }
    private function dateValue(mixed $value, bool $withTime): mixed
    {
        if ($value instanceof \DateTimeInterface) return $value->format($withTime ? 'Y-m-d H:i:s' : 'Y-m-d');
        $timestamp = strtotime((string)$value);
        return $timestamp === false ? $value : date($withTime ? 'Y-m-d H:i:s' : 'Y-m-d',$timestamp);
    }
    private function slug(string $value): string
    {
        $value = trim($value); if ($value === '') return '';
        if (function_exists('iconv')) { $ascii = @iconv('UTF-8','ASCII//TRANSLIT//IGNORE',$value); if (is_string($ascii)&&$ascii!=='') $value=$ascii; }
        $value = $this->lower($value);
        $value = preg_replace('/[^\p{L}\p{N}]+/u','-',$value) ?? $value;
        return trim($value,'-');
    }
    private function isEmptyValue(mixed $value): bool { return $value === null || (is_string($value)&&trim($value)===''); }

    /** @return list<string> */
    private function runRowValidators(DomainImportPreset $preset, array $row, int $rowNumber): array
    {
        $errors=[];
        foreach ($preset->rowValidators() as $validator) foreach ((array)$validator($row,$rowNumber) as $error) { $error=trim((string)$error); if ($error!=='') $errors[]=$error; }
        return array_values(array_unique($errors));
    }

    /** @param list<array<string,mixed>> $rows @param list<string> $uniqueBy @param array<string,int> $seen @return array{0:list<array<string,mixed>>,1:list<array{row:int,errors:list<string>,data:array<string,mixed>}>,2:int} */
    private function applyDuplicatePolicy(array $rows, array $uniqueBy, string $policy, array &$seen): array
    {
        if ($uniqueBy === [] || $policy === 'allow') return [$rows,[],0];
        $valid=[]; $failed=[]; $skipped=0;
        foreach ($rows as $row) {
            $parts=[]; $hasValue=false;
            foreach ($uniqueBy as $column) { $value=$this->lower(trim((string)($row[$column] ?? ''))); $parts[]=$value; $hasValue=$hasValue||$value!==''; }
            if (!$hasValue) { $valid[]=$row; continue; }
            $key=implode("\x1F",$parts); $rowNumber=(int)($row['_mnb_excel_row'] ?? 0);
            if (isset($seen[$key])) {
                if ($policy==='skip') { $skipped++; continue; }
                $failed[]=['row'=>$rowNumber,'errors'=>['Duplicate domain record for unique columns: '.implode(', ',$uniqueBy).'. First seen at row '.$seen[$key].'.'],'data'=>$row];
                continue;
            }
            $seen[$key]=$rowNumber; $valid[]=$row;
        }
        return [$valid,$failed,$skipped];
    }

    /** @param array<string,list<string>> $base @param array<string,mixed> $overrides @return array<string,list<string>> */
    private function mergeAliases(array $base, array $overrides): array
    {
        foreach ($overrides as $target=>$values) $base[(string)$target]=array_values(array_unique(array_merge($base[(string)$target] ?? [],array_map('strval',(array)$values))));
        return $base;
    }
    /** @param list<string> $sourceColumns */
    private function findSourceColumn(array $sourceColumns, string $wanted): ?string
    {
        $needle=$this->normalizeHeader($wanted);
        foreach ($sourceColumns as $source) if ($this->normalizeHeader($source)===$needle) return $source;
        return null;
    }
    private function normalizeHeader(string $value): string { $value=$this->lower(trim($value)); return preg_replace('/[^\p{L}\p{N}]+/u','',$value) ?? $value; }
    private function lower(string $value): string { return function_exists('mb_strtolower') ? mb_strtolower($value,'UTF-8') : strtolower($value); }
    private function upper(string $value): string { return function_exists('mb_strtoupper') ? mb_strtoupper($value,'UTF-8') : strtoupper($value); }
    /** @return list<string> */
    private function stringList(mixed $values): array
    {
        if (is_string($values)) $values=array_map('trim',explode(',',$values));
        if (!is_array($values)) return [];
        return array_values(array_unique(array_filter(array_map(static fn(mixed $v): string=>trim((string)$v),$values),static fn(string $v): bool=>$v!=='')));
    }
}
