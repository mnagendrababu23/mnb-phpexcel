<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Small JSON manifest for resumable large imports.
 *
 * The manifest is intentionally framework-neutral, so it works in plain PHP,
 * cron, Symfony Console, Laravel queues, or shared-hosting CLI jobs.
 */
final class LargeImportManifest
{
    /** @var array<string,mixed> */
    private array $data = [];

    public function __construct(private string $path)
    {
    }

    public static function defaultPath(string $sourcePath, string $table = 'import', int|string $sheet = 1): string
    {
        $base = sha1((realpath($sourcePath) ?: $sourcePath) . '|' . $table . '|' . (string) $sheet);
        return rtrim(sys_get_temp_dir(), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . 'mnb-large-import-' . $base . '.json';
    }

    public function path(): string
    {
        return $this->path;
    }

    /** @return array<string,mixed> */
    public function load(): array
    {
        if (!is_file($this->path)) {
            $this->data = [];
            return [];
        }

        $json = @file_get_contents($this->path);
        if ($json === false) {
            throw MnbExcelException::withCode('Unable to read large import manifest: ' . $this->path, ErrorCode::FILE_READ_FAILED, ['path' => $this->path]);
        }

        $decoded = json_decode($json, true);
        if (!is_array($decoded)) {
            throw MnbExcelException::withCode('Large import manifest is invalid JSON: ' . $this->path, ErrorCode::JSON_INVALID, ['path' => $this->path]);
        }

        $this->data = $decoded;
        return $this->data;
    }

    /** @param array<string,mixed> $data */
    public function start(array $data): void
    {
        $now = gmdate('c');
        $this->data = array_merge([
            'status' => 'running',
            'created_at' => $now,
            'updated_at' => $now,
            'runs' => 0,
            'chunks_completed' => 0,
            'rows_scanned' => 0,
            'valid_rows' => 0,
            'failed_rows' => 0,
            'inserted_rows' => 0,
            'last_row_number' => 0,
            'last_stop_reason' => null,
        ], $data);
        $this->save();
    }

    /** @param array<string,mixed> $patch */
    public function update(array $patch): void
    {
        if ($this->data === []) {
            $this->load();
        }
        $this->data = array_merge($this->data, $patch, ['updated_at' => gmdate('c')]);
        $this->save();
    }

    /** @return array<string,mixed> */
    public function data(): array
    {
        if ($this->data === []) {
            $this->load();
        }
        return $this->data;
    }

    private function save(): void
    {
        $json = json_encode($this->data, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
        if ($json === false) {
            throw MnbExcelException::withCode('Unable to encode large import manifest.', ErrorCode::JSON_ENCODE_FAILED, ['path' => $this->path]);
        }

        AtomicFileWriter::writeViaTemp($this->path, static function (string $tmp) use ($json): void {
            if (@file_put_contents($tmp, $json . PHP_EOL, LOCK_EX) === false) {
                throw MnbExcelException::withCode('Unable to write large import manifest.', ErrorCode::FILE_WRITE_FAILED);
            }
        });
    }
}
