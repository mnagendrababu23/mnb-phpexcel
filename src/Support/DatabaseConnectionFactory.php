<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use PDO;
use Throwable;

final class DatabaseConnectionFactory
{
    private function __construct()
    {
    }

    /**
     * Create or return a PDO connection from an existing PDO, .env path, PHP config file,
     * DSN string, config array, constants, or runtime environment variables.
     *
     * @param PDO|array<string,mixed>|string|null $source
     * @param array<string,mixed> $overrides
     */
    public static function make(PDO|array|string|null $source = null, array $overrides = []): PDO
    {
        if ($source instanceof PDO) {
            self::applyDefaultAttributes($source);
            return $source;
        }

        if (!class_exists(PDO::class)) {
            throw MnbExcelException::withCode('PDO extension is not enabled.', ErrorCode::EXTENSION_MISSING, ['extension' => 'pdo']);
        }

        $config = DatabaseConfigResolver::resolve($source, $overrides);
        $options = self::pdoOptions($config);

        try {
            $pdo = new PDO(
                (string) $config['dsn'],
                $config['username'] ?? null,
                $config['password'] ?? null,
                $options
            );
            self::applyDefaultAttributes($pdo);
            return $pdo;
        } catch (Throwable $e) {
            throw MnbExcelException::withCode(
                'Database connection failed: ' . $e->getMessage(),
                ErrorCode::DB_CONNECTION_FAILED,
                ['driver' => $config['driver'] ?? null, 'dsn' => self::safeDsn((string) ($config['dsn'] ?? ''))],
                $e
            );
        }
    }

    /**
     * @param array<string,mixed> $config
     * @return array<int,mixed>
     */
    private static function pdoOptions(array $config): array
    {
        $options = is_array($config['options'] ?? null) ? $config['options'] : [];
        $options[PDO::ATTR_ERRMODE] = $options[PDO::ATTR_ERRMODE] ?? PDO::ERRMODE_EXCEPTION;
        $options[PDO::ATTR_DEFAULT_FETCH_MODE] = $options[PDO::ATTR_DEFAULT_FETCH_MODE] ?? PDO::FETCH_ASSOC;

        if (defined('PDO::ATTR_EMULATE_PREPARES') && !array_key_exists(PDO::ATTR_EMULATE_PREPARES, $options)) {
            $options[PDO::ATTR_EMULATE_PREPARES] = false;
        }

        return $options;
    }

    private static function applyDefaultAttributes(PDO $pdo): void
    {
        try {
            $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            $pdo->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_ASSOC);
        } catch (Throwable) {
            // Some PDO drivers can reject attributes after connection. Keep the existing PDO usable.
        }
    }

    private static function safeDsn(string $dsn): string
    {
        return preg_replace('/(password|pwd)=([^;]+)/i', '$1=***', $dsn) ?? $dsn;
    }
}
