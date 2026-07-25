<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

/**
 * Resolves database connection details from common application sources without
 * requiring a framework or dotenv dependency.
 */
final class DatabaseConfigResolver
{
    /** @var list<string> */
    private const DEFAULT_ENV_NAMES = ['.env', '.env.local'];

    private function __construct()
    {
    }

    /**
     * Resolve database configuration from an array, .env file, PHP config file,
     * DSN string, constants, getenv()/$_ENV, or the nearest app .env.
     *
     * @param array<string,mixed>|string|null $source
     * @param array<string,mixed> $overrides
     * @return array<string,mixed>
     */
    public static function resolve(array|string|null $source = null, array $overrides = []): array
    {
        $raw = [];

        if (is_array($source)) {
            $raw = self::normalizeArrayKeys($source);
            $file = $raw['env_path'] ?? $raw['config_path'] ?? $raw['file'] ?? null;
            if (is_string($file) && $file !== '') {
                $raw = array_merge(self::readFileConfig($file), $raw);
            }
        } elseif (is_string($source) && trim($source) !== '') {
            $source = trim($source);
            if (is_file($source)) {
                $raw = self::readFileConfig($source);
            } elseif (self::looksLikeDsn($source)) {
                $raw = ['dsn' => $source];
            } else {
                // Treat non-file strings as a DSN-like value. Validation later will produce a clear error.
                $raw = ['dsn' => $source];
            }
        } else {
            $raw = self::autoEnvironmentConfig($overrides);
        }

        $raw = array_merge(self::constantConfig(), self::runtimeEnvConfig(), $raw, self::normalizeArrayKeys($overrides));
        $raw = self::normalizeAliases($raw);

        return self::buildConfig($raw);
    }

    /**
     * Return all candidate .env/config paths checked during auto discovery.
     *
     * @param array<string,mixed> $options
     * @return list<string>
     */
    public static function discoveredPaths(array $options = []): array
    {
        $paths = [];
        $explicit = $options['env_path'] ?? $options['config_path'] ?? null;
        if (is_string($explicit) && $explicit !== '') {
            return [$explicit];
        }

        $base = (string) ($options['base_path'] ?? getcwd() ?: '.');
        $base = realpath($base) ?: $base;
        $names = $options['env_names'] ?? self::DEFAULT_ENV_NAMES;
        if (is_string($names)) {
            $names = [$names];
        }
        if (!is_array($names)) {
            $names = self::DEFAULT_ENV_NAMES;
        }

        $current = $base;
        for ($i = 0; $i < 6; $i++) {
            foreach ($names as $name) {
                $name = trim((string) $name);
                if ($name !== '') {
                    $paths[] = rtrim($current, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $name;
                }
            }
            foreach (['config/database.php', 'database.php'] as $relative) {
                $paths[] = rtrim($current, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . str_replace('/', DIRECTORY_SEPARATOR, $relative);
            }
            $parent = dirname($current);
            if ($parent === $current || $parent === '.') {
                break;
            }
            $current = $parent;
        }

        return array_values(array_unique($paths));
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function summary(array $options = []): array
    {
        $paths = self::discoveredPaths($options);
        $existing = array_values(array_filter($paths, static fn(string $path): bool => is_file($path)));

        return [
            'status' => 'ok',
            'checked_paths' => $paths,
            'existing_paths' => $existing,
            'has_runtime_env' => self::runtimeEnvConfig() !== [],
            'has_db_constants' => self::constantConfig() !== [],
        ];
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    private static function autoEnvironmentConfig(array $options = []): array
    {
        foreach (self::discoveredPaths($options) as $path) {
            if (is_file($path)) {
                return self::readFileConfig($path);
            }
        }
        return [];
    }

    /** @return array<string,mixed> */
    private static function readFileConfig(string $path): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('Database config file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        if ($extension === 'php') {
            /** @var mixed $config */
            $config = require $path;
            if (!is_array($config)) {
                throw MnbExcelException::withCode('Database PHP config file must return an array: ' . $path, ErrorCode::DB_CONFIG_INVALID, ['path' => $path]);
            }
            return self::normalizeArrayKeys($config);
        }

        return self::parseEnvFile($path);
    }

    /** @return array<string,mixed> */
    private static function parseEnvFile(string $path): array
    {
        $lines = @file($path, FILE_IGNORE_NEW_LINES);
        if ($lines === false) {
            throw MnbExcelException::withCode('Unable to read database .env file: ' . $path, ErrorCode::FILE_READ_FAILED, ['path' => $path]);
        }

        $env = [];
        foreach ($lines as $line) {
            $line = trim($line);
            if ($line === '' || str_starts_with($line, '#')) {
                continue;
            }
            if (str_starts_with($line, 'export ')) {
                $line = trim(substr($line, 7));
            }
            $pos = strpos($line, '=');
            if ($pos === false) {
                continue;
            }
            $key = trim(substr($line, 0, $pos));
            $value = trim(substr($line, $pos + 1));
            if ($key === '') {
                continue;
            }
            $env[$key] = self::unquoteEnvValue($value);
        }

        return self::normalizeArrayKeys($env);
    }

    private static function unquoteEnvValue(string $value): string
    {
        if ($value === '') {
            return '';
        }
        $quote = $value[0];
        if (($quote === '"' || $quote === "'") && str_ends_with($value, $quote)) {
            $value = substr($value, 1, -1);
            if ($quote === '"') {
                $value = stripcslashes($value);
            }
        } else {
            $hash = strpos($value, ' #');
            if ($hash !== false) {
                $value = rtrim(substr($value, 0, $hash));
            }
        }
        return $value;
    }

    /** @return array<string,mixed> */
    private static function runtimeEnvConfig(): array
    {
        $keys = [
            'DB_DSN', 'DB_CONNECTION', 'DB_DRIVER', 'DB_HOST', 'DB_PORT', 'DB_DATABASE', 'DB_USERNAME', 'DB_USER', 'DB_PASSWORD', 'DB_PASS', 'DB_CHARSET', 'DB_SOCKET',
            'MNB_DB_DSN', 'MNB_DB_CONNECTION', 'MNB_DB_DRIVER', 'MNB_DB_HOST', 'MNB_DB_PORT', 'MNB_DB_DATABASE', 'MNB_DB_USERNAME', 'MNB_DB_USER', 'MNB_DB_PASSWORD', 'MNB_DB_PASS', 'MNB_DB_CHARSET', 'MNB_DB_SOCKET',
            'MNB_EXCEL_DB_DSN', 'MNB_EXCEL_DB_CONNECTION', 'MNB_EXCEL_DB_DRIVER', 'MNB_EXCEL_DB_HOST', 'MNB_EXCEL_DB_PORT', 'MNB_EXCEL_DB_DATABASE', 'MNB_EXCEL_DB_USERNAME', 'MNB_EXCEL_DB_USER', 'MNB_EXCEL_DB_PASSWORD', 'MNB_EXCEL_DB_PASS', 'MNB_EXCEL_DB_CHARSET', 'MNB_EXCEL_DB_SOCKET',
        ];

        $raw = [];
        foreach ($keys as $key) {
            $value = $_ENV[$key] ?? $_SERVER[$key] ?? getenv($key);
            if ($value !== false && $value !== null) {
                $raw[$key] = $value;
            }
        }
        return self::normalizeArrayKeys($raw);
    }

    /** @return array<string,mixed> */
    private static function constantConfig(): array
    {
        $map = [
            'MNB_EXCEL_DB_DSN' => 'dsn',
            'MNB_EXCEL_DB_CONNECTION' => 'connection',
            'MNB_EXCEL_DB_DRIVER' => 'driver',
            'MNB_EXCEL_DB_HOST' => 'host',
            'MNB_EXCEL_DB_PORT' => 'port',
            'MNB_EXCEL_DB_DATABASE' => 'database',
            'MNB_EXCEL_DB_USERNAME' => 'username',
            'MNB_EXCEL_DB_USER' => 'username',
            'MNB_EXCEL_DB_PASSWORD' => 'password',
            'MNB_EXCEL_DB_PASS' => 'password',
            'MNB_EXCEL_DB_CHARSET' => 'charset',
            'MNB_EXCEL_DB_SOCKET' => 'socket',
            'DB_DSN' => 'dsn',
            'DB_CONNECTION' => 'connection',
            'DB_DRIVER' => 'driver',
            'DB_HOST' => 'host',
            'DB_PORT' => 'port',
            'DB_DATABASE' => 'database',
            'DB_USERNAME' => 'username',
            'DB_USER' => 'username',
            'DB_PASSWORD' => 'password',
            'DB_PASS' => 'password',
            'DB_CHARSET' => 'charset',
            'DB_SOCKET' => 'socket',
        ];

        $raw = [];
        foreach ($map as $constant => $key) {
            if (defined($constant)) {
                $raw[$key] = constant($constant);
            }
        }
        return $raw;
    }

    /** @param array<string,mixed> $input @return array<string,mixed> */
    private static function normalizeArrayKeys(array $input): array
    {
        $normalized = [];
        foreach ($input as $key => $value) {
            if (is_int($key)) {
                continue;
            }
            $normalized[strtolower((string) $key)] = $value;
        }
        return $normalized;
    }

    /** @param array<string,mixed> $raw @return array<string,mixed> */
    private static function normalizeAliases(array $raw): array
    {
        $aliases = [
            'db_dsn' => 'dsn', 'mnb_db_dsn' => 'dsn', 'mnb_excel_db_dsn' => 'dsn',
            'db_connection' => 'connection', 'db_driver' => 'driver', 'mnb_db_connection' => 'connection', 'mnb_db_driver' => 'driver', 'mnb_excel_db_connection' => 'connection', 'mnb_excel_db_driver' => 'driver',
            'db_host' => 'host', 'mnb_db_host' => 'host', 'mnb_excel_db_host' => 'host',
            'db_port' => 'port', 'mnb_db_port' => 'port', 'mnb_excel_db_port' => 'port',
            'db_database' => 'database', 'database_name' => 'database', 'dbname' => 'database', 'mnb_db_database' => 'database', 'mnb_excel_db_database' => 'database',
            'db_username' => 'username', 'db_user' => 'username', 'user' => 'username', 'mnb_db_username' => 'username', 'mnb_db_user' => 'username', 'mnb_excel_db_username' => 'username', 'mnb_excel_db_user' => 'username',
            'db_password' => 'password', 'db_pass' => 'password', 'pass' => 'password', 'mnb_db_password' => 'password', 'mnb_db_pass' => 'password', 'mnb_excel_db_password' => 'password', 'mnb_excel_db_pass' => 'password',
            'db_charset' => 'charset', 'mnb_db_charset' => 'charset', 'mnb_excel_db_charset' => 'charset',
            'db_socket' => 'socket', 'mnb_db_socket' => 'socket', 'mnb_excel_db_socket' => 'socket',
        ];

        foreach ($aliases as $from => $to) {
            if (array_key_exists($from, $raw) && !array_key_exists($to, $raw)) {
                $raw[$to] = $raw[$from];
            }
        }

        if (!isset($raw['driver']) && isset($raw['connection'])) {
            $raw['driver'] = $raw['connection'];
        }

        return $raw;
    }

    /** @param array<string,mixed> $raw @return array<string,mixed> */
    private static function buildConfig(array $raw): array
    {
        $driver = strtolower(trim((string) ($raw['driver'] ?? '')));
        $dsn = trim((string) ($raw['dsn'] ?? ''));
        $database = (string) ($raw['database'] ?? '');

        if ($dsn === '' && $driver === '' && $database !== '' && (str_contains($database, DIRECTORY_SEPARATOR) || $database === ':memory:' || str_ends_with(strtolower($database), '.sqlite'))) {
            $driver = 'sqlite';
        }
        if ($driver === '') {
            $driver = $dsn !== '' ? strtolower(strtok($dsn, ':') ?: 'mysql') : 'mysql';
        }
        if ($driver === 'mariadb') {
            $driver = 'mysql';
        }

        if ($dsn === '') {
            $dsn = self::buildDsn($driver, $raw);
        }
        if ($dsn === '') {
            throw MnbExcelException::withCode('Database connection details are incomplete. Provide DB_DSN or DB_CONNECTION/DB_HOST/DB_DATABASE.', ErrorCode::DB_CONFIG_INVALID, ['driver' => $driver]);
        }

        $username = $raw['username'] ?? null;
        $password = $raw['password'] ?? null;
        $options = is_array($raw['options'] ?? null) ? $raw['options'] : [];

        return [
            'driver' => $driver,
            'dsn' => $dsn,
            'username' => $username === null ? null : (string) $username,
            'password' => $password === null ? null : (string) $password,
            'options' => $options,
            'database' => $database,
            'host' => isset($raw['host']) ? (string) $raw['host'] : null,
            'port' => isset($raw['port']) ? (int) $raw['port'] : null,
            'charset' => isset($raw['charset']) ? (string) $raw['charset'] : null,
            'source' => isset($raw['file']) ? (string) $raw['file'] : null,
        ];
    }

    /** @param array<string,mixed> $raw */
    private static function buildDsn(string $driver, array $raw): string
    {
        $database = (string) ($raw['database'] ?? '');
        $host = (string) ($raw['host'] ?? '127.0.0.1');
        $charset = (string) ($raw['charset'] ?? ($driver === 'pgsql' ? 'utf8' : 'utf8mb4'));
        $port = trim((string) ($raw['port'] ?? ''));
        $socket = trim((string) ($raw['socket'] ?? ''));

        return match ($driver) {
            'sqlite' => 'sqlite:' . ($database !== '' ? $database : ':memory:'),
            'mysql' => 'mysql:' . ($socket !== '' ? 'unix_socket=' . $socket : 'host=' . $host) . ($port !== '' ? ';port=' . $port : '') . ($database !== '' ? ';dbname=' . $database : '') . ';charset=' . $charset,
            'pgsql' => 'pgsql:host=' . $host . ($port !== '' ? ';port=' . $port : '') . ($database !== '' ? ';dbname=' . $database : ''),
            'sqlsrv' => 'sqlsrv:Server=' . $host . ($port !== '' ? ',' . $port : '') . ($database !== '' ? ';Database=' . $database : ''),
            default => '',
        };
    }

    private static function looksLikeDsn(string $value): bool
    {
        return preg_match('/^[A-Za-z0-9_]+:/', $value) === 1;
    }
}
