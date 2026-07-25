<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Support\DatabaseConnectionFactory;

final class DatabaseIntegrationSuite
{
    /** @return array<string,mixed> */
    public static function plan(array $options = []): array
    {
        return [
            'status' => 'ready',
            'drivers' => [
                'mysql' => [
                    'env_prefix' => 'MNB_EXCEL_TEST_MYSQL_',
                    'required_keys' => ['DSN', 'USERNAME', 'PASSWORD'],
                    'example' => [
                        'MNB_EXCEL_TEST_MYSQL_DSN=mysql:host=127.0.0.1;port=3306;dbname=mnb_excel_test;charset=utf8mb4',
                        'MNB_EXCEL_TEST_MYSQL_USERNAME=root',
                        'MNB_EXCEL_TEST_MYSQL_PASSWORD=secret',
                    ],
                ],
                'pgsql' => [
                    'env_prefix' => 'MNB_EXCEL_TEST_PGSQL_',
                    'required_keys' => ['DSN', 'USERNAME', 'PASSWORD'],
                    'example' => [
                        'MNB_EXCEL_TEST_PGSQL_DSN=pgsql:host=127.0.0.1;port=5432;dbname=mnb_excel_test',
                        'MNB_EXCEL_TEST_PGSQL_USERNAME=postgres',
                        'MNB_EXCEL_TEST_PGSQL_PASSWORD=secret',
                    ],
                ],
            ],
            'checks' => ['connect', 'create_temp_table', 'large_import_small_fixture', 'duplicate_skip', 'duplicate_update', 'drop_temp_table'],
            'note' => 'Integration tests are skipped unless driver-specific environment variables are configured.',
        ];
    }

    /** @return array<string,mixed> */
    public static function check(array $options = []): array
    {
        $drivers = (array) ($options['drivers'] ?? ['mysql', 'pgsql']);
        $results = [];
        $summary = ['available' => 0, 'skipped' => 0, 'failed' => 0];

        foreach ($drivers as $driver) {
            $driver = strtolower((string) $driver);
            $prefix = $driver === 'pgsql' ? 'MNB_EXCEL_TEST_PGSQL_' : 'MNB_EXCEL_TEST_MYSQL_';
            $dsn = getenv($prefix . 'DSN') ?: null;
            $username = getenv($prefix . 'USERNAME') ?: '';
            $password = getenv($prefix . 'PASSWORD') ?: '';

            if (!$dsn) {
                $summary['skipped']++;
                $results[$driver] = [
                    'status' => 'skipped',
                    'reason' => $prefix . 'DSN is not configured.',
                    'required_env_prefix' => $prefix,
                ];
                continue;
            }

            try {
                $pdo = DatabaseConnectionFactory::make([
                    'dsn' => $dsn,
                    'username' => $username,
                    'password' => $password,
                ]);
                $results[$driver] = [
                    'status' => 'available',
                    'dsn_driver' => $pdo->getAttribute(\PDO::ATTR_DRIVER_NAME),
                    'server_version' => (string) $pdo->getAttribute(\PDO::ATTR_SERVER_VERSION),
                ];
                $summary['available']++;
            } catch (\Throwable $e) {
                $summary['failed']++;
                $results[$driver] = [
                    'status' => 'failed',
                    'error' => $e->getMessage(),
                ];
            }
        }

        return [
            'status' => $summary['failed'] > 0 ? 'failed' : ($summary['available'] > 0 ? 'available' : 'skipped'),
            'summary' => $summary,
            'results' => $results,
            'plan' => self::plan(),
        ];
    }
}
