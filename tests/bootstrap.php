<?php

declare(strict_types=1);

/**
 * Shared bootstrap for MNB PHPExcel smoke tests.
 *
 * Each *SmokeTest.php file is executed as a standalone script by
 * tools/run-smoke-tests.php (php <file>, checked by exit code), so every
 * test file requires this bootstrap directly:
 *
 *     require __DIR__ . '/bootstrap.php';
 *
 * Failures throw SmokeTestFailure, which is caught here and turned into a
 * non-zero exit code plus a readable message on STDERR.
 */

$composerAutoload = __DIR__ . '/../vendor/autoload.php';
if (is_file($composerAutoload)) {
    require $composerAutoload;
} else {
    spl_autoload_register(static function (string $class): void {
        $prefix = 'Mnb\\PHPExcel\\';
        if (!str_starts_with($class, $prefix)) {
            return;
        }

        $relative = substr($class, strlen($prefix));
        $path = __DIR__ . '/../src/' . str_replace('\\', DIRECTORY_SEPARATOR, $relative) . '.php';
        if (is_file($path)) {
            require $path;
        }
    });
}

final class SmokeTestFailure extends \RuntimeException
{
}

/**
 * @param mixed $actual
 * @param mixed $expected
 */
function smoke_assert_equals($expected, $actual, string $message): void
{
    if ($expected !== $actual) {
        throw new SmokeTestFailure(sprintf(
            '%s (expected %s, got %s)',
            $message,
            var_export($expected, true),
            var_export($actual, true)
        ));
    }
}

function smoke_assert_true(bool $condition, string $message): void
{
    if (!$condition) {
        throw new SmokeTestFailure($message);
    }
}

function smoke_assert_contains(string $needle, string $haystack, string $message): void
{
    if (!str_contains($haystack, $needle)) {
        throw new SmokeTestFailure($message . ' (looking for "' . $needle . '")');
    }
}

/**
 * Creates a fresh temp directory for a test run and returns its path.
 * Registers automatic cleanup on shutdown so smoke tests do not leak files.
 */
function smoke_temp_dir(string $prefix): string
{
    $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb_phpexcel_smoke_' . $prefix . '_' . bin2hex(random_bytes(4));
    if (!mkdir($dir, 0777, true) && !is_dir($dir)) {
        throw new SmokeTestFailure('Unable to create temp directory: ' . $dir);
    }

    register_shutdown_function(static function () use ($dir): void {
        smoke_remove_directory($dir);
    });

    return $dir;
}

function smoke_remove_directory(string $dir): void
{
    if (!is_dir($dir)) {
        return;
    }

    $items = scandir($dir) ?: [];
    foreach ($items as $item) {
        if ($item === '.' || $item === '..') {
            continue;
        }
        $path = $dir . DIRECTORY_SEPARATOR . $item;
        if (is_dir($path)) {
            smoke_remove_directory($path);
        } else {
            @unlink($path);
        }
    }

    @rmdir($dir);
}

/**
 * Runs a smoke test's body and standardizes pass/fail reporting + exit code.
 */
function smoke_run(string $name, callable $body): void
{
    echo '  - ' . $name . '... ';
    try {
        $body();
        echo "OK\n";
    } catch (\Throwable $e) {
        echo "FAILED\n";
        fwrite(STDERR, '    ' . get_class($e) . ': ' . $e->getMessage() . "\n");
        exit(1);
    }
}
