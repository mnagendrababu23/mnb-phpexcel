<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$autoload = $root . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';
if (is_file($autoload)) {
    require $autoload;
} else {
    spl_autoload_register(static function (string $class) use ($root): void {
        $prefix = 'Mnb\\PHPExcel\\';
        if (!str_starts_with($class, $prefix)) {
            return;
        }
        $relative = substr($class, strlen($prefix));
        $path = $root . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR
            . str_replace('\\', DIRECTORY_SEPARATOR, $relative) . '.php';
        if (is_file($path)) {
            require $path;
        }
    });
}

$result = Mnb\PHPExcel\MnbExcel::releaseReadiness($root);
echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES), PHP_EOL;
exit(($result['status'] ?? 'failed') === 'failed' ? 1 : 0);
