<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;

final class LargePdoCursor
{
    /** @param array<int|string,mixed> $params @param array<string,mixed> $options @return \Generator<int,array<string,mixed>> */
    public static function rows(PDO $pdo, string $query, array $params = [], array $options = []): \Generator
    {
        try {
            $stmt = $pdo->prepare($query);
            if ($stmt === false) {
                throw MnbExcelException::withCode('Unable to prepare PDO cursor export query.', ErrorCode::SQL_EXPORT_FAILED);
            }
            foreach ($params as $key => $value) {
                $param = is_int($key) ? $key + 1 : (str_starts_with((string) $key, ':') ? (string) $key : ':' . (string) $key);
                $stmt->bindValue($param, $value);
            }
            $stmt->execute();
            while (($row = $stmt->fetch(PDO::FETCH_ASSOC)) !== false) {
                yield $row;
            }
            $stmt->closeCursor();
        } catch (Throwable $e) {
            if ($e instanceof MnbExcelException) {
                throw $e;
            }
            throw MnbExcelException::withCode('PDO cursor export failed: ' . $e->getMessage(), ErrorCode::SQL_EXPORT_FAILED, [], $e);
        }
    }
}
