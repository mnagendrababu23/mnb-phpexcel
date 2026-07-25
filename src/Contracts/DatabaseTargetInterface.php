<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Contracts;

interface DatabaseTargetInterface
{
    /** @param list<array<string,mixed>> $rows @param array<string,mixed> $options */
    public function insertRows(array $rows, array $options = []): int;
}
