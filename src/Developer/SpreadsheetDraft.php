<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Developer;

use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Simple fluent workbook facade. Advanced users can call builder(). */
final class SpreadsheetDraft
{
    /** @var array<string,array<int|string,mixed>> */
    private array $sheets = [];

    /** @param array<int|string,mixed> $rows */
    public function sheet(string $name, array $rows): self
    {
        $name = trim($name);
        if ($name === '') {
            throw new MnbExcelException('Sheet name cannot be empty.');
        }
        $clone = clone $this;
        $clone->sheets[$name] = $rows;
        return $clone;
    }

    /** @param array<int|string,mixed> $rows */
    public function addSheet(string $name, array $rows): self
    {
        return $this->sheet($name, $rows);
    }

    public function builder(): WorkbookBuilder
    {
        return $this->sheets === []
            ? WorkbookBuilder::fromArray([])
            : WorkbookBuilder::fromWorkbookArray($this->sheets);
    }

    public function save(string $path): string
    {
        return $this->builder()->save($path);
    }
}
