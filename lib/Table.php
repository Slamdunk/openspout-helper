<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use Countable;
use OpenSpout\Writer\AutoFilter;
use OpenSpout\Writer\Common\Entity\Sheet;

final class Table implements Countable
{
    private Sheet $activeSheet;

    /** @var 0|positive-int */
    private int $rowEnd;

    /** @var 0|positive-int */
    private int $rowCurrent;

    /** @var 0|positive-int */
    private int $rowStart;

    /** @var null|0|positive-int */
    private ?int $dataRowStart = null;

    /** @var 0|positive-int */
    private int $columnStart   = 0;

    /** @var 0|positive-int */
    private int $columnEnd     = 0;

    /** @var 0|positive-int */
    private int $columnCurrent = 0;

    private string $heading;

    /** @var iterable<int, array<string, null|float|int|string>> */
    private iterable $data;
    private ColumnCollection $columnCollection;
    private bool $freezePanes = true;
    private int $fontSize     = 8;
    private ?int $rowHeight   = null;
    private bool $textWrap    = false;

    /** @var array<int, string> */
    private array $writtenColumn       = [];
    private ?int $count                = null;

    /** @param iterable<int, array<string, null|float|int|string>> $data */
    public function __construct(Sheet $activeSheet, string $heading, iterable $data)
    {
        $this->activeSheet = $activeSheet;
        $this->heading     = $heading;
        $this->data        = $data;

        $this->columnCollection = new ColumnCollection();

        $this->rowStart   =
        $this->rowEnd     =
        $this->rowCurrent = $activeSheet->getWrittenRowCount();
    }

    public function getActiveSheet(): Sheet
    {
        return $this->activeSheet;
    }

    /** @return 0|positive-int */
    public function getDataRowStart(): int
    {
        \assert(null !== $this->dataRowStart);

        return $this->dataRowStart;
    }

    public function flagDataRowStart(): void
    {
        $this->dataRowStart = $this->rowCurrent;
    }

    /** @return 0|positive-int */
    public function getRowStart(): int
    {
        return $this->rowStart;
    }

    /** @return 0|positive-int */
    public function getRowEnd(): int
    {
        return $this->rowEnd;
    }

    /** @return 0|positive-int */
    public function getRowCurrent(): int
    {
        return $this->rowCurrent;
    }

    public function incrementRow(): void
    {
        $this->rowEnd = \max($this->rowEnd, $this->rowCurrent);
        ++$this->rowCurrent;
    }

    /** @return 0|positive-int */
    public function getColumnStart(): int
    {
        return $this->columnStart;
    }

    /** @return 0|positive-int */
    public function getColumnEnd(): int
    {
        return $this->columnEnd;
    }

    public function incrementColumn(): void
    {
        $this->columnEnd = \max($this->columnEnd, $this->columnCurrent);
        ++$this->columnCurrent;
    }

    public function getHeading(): string
    {
        return $this->heading;
    }

    /** @return iterable<int, array<string, null|float|int|string>> */
    public function getData(): iterable
    {
        return $this->data;
    }

    public function setColumnCollection(ColumnCollection $columnCollection): void
    {
        $this->columnCollection = $columnCollection;
    }

    public function getColumnCollection(): ColumnCollection
    {
        return $this->columnCollection;
    }

    public function setFreezePanes(bool $freezePanes): void
    {
        $this->freezePanes = $freezePanes;
    }

    public function getFreezePanes(): bool
    {
        return $this->freezePanes;
    }

    public function setFontSize(int $fontSize): void
    {
        $this->fontSize = $fontSize;
    }

    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    public function setRowHeight(?int $rowHeight): void
    {
        $this->rowHeight = $rowHeight;
    }

    public function getRowHeight(): ?int
    {
        return $this->rowHeight;
    }

    public function setTextWrap(bool $textWrap): void
    {
        $this->textWrap = $textWrap;
    }

    public function getTextWrap(): bool
    {
        return $this->textWrap;
    }

    /** @param array<int, string> $writtenColumn */
    public function setWrittenColumn(array $writtenColumn): void
    {
        $this->writtenColumn = $writtenColumn;
    }

    /** @return array<int, string> */
    public function getWrittenColumn(): array
    {
        return $this->writtenColumn;
    }

    public function setCount(int $count): void
    {
        $this->count = $count;
    }

    public function count(): int
    {
        if (null === $this->count) {
            throw new Exception(\sprintf('%s::setCount() have not been called yet', __CLASS__));
        }

        return $this->count;
    }

    public function isEmpty(): bool
    {
        return 0 === $this->count();
    }

    public function splitTableOnNewWorksheet(Sheet $sheet): self
    {
        $newTable = new self(
            $sheet,
            $this->getHeading(),
            $this->getData()
        );
        $newTable->setColumnCollection($this->getColumnCollection());
        $newTable->setFreezePanes($this->getFreezePanes());

        return $newTable;
    }

    public function enableAutoFilter(): void
    {
        if ($this->isEmpty() || 0 === $this->getDataRowStart()) {
            return;
        }

        $minCol = $this->getColumnStart();
        $minRow = $this->getDataRowStart() - 1; // header row
        $maxCol = $this->getColumnEnd();
        $maxRow = $this->getRowEnd();

        $autoFilter = new AutoFilter($minCol, $minRow + 1, $maxCol, $maxRow + 1);
        $this->activeSheet->setAutoFilter($autoFilter);
    }
}
