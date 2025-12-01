<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use Countable;
use OpenSpout\Writer\AutoFilter;
use OpenSpout\Writer\Common\Entity\Sheet;

final class Table implements Countable
{
    private Sheet $activeSheet;

    /** @var non-negative-int */
    private int $rowEnd;

    /** @var non-negative-int */
    private int $rowCurrent;

    /** @var non-negative-int */
    private int $rowStart;

    /** @var null|non-negative-int */
    private ?int $dataRowStart = null;

    /** @var non-negative-int */
    private int $columnStart   = 0;

    /** @var non-negative-int */
    private int $columnEnd     = 0;

    /** @var non-negative-int */
    private int $columnCurrent = 0;

    private string $heading;

    /** @var iterable<int, array<string, null|float|int|string>> */
    private iterable $data;
    private ColumnCollection $columnCollection;
    private bool $freezePanes = true;

    /** @var non-negative-int */
    private int $fontSize     = 8;

    /** @var null|non-negative-int */
    private ?int $rowHeight   = null;
    private bool $textWrap    = false;

    /** @var array<int, string> */
    private array $writtenColumn       = [];

    /** @var null|non-negative-int */
    private ?int $count                = null;

    /** @param iterable<int, array<string, null|float|int|string>> $data */
    public function __construct(Sheet $activeSheet, string $heading, iterable $data)
    {
        $this->activeSheet = $activeSheet;
        $this->heading     = $heading;
        $this->data        = $data;

        $this->columnCollection = new ColumnCollection();

        $this->rowStart
        = $this->rowEnd
        = $this->rowCurrent = $activeSheet->getWrittenRowCount();
    }

    public function getActiveSheet(): Sheet
    {
        return $this->activeSheet;
    }

    /** @return non-negative-int */
    public function getDataRowStart(): int
    {
        \assert(null !== $this->dataRowStart);

        return $this->dataRowStart;
    }

    public function flagDataRowStart(): void
    {
        $this->dataRowStart = $this->rowCurrent;
    }

    /** @return non-negative-int */
    public function getRowStart(): int
    {
        return $this->rowStart;
    }

    /** @return non-negative-int */
    public function getRowEnd(): int
    {
        return $this->rowEnd;
    }

    /** @return non-negative-int */
    public function getRowCurrent(): int
    {
        return $this->rowCurrent;
    }

    public function incrementRow(): void
    {
        $this->rowEnd = \max($this->rowEnd, $this->rowCurrent);
        ++$this->rowCurrent;
    }

    /** @return non-negative-int */
    public function getColumnStart(): int
    {
        return $this->columnStart;
    }

    /** @return non-negative-int */
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

    /** @param non-negative-int $fontSize */
    public function setFontSize(int $fontSize): void
    {
        $this->fontSize = $fontSize;
    }

    /** @return non-negative-int */
    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    /** @param null|non-negative-int $rowHeight */
    public function setRowHeight(?int $rowHeight): void
    {
        $this->rowHeight = $rowHeight;
    }

    /** @return null|non-negative-int */
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

    /** @param non-negative-int $count */
    public function setCount(int $count): void
    {
        $this->count = $count;
    }

    /** @return non-negative-int */
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
