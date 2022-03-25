<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\Tests;

use OpenSpout\Common\Helper\StringHelper;
use OpenSpout\Writer\Common\Entity\Sheet;
use OpenSpout\Writer\Common\Manager\SheetManager;
use PHPUnit\Framework\TestCase;
use Slam\OpenspoutHelper\Exception;
use Slam\OpenspoutHelper\Table;

final class TableTest extends TestCase
{
    private const EXCEL_DATA = [['col' => 'a'], ['col' => 'b']];
    private Sheet $activeSheet;
    private Table $table;
    private SheetManager $sheetManager;

    protected function setUp(): void
    {
        $this->sheetManager = new SheetManager(new StringHelper(false));
        $this->activeSheet  = new Sheet(1, '1', $this->sheetManager);
        $this->table        = new Table(
            $this->activeSheet,
            'My Heading',
            self::EXCEL_DATA
        );
    }

    public function testRowAndColumn(): void
    {
        self::assertSame($this->activeSheet, $this->table->getActiveSheet());
        self::assertSame('My Heading', $this->table->getHeading());
        self::assertSame(self::EXCEL_DATA, $this->table->getData());

        $this->table->incrementRow();
        $this->table->flagDataRowStart();
        $this->table->incrementRow();

        self::assertSame(0, $this->table->getRowStart());
        self::assertSame(1, $this->table->getDataRowStart());
        self::assertSame(1, $this->table->getRowEnd());

        $this->table->incrementColumn();
        $this->table->incrementColumn();

        self::assertSame(0, $this->table->getColumnStart());
        self::assertSame(1, $this->table->getColumnEnd());

        $this->table->setCount(0);
        self::assertCount(0, $this->table);
        self::assertTrue($this->table->isEmpty());

        $this->table->setCount(5);
        self::assertCount(5, $this->table);
        self::assertFalse($this->table->isEmpty());

        self::assertTrue($this->table->getFreezePanes());
        $this->table->setFreezePanes(false);
        self::assertFalse($this->table->getFreezePanes());

        self::assertEmpty($this->table->getWrittenColumn());
        $columns = [
            2 => 'column_1',
            3 => 'column_2',
        ];
        $this->table->setWrittenColumn($columns);
        self::assertSame($columns, $this->table->getWrittenColumn());
    }

    public function testTableCountMustBeSet(): void
    {
        $this->expectException(Exception::class);

        $this->table->count();
    }

    public function testSplitTableIfNeeded(): void
    {
        $this->table->setFreezePanes(false);
        $newTable = $this->table->splitTableOnNewWorksheet(new Sheet(2, '2', $this->sheetManager));

        self::assertNotSame($this->table, $newTable);

        // The starting row must be the first of the new sheet
        self::assertSame(0, $newTable->getRowStart());
        self::assertSame(0, $newTable->getRowEnd());

        // The starting column must be the same of the previous sheet
        self::assertSame(0, $newTable->getColumnStart());
        self::assertSame(0, $newTable->getColumnEnd());

        self::assertSame($this->table->getFreezePanes(), $newTable->getFreezePanes());
    }

    public function testFontRowAttributes(): void
    {
        self::assertSame(8, $this->table->getFontSize());
        self::assertNull($this->table->getRowHeight());
        self::assertFalse($this->table->getTextWrap());

        $this->table->setFontSize($fontSize   = \mt_rand(10, 100));
        $this->table->setRowHeight($rowHeight = \mt_rand(10, 100));
        $this->table->setTextWrap(true);

        self::assertSame($fontSize, $this->table->getFontSize());
        self::assertSame($rowHeight, $this->table->getRowHeight());
        self::assertTrue($this->table->getTextWrap());
    }
}
