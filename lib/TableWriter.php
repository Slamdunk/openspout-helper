<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\XLSX\Entity\SheetView;
use OpenSpout\Writer\XLSX\Writer;

final class TableWriter
{
    public const COLOR_HEADER_FONT = 'FFFFFF';
    public const COLOR_HEADER_FILL = '4472C4';
    public const COLOR_ODD_FILL    = 'D9E1F2';

    public const COLUMN_DEFAULT_WIDTH = 10;

    /**
     * @var array<string, CellStyleSpec>
     */
    private array $styles;

    public function __construct(
        private string $emptyTableMessage = '',
        private int $rowsPerSheet = 262144
    ) {
    }

    /**
     * @return Table[]
     */
    public function writeTable(Writer $writer, Table $table): array
    {
        $defaultStyle = new Style();
        $defaultStyle->setFontSize($table->getFontSize());
        $defaultStyle->setShouldWrapText($table->getTextWrap());

        $writer->getOptions()->DEFAULT_ROW_STYLE = $defaultStyle;
        // if (null !== ($rowHeight = $table->getRowHeight())) {
        //     $writer->setDefaultRowHeight($rowHeight);
        // }
        $tables = [$table];

        $count      = 0;
        $headingRow = true;
        foreach ($table->getData() as $row) {
            ++$count;

            if ($table->getRowCurrent() >= $this->rowsPerSheet) {
                $table->setCount($count - 1);
                $count = 1;

                $table      = $table->splitTableOnNewWorksheet($writer->addNewSheetAndMakeItCurrent());
                $tables[]   = $table;
                $headingRow = true;
            }

            if ($headingRow) {
                $this->consumeFirstRow($table, $row);

                $columnKeys = \array_keys($row);
                $this->writeTableProperties($writer, $table, $columnKeys);
                $this->writeTableHeading($writer, $table);
                $this->writeColumnsHeading($writer, $table, $columnKeys);

                $headingRow = false;
            }

            $this->writeRow($writer, $table, $row, $count);
        }
        $table->setCount($count);

        if (\count($tables) > 1) {
            \reset($tables);
            $table      = \current($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = \substr($firstSheet->getName(), 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->setName(\sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal));
            }
        }

        if (0 === $tables[0]->count()) {
            $this->writeTableHeading($writer, $table);
            $writer->addRow(new Row([], null));
            $table->incrementRow();
            $writer->addRow(new Row([new Cell\StringCell($this->emptyTableMessage, null)], null));
            $table->incrementRow();
        }

        $table->setCount($count);

        return $tables;
    }

    /**
     * @param string[] $columnKeys
     */
    private function writeTableProperties(Writer $writer, Table $table, array $columnKeys): void
    {
        $this->generateStyles($table, $columnKeys);
        $columnCollection = $table->getColumnCollection();
        foreach ($columnKeys as $columnIndex => $columnKey) {
            $width = self::COLUMN_DEFAULT_WIDTH;
            if (null !== ($column = $columnCollection[$columnKey] ?? null)) {
                $width = $column->getWidth();
            }

            $writer->getOptions()->setColumnWidth($width, $columnIndex + 1);
            $table->incrementColumn();
        }

        if ($table->getFreezePanes()) {
            $table->getActiveSheet()->setSheetView(
                (new SheetView())
                    ->setFreezeRow(3)
            );
        }
    }

    private function writeTableHeading(Writer $writer, Table $table): void
    {
        $style = new Style();
        $style->setShouldWrapText(false);
        $style->setFontSize($table->getFontSize() + 2);

        $cell = new Cell\StringCell($table->getHeading(), $style);
        $writer->addRow(new Row([$cell], null));

        $table->incrementRow();
    }

    /**
     * @param string[] $columnKeys
     */
    private function writeColumnsHeading(Writer $writer, Table $table, array $columnKeys): void
    {
        $columnCollection = $table->getColumnCollection();
        $writtenColumn    = [];
        $titles           = [];
        foreach ($columnKeys as $columnIndex => $columnKey) {
            $newTitle = \ucwords(\str_replace('_', ' ', $columnKey));
            if (null !== ($column = $columnCollection[$columnKey] ?? null)) {
                $newTitle = $column->getHeading();
            }

            $writtenColumn[$columnIndex] = $columnKey;
            $titles[$columnKey]          = $newTitle;
        }

        $this->writeRow($writer, $table, $titles, 0);

        $table->setWrittenColumn($writtenColumn);
        $table->flagDataRowStart();
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeRow(Writer $writer, Table $table, array $row, int $odd): void
    {
        $isTitle = 0 === $odd;
        $cells   = [];
        foreach ($row as $key => $content) {
            if (
                ! $isTitle
                && null !== $content
                && ($cellStyle = $this->styles[$key]->cellStyle) instanceof ContentDecoratorInterface
            ) {
                $content = $cellStyle->decorate($content);
            }

            $cellStyleSpec = $this->styles[$key];
            $style         = $isTitle
                ? $cellStyleSpec->headerStyle
                : (
                    (1 === ($odd % 2))
                    ? $cellStyleSpec->zebraLightStyle
                    : $cellStyleSpec->zebraDarkStyle
                )
            ;

            $cells[] = Cell::fromValue($content, $style);
        }

        $writer->addRow(new Row($cells, null));

        $table->incrementRow();
    }

    /**
     * @param string[] $columnKeys
     */
    private function generateStyles(Table $table, array $columnKeys): void
    {
        $columnCollection = $table->getColumnCollection();
        $this->styles     = [];
        foreach ($columnKeys as $columnKey) {
            $header = new Style();
            $header->setCellAlignment(CellAlignment::CENTER);
            $header->setShouldWrapText(true);
            $header->setBackgroundColor(self::COLOR_HEADER_FILL);
            $header->setFontSize($table->getFontSize());
            $header->setFontBold();
            $header->setFontColor(self::COLOR_HEADER_FONT);

            $zebraLight = new Style();
            $zebraLight->setFontSize($table->getFontSize());

            $zebraDark = new Style();
            $zebraDark->setFontSize($table->getFontSize());
            $zebraDark->setBackgroundColor(self::COLOR_ODD_FILL);

            $cellStyle = ($columnCollection[$columnKey] ?? null)?->getCellStyle();

            if (null !== $cellStyle) {
                $cellStyle->styleCell($zebraLight);
                $cellStyle->styleCell($zebraDark);
            }

            $this->styles[$columnKey] = new CellStyleSpec(
                $cellStyle,
                $header,
                $zebraLight,
                $zebraDark,
            );
        }
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function consumeFirstRow(Table $table, array $row): void
    {
        foreach ($table->getColumnCollection() as $column) {
            $style = $column->getCellStyle();
            if ($style instanceof ContentConsumerInterface && isset($row[$column->getKey()])) {
                $style->consume($row[$column->getKey()]);
            }
        }
    }
}
