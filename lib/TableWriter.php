<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Writer\XLSX\Writer;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception as PhpspreadsheetException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;

final class TableWriter
{
    public const COLOR_HEADER_FONT = 'FFFFFF';
    public const COLOR_HEADER_FILL = '4472C4';
    public const COLOR_ODD_FILL    = 'D9E1F2';

    public const COLUMN_DEFAULT_WIDTH = 10;

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
        
        $writer->setDefaultRowStyle($defaultStyle);
        // if (null !== ($rowHeight = $table->getRowHeight())) {
        //     $writer->setDefaultRowHeight($rowHeight);
        // }

        $this->writeTableHeading($writer, $table);
        $tables = [$table];

        $count      = 0;
        $headingRow = true;
        foreach ($table->getData() as $row) {
            ++$count;

            if ($table->getRowCurrent() >= $this->rowsPerSheet) {
                $table->setCount($count - 1);
                $count = 1;

                $table    = $table->splitTableOnNewWorksheet($writer->addNewSheetAndMakeItCurrent());
                $tables[] = $table;
                $this->writeTableHeading($writer, $table);
                $headingRow = true;
            }

            if ($headingRow) {
                $this->writeColumnsHeading($writer, $table, $row);

                $headingRow = false;
            }

            $this->writeRow($writer, $table, $row, false);
        }
        $table->setCount($count);

        if (\count($tables) > 1) {
            \reset($tables);
            $table      = \current($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = \mb_substr($firstSheet->getName(), 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->setName(\sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal));
            }
        }

        foreach ($tables as $table) {
            $columnCollection = $table->getColumnCollection();
            foreach ($table->getWrittenColumn() as $columnIndex => $columnKey) {
                if (! isset($columnCollection[$columnKey])) {
                    continue;
                }

                $dataRowStart = $table->getDataRowStart();
                \assert(null !== $dataRowStart);
                $columnCollection[$columnKey]->getCellStyle()->styleCell($table->getActiveSheet()->getStyleByColumnAndRow(
                    $columnIndex,
                    $dataRowStart,
                    $columnIndex,
                    $table->getRowEnd()
                ));
            }
        }

//        if ($table->getFreezePanes()) {
//            foreach ($tables as $table) {
//                $table->getActiveSheet()->freezePaneByColumnAndRow(1, 2 + $table->getRowStart());
//            }
//        }

        if (0 !== $tables[0]->count()) {
//            $conditional = $this->getZebraStripingStyle();
//            foreach ($tables as $table) {
//                $activeSheet = $table->getActiveSheet();
//                $activeSheet->setAutoFilterByColumnAndRow(
//                    $table->getColumnStart(),
//                    $table->getDataRowStart() - 1,
//                    $table->getColumnEnd(),
//                    $table->getRowEnd()
//                );
//                $activeSheet->getStyleByColumnAndRow(
//                    $table->getColumnStart(),
//                    $table->getDataRowStart(),
//                    $table->getColumnEnd(),
//                    $table->getRowEnd()
//                )->setConditionalStyles([$conditional]);
//                $activeSheet->setSelectedCellByColumnAndRow(
//                    $table->getColumnStart(),
//                    $table->getDataRowStart()
//                );
//            }
        } else {
            $writer->addRow(new Row([], null));
            $table->incrementRow();
            $writer->addRow(new Row([new Cell($this->emptyTableMessage)], null));
            $table->incrementRow();
        }

        $table->setCount($count);

        return $tables;
    }

    private function writeTableHeading(Writer $writer, Table $table): void
    {
//        $table->resetColumn();
//        $table->getActiveSheet()->setCellValueExplicitByColumnAndRow(
//            $table->getColumnCurrent(),
//            $table->getRowCurrent(),
//            $table->getHeading(),
//            DataType::TYPE_STRING
//        );
//
//        $headingStyle = $table->getActiveSheet()->getStyleByColumnAndRow(
//            $table->getColumnCurrent(),
//            $table->getRowCurrent()
//        );
//        $headingStyle->getAlignment()->setWrapText(false);
//        $headingStyle->getFont()->setSize($table->getFontSize() + 2);

        $writer->addRow(new Row([new Cell($table->getHeading())], null));

        $table->incrementRow();
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeColumnsHeading(Writer $writer, Table $table, array $row): void
    {
        $columnCollection = $table->getColumnCollection();
        $columnKeys       = \array_keys($row);

        $writtenColumn = [];
        $titles        = [];
        foreach ($columnKeys as $columnIndex => $columnKey) {
            $width    = self::COLUMN_DEFAULT_WIDTH;
            $newTitle = \ucwords(\str_replace('_', ' ', $columnKey));

            if (null !== ($column = $columnCollection[$columnKey] ?? null)) {
                $width    = $column->getWidth();
                $newTitle = $column->getHeading();
            }

//            $table->getActiveSheet()->getColumnDimensionByColumn($table->getColumnCurrent())->setWidth($width);
            $writtenColumn[$columnIndex] = $columnKey;
            $titles[$columnKey]          = $newTitle;

            $table->incrementColumn();
        }

        $this->writeRow($writer, $table, $titles, true);

        $table->setWrittenColumn($writtenColumn);
        $table->flagDataRowStart();
    }

    /**
     * @param array<string, null|float|int|string> $row
     */
    private function writeRow(Writer $writer, Table $table, array $row, bool $isTitle): void
    {
        $cells = [];
        foreach ($row as $key => $content) {
            $content  = null !== $content
                ? (string) $content
                : null
            ;
//            $dataType = DataType::TYPE_STRING;
//            if (null === $content) {
//                $dataType = DataType::TYPE_NULL;
//            } elseif (
//                ! $isTitle
//                && 0 !== ($columnCollection = $table->getColumnCollection())->count()
//                && isset($columnCollection[$key])
//            ) {
//                $cellStyle = $columnCollection[$key]->getCellStyle();
//                $dataType  = $cellStyle->getDataType();
//                if ($cellStyle instanceof ContentDecoratorInterface) {
//                    $content = $cellStyle->decorate($content);
//                }
//            }

            $cells[] = new Cell($content);
        }
        
        $writer->addRow(new Row($cells, null));

//        if (null !== ($rowHeight = $table->getRowHeight())) {
//            $sheet->getRowDimension($table->getRowCurrent())->setRowHeight($rowHeight);
//        }

//        if ($isTitle) {
//            $titleStyle = $sheet->getStyleByColumnAndRow(
//                $table->getColumnStart(),
//                $table->getRowCurrent(),
//                $table->getColumnEnd(),
//                $table->getRowCurrent(),
//            );
//            $alignment = $titleStyle->getAlignment();
//            $alignment->setHorizontal(Alignment::HORIZONTAL_CENTER);
//            $alignment->setVertical(Alignment::VERTICAL_CENTER);
//            $alignment->setWrapText(true);
//            $font = $titleStyle->getFont();
//            $font->getColor()->setARGB(self::COLOR_HEADER_FONT);
//            $font->setBold(true);
//            $fill = $titleStyle->getFill();
//            $fill->setFillType(Fill::FILL_SOLID);
//            $fill->getStartColor()->setARGB(self::COLOR_HEADER_FILL);
//            $fill->getEndColor()->setARGB(self::COLOR_HEADER_FILL);
//        }

        $table->incrementRow();
    }

    private function getZebraStripingStyle(): Conditional
    {
        $conditional = new Conditional();
        $conditional->setConditionType(Conditional::CONDITION_EXPRESSION);
        $conditional->setOperatorType(Conditional::OPERATOR_EQUAL);
        $conditional->addCondition('MOD(ROW(),2)=0');
        $style = $conditional->getStyle();
        $fill  = $style->getFill();
        $fill->setFillType(Fill::FILL_SOLID);
        $fill->getStartColor()->setARGB(self::COLOR_ODD_FILL);
        $fill->getEndColor()->setARGB(self::COLOR_ODD_FILL);

        return $conditional;
    }
}
