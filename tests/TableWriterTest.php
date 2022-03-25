<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\Tests;

use OpenSpout\Writer\XLSX\Writer;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PHPUnit\Framework\TestCase;
use Slam\OpenspoutHelper\CellStyle;
use Slam\OpenspoutHelper\Column;
use Slam\OpenspoutHelper\ColumnCollection;
use Slam\OpenspoutHelper\Table;
use Slam\OpenspoutHelper\TableWriter;

final class TableWriterTest extends TestCase
{
    private string $filename;

    protected function setUp(): void
    {
        $this->filename = __DIR__ . '/tmp/test.xlsx';
        @\unlink($this->filename);
    }

    public function testPostGenerationDetails(): void
    {
        $XLSXWriter  = new Writer();
        $XLSXWriter->openToFile($this->filename);
        $heading = \uniqid('Heading_');
        $table   = new Table($XLSXWriter->getCurrentSheet(), $heading, [
            ['description' => 'AAA'],
            ['description' => 'BBB'],
        ]);

        (new TableWriter())->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();

        self::assertSame(0, $table->getRowStart());
        self::assertSame(3, $table->getRowEnd());

        self::assertSame(2, $table->getDataRowStart());

        self::assertSame(0, $table->getColumnStart());
        self::assertSame(0, $table->getColumnEnd());

        self::assertCount(2, $table);
        self::assertSame([0 => 'description'], $table->getWrittenColumn());

        $sheet = (new Xlsx())->load($this->filename)->getActiveSheet();

        self::assertSame($heading, (string) $sheet->getCellByColumnAndRow(1, 1)->getValue());
        self::assertSame('Description', (string) $sheet->getCellByColumnAndRow(1, 2)->getValue());
        self::assertSame('AAA', (string) $sheet->getCellByColumnAndRow(1, 3)->getValue());
        self::assertSame('BBB', (string) $sheet->getCellByColumnAndRow(1, 4)->getValue());
    }

    public function testHandleEncoding(): void
    {
        $textWithSpecialCharacters = \implode(' # ', [
            '€',
            'VIA MARTIRI DELLA LIBERTà 2',
            'FISSO20+OPZ.I¢CASA EURIB 3',
            'FISSO 20+ OPZIONE I°CASA EUR 6',
            '1° MAGGIO',
            'GIÀ XXXXXXX YYYYYYYYYYY',
            'FINANZIAMENTO 13/14¬ MENSILITà',

            'A \'\\|!"£$%&/()=?^àèìòùáéíóúÀÈÌÒÙÁÉÍÓÚ<>*ç°§[]@#{},.-;:_~` Z',
        ]);
        $heading = \sprintf('%s: %s', \uniqid('Heading_'), $textWithSpecialCharacters);
        $data    = \sprintf('%s: %s', \uniqid('Data_'), $textWithSpecialCharacters);

        $XLSXWriter  = new Writer();
        $XLSXWriter->openToFile($this->filename);
        $activeSheet = $XLSXWriter->getCurrentSheet();
        $activeSheet->setName(\uniqid());
        $table   = new Table($activeSheet, $heading, [
            ['description' => $data],
        ]);

        (new TableWriter())->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();

        $sheet = (new Xlsx())->load($this->filename)->getActiveSheet();
        self::assertSame($activeSheet->getName(), $sheet->getTitle());

        // Heading
        self::assertSame($heading, (string) $sheet->getCell('A1')->getValue());

        // Data
        self::assertSame($data, (string) $sheet->getCell('A3')->getValue());
    }

    public function testCellStyles(): void
    {
        $XLSXWriter  = new Writer();
        $XLSXWriter->openToFile($this->filename);

        $columnCollection = new ColumnCollection(...[
            new Column('disorder', 'Foo50', 50, new CellStyle\Text()),

            new Column('my_text', 'Foo1', 11, new CellStyle\Text()),
            new Column('my_perc', 'Foo2', 12, new CellStyle\Percentage()),
            new Column('my_inte', 'Foo3', 13, new CellStyle\Integer()),
            new Column('my_date', 'Foo4', 14, new CellStyle\Date()),
            new Column('my_amnt', 'Foo5', 15, new CellStyle\Amount()),
            new Column('my_itfc', 'Foo6', 16, new CellStyle\Text()),
            new Column('my_nodd', 'Foo7', 17, new CellStyle\Date()),
            new Column('my_padd', 'Foo8', 18, new CellStyle\PaddedInteger()),

            new Column(\uniqid('non_existent_'), 'Foo8', 18, new CellStyle\PaddedInteger()),
        ]);

        $table = new Table($XLSXWriter->getCurrentSheet(), \uniqid('Heading_'), [
            [
                'my_text' => 'text',
                'my_perc' => 3.45,
                'my_inte' => 1234567.8,
                'my_date' => '2017-03-02',
                'my_amnt' => 1234567.89,
                'my_itfc' => 'AABB',
                'my_nodd' => null,
                'my_padd' => '0123',

                'disorder'  => 'disorder',
                'no_column' => 'no_column',
            ],
        ]);
        $table->setColumnCollection($columnCollection);

        (new TableWriter())->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();

        $firstSheet = (new Xlsx())->load($this->filename)->getActiveSheet();

        $expectedContent = [
            'A1' => $table->getHeading(),

            'A2' => 'Foo1',
            'B2' => 'Foo2',
            'C2' => 'Foo3',
            'D2' => 'Foo4',
            'E2' => 'Foo5',
            'F2' => 'Foo6',
            'G2' => 'Foo7',
            'H2' => 'Foo8',
            'I2' => 'Foo50',
            'J2' => 'No Column',

            'A3' => 'text',
            'B3' => 3.45,
            'C3' => 1234567.8,
            'D3' => 42796,
            'E3' => 1234567.89,
            'F3' => 'AABB',
            'G3' => null,
            'H3' => 123,
            'I3' => 'disorder',
            'J3' => 'no_column',
        ];

        $expectedDataType = [
            'A1' => DataType::TYPE_INLINE,

            'A2' => DataType::TYPE_INLINE,
            'B2' => DataType::TYPE_INLINE,
            'C2' => DataType::TYPE_INLINE,
            'D2' => DataType::TYPE_INLINE,
            'E2' => DataType::TYPE_INLINE,
            'F2' => DataType::TYPE_INLINE,
            'G2' => DataType::TYPE_INLINE,
            'H2' => DataType::TYPE_INLINE,
            'I2' => DataType::TYPE_INLINE,
            'J2' => DataType::TYPE_INLINE,

            'A3' => DataType::TYPE_INLINE,
            'B3' => DataType::TYPE_NUMERIC,
            'C3' => DataType::TYPE_NUMERIC,
            'D3' => DataType::TYPE_NUMERIC,
            'E3' => DataType::TYPE_NUMERIC,
            'F3' => DataType::TYPE_INLINE,
            'G3' => DataType::TYPE_NULL,
            'H3' => DataType::TYPE_NUMERIC,
            'I3' => DataType::TYPE_INLINE,
            'J3' => DataType::TYPE_INLINE,
        ];

        $expectedNumberFormat = [
            'A1' => NumberFormat::FORMAT_GENERAL,

            'A2' => NumberFormat::FORMAT_GENERAL,
            'B2' => NumberFormat::FORMAT_GENERAL,
            'C2' => NumberFormat::FORMAT_GENERAL,
            'D2' => NumberFormat::FORMAT_GENERAL,
            'E2' => NumberFormat::FORMAT_GENERAL,
            'F2' => NumberFormat::FORMAT_GENERAL,
            'G2' => NumberFormat::FORMAT_GENERAL,
            'H2' => NumberFormat::FORMAT_GENERAL,
            'I2' => NumberFormat::FORMAT_GENERAL,
            'J2' => NumberFormat::FORMAT_GENERAL,

            'A3' => NumberFormat::FORMAT_GENERAL,
            'B3' => CellStyle\Percentage::FORMATCODE,
            'C3' => CellStyle\Integer::FORMATCODE,
            'D3' => NumberFormat::FORMAT_DATE_DDMMYYYY,
            'E3' => CellStyle\Amount::FORMATCODE,
            'F3' => NumberFormat::FORMAT_GENERAL,
            'G3' => NumberFormat::FORMAT_DATE_DDMMYYYY,
            'H3' => '0000',
            'I3' => NumberFormat::FORMAT_GENERAL,
            'J3' => NumberFormat::FORMAT_GENERAL,
        ];

        $expectedWidths = [
            'A' => 11,
            'B' => 12,
            'C' => 13,
            'D' => 14,
            'E' => 15,
            'F' => 16,
            'G' => 17,
            'H' => 18,
            'I' => 50,
            'J' => TableWriter::COLUMN_DEFAULT_WIDTH,
        ];

        $actualContent      = [];
        $actualDataType     = [];
        $actualNumberFormat = [];
        $actualWidths       = [];
        foreach ($expectedContent as $coordinate => $content) {
            $cell                            = $firstSheet->getCell($coordinate);
            $actualContent[$coordinate]      = (($value = $cell->getValue()) instanceof RichText)
                ? (string) $value
                : $value
            ;
            $actualDataType[$coordinate]     = $cell->getDataType();
            $actualNumberFormat[$coordinate] = $cell->getStyle()->getNumberFormat()->getFormatCode();
            $actualWidths[$coordinate[0]]    = (int) $firstSheet->getColumnDimension($coordinate[0])->getWidth();
        }

        self::assertSame($expectedContent, $actualContent);
        self::assertSame($expectedDataType, $actualDataType);
        self::assertSame($expectedNumberFormat, $actualNumberFormat);
        self::assertSame($expectedWidths, $actualWidths);
    }

    public function testTablePagination(): void
    {
        $XLSXWriter  = new Writer();
        $XLSXWriter->openToFile($this->filename);
        $worksheet = $XLSXWriter->getCurrentSheet();
        $worksheet->setName('names');
        $table = new Table($worksheet, \uniqid('Heading_'), [
            ['description' => 'AAA'],
            ['description' => 'BBB'],
            ['description' => 'CCC'],
            ['description' => 'DDD'],
            ['description' => 'EEE'],
        ]);

        $tables = (new TableWriter('', 5))->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();

        self::assertCount(2, $tables);
        self::assertSame(3, $tables[0]->count());
        self::assertSame(2, $tables[1]->count());

        $sheets     = (new Xlsx())->load($this->filename)->getAllSheets();
        $firstSheet = $sheets[0];

        $expected   = [
            'A1' => $table->getHeading(),
            'A2' => 'Description',
            'A3' => 'AAA',
            'A4' => 'BBB',
            'A5' => 'CCC',
            'A6' => '',
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = (string) $firstSheet->getCell($cell)->getValue();
        }
        self::assertSame($expected, $actual);

        $secondSheet = $sheets[1];
        $expected    = [
            'A1' => $tables[1]->getHeading(),
            'A2' => 'Description',
            'A3' => 'DDD',
            'A4' => 'EEE',
            'A5' => '',
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = (string) $secondSheet->getCell($cell)->getValue();
        }
        self::assertSame($expected, $actual);

        self::assertStringContainsString('names (', $firstSheet->getTitle());
        self::assertStringContainsString('names (', $secondSheet->getTitle());
    }

    public function testEmptyTable(): void
    {
        $emptyTableMessage = \uniqid('no_data_');
        $XLSXWriter        = new Writer();
        $XLSXWriter->openToFile($this->filename);

        $table = new Table($XLSXWriter->getCurrentSheet(), \uniqid(), []);

        (new TableWriter($emptyTableMessage))->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();
        $firstSheet = (new Xlsx())->load($this->filename)->getActiveSheet();

        $expected   = [
            'A1' => $table->getHeading(),
            'A2' => '',
            'A3' => $emptyTableMessage,
            'A4' => '',
        ];

        $actual = [];
        foreach ($expected as $cell => $content) {
            $actual[$cell] = (string) $firstSheet->getCell($cell)->getValue();
        }

        self::assertSame($expected, $actual);
    }

    public function testFontRowAttributesUsage(): void
    {
        $XLSXWriter = new Writer();
        $XLSXWriter->openToFile($this->filename);
        $table  = new Table($XLSXWriter->getCurrentSheet(), \uniqid(), [
            [
                'name'    => 'Foo',
                'surname' => 'Bar',
            ],
            [
                'name'    => 'Baz',
                'surname' => 'Xxx',
            ],
        ]);

        $table->setFontSize(12);
        $table->setRowHeight(33);
        $table->setTextWrap(true);

        (new TableWriter())->writeTable($XLSXWriter, $table);
        $XLSXWriter->close();
        $firstSheet = (new Xlsx())->load($this->filename)->getActiveSheet();

        $cell       = $firstSheet->getCell('A3');
        $style      = $cell->getStyle();

        self::assertSame('Foo', (string) $cell->getValue());
        self::assertSame(12, (int) $style->getFont()->getSize());
        // self::assertSame(33, (int) $firstSheet->getRowDimension($cell->getRow())->getRowHeight());
        self::assertTrue($style->getAlignment()->getWrapText());
    }
}
