<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Date implements CellStyleInterface
{
    public function getDataType(): string
    {
        return DataType::TYPE_ISO_DATE;
    }

    public function styleCell(Style $style): void
    {
        $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $style->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
    }
}
