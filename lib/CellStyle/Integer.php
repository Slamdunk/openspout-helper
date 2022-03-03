<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Integer implements CellStyleInterface
{
    public const FORMATCODE = '#,##0';

    public function getDataType(): string
    {
        return DataType::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $style->getNumberFormat()->setFormatCode(self::FORMATCODE);
    }
}
