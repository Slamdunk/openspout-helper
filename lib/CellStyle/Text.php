<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Text implements CellStyleInterface
{
    public function getDataType(): string
    {
        return DataType::TYPE_STRING;
    }

    public function styleCell(Style $style): void
    {
        $style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    }
}
