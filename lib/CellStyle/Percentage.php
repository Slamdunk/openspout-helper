<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Percentage implements CellStyleInterface
{
    public const FORMATCODE = '#,##0.000';

    public function getDataType(): string
    {
        return Cell\NumericCell::class;
    }

    public function styleCell(Style $style): void
    {
        $style->setFormat(self::FORMATCODE);
    }
}
