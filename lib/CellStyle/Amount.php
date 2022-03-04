<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Amount implements CellStyleInterface
{
    public const FORMATCODE = '#,##0.00';

    public function getDataType(): int
    {
        return Cell::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->setFormat(self::FORMATCODE);
    }
}
