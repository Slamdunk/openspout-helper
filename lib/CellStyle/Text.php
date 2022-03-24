<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\CellStyleInterface;

final class Text implements CellStyleInterface
{
    public function getDataType(): string
    {
        return Cell\StringCell::class;
    }

    public function styleCell(Style $style): void
    {
        $style->setCellAlignment(CellAlignment::LEFT);
    }
}
