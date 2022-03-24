<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\Style;

interface CellStyleInterface
{
    /**
     * @return class-string<Cell>
     */
    public function getDataType(): string;

    public function styleCell(Style $style): void;
}
