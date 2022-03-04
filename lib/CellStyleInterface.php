<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use OpenSpout\Common\Entity\Style\Style;

interface CellStyleInterface
{
    public function getDataType(): int;

    public function styleCell(Style $style): void;
}
