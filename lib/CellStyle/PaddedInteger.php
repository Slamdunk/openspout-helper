<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentConsumerInterface;

final class PaddedInteger implements ContentConsumerInterface
{
    private int $maxLength = 0;

    public function getDataType(): int
    {
        return Cell::TYPE_NUMERIC;
    }

    public function styleCell(Style $style): void
    {
        $style->setCellAlignment(CellAlignment::CENTER);
        $style->setFormat(\str_repeat('0', $this->maxLength));
    }

    public function consume(mixed $content): void
    {
        if (\is_string($content)) {
            $this->maxLength = \max($this->maxLength, \strlen($content));
        }
    }
}
