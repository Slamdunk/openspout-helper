<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentConsumerInterface;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class PaddedInteger implements ContentConsumerInterface, ContentDecoratorInterface
{
    private int $maxLength = 0;

    public function getDataType(): string
    {
        return Cell\NumericCell::class;
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

    public function decorate(mixed $content): int
    {
        \assert(\is_string($content));

        return (int) $content;
    }
}
