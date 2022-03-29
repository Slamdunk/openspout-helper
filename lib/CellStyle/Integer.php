<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Integer implements ContentDecoratorInterface
{
    public const FORMATCODE = '#,##0';

    public function styleCell(Style $style): void
    {
        $style->setCellAlignment(CellAlignment::CENTER);
        $style->setFormat(self::FORMATCODE);
    }

    public function decorate(string|int|float $content): int
    {
        \assert(\is_string($content) && \is_numeric($content) && ! \str_contains($content, '.'));

        return (int) $content;
    }
}
