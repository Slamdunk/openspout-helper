<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Amount implements ContentDecoratorInterface
{
    public const FORMATCODE = '#,##0.00';

    public function styleCell(Style $style): void
    {
        $style->setFormat(self::FORMATCODE);
    }

    public function decorate(string|int|float $content): float
    {
        \assert(\is_numeric($content));

        return (float) $content;
    }
}
