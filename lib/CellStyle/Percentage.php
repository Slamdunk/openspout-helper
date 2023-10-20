<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Percentage implements ContentDecoratorInterface
{
    public const FORMATCODE = '#,##0.000';

    public function styleCell(Style $style): void
    {
        $style->setFormat(self::FORMATCODE);
    }

    public function decorate(float|int|string $content): float
    {
        \assert(\is_numeric($content));

        return (float) $content;
    }
}
