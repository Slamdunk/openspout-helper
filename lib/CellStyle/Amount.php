<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Amount implements ContentDecoratorInterface
{
    public const string FORMATCODE = '#,##0.00';

    public function styleCell(Style $style): Style
    {
        return $style->withFormat(self::FORMATCODE);
    }

    public function decorate(float|int|string $content): float
    {
        \assert(\is_numeric($content));

        return (float) $content;
    }
}
