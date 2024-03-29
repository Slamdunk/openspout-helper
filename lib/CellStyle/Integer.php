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

    public function decorate(float|int|string $content): int
    {
        \assert(\is_numeric($content) && ! \str_contains((string) $content, '.'));

        return (int) $content;
    }
}
