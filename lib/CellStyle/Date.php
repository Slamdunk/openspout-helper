<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use DateTimeImmutable;
use DateTimeZone;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Date implements ContentDecoratorInterface
{
    public const FORMATCODE = 'dd/mm/yyyy';

    public function styleCell(Style $style): void
    {
        $style->setCellAlignment(CellAlignment::CENTER);
        $style->setFormat(self::FORMATCODE);
    }

    public function decorate(float|int|string $content): DateTimeImmutable
    {
        \assert(\is_string($content));

        return new DateTimeImmutable($content, new DateTimeZone('UTC'));
    }
}
