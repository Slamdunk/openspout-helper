<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper\CellStyle;

use DateTimeImmutable;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\Style;
use Slam\OpenspoutHelper\ContentDecoratorInterface;

final class Date implements ContentDecoratorInterface
{
    public const FORMATCODE = 'dd/mm/yyyy';

    public function getDataType(): int
    {
        return Cell::TYPE_DATE;
    }

    public function styleCell(Style $style): void
    {
        $style->setCellAlignment(CellAlignment::CENTER);
        $style->setFormat(self::FORMATCODE);
    }

    public function decorate(mixed $content): mixed
    {
        if (! \is_string($content)) {
            return $content;
        }

        return new DateTimeImmutable($content);
    }
}
