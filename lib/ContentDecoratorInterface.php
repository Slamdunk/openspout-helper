<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use DateTimeImmutable;

interface ContentDecoratorInterface extends CellStyleInterface
{
    public function decorate(string|int|float $content): null|string|int|float|DateTimeImmutable;
}
