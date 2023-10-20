<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use DateTimeImmutable;

interface ContentDecoratorInterface extends CellStyleInterface
{
    public function decorate(float|int|string $content): null|DateTimeImmutable|float|int|string;
}
