<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

interface ContentConsumerInterface extends CellStyleInterface
{
    public function consume(mixed $content): void;
}
