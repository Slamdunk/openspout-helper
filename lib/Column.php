<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

final class Column implements ColumnInterface
{
    public function __construct(
        private string $key,
        private string $heading,
        private int $width,
        private CellStyleInterface $cellStyle
    ) {
    }

    public function getKey(): string
    {
        return $this->key;
    }

    public function getHeading(): string
    {
        return $this->heading;
    }

    public function getWidth(): int
    {
        return $this->width;
    }

    public function getCellStyle(): CellStyleInterface
    {
        return $this->cellStyle;
    }
}
