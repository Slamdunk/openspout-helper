<?php

declare(strict_types=1);

namespace Slam\OpenspoutHelper;

use OpenSpout\Common\Entity\Style\Style;

/**
 * @internal
 */
final class CellStyleSpec
{
    public function __construct(
        public readonly ?CellStyleInterface $cellStyle,
        public readonly Style $headerStyle,
        public readonly Style $zebraLightStyle,
        public readonly Style $zebraDarkStyle,
    ) {
    }
}
