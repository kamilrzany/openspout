<?php

declare(strict_types=1);

namespace OpenSpout\Writer\Common;

/**
 * @internal
 */
final class ColumnWidth
{
    /**
     * @param positive-int $start
     * @param positive-int $end
     * @param positive-int $outlineLevel
     */
    public function __construct(
        public readonly int $start,
        public readonly int $end,
        public readonly float $width,
        public readonly ?int $outlineLevel = null,
        public readonly bool $collapsed = false,
        public readonly bool $hidden = false
    ) {}
}
