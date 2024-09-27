<?php

declare(strict_types=1);

namespace OpenSpout\Writer\Common;

use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\TempFolderOptionTrait;

abstract class AbstractOptions
{
    use TempFolderOptionTrait;

    public Style $DEFAULT_ROW_STYLE;
    public bool $SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY = true;
    public ?float $DEFAULT_COLUMN_WIDTH = null;
    public ?float $DEFAULT_ROW_HEIGHT = null;

    /** @var ColumnWidth[] Array of min-max-width arrays */
    private array $COLUMN_WIDTHS = [];

    public function __construct()
    {
        $this->DEFAULT_ROW_STYLE = new Style();
    }

    /**
     * @param positive-int ...$columns One or more columns with this width
     */
    final public function setColumnWidth(float $width, ?int $outlineLevel = null, bool $collapsed = false, bool $hidden = false, int ...$columns): void
    {
        // Gather sequences
        $sequence = [];
        foreach ($columns as $column) {
            $sequenceLength = \count($sequence);
            if ($sequenceLength > 0) {
                $previousValue = $sequence[$sequenceLength - 1];
                if ($column !== $previousValue + 1) {
                    $this->setColumnWidthForRange($width, $sequence[0], $previousValue, $outlineLevel, $collapsed, $hidden);
                    $sequence = [];
                }
            }
            $sequence[] = $column;
        }
        $this->setColumnWidthForRange($width, $sequence[0], $sequence[\count($sequence) - 1], $outlineLevel, $collapsed, $hidden);
    }

    /**
     * @param float        $width The width to set
     * @param positive-int $start First column index of the range
     * @param positive-int $end   Last column index of the range
     * @param int|null $outlineLevel Outline level (1-7)
     * @param bool $collapsed Whether the column group is collapsed
     * @param bool $hidden Whether the column is hidden
     */
    final public function setColumnWidthForRange(float $width, int $start, int $end, ?int $outlineLevel = null, bool $collapsed = false, bool $hidden = false): void
    {
        if (null !== $outlineLevel && ($outlineLevel < 1 || $outlineLevel > 7)) {
            throw new \InvalidArgumentException('Outline level must be between 1 and 7');
        }

        $this->COLUMN_WIDTHS[] = new ColumnWidth($start, $end, $width, $outlineLevel, $collapsed, $hidden);
    }

    /**
     * @internal
     *
     * @return ColumnWidth[]
     */
    final public function getColumnWidths(): array
    {
        return $this->COLUMN_WIDTHS;
    }
}
