<?php

namespace Staffim\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

interface WorksheetBuilderInterface
{
    /**
     * @param iterable $data
     *
     * @return string
     */
    public function getWorksheetTitle(iterable $data): string;

    /**
     * @param Worksheet $templateWorksheet
     * @param iterable $data
     */
    public function build(Worksheet $templateWorksheet, iterable $data): void;
}
