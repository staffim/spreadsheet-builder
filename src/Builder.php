<?php

namespace Staffim\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Builder
{
    /**
     * @var array
     */
    protected $worksheetBuilders = [];

    /**
     * @param array $worksheetBuilders
     */
    public function __construct(array $worksheetBuilders = [])
    {
        $this->setWorksheetBuilders($worksheetBuilders);
    }

    /**
     * @param WorksheetBuilderInterface $builder
     */
    public function registerWorksheetBuilder(WorksheetBuilderInterface $builder): void
    {
        $this->worksheetBuilders[] = $builder;
    }

    /**
     * @return WorksheetBuilderInterface[]
     */
    public function getWorksheetBuilders(): array
    {
        return $this->worksheetBuilders;
    }

    /**
     * @param array $builders
     */
    public function setWorksheetBuilders(array $builders): void
    {
        $this->worksheetBuilders = [];

        foreach ($builders as $builder) {
            $this->registerWorksheetBuilder($builder);
        }
    }

    /**
     * @param iterable $worksheetsData
     *
     * @return Spreadsheet
     *
     * @throws Exception
     */
    public function build(iterable $worksheetsData): Spreadsheet
    {
        if (count($this->getWorksheetBuilders()) < 1) {
            throw new \RuntimeException('Worksheet builders are missing');
        }

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');

        /**
         * @var WorksheetBuilderInterface $worksheetBuilder
         */
        foreach ($this->getWorksheetBuilders() as $index => $worksheetBuilder) {
            $spreadsheet->setActiveSheetIndex($index);

            if (!array_key_exists($index, $worksheetsData)) {
                throw new \RuntimeException(sprintf('Data for sheet %s not found', get_class($worksheetBuilder)));
            }
            $data = $worksheetsData[$index];

            $worksheetBuilder->build($spreadsheet->getActiveSheet(), $data);
        }

        return $spreadsheet;
    }
}
