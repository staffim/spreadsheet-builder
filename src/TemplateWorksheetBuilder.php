<?php

namespace Staffim\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Reader;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer;

class TemplateWorksheetBuilder implements WorksheetBuilderInterface
{
    /**
     * @var string
     */
    protected $template;

    /**
     * @var string
     */
    protected $templateWorksheetName;

    /**
     * @var array
     */
    protected $cellRange;

    public function __construct(
        string $template,
        string $worksheetName,
        string $cellRange = 'A1:Z30'
    )
    {
        $reader = new Reader\Xlsx();
        $this->template = $reader->load($template);
        $this->templateWorksheetName = $worksheetName;
        $this->cellRange = Coordinate::getRangeBoundaries($cellRange);
    }

    /**
     * @param iterable $data
     *
     * @return string
     */
    public function getWorksheetTitle(iterable $data): string
    {
        return 'worksheet title';
    }

    public function build(Worksheet $worksheet, iterable $data): void
    {
        $writer = new Writer\Xlsx($this->template);
        $spreadsheet = $worksheet->getParent();
        $spreadsheet->removeSheetByIndex($spreadsheet->getActiveSheetIndex());
        $worksheet = $writer->getSpreadsheet()->getSheetByName($this->templateWorksheetName);

        $iterator = $worksheet->getRowIterator($this->cellRange[0][1], $this->cellRange[1][1]);

        foreach ($iterator as $rowNumber => $row) {
            $cellsIterator = $row->getCellIterator($this->cellRange[0][0], $this->cellRange[1][0]);
            foreach ($cellsIterator as $columnIndex => $cell) {
                if ($cell->getFormattedValue() && preg_match_all('/{(?<placeholder>[a-zA-Z_-]*)}/i', $cell->getFormattedValue(), $matches)) {
                    foreach ($matches['placeholder'] as $placeholder) {
                        $value = $this->replacePlaceholders($cell->getFormattedValue(), $placeholder, $data[$placeholder] ?? '');
                        if ($cell->getValue() instanceof RichText) {
                            /** @var RichText $value */
                            $value = $cell->getValue();
                            foreach ($value->getRichTextElements() as $richTextElement) {
                                $richTextElement->setText($this->replacePlaceholders($richTextElement->getText(), $placeholder, $data[$placeholder] ?? ''));
                            }
                        }
                        $cell->setValue($value);
                    }
                }
            }
        }

        $worksheet->setTitle($this->getWorksheetTitle($data));
        $spreadsheet->addExternalSheet($worksheet);
        $spreadsheet->setActiveSheetIndexByName($this->getWorksheetTitle($data));
    }

    protected function replacePlaceholders($text, $placeholder, $value): string
    {
        return preg_replace(sprintf('/{%s}/i', $placeholder), $value, $text);
    }
}
