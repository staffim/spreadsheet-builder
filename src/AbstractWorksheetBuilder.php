<?php

namespace Experium\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

abstract class AbstractWorksheetBuilder
{
    protected $images = [];

    private const CELL_PADDING = 5;

    /**
     * @param iterable $data
     *
     * @return string
     */
    abstract public function getTableTitle(iterable $data): string;

    /**
     * @param iterable $data
     *
     * @return string
     */
    abstract public function getWorksheetTitle(iterable $data): string;

    /**
     * @return array
     */
    abstract protected function getColumnsSettings(iterable $data): array;

    /**
     * @param Worksheet $sheet
     * @param iterable $data
     *
     * @throws Exception
     */
    public function build(Worksheet $sheet, iterable $data)
    {
        if (count($this->getColumnsSettings($data)) <= 0) {
            throw new \RuntimeException('Columns settings is missing');
        }

        $sheet->setTitle($this->getWorksheetTitle($data));
        $this->buildTitle($sheet, $data);
        $this->buildTable($sheet, $data);
    }

    /**
     * @param string $name
     *
     * @return array|mixed
     */
    public function getStyle(string $name)
    {
        return $this->getStyles()[$name] ?? [];
    }

    /**
     * @param Worksheet $sheet
     * @param iterable $data
     *
     * @throws Exception
     */
    protected function buildTitle(Worksheet $sheet, iterable $data): void
    {
        $tableTitle = $this->getTableTitle($data);

        if ($this->getTableFirstRowNumber() > 1 && (bool) $tableTitle) {
            $titleRowNumber = $this->getTableFirstRowNumber() - 1;
            $sheet->setCellValueByColumnAndRow(1, $titleRowNumber, $tableTitle);
            $sheet->mergeCellsByColumnAndRow(1, $titleRowNumber, count($this->getColumnsSettings($data)), $titleRowNumber);
            $sheet->getStyleByColumnAndRow(1, $titleRowNumber)->applyFromArray($this->getStyle('title'));
        }
    }

    /**
     * @param Worksheet $sheet
     * @param iterable $data
     *
     * @return int
     */
    protected function buildTable(Worksheet $sheet, iterable $data)
    {
        $sheet->setShowGridlines(false);

        $rowNumber = $firstTableRowNumber = $this->buildTableHeader($sheet, $data);

        $sheet->freezePaneByColumnAndRow(1, $rowNumber);

        foreach ($data as $dataItem) {
            foreach ($this->getColumnsSettings($data) as $columnIndex => $columnsSetting) {
                $column = $columnIndex + 1;
                $sheet->setCellValueByColumnAndRow(
                    $column,
                    $rowNumber,
                    is_callable($columnsSetting['value']) ? $columnsSetting['value']($dataItem, $column, $rowNumber) : $columnsSetting['value']
                );

                $this->appendImages($column, $rowNumber, $sheet);
            }
            ++$rowNumber;
        }

        --$rowNumber;

        foreach ($this->getColumnsSettings($data) as $columnIndex => $columnsSetting) {
            $this->applyColumnStyle($columnsSetting, $sheet, $columnIndex);
        }

        $sheet->getStyleByColumnAndRow(1, $firstTableRowNumber, $this->getColumnsCount($data), $rowNumber)
            ->applyFromArray($this->getStyle('cell'));

        $this->images = [];

        return $rowNumber;
    }

    /**
     * @param Worksheet $sheet
     * @param iterable $data
     *
     * @return int
     *
     * @throws Exception
     */
    protected function buildTableHeader(Worksheet $sheet, iterable $data): int
    {
        $rowNumber = $this->getTableFirstRowNumber();

        $tableTitle = $this->getTableTitle($data);

        if ($this->getTableFirstRowNumber() > 1 && (bool) $tableTitle) {
            $titleRowNumber = $this->getTableFirstRowNumber() - 1;
            $sheet->setCellValueByColumnAndRow(1, $titleRowNumber, $tableTitle);
            $sheet->mergeCellsByColumnAndRow(1, $titleRowNumber, count($this->getColumnsSettings($data)), $titleRowNumber);
        }

        $sheet->getRowDimension($rowNumber)->setRowHeight(50);
        $this->setColumnDimensions($sheet, $data);

        $titles = $this->getHeaderTitles($data);

        $sheet->getStyleByColumnAndRow(1, $rowNumber, count($titles), $rowNumber)->applyFromArray($this->getStyle('header'));
        foreach ($titles as $column => $title) {
            $columnNumber = $column + 1;
            $sheet->setCellValueByColumnAndRow($columnNumber, $rowNumber, $title);
        }
        ++$rowNumber;

        return $rowNumber;
    }

    /**
     * @return array
     */
    protected function getStyles(): array
    {
        return [
            'title' => [
                'font' => [
                    'bold' => true,
                ],
            ],
            'header' => [
                'font' => [
                    'bold' => true,
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => ['argb' => 'FFd6dce5'],
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_LEFT,
                    'vertical' => Alignment::VERTICAL_CENTER,
                    'wrapText' => true,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => Color::COLOR_BLACK],
                    ],
                ],
            ],
            'cell' => [
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_LEFT,
                    'vertical' => Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => 'FF000000'],
                    ],
                ],
            ],
        ];
    }

    /**
     * @return int
     */
    protected function getTableFirstRowNumber(): int
    {
        return 1;
    }

    /**
     * @param iterable $data
     *
     * @return int
     */
    protected function getColumnsCount(iterable $data): int
    {
        return count($this->getColumnsSettings($data));
    }

    /**
     * @param Worksheet $sheet
     * @param iterable $data
     */
    protected function setColumnDimensions(Worksheet $sheet, iterable $data): void
    {
        foreach ($this->getColumnsSettings($data) as $columnIndex => $columnSetting) {
            $columnNumber = $columnIndex + 1;
            if (array_key_exists('width', $columnSetting)) {
                $sheet->getColumnDimensionByColumn($columnNumber)->setWidth($columnSetting['width']);
            } else {
                $sheet->getColumnDimensionByColumn($columnNumber)->setAutoSize(true);
            }
        }
    }

    /**
     * @param iterable $data
     *
     * @return array
     */
    protected function getHeaderTitles(iterable $data): array
    {
        return array_map(static function (array $columnSettings) {
            return $columnSettings['title'];
        }, $this->getColumnsSettings($data));
    }

    protected function findImage($content, $column, $row): string
    {
        if (preg_match('/(<img[^>]+>)/i', $content, $imageMatches)) {
            $image = $imageMatches[1];

            $content = str_replace($image, '', $content);

            if (preg_match('/< *img[^>]*src *= *["\']?([^"\']*)/i', $image, $srcMatches)) {
                $src = $srcMatches[1];
                [$imageInfo, $imageData] = explode(',', $src);

                if (preg_match('/\/(.*);/', $imageInfo, $extMatches)) {
                    $extension = $extMatches[1];
                    $this->createImageFromData($imageData, $extension, $column, $row);
                }
            }
        }

        return $content;
    }

    protected function createImageFromData($data, $extension, $column, $row)
    {
        $image = imagecreatefromstring(base64_decode($data));

        $postfix = $this->getImageOptionPostfix($extension);

        $drawing = new MemoryDrawing();
        $drawing->setImageResource($image);
        $drawing->setRenderingFunction(constant(MemoryDrawing::class . '::RENDERING_' . $postfix));
        $drawing->setMimeType(constant(MemoryDrawing::class . '::MIMETYPE_' . $postfix));

        $this->images[$row][$column] = $drawing;
    }

    private function getImageOptionPostfix(string $extension): string
    {
        if (in_array($extension, ['jpg', 'jpeg'])) {
            return 'JPEG';
        }

        if (in_array($extension, ['gif', 'png'])) {
            return mb_strtoupper($extension);
        }

        return 'DEFAULT';
    }

    protected function applyColumnStyle($settings, Worksheet $sheet, $column)
    {
        $columnNumber = $column + 1;
        $maxRow = $sheet->getHighestRow();

        if (array_key_exists('style', $settings)) {
            $style = is_callable($settings['style']) ? $settings['style']($columnNumber) : $settings['style'];

            if (!is_array($style)) {
                throw new \RuntimeException(sprintf(
                    'Style for column [%s] should be array or callable that return array',
                    json_encode($settings)
                ));
            }
            $sheet->getStyleByColumnAndRow($columnNumber, $this->getTableFirstRowNumber(), $columnNumber, $maxRow)
                ->applyFromArray($style);
        }

        if (array_key_exists('numberFormat', $settings)) {
            $sheet->getStyleByColumnAndRow($columnNumber, $this->getTableFirstRowNumber(), $columnNumber, $maxRow)
                ->getNumberFormat()->setFormatCode($settings['numberFormat']);
        }

        if (array_key_exists('conditionalStyle', $settings)) {
            $sheet->getStyleByColumnAndRow($columnNumber, $this->getTableFirstRowNumber(), $columnNumber, $maxRow)
                ->setConditionalStyles($settings['conditionalStyle']);
        }
    }

    protected function appendImages($column, $row, Worksheet $sheet)
    {
        if (isset($this->images[$row][$column])) {
            $rowDimension = $sheet->getRowDimension($row);
            $height = $rowDimension->getRowHeight() > 0 ? $rowDimension->getRowHeight() + self::CELL_PADDING : self::CELL_PADDING;
            $coordinates = $sheet->getCellByColumnAndRow($column, $row)->getCoordinate();
            /** @var Drawing $drawing */
            $drawing = $this->images[$row][$column];
            $drawing->setCoordinates($coordinates);
            $drawing->setOffsetY($height);
            $drawing->setOffsetX(self::CELL_PADDING);
            $drawing->setWorksheet($sheet);
            $rowDimension->setRowHeight($height + $drawing->getHeight());
        }
    }
}
