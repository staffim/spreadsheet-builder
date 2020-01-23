<?php

namespace Staffim\Tests;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PHPUnit\Framework\TestCase;
use Staffim\SpreadsheetBuilder\Builder;
use Staffim\SpreadsheetBuilder\TemplateWorksheetBuilder;
use Staffim\Tests\Model\TestTableWorksheetBuilder;

class WorksheetBuilderTestextends extends TestCase
{
    public function testBuildByTemplate()
    {
        $cells = [
            'A3' => 'Title for tests',
            'C3' => 'Content for tests  
Next line',
            'H4' => 'Test author',
            'F3' => '',
            'J3' => '{content}',
        ];

        $data = [
            'title' => 'Title for tests',
            'content' => 'Content for tests',
            'author' => 'Test author',
        ];

        $builder = new Builder();
        $builder->registerWorksheetBuilder(new TemplateWorksheetBuilder(__DIR__ . '/Resources/rich_payment_template.xlsx', 'payment', 'A1:I4'));
        $spreadsheet = $builder->build([$data]);

        foreach ($cells as $coordinates => $cell) {
            $this->assertEquals($cell, $spreadsheet->getSheetByName('worksheet title')->getCell($coordinates)->getFormattedValue());
        }
    }

    public function testBuildByTable()
    {
        $headers = ['id', 'name', 'phone', 'email'];
        $table = [
            [1, 'Богдан', '79998887766', 'bogdan@example.com'],
            [2, 'Янина', '79993332211', 'yanina@example.com'],
            [3, 'Игнат', '79995551133', 'ignat@example.com'],
        ];

        $builder = new Builder();
        $builder->registerWorksheetBuilder(new TestTableWorksheetBuilder());
        $spreadsheet = $builder->build([$table]);

        $worksheet = $spreadsheet->getSheetByName('test worksheet');
        foreach ($worksheet->getRowIterator(1, 4) as $rowIndex => $row) {
            /** @var Cell $cell */
            $columnNumber = 0;
            foreach ($row->getCellIterator('A', 'D') as $columnIndex => $cell) {
                $this->assertEquals(array_merge([$headers], $table)[$rowIndex - 1][$columnNumber], $cell->getFormattedValue());
                $columnNumber++;
            }
        }
    }
}
