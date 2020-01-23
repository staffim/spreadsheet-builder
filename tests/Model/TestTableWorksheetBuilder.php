<?php

namespace Staffim\Tests\Model;

use Staffim\SpreadsheetBuilder\AbstractTableWorksheetBuilder;

class TestTableWorksheetBuilder extends AbstractTableWorksheetBuilder
{
    public function getTableTitle(iterable $data): string
    {
        return 'test table';
    }

    public function getWorksheetTitle(iterable $data): string
    {
        return 'test worksheet';
    }

    protected function getColumnsSettings(iterable $data): array
    {
        return [
            [
                'title' => 'id',
                'value' => static function ($item) {
                    return $item[0];
                },
            ],
            [
                'title' => 'name',
                'value' => static function ($item) {
                    return $item[1];
                },
            ],
            [
                'title' => 'phone',
                'value' => static function ($item) {
                    return $item[2];
                },
            ],
            [
                'title' => 'email',
                'value' => static function ($item) {
                    return $item[3];
                },
            ],
        ];
    }
}
