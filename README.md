# spreadsheet-builder

Small library for [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) that helps to build xlsx tables by configuring data columns.


##### Examples

Creating worksheet builder
```php
<?php

namespace Acme\Xlsx;

use Experium\SpreadsheetBuilder\AbstractWorksheetBuilder;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class UsersListWorksheetBuilder extends AbstractWorksheetBuilder {

   public function getTableTitle(iterable $data): string
   {
       return sprintf('Users list at %s', (new \DateTime())->format('d.m.Y'));
   }

   public function getWorksheetTitle(iterable $data): string
   {
       return 'Users list';
   }

   protected function getColumnsSettings(iterable $data): array
   {
       return [
           [
               'title' => 'ID',
               'value' => static function (User $user) {
                   return $user->getId();
               },
               'width' => 5,
           ],
           [
               'title' => 'Email',
               'value' => static function (User $user) {
                   return $user->getEmail();
               },
               'width' => 20,
           ],
           [
              'title' => 'About',
              'value' => static function (User $user) {
                  return $user->getAbout();
              },
              'width' => 25,
              'style' => [// all available phpspreadsheet's styles
                'alignment' => [
                    'wrapText' => true,
                    'vertical' => Alignment::VERTICAL_TOP,
                ],
              ],
          ],
       ];
   }
}
```

building whole spreadsheet

```php

use Experium\SpreadsheetBuilder\Builder;

/// ....

$builder = new Builder([
    new \Acme\Xlsx\UsersListWorksheetBuilder(),
])

$users = $usersManager->fetchUsers();

$spreadsheet = $builder->build([$users]);

```
