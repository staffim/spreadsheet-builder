# Xlsx spreadsheet builder

Small library for [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) that helps to build xlsx tables by configuring data columns.

#### Installation

```
composer require staffim/spreadsheet-builder
```

#### Examples

Creating worksheet builder
```php
<?php

namespace Acme\Xlsx;

use Staffim\SpreadsheetBuilder\AbstractWorksheetBuilder;
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

use Staffim\SpreadsheetBuilder\Builder;

/// ....

$builder = new Builder([
    new \Acme\Xlsx\UsersListWorksheetBuilder(),
]);

$users = $usersManager->fetchUsers();

$spreadsheet = $builder->build([$users]);

```

building from template xlsx with placeholders (this example replaced {title}, {content}, {foo} placeholders to $data array values in A1:D2 range of cells)

template before

![Template Before](https://i.ibb.co/qp2XQNX/ss-before.png)

```

$data = [
    'title' => 'Example title',
    'content' => 'Content for example',
    'foo' => 'Bar',
];

$builder = new Builder([
    new TemplateWorksheetBuilder('/path/from/template.xlsx', 'payment', 'A1:D2'),
]);
$spreadsheet = $builder->build([$data]);

```

Result after

![Template After](https://i.ibb.co/tzFFDmx/ss-after.png)

#### Working with html

Use `\Staffim\SpreadsheetBuilder\RichTextToHtmlConverter` for converting html to
`RichText` or vice versa:

```php
$converter = new \Staffim\SpreadsheetBuilder\RichTextToHtmlConverter(
     [
         new BoldConverter(),
         new ItalicConverter(),
         new UnderlineConverter(),
         new ColorConverter(),
     ]
 );
};

$html = '<span style="color: brown; font-weight: bold">bold<br/></span>
<i style="color: #ffcc01">italic</i><br/>
<b style="text-decoration: underline">underline111</b>
<span style="font-weight: bold; color: #AA0000">bold red</span>';

$richText = $converter->convertFromHtml($html);
```

See [tests](https://github.com/staffim/spreadsheet-builder/tree/master/tests) for more examples
