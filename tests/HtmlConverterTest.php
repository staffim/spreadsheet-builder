<?php

namespace Staffim\Tests;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPUnit\Framework\TestCase;
use Staffim\SpreadsheetBuilder\Converter\BoldConverter;
use Staffim\SpreadsheetBuilder\Converter\ColorConverter;
use Staffim\SpreadsheetBuilder\Converter\ItalicConverter;
use Staffim\SpreadsheetBuilder\Converter\UnderlineConverter;
use Staffim\SpreadsheetBuilder\RichTextToHtmlConverter;

class HtmlConverterTest extends TestCase
{
    public function testRunnableConverter()
    {
        $spreadsheet = IOFactory::load(__DIR__ . '/resources/rich_text.xlsx');
        $worksheet = $spreadsheet->getActiveSheet();

        $html = $this->getConverter()->covertToHtml($worksheet->getCell('A1')->getValue());
        $this->assertEquals(
            trim(
                "<span style=\"text-decoration:underline\">Underlined<br />
</span><span style=\"font-style:italic; color:#CE181E\">Red cursive<br />
</span><span style=\"font-weight:bold\">Bold</span>"
            ),
            $html
        );
    }

    public function testBuildBoldText()
    {
        $text = 'before <span style="font-weight: bold">bold</span><b>BOLD</b> after';

        /**
         * @var RichText $richText
         */
        $richText = $this->getConverter()->convertFromHtml($text);

        $elements = $richText->getRichTextElements();
        $this->assertCount(4, $elements);

        $this->assertEquals($elements[0]->getText(), 'before ');
        $this->assertFalse($elements[0]->getFont()->getBold());

        $this->assertEquals($elements[1]->getText(), 'bold');
        $this->assertTrue($elements[1]->getFont()->getBold());

        $this->assertEquals($elements[2]->getText(), 'BOLD');
        $this->assertTrue($elements[2]->getFont()->getBold());

        $this->assertEquals($elements[3]->getText(), ' after');
        $this->assertFalse($elements[3]->getFont()->getBold());
    }

    public function testBuildColoredText()
    {
        $text = '<span>before <span style="font-weight: bold; color: #AA0000">bold red</span></span>';

        /**
         * @var RichText $richText
         */
        $richText = $this->getConverter()->convertFromHtml($text);

        $elements = $richText->getRichTextElements();
        $this->assertCount(2, $elements);

        $coloredElement = $elements[1];

        $this->assertEquals('AA0000', $coloredElement->getFont()->getColor()->getRGB());
    }

    public function testConversionConsistence()
    {
        $text = 'before <span style="font-weight: bold">bold</span><b>BOLD</b> after';

        $richText = $this->getConverter()->convertFromHtml($text);

        $convertedHtml = $this->getConverter()->covertToHtml($richText);
        $this->assertEquals('before <span style="font-weight:bold">bold</span><span style="font-weight:bold">BOLD</span> after', $convertedHtml);
        $secondRichText = $this->getConverter()->convertFromHtml($convertedHtml);

        $elements = $secondRichText->getRichTextElements();
        $this->assertCount(4, $elements);

        $this->assertEquals($richText, $secondRichText);
    }

    public function testBuildFileFromHtml()
    {
        $html = '<span style="color: brown; font-weight: bold">bold<br/></span><i style="color: #ffcc01">italic</i><br/>
<b style="text-decoration: underline">underline111</b>';

        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $richText = $this->getConverter()->convertFromHtml($html);

        $worksheet->getCell('A1')->setValue($richText);

        $writer = new Xlsx($spreadsheet);
        $resultFileName = sys_get_temp_dir() . '/' . 'test building.xlsx';
        $writer->save($resultFileName);
        // echo 'file saved in ' . $resultFileName;

        $this->assertNotNull($richText);
    }

    protected function getConverter()
    {
        return new RichTextToHtmlConverter(
            [
                new BoldConverter(),
                new ItalicConverter(),
                new UnderlineConverter(),
                new ColorConverter(),
            ]
        );
    }
}
