<?php

namespace Staffim\SpreadsheetBuilder\Converter;

use PhpOffice\PhpSpreadsheet\RichText\ITextElement;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Style\Font;

class UnderlineConverter implements RichTextElementConverterInterface
{
    public function matchTextElement(ITextElement $element): bool
    {
        return $element->getFont()->getUnderline() === Font::UNDERLINE_SINGLE;
    }

    public function toHtmlStyle(ITextElement $element): array
    {
        return [
            'text-decoration' => 'underline',
        ];
    }

    public function matchHtml(string $elementName, array $elementStyle = []): bool
    {
        return array_key_exists('text-decoration', $elementStyle) && $elementStyle['text-decoration'] === 'underline';
    }

    public function buildTextElement(Run $element, array $elementStyle = []): Run
    {
        $element->getFont()->setUnderline(Font::UNDERLINE_SINGLE);

        return $element;
    }
}
