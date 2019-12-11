<?php

namespace Staffim\SpreadsheetBuilder\Converter;

use PhpOffice\PhpSpreadsheet\RichText\ITextElement;
use PhpOffice\PhpSpreadsheet\RichText\Run;

class ItalicConverter implements RichTextElementConverterInterface
{
    public function matchTextElement(ITextElement $element): bool
    {
        $font = $element->getFont();

        return $font && $font->getItalic();
    }

    public function toHtmlStyle(ITextElement $element): array
    {
        return [
            'font-style' => 'italic',
        ];
    }

    public function matchHtml(string $elementName, array $elementStyle = []): bool
    {
        if (in_array($elementName, ['em', 'i'])) {
            return true;
        }

        return array_key_exists('font-style', $elementStyle) && $elementStyle['font-style'] === 'italic';
    }

    public function buildTextElement(Run $element, array $elementStyle = []): Run
    {
        $element->getFont()->setItalic(true);

        return $element;
    }
}
