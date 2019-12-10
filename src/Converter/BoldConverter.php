<?php

namespace Staffim\SpreadsheetBuilder\Converter;

use PhpOffice\PhpSpreadsheet\RichText\ITextElement;
use PhpOffice\PhpSpreadsheet\RichText\Run;

class BoldConverter implements RichTextElementConverterInterface
{
    public function matchTextElement(ITextElement $element): bool
    {
        return $element->getFont()->getBold();
    }


    public function toHtmlStyle(ITextElement $element): array
    {
        return [
            'font-weight' => 'bold',
        ];
    }

    public function matchHtml(string $elementName, array $elementStyle = []): bool
    {
        if (in_array($elementName, ['b', 'strong'])) {
            return true;
        }

        return array_key_exists('font-weight', $elementStyle) && $elementStyle['font-weight'] === 'bold';
    }

    public function buildTextElement(Run $element, array $elementStyle = []): Run
    {
        $element->getFont()->setBold(true);

        return $element;
    }
}
