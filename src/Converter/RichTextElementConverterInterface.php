<?php

namespace Staffim\SpreadsheetBuilder\Converter;

use PhpOffice\PhpSpreadsheet\RichText\ITextElement;
use PhpOffice\PhpSpreadsheet\RichText\Run;

interface RichTextElementConverterInterface
{
    public function matchTextElement(ITextElement $element): bool;

    public function matchHtml(string $elementName, array $elementStyle = []): bool;

    public function toHtmlStyle(ITextElement $element): array;

    public function buildTextElement(Run $element, array $elementStyle = []): Run;
}
