<?php

namespace Staffim\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use Staffim\SpreadsheetBuilder\Converter\RichTextElementConverterInterface;

class RichTextToHtmlConverter
{
    /**
     * @var RichTextElementConverterInterface[]
     */
    private $converters = [];

    /**
     * @param RichTextElementConverterInterface[] $converters
     */
    public function __construct(array $converters = [])
    {
        foreach ($converters as $converter) {
            $this->registerConverter($converter);
        }
    }

    public function registerConverter(RichTextElementConverterInterface $converter)
    {
        $this->converters[] = $converter;
    }

    /**
     * @return RichTextElementConverterInterface[]
     */
    public function getConverters(): array
    {
        return $this->converters;
    }

    public function covertToHtml($value): string
    {
        $result = '';
        if (!$value instanceof RichText) {
            return nl2br(trim($value));
        }

        foreach ($value->getRichTextElements() as $richTextElement) {
            $style = [];
            foreach ($this->getConverters() as $converter) {
                if (!$converter->matchTextElement($richTextElement)) {
                    continue;
                }

                $style = array_merge($style, $converter->toHtmlStyle($richTextElement));
            }

            $textContent = $richTextElement->getText();
            $result .= nl2br(count($style) > 0 ? sprintf('<span style="%s">%s</span>', $this->buildStyleString($style), $textContent) : $textContent);
        }

        return trim($result);
    }

    public function convertFromHtml(string $string): RichText
    {
        $result = new RichText();

        $html = '<?xml encoding="utf-8" ?><body>' . $this->prepareNewLines($string) . '</body>';
        $document = new \DOMDocument();
        $document->loadHtml($html);

        $body = $document->getElementsByTagName('body')->item(0);

        $parts = [];
        foreach ($body->childNodes as $childNode) {
            $parts = array_merge($parts, $this->createTextElement($childNode));
        }

        if (count($parts) > 0) {
            $result->setRichTextElements($parts);
        } else {
            $result->createText('');
        }

        return $result;
    }

    protected function createTextElement(\DOMNode $node, string $parentName = '', array $parentStyle = []): array
    {
        $result = [];

        if ($node->hasChildNodes()) {
            $style = $this->getNodeStyle($node);
            foreach ($node->childNodes as $childNode) {
                $result = array_merge($result, $this->createTextElement($childNode, $node->nodeName, $style));
            }
        } else {
            if (!$node->textContent) {
                return $result;
            }
            $run = new Run($node->textContent);
            foreach ($this->getConverters() as $converter) {
                if ($converter->matchHtml(mb_strtolower($parentName), $parentStyle)) {
                    $converter->buildTextElement($run, $parentStyle);
                }
            }

            $result = [$run];
        }

        return $result;
    }

    protected function getNodeStyle(\DOMNode $node): array
    {
        if (!$node->hasAttributes()) {
            return [];
        }

        $styleAttribute = $node->attributes->getNamedItem('style');
        if (!$styleAttribute) {
            return [];
        }

        return $this->parseStyleFromString($styleAttribute->textContent);
    }

    protected function parseStyleFromString(string $style): array
    {
        $result = [];

        foreach (explode(';', mb_strtolower($style)) as $item) {
            if (strpos($item, ':') === false) {
                continue;
            }
            list($name, $value) = explode(':', $item);
            $result[trim($name)] = trim($value);
        }

        return $result;
    }

    protected function buildStyleString(array $style): string
    {
        return implode(
            '; ',
            array_map(
                static function (string $name, string $value) {
                    return $value . ':' . $name;
                },
                $style,
                array_keys($style)
            )
        );
    }

    protected function prepareNewLines(string $value): string
    {
        return preg_replace(['#<br\s*/?>\n#i', '#<br\s*/?>#i'], "\n", $value);
    }
}
