<?php

namespace WrkLst\DocxMustache;

class MustacheRender
{
    public static function Render($items, $mustache_template, $clean_tags = true)
    {
        if ($clean_tags) {
            $mustache_template = self::TagCleaner($mustache_template);
        }

        $m = new \Mustache_Engine(['escape' => function ($value) {
            if (str_replace('*[[DONOTESCAPE]]*', '', $value) != $value) {
                $value = str_replace('&', '&amp;', $value);

                return str_replace('*[[DONOTESCAPE]]*', '', $value);
            }

            return htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
        }]);

        return $m->render($mustache_template, $items);
    }

    /**
     * Clean mustache tags that Word has split across multiple <w:r> runs.
     * Uses DOM parsing to safely shift text between <w:t> elements
     * without removing or corrupting XML structure.
     */
    public static function TagCleaner($content)
    {
        $doc = new \DOMDocument();
        $doc->loadXML($content);

        $xpath = new \DOMXPath($doc);
        $xpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        $textNodes = $xpath->query('//w:t');

        $nodes = [];
        for ($i = 0; $i < $textNodes->length; $i++) {
            $nodes[] = $textNodes->item($i);
        }

        $i = 0;
        while ($i < count($nodes)) {
            $text = $nodes[$i]->textContent;

            if (self::hasIncompleteMustacheTag($text)) {
                $j = $i + 1;
                while ($j < count($nodes) && self::hasIncompleteMustacheTag($text)) {
                    $text .= $nodes[$j]->textContent;
                    $nodes[$j]->textContent = '';
                    $j++;
                }
                $nodes[$i]->textContent = $text;
                if (!$nodes[$i]->hasAttribute('xml:space')) {
                    $nodes[$i]->setAttribute('xml:space', 'preserve');
                }
            }

            $i++;
        }

        return $doc->saveXML();
    }

    /**
     * Check if text contains an incomplete mustache tag that needs
     * content from subsequent <w:t> elements to be complete.
     */
    private static function hasIncompleteMustacheTag($text)
    {
        $len = strlen($text);
        if ($len === 0) {
            return false;
        }

        // Ends with a single { (potential start of {{)
        if ($text[$len - 1] === '{' && ($len < 2 || $text[$len - 2] !== '{')) {
            return true;
        }

        // Check for {{ without a matching }}
        $searchFrom = 0;
        while (($pos = strpos($text, '{{', $searchFrom)) !== false) {
            $closePos = strpos($text, '}}', $pos + 2);
            if ($closePos === false) {
                return true;
            }
            $searchFrom = $closePos + 2;
        }

        return false;
    }
}
