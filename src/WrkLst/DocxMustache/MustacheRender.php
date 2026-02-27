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
     * Finds <w:t> elements via regex, applies merge logic, and writes back
     * only the changed text via substr_replace — preserving the original
     * XML byte-for-byte everywhere else.
     */
    public static function TagCleaner($content)
    {
        // Find all <w:t> elements with their text and positions
        if (!preg_match_all('/<w:t([^>]*)>([^<]*)<\/w:t>/', $content, $matches, PREG_OFFSET_CAPTURE)) {
            return $content;
        }

        // Extract raw text content (already XML-encoded, but { } are never encoded)
        $texts = [];
        for ($i = 0; $i < count($matches[0]); $i++) {
            $texts[$i] = $matches[2][$i][0];
        }

        // Apply merge logic: shift text from later elements into earlier ones
        $newTexts = $texts;
        $i = 0;
        while ($i < count($newTexts)) {
            if (self::hasIncompleteMustacheTag($newTexts[$i])) {
                $j = $i + 1;
                while ($j < count($newTexts) && self::hasIncompleteMustacheTag($newTexts[$i])) {
                    $newTexts[$i] .= $newTexts[$j];
                    $newTexts[$j] = '';
                    $j++;
                }
            }
            $i++;
        }

        // Apply changes in reverse order so offsets stay valid
        for ($i = count($matches[0]) - 1; $i >= 0; $i--) {
            if ($texts[$i] === $newTexts[$i]) {
                continue;
            }

            $attrs = $matches[1][$i][0];
            if ($newTexts[$i] !== '' && strpos($attrs, 'xml:space') === false) {
                $attrs = ' xml:space="preserve"' . $attrs;
            }

            $replacement = '<w:t' . $attrs . '>' . $newTexts[$i] . '</w:t>';
            $content = substr_replace($content, $replacement, $matches[0][$i][1], strlen($matches[0][$i][0]));
        }

        return $content;
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
