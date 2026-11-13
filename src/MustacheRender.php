<?php

namespace WrkLst\DocxMustache;

class MustacheRender
{
    /** Prefix of a value that already is valid OpenXML/HTML and must not be escaped. */
    private const DO_NOT_ESCAPE = '*[[DONOTESCAPE]]*';

    public static function render($items, $mustache_template, $clean_tags = true)
    {
        if ($clean_tags) {
            $mustache_template = self::tagCleaner($mustache_template);
        }

        $m = new \Mustache_Engine(['escape' => self::escape(...)]);

        return $m->render($mustache_template, $items);
    }

    /**
     * Escape a value for the XML of the document.
     *
     * Mustache hands over whatever a tag resolves to. Scalars and null are cast to
     * string. An array means a tag points at non-scalar data (e.g. a list used
     * without a section); it renders as a visible marker listing the array's first
     * keys, so the broken template tag is obvious in the generated document.
     */
    private static function escape($value)
    {
        if (is_array($value)) {
            $keys = array_slice(array_keys($value), 0, 5);
            $value = 'Array('.implode(', ', $keys).(count($value) > 5 ? ', ...' : '').') Value used';
        } elseif (! is_string($value)) {
            $value = is_scalar($value) ? (string) $value : '';
        }

        if (str_contains($value, self::DO_NOT_ESCAPE)) {
            return str_replace([self::DO_NOT_ESCAPE, '&'], ['', '&amp;'], $value);
        }

        return htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
    }

    /**
     * Clean mustache tags that Word has split across multiple <w:r> runs.
     * Finds <w:t> elements via regex, applies merge logic, and writes back
     * only the changed text via substr_replace — preserving the original
     * XML byte-for-byte everywhere else.
     */
    public static function tagCleaner($content)
    {
        // Find all <w:t> elements with their text and positions
        if (! preg_match_all('/<w:t([^>]*)>([^<]*)<\/w:t>/', $content, $matches, PREG_OFFSET_CAPTURE)) {
            return $content;
        }

        // Raw text content (already XML-encoded, but { } are never encoded)
        $texts = array_column($matches[2], 0);

        // Merge logic: shift text from later elements into earlier ones
        $newTexts = $texts;
        $count = count($newTexts);
        for ($i = 0; $i < $count; $i++) {
            for ($j = $i + 1; $j < $count && self::hasIncompleteMustacheTag($newTexts[$i]); $j++) {
                $newTexts[$i] .= $newTexts[$j];
                $newTexts[$j] = '';
            }
        }

        // Apply changes in reverse order so offsets stay valid
        for ($i = $count - 1; $i >= 0; $i--) {
            if ($texts[$i] === $newTexts[$i]) {
                continue;
            }

            $attrs = $matches[1][$i][0];
            if ($newTexts[$i] !== '' && ! str_contains($attrs, 'xml:space')) {
                $attrs = ' xml:space="preserve"'.$attrs;
            }

            $replacement = '<w:t'.$attrs.'>'.$newTexts[$i].'</w:t>';
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
        // Ends with a single { (potential start of {{)
        if (str_ends_with($text, '{') && ! str_ends_with($text, '{{')) {
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
