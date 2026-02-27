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

    public static function TagCleaner($content)
    {
        // Word splits mustache tags across multiple <w:r> runs when formatting
        // or language attributes change mid-tag. We need to reassemble them by
        // shifting text between <w:t> elements, never removing XML structure.

        // Step 1: Merge split opening braces {</w:t>...<w:t>{ → shift first { to second position
        // Only across a single run boundary (no other <w:t> elements in between)
        $content = preg_replace(
            '/\{(?!\{)<\/w:t>((?:(?!<\/?w:t[ >])[\s\S])*?)<w:t([^>]*)>\{/',
            '</w:t>$1<w:t$2>{{',
            $content
        );

        // Step 2: Merge split closing braces }</w:t>...<w:t>} → shift first } to second position
        // Only across a single run boundary
        $content = preg_replace(
            '/\}(?!\})<\/w:t>((?:(?!<\/?w:t[ >])[\s\S])*?)<w:t([^>]*)>\}/',
            '</w:t>$1<w:t$2>}}',
            $content
        );

        // Step 3: Handle mustache tags where tag name spans multiple XML runs.
        // Shift text from intermediate <w:t> elements into the first one, emptying them
        // but preserving all XML structure (runs, properties, etc.).
        $content = preg_replace_callback(
            '/(\{\{)([\s\S]*?)(\}\})/',
            function ($match) {
                // Only process if there are intermediate <w:t> elements
                if (strpos($match[2], '<w:t') === false) {
                    return $match[0];
                }
                $extractedText = '';
                $modified = preg_replace_callback(
                    '/<w:t([^>]*)>([^<]*)<\/w:t>/',
                    function ($m) use (&$extractedText) {
                        $extractedText .= $m[2];
                        return '<w:t' . $m[1] . '></w:t>';
                    },
                    $match[2]
                );
                return $match[1] . $extractedText . $match[3] . $modified;
            },
            $content
        );

        return $content;
    }
}
