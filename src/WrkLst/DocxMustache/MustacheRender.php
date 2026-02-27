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
        // or language attributes change mid-tag. We need to merge these back.

        // Step 1: Merge split opening braces: {</w:t>...<w:t>{ → {{
        $content = preg_replace(
            '/\{(?!\{)<\/w:t>[\s\S]*?<w:t[^>]*>\{/',
            '{{',
            $content
        );

        // Step 2: Merge split closing braces: }</w:t>...<w:t>} → }}
        $content = preg_replace(
            '/\}(?!\})<\/w:t>[\s\S]*?<w:t[^>]*>\}/',
            '}}',
            $content
        );

        // Step 3: Handle mustache tags where content spans multiple XML runs.
        // Match {{ through any intermediate XML to }} and strip XML from inner content.
        $content = preg_replace_callback(
            '/(\{\{)([\s\S]*?)(\}\})/',
            function ($match) {
                return $match[1] . strip_tags($match[2]) . $match[3];
            },
            $content
        );

        return $content;
    }
}
