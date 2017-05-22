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
                return str_replace('*[[DONOTESCAPE]]*', '', $value);
            }

            return htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
        }]);

        return $m->render($mustache_template, $items);
    }

    public static function TagCleaner($content)
    {
        //kills all xml tags within curly mustache brackets
        //this is necessary, as word might produce unnecesary xml tags inbetween curly backets.

        //this regex needs either to be improved or it needs to be replace with a method that is aware of the xml
        // as the regex can mess up the xml badly if the pattern does not coem with the expected content

        return preg_replace_callback(
            '/{{(.*?)}}/',
            function ($match) {
                return strip_tags($match[0]);
            },
            preg_replace("/(?<!{){(?!{)<\/w:t>[\s\S]*?<w:t>{/", "{{", $content)
        );

        /*
        {{...value with tags...}}  --> {{value without tags}}
        <w:t>{</w:t>...<w:t>{  --> <w:t>{{
        }}{</w:t>...<w:t>{  --> }}{{
        */
    }
}
