<?php

namespace WrkLst\DocxMustache;

class MustacheRender
{
    public function Render($items, $mustache_template, $clean_tags = true)
    {
        if ($clean_tags) {
            $mustache_template = $this->TagCleaner($mustache_template);
        }

        $m = new \Mustache_Engine(array('escape' => function($value) {
            if (str_replace('*[[DONOTESCAPE]]*', '', $value) != $value) {
                            return str_replace('*[[DONOTESCAPE]]*', '', $value);
            }
            return htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
        }));
        return $m->render($mustache_template, $items);
    }

    protected function TagCleaner($content)
    {
        //kills all xml tags within curly mustache brackets
        //this is necessary, as word might produce unnecesary xml tage inbetween curly backets.
        return preg_replace_callback(
            '/{{(.*?)}}/',
            function($match) {
                return strip_tags($match[0]);
            },
            $content
        );
    }
}