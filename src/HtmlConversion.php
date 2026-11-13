<?php

namespace WrkLst\DocxMustache;

/**
 * Turns the few supported HTML tags inside a text run into OpenXML runs.
 * Works on strings on purpose: the document has to stay byte-identical everywhere else.
 */
class HtmlConversion
{
    private const LINE_BREAKS = ['&lt;br /&gt;', '&lt;br/&gt;', '&lt;br&gt;', '<br />', '<br/>', '<br>'];

    private const TEXT_RUN_OPEN = '<w:t xml:space="preserve">';

    /**
     * @param  string  $value
     */
    public static function convert($value)
    {
        $value = str_replace(self::LINE_BREAKS, '</w:t><w:br/>'.self::TEXT_RUN_OPEN, $value);

        foreach (['b', 'i', 'u', 'strong', 'em'] as $tag) {
            $value = self::convertHtmlToOpenXMLTag($value, $tag);
        }

        return $value;
    }

    public static function convertHtmlToOpenXMLTag($value, $tag = 'b')
    {
        // OpenXML run property of the tag, and the property of the surrounding run it overrides
        [$tag_ooxml, $loose_formatting] = match ($tag) {
            'u'           => ['u w:val="single" ', ''],
            'i', 'em'     => ['i ', '<w:i w:val="0"/>'],
            default       => ['b ', ''],
        };

        while (count($open = explode('<'.$tag.'>', $value, 2)) > 1) {
            [$before, $rest] = $open;
            $close = explode('</'.$tag.'>', $rest, 2);

            // styling of the run the tag sits in
            $wrPr_open = strrpos($before, '<w:rPr>');
            $wrPr_close = strrpos($before, '</w:rPr>', $wrPr_open);
            $style = substr($before, ($wrPr_open + 7), ($wrPr_close - ($wrPr_open + 7)));

            $value = $before
                // close the run, open a styled one for the tagged text
                .'</w:t></w:r>'
                .'<w:r><w:rPr><w:'.$tag_ooxml.'/>'.str_replace($loose_formatting, '', $style).'</w:rPr>'.self::TEXT_RUN_OPEN
                .$close[0]
                // close it and continue with the styling from before
                .'</w:t></w:r>'
                .'<w:r><w:rPr>'.$style.'</w:rPr>'.self::TEXT_RUN_OPEN
                .($close[1] ?? '');
        }

        return str_replace('<w:t>', self::TEXT_RUN_OPEN, $value);
    }
}
