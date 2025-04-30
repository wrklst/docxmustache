<?php

namespace WrkLst\DocxMustache;

class HtmlConversion
{
    /**
     * @param string $value
     */
    public static function convert($value)
    {
        $line_breaks = ['&lt;br /&gt;', '&lt;br/&gt;', '&lt;br&gt;', '<br />', '<br/>', '<br>'];
        $value = str_replace($line_breaks, '</w:t><w:br/><w:t xml:space="preserve">', $value);

        $value = self::convertHtmlToOpenXMLTag($value);

        return $value;
    }

    public static function convertHtmlToOpenXMLTag($html, $tag = null)
    {
        $doc = new \DOMDocument();
        libxml_use_internal_errors(true); // suppress warnings for malformed HTML
        $doc->loadHTML(mb_convert_encoding('<div>' . $html . '</div>', 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();

        $body = $doc->getElementsByTagName('div')->item(0);
        $result = self::convertNodeToOpenXML($body);

        return $result;
    }

    private static function convertNodeToOpenXML(\DOMNode $node, $inheritedStyles = [])
    {
        $output = '';

        foreach ($node->childNodes as $child) {
            if ($child instanceof \DOMText) {
                $text = htmlspecialchars($child->nodeValue, ENT_QUOTES | ENT_XML1, 'UTF-8');
                $output .= self::buildRun($text, $inheritedStyles);
            } elseif ($child instanceof \DOMElement) {
                $tag = strtolower($child->nodeName);
                $newStyles = $inheritedStyles;

                switch ($tag) {
                    case 'b':
                    case 'strong':
                        $newStyles['b'] = true;
                        break;
                    case 'i':
                    case 'em':
                        $newStyles['i'] = true;
                        break;
                    case 'u':
                        $newStyles['u'] = true;
                        break;
                    case 'br':
                        $output .= '</w:t><w:br/><w:t xml:space="preserve">';
                        continue 2; // skip child recursion
                    default:
                        // For unknown tags, recurse without adding formatting
                        break;
                }

                $output .= self::convertNodeToOpenXML($child, $newStyles);
            }
        }

        return $output;
    }

    private static function buildRun($text, $styles = [])
    {
        $rPr = '';

        if (!empty($styles['b'])) {
            $rPr .= '<w:b/>';
        }
        if (!empty($styles['i'])) {
            $rPr .= '<w:i/>';
        }
        if (!empty($styles['u'])) {
            $rPr .= '<w:u w:val="single"/>';
        }

        return '</w:t></w:r><w:r><w:rPr>' . $rPr . '</w:rPr><w:t xml:space="preserve">' . $text;
    }
}
