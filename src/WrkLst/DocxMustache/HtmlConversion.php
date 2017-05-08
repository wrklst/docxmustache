<?php

namespace WrkLst\DocxMustache;

use Exception;
use Illuminate\Support\Facades\Log;

class HtmlConversion
{
    /**
     * @param string $value
     */
    public function convert($value)
    {
        $line_breaks = array("&lt;br /&gt;", "&lt;br/&gt;", "&lt;br&gt;", "<br />", "<br/>", "<br>");
        $value = str_replace($line_breaks, '<w:br/>', $value);

        $value = $this->convertHtmlToOpenXMLTag($value, "b");
        $value = $this->convertHtmlToOpenXMLTag($value, "i");
        $value = $this->convertHtmlToOpenXMLTag($value, "u");

        return $value;
    }

    protected function convertHtmlToOpenXMLTag($value, $tag = "b")
    {
        $value_array = array();
        $run_again = false;
        //this could be used instead if html was already escaped
        /*
        $bo = "&lt;";
        $bc = "&gt;";
        */
        $bo = "<";
        $bc = ">";

        //get first BOLD
        $tag_open_values = explode($bo.$tag.$bc, $value, 2);

        if (count($tag_open_values) > 1)
        {
            //save everything before the bold and close it
            $value_array[] = $tag_open_values[0];
            $value_array[] = '</w:t></w:r>';

            //define styling parameters
            $wrPr_open = strrpos($tag_open_values[0], '<w:rPr>');
            $wrPr_close = strrpos($tag_open_values[0], '</w:rPr>', $wrPr_open);
            $neutral_style = '<w:r><w:rPr>'.substr($tag_open_values[0], ($wrPr_open + 7), ($wrPr_close - ($wrPr_open + 7))).'</w:rPr><w:t>';
            $tagged_style = '<w:r><w:rPr><w:'.$tag.'/>'.substr($tag_open_values[0], ($wrPr_open + 7), ($wrPr_close - ($wrPr_open + 7))).'</w:rPr><w:t>';

            //open new text run and make it bold, include previous styling
            $value_array[] = $tagged_style;
            //get everything before bold close and after
            $tag_close_values = explode($bo.'/'.$tag.$bc, $tag_open_values[1], 2);
            //add bold text
            $value_array[] = $tag_close_values[0];
            //close bold run
            $value_array[] = '</w:t></w:r>';
            //open run for after bold
            $value_array[] = $neutral_style;
            $value_array[] = $tag_close_values[1];

            $run_again = true;
        } else {
            $value_array[] = $tag_open_values[0];
        }

        $value = implode('', $value_array);

        if ($run_again) {
                    $value = $this->convertHtmlToOpenXMLTag($value, $tag);
        }

        return $value;
    }
}
