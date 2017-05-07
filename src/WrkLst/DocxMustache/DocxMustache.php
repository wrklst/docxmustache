<?PHP

namespace WrkLst\DocxMustache;

use Exception;
use Illuminate\Support\Facades\Log;

//Custom DOCX template class to change content based on mustache templating engine.
class DocxMustache
{
    public $items;
    public $word_doc;
    public $template_file_name;
    public $template_file;
    public $local_path;
    public $storageDisk;
    public $storagePathPrefix;
    public $zipper;
    public $imageManipulation;
    public $verbose;

    public function __construct($items, $local_template_file)
    {
        $this->items = $items;
        $this->template_file_name = basename($local_template_file);
        $this->template_file = $local_template_file;
        $this->word_doc = false;
        $this->zipper = new \Chumper\Zipper\Zipper;

        //name of disk for storage
        $this->storageDisk = 'local';

        //prefix within your storage path
        $this->storagePathPrefix = 'app/';

        //if you use img urls that support manipulation via parameter
        $this->imageManipulation = ''; //'&w=1800';

        $this->verbose = false;
    }

    public function execute()
    {
        $this->copyTmplate();
        $this->readTeamplate();
    }

    /**
     * @param string $file
     */
    protected function storagePath($file)
    {
        return storage_path($file);
    }

    /**
     * @param string $msg
     */
    protected function log($msg)
    {
        //introduce logging method here to keep track of process
        // can be overwritten in extended class to log with custom preocess logger
        if ($this->verbose)
            Log::error($msg);
    }

    public function cleanUpTmpDirs()
    {
        $now = time();
        $isExpired = ($now - (60 * 240));
        $disk = \Storage::disk($this->storageDisk);
        $all_dirs = $disk->directories($this->storagePathPrefix.'DocxMustache');
        foreach ($all_dirs as $dir) {
            //delete dirs older than 20min
            if ($disk->lastModified($dir) < $isExpired)
            {
                $disk->deleteDirectory($dir);
            }
        }
    }

    public function getTmpDir()
    {
        $this->cleanUpTmpDirs();
        $path = $this->storagePathPrefix.'DocxMustache/'.uniqid($this->template_file).'/';
        \File::makeDirectory($this->storagePath($path), 0775, true);
        return $path;
    }

    public function copyTmplate()
    {
        $this->log('Get Copy of Template');
        $this->local_path = $this->getTmpDir();
        \Storage::disk($this->storageDisk)->copy($this->storagePathPrefix.$this->template_file, $this->local_path.$this->template_file_name);
    }

    public function readTeamplate()
    {
        $this->log('Analyze Template');
        //get the main document out of the docx archive
        $this->zipper->make($this->storagePath($this->local_path.$this->template_file_name))
            ->extractTo($this->storagePath($this->local_path), array('word/document.xml'), \Chumper\Zipper\Zipper::WHITELIST);

        //if the main document exists
        if ($this->word_doc = \Storage::disk($this->storageDisk)->get($this->local_path.'word/document.xml'))
        {
            $this->log('Merge Data into Template');
            $this->word_doc = $this->MustacheTagCleaner($this->word_doc);
            $this->word_doc = $this->MustacheRender($this->items, $this->word_doc);
            $this->word_doc = $this->convertHtmlToOpenXML($this->word_doc);

            $this->ImageReplacer();

            $this->log('Compact Template with Data');
            //store new content
            \Storage::disk($this->storageDisk)
                ->put($this->local_path.'word/document.xml', $this->word_doc);
            //add new content to word doc
            $this->zipper->folder('word')
                ->add($this->storagePath($this->local_path.'word/document.xml'))
                ->close();
        } else
        {
            throw new Exception('docx has no main xml doc.');
        }
    }

    protected function MustacheTagCleaner($content)
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

    protected function MustacheRender($items, $mustache_template)
    {
        $m = new \Mustache_Engine(array('escape' => function($value) {
            if (str_replace('*[[DONOTESCAPE]]*', '', $value) != $value) {
                            return str_replace('*[[DONOTESCAPE]]*', '', $value);
            }
            return htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
        }));
        return $m->render($mustache_template, $items);

    }

    protected function AddContentType($imageCt = "jpeg")
    {
        //get content type file from archive
        $this->zipper->make($this->storagePath($this->local_path.$this->template_file_name))
            ->extractTo($this->storagePath($this->local_path), array('[Content_Types].xml'), \Chumper\Zipper\Zipper::WHITELIST);

        // load content type file xml
        $ct_file = simplexml_load_file($this->storagePath($this->local_path.'[Content_Types].xml'));

        //check if content type for jpg has been set
        $i = 0;
        $ct_already_set = false;
        foreach ($ct_file as $ct)
        {
            if ((string) $ct_file->Default[$i]['Extension'] == $imageCt) {
                            $ct_already_set = true;
            }
            $i++;
        }

        //if content type for jpg has not been set, add it to xml
        // and save xml to file and add it to the archive
        if (!$ct_already_set)
        {
            $sxe = $ct_file->addChild('Default');
            $sxe->addAttribute('Extension', $imageCt);
            $sxe->addAttribute('ContentType', 'image/'.$imageCt);

            if ($ct_file_xml = $ct_file->asXML())
            {
                \Storage::disk($this->storageDisk)->put($this->local_path.'[Content_Types].xml', $ct_file_xml);
                $this->zipper->add($this->storagePath($this->local_path.'[Content_Types].xml'));
            } else
            {
                throw new Exception('Cannot generate xml for [Content_Types].xml.');
            }
        }
    }

    protected function FetchReplaceableImages(&$main_file, $ns)
    {
        //set up basic arrays to keep track of imgs
        $imgs = array();
        $imgs_replaced = array(); // so they can later be removed from media and relation file.
        $newIdCounter = 1;

        //iterate through all drawing containers of the xml document
        foreach ($main_file->xpath('//w:drawing') as $k=>$drawing)
        {
            $ueid = "wrklstId".$newIdCounter;
            $wasId = (string) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])["embed"];
            $imgs_replaced[$wasId] = $wasId;
            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])["embed"] = $ueid;

            $cx = (int) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()["cx"];
            $cy = (int) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()["cy"];

            //figure out if there is a URL saved in the description field of the img
            $img_url = $this->analyseImgUrlString((string) $drawing->children($ns['wp'])->xpath('wp:docPr')[0]->attributes()["descr"]);
            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->xpath('wp:docPr')[0]->attributes()["descr"] = $img_url["rest"];

            //check https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            // for EMUs calculation
            /*
            295px @72 dpi = 1530350 EMUs = Multiplier for 72dpi pixels 5187.627118644067797
            413px @72 dpi = 2142490 EMUs = Multiplier for 72dpi pixels 5187.627118644067797

            */

            //if there is a url, save this img as a img to be replaced
            if (trim($img_url["url"]))
            {
                $imgs[] = array(
                    "cx" => $cx,
                    "cy" => $cy,
                    "width" => (int) ($cx / 5187.627118644067797),
                    "height" => (int) ($cy / 5187.627118644067797),
                    "wasId" => $wasId,
                    "id" => $ueid,
                    "url" => $img_url["url"],
                );

                $newIdCounter++;
            }
        }
        return array(
            'imgs' => $imgs,
            'imgs_replaced' => $imgs_replaced
        );
    }

    protected function RemoveReplaceImages($imgs_replaced, &$rels_file)
    {
        //iterate through replaced images and clean rels files from them
        foreach ($imgs_replaced as $img_replaced)
        {
            $i = 0;
            foreach ($rels_file as $rel)
            {
                if ((string) $rel->attributes()['Id'] == $img_replaced)
                {
                    $this->zipper->remove('word/'.(string) $rel->attributes()['Target']);
                    unset($rels_file->Relationship[$i]);
                }
                $i++;
            }
        }
    }

    protected function InsertImages($ns, &$imgs, &$rels_file, &$main_file)
    {
        //define what images are allowed
        $allowed_imgs = array(
            'image/gif' => '.gif',
            'image/jpeg' => '.jpeg',
            'image/png' => '.png',
            'image/bmp' => '.bmp',
        );

        //iterate through replacable images
        foreach ($imgs as $k=>$img)
        {
            //get file type of img and test it against supported imgs
            if ($img_file_handle = fopen($img['url'].$this->imageManipulation, "rb"))
            {
                $img_data = stream_get_contents($img_file_handle);
                fclose($img_file_handle);
                $fi = new \finfo(FILEINFO_MIME);

                $image_mime = strstr($fi->buffer($img_data), ';', true);
                //dd($image_mime);
                if (isset($allowed_imgs[$image_mime]))
                {
                    $imgs[$k]['img_file_src'] = str_replace("wrklstId", "wrklst_image", $img['id']).$allowed_imgs[$image_mime];
                    $imgs[$k]['img_file_dest'] = str_replace("wrklstId", "wrklst_image", $img['id']).'.jpeg';

                    \Storage::disk($this->storageDisk)->put($this->local_path.'word/media/'.$imgs[$k]['img_file_src'], $img_data);

                    //rework img to new size and jpg format
                    $img_rework = \Image::make($this->storagePath($this->local_path.'word/media/'.$imgs[$k]['img_file_src']));
                    $w = $img['width'];
                    $h = $img['height'];
                    if ($w > $h) {
                                            $h = null;
                    } else {
                                            $w = null;
                    }
                    $img_rework->resize($w, $h, function($constraint) {
                        $constraint->aspectRatio();
                        $constraint->upsize();
                    });
                    $new_height = $img_rework->height();
                    $new_width = $img_rework->width();
                    $img_rework->save($this->storagePath($this->local_path.'word/media/'.$imgs[$k]['img_file_dest']));

                    $this->zipper->folder('word/media')->add($this->storagePath($this->local_path.'word/media/'.$imgs[$k]['img_file_dest']));

                    $sxe = $rels_file->addChild('Relationship');
                    $sxe->addAttribute('Id', $img['id']);
                    $sxe->addAttribute('Type', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                    $sxe->addAttribute('Target', "media/".$imgs[$k]['img_file_dest']);

                    //update height and width of image in document.xml
                    $new_height_emus = (int) ($new_height * 5187.627118644067797);
                    $new_width_emus = (int) ($new_width * 5187.627118644067797);
                    foreach ($main_file->xpath('//w:drawing') as $k=>$drawing)
                    {
                        if ($img['id'] == $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])["embed"])
                        {
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()["cx"] = $new_width_emus;
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()["cy"] = $new_height_emus;

                            //the following also changes the contraints of the container for the img.
                            // probably not wanted, as this will make images larger than the constraints of the placeholder
                            /*
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->inline->extent->attributes()["cx"] = $new_width_emus;
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->inline->extent->attributes()["cy"] = $new_height_emus;
                            */
                            break;
                        }
                    }
                }
            }
        }
    }

    protected function ImageReplacer()
    {
        $this->log('Merge Images into Template');

        //load main doc xml
        $main_file = simplexml_load_string($this->word_doc);

        //get all namespaces of the document
        $ns = $main_file->getNamespaces(true);

        $replaceableImage = $this->FetchReplaceableImages($main_file, $ns);
        $imgs = $replaceableImage['imgs'];
        $imgs_replaced = $replaceableImage['imgs_replaced'];


        //get relation xml file for img relations
        $this->zipper->make($this->storagePath($this->local_path.$this->template_file_name))
            ->extractTo($this->storagePath($this->local_path), array('word/_rels/document.xml.rels'), \Chumper\Zipper\Zipper::WHITELIST);

        //load img relations into xml
        $rels_file = simplexml_load_file($this->storagePath($this->local_path.'word/_rels/document.xml.rels'));

        $this->RemoveReplaceImages($imgs_replaced, $rels_file);

        //add jpg content type if not set
        $this->AddContentType('jpeg');

        $this->InsertImages($ns, $imgs, $rels_file, $main_file);

        if ($rels_file_xml = $rels_file->asXML())
        {
            \Storage::disk($this->storageDisk)->put($this->local_path.'word/_rels/document.xml.rels', $rels_file_xml);
            $this->zipper->folder('word/_rels')->add($this->storagePath($this->local_path.'word/_rels/document.xml.rels'));
        } else
        {
            throw new Exception('Cannot generate xml for word/_rels/document.xml.rels.');
        }

        if ($main_file_xml = $main_file->asXML())
        {
            $this->word_doc = $main_file_xml;
        } else
        {
            throw new Exception('Cannot generate xml for word/document.xml.');
        }
    }

    /**
     * @param string $string
     */
    protected function analyseImgUrlString($string)
    {
        $start = "[IMG-REPLACE]";
        $end = "[/IMG-REPLACE]";
        $string = ' '.$string;
        $ini = strpos($string, $start);
        if ($ini == 0)
        {
            $url = '';
            $rest = $string;
        } else
        {
            $ini += strlen($start);
            $len = ((strpos($string, $end, $ini)) - $ini);
            $url = substr($string, $ini, $len);

            $ini = strpos($string, $start);
            $len = strpos($string, $end, $ini + strlen($start)) + strlen($end);
            $rest = substr($string, 0, $ini).substr($string, $len);
        }
        return array(
            "url" => $url,
            "rest" => $rest,
        );
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

    /**
     * @param string $value
     */
    protected function convertHtmlToOpenXML($value)
    {
        $line_breaks = array("&lt;br /&gt;", "&lt;br/&gt;", "&lt;br&gt;", "<br />", "<br/>", "<br>");
        $value = str_replace($line_breaks, '<w:br/>', $value);

        $value = $this->convertHtmlToOpenXMLTag($value, "b");
        $value = $this->convertHtmlToOpenXMLTag($value, "i");
        $value = $this->convertHtmlToOpenXMLTag($value, "u");

        return $value;
    }

    public function saveAsPdf()
    {
        $this->log('Converting DOCX to PDF');
        //convert to pdf with libre office
        $command = "soffice --headless --convert-to pdf ".$this->storagePath($this->local_path.$this->template_file_name).' --outdir '.$this->storagePath($this->local_path);
        $process = new \Symfony\Component\Process\Process($command);
        $process->start();
        while ($process->isRunning()) {
            //wait until process is ready
        }
        // executes after the command finishes
        if (!$process->isSuccessful()) {
            throw new \Symfony\Component\Process\Exception\ProcessFailedException($process);
        } else
        {
            $path_parts = pathinfo($this->storagePath($this->local_path.$this->template_file_name));
            return $this->storagePath($this->local_path.$path_parts['filename'].'pdf');
        }
    }
}
