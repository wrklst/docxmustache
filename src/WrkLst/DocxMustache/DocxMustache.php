<?php

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
        $this->zipper = new \Chumper\Zipper\Zipper();

        //name of disk for storage
        $this->storageDisk = 'local';

        //prefix within your storage path
        $this->storagePathPrefix = 'app/';

        //if you use img urls that support manipulation via parameter
        $this->imageManipulation = ''; //'&w=1800';

        $this->verbose = false;
    }

    public function Execute()
    {
        $this->CopyTmplate();
        $this->ReadTeamplate();
    }

    /**
     * @param string $file
     */
    public function StoragePath($file)
    {
        return storage_path($file);
    }

    /**
     * @param string $msg
     */
    protected function Log($msg)
    {
        //introduce logging method here to keep track of process
        // can be overwritten in extended class to log with custom preocess logger
        if ($this->verbose) {
            Log::error($msg);
        }
    }

    public function CleanUpTmpDirs()
    {
        $now = time();
        $isExpired = ($now - (60 * 240));
        $disk = \Storage::disk($this->storageDisk);
        $all_dirs = $disk->directories($this->storagePathPrefix.'DocxMustache');
        foreach ($all_dirs as $dir) {
            //delete dirs older than 20min
            if ($disk->lastModified($dir) < $isExpired) {
                $disk->deleteDirectory($dir);
            }
        }
    }

    public function GetTmpDir()
    {
        $this->CleanUpTmpDirs();
        $path = $this->storagePathPrefix.'DocxMustache/'.uniqid($this->template_file).'/';
        \File::makeDirectory($this->StoragePath($path), 0775, true);

        return $path;
    }

    public function CopyTmplate()
    {
        $this->Log('Get Copy of Template');
        $this->local_path = $this->GetTmpDir();
        \Storage::disk($this->storageDisk)->copy($this->storagePathPrefix.$this->template_file, $this->local_path.$this->template_file_name);
    }

    protected function exctractOpenXmlFile($file)
    {
        $this->zipper->make($this->StoragePath($this->local_path.$this->template_file_name))
            ->extractTo($this->StoragePath($this->local_path), [$file], \Chumper\Zipper\Zipper::WHITELIST);
    }

    protected function ReadOpenXmlFile($file, $type = 'file')
    {
        $this->exctractOpenXmlFile($file);
        $this->tempChangeProblematicChars($file);

        if ($type == 'file') {
            if ($file_contents = \Storage::disk($this->storageDisk)->get($this->local_path.$file)) {
                return $file_contents;
            } else {
                throw new Exception('Cannot not read file '.$file);
            }
        } else {
            if ($xml_object = simplexml_load_file($this->StoragePath($this->local_path.$file))) {
                return $xml_object;
            } else {
                throw new Exception('Cannot load XML Object from file '.$file);
            }
        }
    }

    protected function SaveOpenXmlFile($file, $folder, $content)
    {
        $content = str_replace('%%KAUFUND%%','&',$content);

        \Storage::disk($this->storageDisk)
            ->put($this->local_path.$file, $content);
        //add new content to word doc
        if ($folder) {
            $this->zipper->folder($folder)
                ->add($this->StoragePath($this->local_path.$file));
        } else {
            $this->zipper
                ->add($this->StoragePath($this->local_path.$file));
        }
    }

    protected function SaveOpenXmlObjectToFile($xmlObject, $file, $folder)
    {
        if ($xmlString = $xmlObject->asXML()) {
            $this->SaveOpenXmlFile($file, $folder, $xmlString);
            $this->tempChangeProblematicChars($file, true);
        } else {
            throw new Exception('Cannot generate xml for '.$file);
        }
    }

    public function ReadTeamplate()
    {
        $this->Log('Analyze Template');
        //get the main document out of the docx archive
        $this->word_doc = $this->ReadOpenXmlFile('word/document.xml', 'file');

        $this->Log('Merge Data into Template');

        $this->word_doc = MustacheRender::render($this->items, $this->word_doc);

        $this->word_doc = HtmlConversion::convert($this->word_doc);

        $this->ImageReplacer();

        $this->Log('Compact Template with Data');

        $this->SaveOpenXmlFile('word/document.xml', 'word', $this->word_doc);
        $this->zipper->close();
    }

    protected function tempChangeProblematicChars($file, $reverse = false)
    {
        $content = \Storage::disk($this->storageDisk)
            ->get($this->local_path.$file);
        if(!$reverse)
            $content = str_replace('&','%%KAUFUND%%',$content);
        else
            $content = str_replace('%%KAUFUND%%','&',$content);
        \Storage::disk($this->storageDisk)
            ->put($this->local_path.$file, $content);
    }

    protected function AddContentType($imageCt = 'jpeg')
    {
        $ct_file = $this->ReadOpenXmlFile('[Content_Types].xml', 'object');

        if (! ($ct_file instanceof \Traversable)) {
            throw new Exception('Cannot traverse through [Content_Types].xml.');
        }

        //check if content type for jpg has been set
        $i = 0;
        $ct_already_set = false;
        foreach ($ct_file as $ct) {
            if ((string) $ct_file->Default[$i]['Extension'] == $imageCt) {
                $ct_already_set = true;
            }
            $i++;
        }

        //if content type for jpg has not been set, add it to xml
        // and save xml to file and add it to the archive
        if (! $ct_already_set) {
            $sxe = $ct_file->addChild('Default');
            $sxe->addAttribute('Extension', $imageCt);
            $sxe->addAttribute('ContentType', 'image/'.$imageCt);
            $this->SaveOpenXmlObjectToFile($ct_file, '[Content_Types].xml', false);
        }
    }

    protected function FetchReplaceableImages(&$main_file, $ns)
    {
        //set up basic arrays to keep track of imgs
        $imgs = [];
        $imgs_replaced = []; // so they can later be removed from media and relation file.
        $newIdCounter = 1;

        //iterate through all drawing containers of the xml document
        foreach ($main_file->xpath('//w:drawing') as $k=>$drawing) {
            //figure out if there is a URL saved in the description field of the img
            $img_url = $this->AnalyseImgUrlString($drawing->children($ns['wp'])->xpath('wp:docPr')[0]->attributes()['descr']);
            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->xpath('wp:docPr')[0]->attributes()['descr'] = $img_url['rest'];

            //if there is a url, save this img as a img to be replaced
            if ($img_url['valid']) {
                $ueid = 'wrklstId'.$newIdCounter;
                $wasId = (string) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])['embed'];

                //get dimensions
                $cx = (int) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()['cx'];
                $cy = (int) $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])->xfrm->ext->attributes()['cy'];

                //remember img as being replaced
                $imgs_replaced[$wasId] = $wasId;

                //set new img id
                $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])['embed'] = $ueid;

                $imgs[] = [
                    'cx'     => (int) $cx,
                    'cy'     => (int) $cy,
                    'wasId'  => $wasId,
                    'id'     => $ueid,
                    'url'    => $img_url['url'],
                    'path'    => $img_url['path'],
                    'mode'    => $img_url['mode'],
                ];

                $newIdCounter++;
            }
        }

        return [
            'imgs'          => $imgs,
            'imgs_replaced' => $imgs_replaced,
        ];
    }

    protected function RemoveReplaceImages($imgs_replaced, &$rels_file)
    {
        //TODO: check if the same img is used at a different position int he file as well, as otherwise broken images are produced.
        //iterate through replaced images and clean rels files from them
        foreach ($imgs_replaced as $img_replaced) {
            $i = 0;
            foreach ($rels_file as $rel) {
                if ((string) $rel->attributes()['Id'] == $img_replaced) {
                    $this->zipper->remove('word/'.(string) $rel->attributes()['Target']);
                    unset($rels_file->Relationship[$i]);
                }
                $i++;
            }
        }
    }

    protected function InsertImages($ns, &$imgs, &$rels_file, &$main_file)
    {
        $docimage = new DocImage();
        $allowed_imgs = $docimage->AllowedContentTypeImages();
        $image_i = 1;
        //iterate through replacable images
        foreach ($imgs as $k=>$img) {
            $this->Log('Merge Images into Template - '.round($image_i / count($imgs) * 100).'%');
            //get file type of img and test it against supported imgs
            if ($imgageData = $docimage->GetImageFromUrl($img['mode'] == 'url' ? $img['url'] : $img['path'], $img['mode'] == 'url' ? $this->imageManipulation : '')) {
                $imgs[$k]['img_file_src'] = str_replace('wrklstId', 'wrklst_image', $img['id']).$allowed_imgs[$imgageData['mime']];
                $imgs[$k]['img_file_dest'] = str_replace('wrklstId', 'wrklst_image', $img['id']).'.jpeg';

                $resampled_img = $docimage->ResampleImage($this, $imgs, $k, $imgageData['data']);

                $sxe = $rels_file->addChild('Relationship');
                $sxe->addAttribute('Id', $img['id']);
                $sxe->addAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                $sxe->addAttribute('Target', 'media/'.$imgs[$k]['img_file_dest']);

                foreach ($main_file->xpath('//w:drawing') as $k=>$drawing) {
                    if ($img['id'] == $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])
                        ->graphic->graphicData->children($ns['pic'])->pic->blipFill->children($ns['a'])
                        ->blip->attributes($ns['r'])['embed']) {
                        $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])
                            ->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])
                            ->xfrm->ext->attributes()['cx'] = $resampled_img['width_emus'];
                        $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->children($ns['a'])
                            ->graphic->graphicData->children($ns['pic'])->pic->spPr->children($ns['a'])
                            ->xfrm->ext->attributes()['cy'] = $resampled_img['height_emus'];
                        //anchor images
                        if (isset($main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->anchor)) {
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->anchor->extent->attributes()['cx'] = $resampled_img['width_emus'];
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->anchor->extent->attributes()['cy'] = $resampled_img['height_emus'];
                        }
                        //inline images
                        elseif (isset($main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->inline)) {
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->inline->extent->attributes()['cx'] = $resampled_img['width_emus'];
                            $main_file->xpath('//w:drawing')[$k]->children($ns['wp'])->inline->extent->attributes()['cy'] = $resampled_img['height_emus'];
                        }

                        break;
                    }
                }
            }
            $image_i++;
        }
    }

    protected function ImageReplacer()
    {
        $this->Log('Merge Images into Template');

        //load main doc xml
        $main_file = simplexml_load_string($this->word_doc);

        //get all namespaces of the document
        $ns = $main_file->getNamespaces(true);

        $replaceableImage = $this->FetchReplaceableImages($main_file, $ns);
        $imgs = $replaceableImage['imgs'];
        $imgs_replaced = $replaceableImage['imgs_replaced'];

        $rels_file = $this->ReadOpenXmlFile('word/_rels/document.xml.rels', 'object');

        //do not remove until it is checked if the same img is used at a different position int he file as well, as otherwise broken images are produced.
        //$this->RemoveReplaceImages($imgs_replaced, $rels_file);

        //add jpg content type if not set
        $this->AddContentType('jpeg');

        $this->InsertImages($ns, $imgs, $rels_file, $main_file);

        $this->SaveOpenXmlObjectToFile($rels_file, 'word/_rels/document.xml.rels', 'word/_rels');

        if ($main_file_xml = $main_file->asXML()) {
            $this->word_doc = $main_file_xml;
        } else {
            throw new Exception('Cannot generate xml for word/document.xml.');
        }
    }

    /**
     * @param string $string
     */
    protected function AnalyseImgUrlString($string)
    {
        $string = (string) $string;
        $start = '[IMG-REPLACE]';
        $end = '[/IMG-REPLACE]';
        $start_local = '[LOCAL_IMG_REPLACE]';
        $end_local = '[/LOCAL_IMG_REPLACE]';
        $valid = false;
        $url = '';
        $path = '';

        if ($string != str_replace($start, '', $string) && $string == str_replace($start.$end, '', $string)) {
            $string = ' '.$string;
            $ini = strpos($string, $start);
            if ($ini == 0) {
                $url = '';
                $rest = $string;
            } else {
                $ini += strlen($start);
                $len = ((strpos($string, $end, $ini)) - $ini);
                $url = substr($string, $ini, $len);

                $ini = strpos($string, $start);
                $len = strpos($string, $end, $ini + strlen($start)) + strlen($end);
                $rest = substr($string, 0, $ini).substr($string, $len);
            }

            $valid = true;

            //TODO: create a better url validity check
            if (! trim(str_replace(['http', 'https', ':', ' '], '', $url)) || $url == str_replace('http', '', $url)) {
                $valid = false;
            }
            $mode = 'url';
        } elseif ($string != str_replace($start_local, '', $string) && $string == str_replace($start_local.$end_local, '', $string)) {
            $string = ' '.$string;
            $ini = strpos($string, $start_local);
            if ($ini == 0) {
                $path = '';
                $rest = $string;
            } else {
                $ini += strlen($start_local);
                $len = ((strpos($string, $end_local, $ini)) - $ini);
                $path = str_replace('..', '', substr($string, $ini, $len));

                $ini = strpos($string, $start_local);
                $len = strpos($string, $end_local, $ini + strlen($start)) + strlen($end_local);
                $rest = substr($string, 0, $ini).substr($string, $len);
            }

            $valid = true;

            //check if path starts with storage path
            if (! starts_with($path, storage_path())) {
                $valid = false;
            }
            $mode = 'path';
        } else {
            $mode = 'nothing';
            $url = '';
            $path = '';
            $rest = str_replace([$start, $end, $start_local, $end_local], '', $string);
        }

        return [
            'mode' => $mode,
            'url'  => trim($url),
            'path' => trim($path),
            'rest' => trim($rest),
            'valid' => $valid,
        ];
    }

    public function SaveAsPdf()
    {
        $this->Log('Converting DOCX to PDF');
        //convert to pdf with libre office
        $command = 'soffice --headless --convert-to pdf '.$this->StoragePath($this->local_path.$this->template_file_name).' --outdir '.$this->StoragePath($this->local_path);
        $process = new \Symfony\Component\Process\Process($command);
        $process->start();
        while ($process->isRunning()) {
            //wait until process is ready
        }
        // executes after the command finishes
        if (! $process->isSuccessful()) {
            throw new \Symfony\Component\Process\Exception\ProcessFailedException($process);
        } else {
            $path_parts = pathinfo($this->StoragePath($this->local_path.$this->template_file_name));

            return $this->StoragePath($this->local_path.$path_parts['filename'].'pdf');
        }
    }
}
