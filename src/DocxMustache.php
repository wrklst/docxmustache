<?php

namespace WrkLst\DocxMustache;

use Exception;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Storage;
use SimpleXMLElement;
use Symfony\Component\Process\Process;
use Throwable;
use ZipArchive;

/**
 * Fills a DOCX template: mustache tags in the text parts, basic HTML styling,
 * and pictures whose alt text names a replacement image.
 */
class DocxMustache
{
    private const MAIN_DOCUMENT = 'word/document.xml';

    private const EMPTY_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n".'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

    private const IMAGE_TAGS = [
        'url' => ['[IMG-REPLACE]', '[/IMG-REPLACE]'],
        'path' => ['[LOCAL_IMG_REPLACE]', '[/LOCAL_IMG_REPLACE]'],
    ];

    public $items;

    public $word_doc;

    public $template_file_name;

    public $template_file;

    public $local_path;

    public $storageDisk;

    public $storagePathPrefix;

    public $useStoragePath = false;

    public $zipper;

    public $imageManipulation;

    public $verbose;

    /** Hours after which the working directories of earlier runs are deleted. Null keeps them forever. */
    public $tmpDirLifetime;

    private $filelist = [];

    private $fileWhitelist = [
        'word/document.xml',
        'word/endnotes.xml',
        'word/footer*.xml',
        'word/footnotes.xml',
        'word/header*.xml',
    ];

    /**
     * @param  array  $items  data for the mustache tags
     * @param  string  $local_template_file  template, relative to $storagePathPrefix on the disk
     * @param  string  $storageDisk  name of a local filesystem disk
     * @param  string  $storagePathPrefix  prefix within that disk
     * @param  string  $imageManipulation  appended to image urls, e.g. '&w=1800'
     * @param  int|float|null  $tmpDirLifetime  hours to keep the working directories of earlier runs, null for ever
     */
    public function __construct($items, $local_template_file, $storageDisk = 'local', $storagePathPrefix = 'app/', $imageManipulation = '', $tmpDirLifetime = 48)
    {
        $this->items = $items;
        $this->template_file_name = basename($local_template_file);
        $this->template_file = $local_template_file;
        $this->word_doc = false;
        $this->zipper = new ZipArchive;
        $this->storageDisk = $storageDisk;
        $this->storagePathPrefix = $storagePathPrefix;
        $this->imageManipulation = $imageManipulation;
        $this->tmpDirLifetime = $tmpDirLifetime;
        $this->verbose = false;
    }

    /**
     * Render the template. The result is the docx at $local_path.$template_file_name on the disk.
     */
    public function execute($dpi = 72)
    {
        $this->copyTemplate();

        $this->log('Analyze Template');
        $this->getAllFilesFromDocx();

        $this->log('Merge Data into Template');
        foreach ($this->filelist as $file) {
            $this->doInplaceMustache($file);
        }

        $parts = $this->imageReplacer($dpi);

        $this->log('Compact Template with Data');
        if ($parts) {
            $this->writeParts($parts);
        }
    }

    /**
     * Absolute path of a file on the storage disk.
     *
     * @param  string  $file
     */
    public function storagePath($file)
    {
        if ($this->useStoragePath) {
            return Storage::disk($this->storageDisk)->path($file);
        }

        return Storage::disk($this->storageDisk)->path('').$file;
    }

    /**
     * Progress messages, only when $verbose is set. Override to log somewhere else.
     *
     * @param  string  $msg
     */
    protected function log($msg)
    {
        if ($this->verbose) {
            Log::error($msg);
        }
    }

    /**
     * Delete the working directories of earlier runs that are older than $tmpDirLifetime hours,
     * generated documents included. Runs before every render.
     */
    public function cleanUpTmpDirs()
    {
        if ($this->tmpDirLifetime === null) {
            return;
        }

        $isExpired = time() - (int) ($this->tmpDirLifetime * 3600);
        $disk = Storage::disk($this->storageDisk);
        foreach ($disk->directories($this->storagePathPrefix.'DocxMustache') as $dir) {
            try {
                if ($disk->lastModified($dir) < $isExpired) {
                    $disk->deleteDirectory($dir);
                }
            } catch (Throwable) {
                // a render running in parallel has just removed it
            }
        }
    }

    /**
     * Create the working directory of this run: <prefix>DocxMustache/<unique id>/.
     *
     * One flat directory per run, so that cleanUpTmpDirs() can judge each by its own age.
     */
    public function getTmpDir()
    {
        $this->cleanUpTmpDirs();
        $path = $this->storagePathPrefix.'DocxMustache/'.uniqid('', true).'/';
        Storage::disk($this->storageDisk)->makeDirectory($path);

        return $path;
    }

    public function copyTemplate()
    {
        $this->log('Get Copy of Template');
        $this->local_path = $this->getTmpDir();
        $source = $this->storagePathPrefix.$this->template_file;
        if (! Storage::disk($this->storageDisk)->copy($source, $this->local_path.$this->template_file_name)) {
            throw new Exception('Cannot copy template: '.$source);
        }
    }

    public function getAllFilesFromDocx()
    {
        $this->openZip();
        $this->filelist = [];
        for ($i = 0; $i < $this->zipper->numFiles; $i++) {
            $name = $this->zipper->getNameIndex($i);
            foreach ($this->fileWhitelist as $pattern) {
                if (fnmatch($pattern, $name)) {
                    $this->filelist[] = $name;
                }
            }
        }
        $this->zipper->close();
    }

    /**
     * Render mustache tags and HTML styling of one part of the docx, in place.
     *
     * @param  string  $file  name of the part, e.g. word/header1.xml
     */
    public function doInplaceMustache($file)
    {
        $this->openZip();
        $content = MustacheRender::render($this->items, $this->zipper->getFromName($file));
        $this->zipper->addFromString($file, HtmlConversion::convert($content));
        $this->zipper->close();
    }

    /**
     * Convert the generated docx to pdf with LibreOffice ("soffice" has to be on the PATH).
     *
     * @return string absolute path of the pdf
     */
    public function saveAsPdf()
    {
        $this->log('Converting DOCX to PDF');
        $docx = $this->storagePath($this->local_path.$this->template_file_name);

        // no timeout: large documents take LibreOffice a while
        (new Process(['soffice', '--headless', '--convert-to', 'pdf', $docx, '--outdir', $this->storagePath($this->local_path)], timeout: null))
            ->mustRun();

        return $this->storagePath($this->local_path.pathinfo($docx, PATHINFO_FILENAME).'.pdf');
    }

    private function openZip()
    {
        $path = $this->storagePath($this->local_path.$this->template_file_name);
        if ($this->zipper->open($path) !== true) {
            throw new Exception('Cannot open zip: '.$path);
        }
    }

    /** Content of a part of the docx, false if there is no such part. */
    private function readPart(string $name): string|false
    {
        $this->openZip();
        $content = $this->zipper->getFromName($name);
        $this->zipper->close();

        return $content;
    }

    /** @param  string|null  $default  xml to use if the docx has no such part */
    private function readXmlPart(string $name, ?string $default = null): SimpleXMLElement
    {
        $xml = simplexml_load_string($this->readPart($name) ?: (string) $default);
        if ($xml === false) {
            throw new Exception('Cannot load XML Object from file '.$name);
        }

        return $xml;
    }

    /** Relationships of a part, e.g. word/_rels/header1.xml.rels for word/header1.xml. */
    private function relsName(string $part): string
    {
        return dirname($part).'/_rels/'.basename($part).'.rels';
    }

    /**
     * @param  array<string, string>  $parts  content by part name
     */
    private function writeParts(array $parts): void
    {
        $this->openZip();
        foreach ($parts as $name => $content) {
            $this->zipper->addFromString($name, $content);
        }
        $this->zipper->close();
    }

    private function xmlToString(SimpleXMLElement $xml, string $name): string
    {
        $string = $xml->asXML();
        if (! $string) {
            throw new Exception('Cannot generate xml for '.$name);
        }

        return $string;
    }

    /**
     * Replace the pictures that carry an image tag in their alt text: in the main document,
     * headers, footers, footnotes and endnotes. Sets $word_doc to the main document.
     *
     * @return array<string, string> the parts of the docx that have to be written
     */
    protected function imageReplacer($dpi)
    {
        $this->log('Load XML Document to Merge Images');

        $previous = libxml_use_internal_errors(true);
        try {
            $documents = [];
            $images = [];
            // the main document first: its pictures keep the numbers they always had
            foreach (array_unique(array_merge([self::MAIN_DOCUMENT], $this->filelist)) as $file) {
                $content = $this->readPart($file);
                if ($file == self::MAIN_DOCUMENT) {
                    if (! $content) {
                        throw new Exception('Cannot read file '.$file);
                    }
                    $this->word_doc = $content;
                }

                $xml = simplexml_load_string((string) $content);
                if ($xml === false) {
                    $xmlerror = '';
                    foreach (libxml_get_errors() as $error) {
                        $xmlerror .= $this->displayXmlError($error, explode("\n", (string) $content));
                    }
                    libxml_clear_errors();
                    $this->log('Error: Could not load XML file. '.$xmlerror);

                    continue;
                }

                $found = $this->fetchReplaceableImages($xml, $file);
                // the main document is always written back, the other parts only if they have image tags
                if ($file == self::MAIN_DOCUMENT || $found['changed']) {
                    $documents[$file] = $xml;
                }
                array_push($images, ...$found['images']);
            }

            if (! $documents) {
                return [];
            }

            $this->log('Merge Images into Template');

            $rels = [];
            if (isset($documents[self::MAIN_DOCUMENT])) {
                $rels[self::MAIN_DOCUMENT] = $this->readXmlPart($this->relsName(self::MAIN_DOCUMENT), self::EMPTY_RELS);
            }

            $parts = $this->addContentType('jpeg') + $this->insertImages($images, $rels, $dpi);
            foreach ($documents as $file => $xml) {
                $parts[$file] = $this->xmlToString($xml, $file);
            }
            foreach ($rels as $file => $xml) {
                $parts[$this->relsName($file)] = $this->xmlToString($xml, $this->relsName($file));
            }
            $this->word_doc = $parts[self::MAIN_DOCUMENT] ?? $this->word_doc;

            return $parts;
        } finally {
            libxml_use_internal_errors($previous);
        }
    }

    /**
     * Register the image content type, unless the docx already knows it.
     *
     * @return array<string, string> the changed [Content_Types].xml, if any
     */
    protected function addContentType($imageCt = 'jpeg')
    {
        $ct_file = $this->readXmlPart('[Content_Types].xml');

        foreach ($ct_file->Default as $default) {
            if ((string) $default['Extension'] == $imageCt) {
                return [];
            }
        }

        $default = $ct_file->addChild('Default');
        $default->addAttribute('Extension', $imageCt);
        $default->addAttribute('ContentType', 'image/'.$imageCt);

        return ['[Content_Types].xml' => $this->xmlToString($ct_file, '[Content_Types].xml')];
    }

    /**
     * Strip the image tags from the alt text of all pictures of a part and collect those to be replaced.
     *
     * @return array{changed: bool, images: array<int, array{part: string, ns: array, drawing: SimpleXMLElement, source: string, manipulation: string}>}
     */
    protected function fetchReplaceableImages(SimpleXMLElement $xml, string $part)
    {
        $ns = $xml->getNamespaces(true);
        $images = [];
        $changed = false;

        foreach ($xml->xpath('//w:drawing') ?: [] as $drawing) {
            $wp = $drawing->children($ns['wp']);
            $docPr = $wp->xpath('wp:docPr')[0] ?? null;
            if ($docPr === null) {
                continue;
            }

            // attributes(): array access on a node reached through children($namespace)
            // would look for the attribute in that namespace
            $descr = (string) $docPr->attributes()['descr'];
            $tag = $this->analyseImgUrlString($descr);
            $docPr->attributes()['descr'] = $tag['rest'];
            $changed = $changed || $descr !== $tag['rest'];

            if ($tag['valid']) {
                $images[] = [
                    'part' => $part,
                    'ns' => $ns,
                    'drawing' => $wp,
                    'source' => $tag['mode'] == 'url' ? $tag['url'] : $tag['path'],
                    'manipulation' => $tag['mode'] == 'url' ? $this->imageManipulation : '',
                ];
            }
        }

        return ['changed' => $changed, 'images' => $images];
    }

    /**
     * Fetch and resample the images, point the pictures at them and resize the pictures to fit.
     * A picture whose image cannot be loaded keeps its placeholder.
     *
     * @param  array<string, SimpleXMLElement>  $rels  relationships by part, loaded as needed
     * @return array<string, string> the new media files
     */
    protected function insertImages(array $images, array &$rels, $dpi)
    {
        $docimage = new DocImage;
        $media = [];

        foreach ($images as $i => $image) {
            $number = $i + 1;
            $this->log('Merge Images into Template - '.round($number / count($images) * 100).'%');

            $ns = $image['ns'];
            $wp = $image['drawing'];
            $pic = $wp->children($ns['a'] ?? '')->graphic->graphicData->children($ns['pic'] ?? '')->pic;
            if (! isset($pic->blipFill, $pic->spPr)) {
                continue; // a chart or shape, not a picture
            }
            $size = $pic->spPr->children($ns['a'])->xfrm->ext->attributes();

            $imageData = $docimage->getImageFromUrl($image['source'], $image['manipulation']);
            if (! $imageData) {
                continue;
            }
            $resampled = $docimage->resampleImage($imageData['data'], (int) $size['cx'], (int) $size['cy'], $dpi);

            // ponytail: the placeholder's own media file stays in the docx (it may be shared by
            // other pictures); drop it once that is checked, if file size ever matters.
            $target = 'media/wrklst_image'.$number.'.jpeg';
            $media['word/'.$target] = $resampled['data'];

            $rels[$image['part']] ??= $this->readXmlPart($this->relsName($image['part']), self::EMPTY_RELS);
            $relationship = $rels[$image['part']]->addChild('Relationship');
            $relationship->addAttribute('Id', 'wrklstId'.$number);
            $relationship->addAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
            $relationship->addAttribute('Target', $target);

            $pic->blipFill->children($ns['a'])->blip->attributes($ns['r'])['embed'] = 'wrklstId'.$number;

            $size['cx'] = $resampled['width_emus'];
            $size['cy'] = $resampled['height_emus'];
            foreach (['anchor', 'inline'] as $position) {
                if (isset($wp->$position)) {
                    $wp->$position->extent->attributes()['cx'] = $resampled['width_emus'];
                    $wp->$position->extent->attributes()['cy'] = $resampled['height_emus'];
                    break;
                }
            }
        }

        return $media;
    }

    /**
     * Readable libxml error, see http://php.net/manual/en/function.libxml-get-errors.php.
     */
    protected function displayXmlError($error, $xml)
    {
        $return = ($xml[$error->line - 1] ?? '')."\n";
        $return .= str_repeat('-', $error->column)."^\n";

        $return .= match ($error->level) {
            LIBXML_ERR_WARNING => "Warning $error->code: ",
            LIBXML_ERR_ERROR => "Error $error->code: ",
            LIBXML_ERR_FATAL => "Fatal Error $error->code: ",
            default => '',
        };

        $return .= trim($error->message).
                    "\n  Line: $error->line".
                    "\n  Column: $error->column";

        if ($error->file) {
            $return .= "\n  File: $error->file";
        }

        return "$return\n\n--------------------------------------------\n\n";
    }

    /**
     * Find an image tag in the alt text of a picture.
     *
     * [IMG-REPLACE]https://...[/IMG-REPLACE] is valid for http(s) urls,
     * [LOCAL_IMG_REPLACE]/abs/path[/LOCAL_IMG_REPLACE] for files inside storage_path().
     *
     * @param  string  $string
     * @return array{mode: string, url: string, path: string, rest: string, valid: bool}
     */
    protected function analyseImgUrlString($string)
    {
        $string = (string) $string;

        foreach (self::IMAGE_TAGS as $mode => [$start, $end]) {
            $open = strpos($string, $start);
            // an empty tag pair anywhere switches the mode off
            if ($open === false || str_contains($string, $start.$end)) {
                continue;
            }
            $from = $open + strlen($start);
            $close = strpos($string, $end, $from);
            if ($close === false) {
                continue;
            }

            $value = substr($string, $from, $close - $from);
            if ($mode == 'url') {
                $valid = (bool) preg_match('#^\s*https?://#', $value);
            } else {
                $value = str_replace('..', '', $value);
                $valid = str_starts_with($value, rtrim(storage_path(), '/').'/');
            }

            return [
                'mode' => $mode,
                'url' => $mode == 'url' ? trim($value) : '',
                'path' => $mode == 'path' ? trim($value) : '',
                'rest' => trim(substr($string, 0, $open).substr($string, $close + strlen($end))),
                'valid' => $valid,
            ];
        }

        return [
            'mode' => 'nothing',
            'url' => '',
            'path' => '',
            'rest' => trim(str_replace(array_merge(...array_values(self::IMAGE_TAGS)), '', $string)),
            'valid' => false,
        ];
    }
}
