<?php

namespace WrkLst\DocxMustache\Tests;

use Illuminate\Support\Facades\Log;
use PHPUnit\Framework\Attributes\DataProvider;
use Symfony\Component\Process\Exception\ProcessFailedException;
use Symfony\Component\Process\Process;
use WrkLst\DocxMustache\DocxMustache;

class DocxMustacheTest extends TestCase
{
    private function local(string $file = 'logo.png'): string
    {
        return '[LOCAL_IMG_REPLACE]'.$this->storage.'/app/'.$file.'[/LOCAL_IMG_REPLACE]';
    }

    /** A 1600x900 image, larger than any placeholder. */
    private function bigImage(): string
    {
        imagegif(imagecreatetruecolor(1600, 900), $this->storage.'/app/big.gif');

        return $this->local('big.gif');
    }

    private function execute(array $items, string $template, int $dpi = 72): DocxMustache
    {
        $docx = new DocxMustache($items, $template);
        $docx->execute($dpi);

        return $docx;
    }

    private function path(DocxMustache $docx): string
    {
        return $docx->storagePath($docx->local_path.$docx->template_file_name);
    }

    public function test_example_template(): void
    {
        $this->template('example_template.docx');

        $docx = $this->execute(['items' => [
            ['name' => 'Someone Other', 'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on', 'img_url' => $this->local()],
            ['name' => 'Person & X', 'captions' => 'plain <b>text</b>', 'img_url' => $this->local()],
        ]], 'example_template.docx');

        $this->assertMatchesRegularExpression('#^app/DocxMustache/[^/]+/$#', $docx->local_path);
        $this->assertSame('example_template.docx', $docx->template_file_name);
        $this->assertFileExists($this->path($docx));

        $document = $this->part($this->path($docx), 'word/document.xml');
        $this->assertSame($document, $docx->word_doc);
        $this->assertNotFalse(simplexml_load_string($document));
        $this->assertStringNotContainsString('{{', $document);
        $this->assertStringContainsString('>Someone Other</w:t>', $document);
        $this->assertStringContainsString('>Person &amp; X</w:t>', $document);
        $this->assertStringContainsString('<w:b/><w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica" w:cs="Helvetica"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">something bold</w:t>', $document);
        $this->assertStringContainsString('<w:br/><w:t xml:space="preserve">and so on</w:t>', $document);
        $this->assertStringContainsString('>plain &lt;b&gt;text&lt;/b&gt;</w:t>', $document);

        // the picture of each row is replaced by the logo
        $this->assertSame(1, substr_count($document, 'r:embed="wrklstId1"'));
        $this->assertSame(1, substr_count($document, 'r:embed="wrklstId2"'));
        $this->assertSame(2, substr_count($document, '<wp:docPr id="1" name="Picture 1" descr=""/>'));
        $this->assertSame([245, 245], array_slice(getimagesizefromstring($this->part($this->path($docx), 'word/media/wrklst_image2.jpeg')), 0, 2));

        // the template itself stays untouched, nothing but the docx is left in the temp dir
        $this->assertStringContainsString('{{name}}', $this->part($this->storage.'/app/private/app/example_template.docx', 'word/document.xml'));
        $this->assertSame([$this->path($docx)], glob(dirname($this->path($docx)).'/*'));
    }

    public function test_headers_and_footers_are_rendered(): void
    {
        $part = fn ($tag) => '<?xml version="1.0"?><w:'.$tag.' '.self::NAMESPACES.'>'.$this->paragraph('{{title}}').'</w:'.$tag.'>';
        $this->template('parts.docx', ['word/header1.xml' => $part('hdr'), 'word/footer1.xml' => $part('ftr')]);

        $docx = $this->execute(['title' => 'In all parts'], 'parts.docx');

        $this->assertStringContainsString('>In all parts</w:t>', $this->part($this->path($docx), 'word/header1.xml'));
        $this->assertStringContainsString('>In all parts</w:t>', $this->part($this->path($docx), 'word/footer1.xml'));
    }

    public function test_images_in_headers_and_footers_are_replaced(): void
    {
        $part = fn ($tag, $body) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n".'<w:'.$tag.' '.self::NAMESPACES.'>'.$body.'</w:'.$tag.'>';
        $untouched = $part('ftr', $this->paragraph('{{title}}').$this->picture('just a description'));
        $this->template('parts.docx', [
            'word/document.xml' => $this->document($this->picture('{{logo}}')),
            'word/header1.xml'  => $part('hdr', $this->picture('{{logo}}')),
            'word/footer1.xml'  => $untouched,
            // a header that has no relationships yet
            'word/header2.xml'  => $part('hdr', $this->picture('{{logo}}', 'anchor')),
        ]);

        $docx = $this->execute(['logo' => $this->local(), 'title' => 'T'], 'parts.docx');
        $path = $this->path($docx);
        $rel = fn ($n) => '<Relationship Id="wrklstId'.$n.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/wrklst_image'.$n.'.jpeg"/>';

        // numbered through all parts, the main document first; each part gets its own relationship
        $this->assertStringContainsString('r:embed="wrklstId1"', $docx->word_doc);
        $this->assertStringContainsString($rel(1), $this->part($path, 'word/_rels/document.xml.rels'));
        foreach ([2 => 'header1', 3 => 'header2'] as $n => $header) {
            $this->assertStringContainsString('r:embed="wrklstId'.$n.'"', $this->part($path, 'word/'.$header.'.xml'));
            $this->assertStringContainsString('<wp:extent cx="1270968" cy="1270968"/>', $this->part($path, 'word/'.$header.'.xml'));
            $rels = $this->part($path, 'word/_rels/'.$header.'.xml.rels');
            $this->assertStringContainsString($rel($n), $rels);
            $this->assertStringNotContainsString('wrklstId1"', $rels);
            $this->assertNotFalse(simplexml_load_string($rels));
            $this->assertNotFalse($this->part($path, 'word/media/wrklst_image'.$n.'.jpeg'));
        }

        // a part without image tags is not touched beyond its mustache tags
        $this->assertSame(str_replace(['{{title}}', '<w:t>'], ['T', '<w:t xml:space="preserve">'], $untouched), $this->part($path, 'word/footer1.xml'));
        $this->assertFalse($this->part($path, 'word/_rels/footer1.xml.rels'));
    }

    public function test_document_without_relationships(): void
    {
        $this->template('norels.docx', ['word/document.xml' => $this->document($this->picture('{{logo}}'))]);
        $zip = new \ZipArchive;
        $zip->open($this->storage.'/app/private/app/norels.docx');
        $zip->deleteName('word/_rels/document.xml.rels');
        $zip->close();

        $docx = $this->execute(['logo' => $this->local()], 'norels.docx');

        $this->assertStringContainsString('Id="wrklstId1"', $this->part($this->path($docx), 'word/_rels/document.xml.rels'));
    }

    public function test_working_directories_are_cleaned_up_after_48_hours(): void
    {
        $this->template('example_template.docx');
        $age = function (DocxMustache $docx, float $hours) {
            touch(dirname($this->path($docx)), time() - (int) ($hours * 3600));

            return $this->path($docx);
        };

        $fresh = $age($this->execute([], 'example_template.docx'), 47);
        $expired = $age($this->execute([], 'example_template.docx'), 49);

        $this->execute([], 'example_template.docx');
        $this->assertFileExists($fresh);
        $this->assertFileDoesNotExist($expired);
        $this->assertDirectoryDoesNotExist(dirname($expired));

        // configurable: as constructor argument or property, null keeps everything
        $docx = new DocxMustache([], 'example_template.docx', tmpDirLifetime: null);
        $this->assertNull($docx->tmpDirLifetime);
        $docx->execute();
        $this->assertFileExists($fresh);

        $docx = new DocxMustache([], 'example_template.docx');
        $this->assertSame(48, $docx->tmpDirLifetime);
        $docx->tmpDirLifetime = 0.5;
        $docx->execute();
        $this->assertFileDoesNotExist($fresh);
        $this->assertFileExists($this->path($docx));
    }

    /** What wrklst-img does: PascalCase overrides of the logger and the working directory. */
    public function test_subclass_overriding_log_and_tmp_dir(): void
    {
        $this->template('example_template.docx');
        $docx = new class([], 'example_template.docx') extends DocxMustache
        {
            public $messages = [];

            protected function Log($msg)
            {
                $this->messages[] = $msg;
            }

            public function GetTmpDir()
            {
                $this->CleanUpTmpDirs();
                \Illuminate\Support\Facades\Storage::disk($this->storageDisk)->makeDirectory('app/custom/');

                return 'app/custom/';
            }
        };

        $docx->Execute(72);

        $this->assertSame('app/custom/', $docx->local_path);
        $this->assertFileExists($this->path($docx));
        $this->assertSame(['Get Copy of Template', 'Analyze Template', 'Merge Data into Template', 'Load XML Document to Merge Images', 'Merge Images into Template', 'Compact Template with Data'], $docx->messages);
    }

    public function test_values_are_not_rendered_twice(): void
    {
        $this->template('braces.docx', ['word/document.xml' => $this->document($this->paragraph('{{title}}|{{open}}'))]);

        $docx = $this->execute(['title' => 'literal {{number}}', 'open' => '{{#x}', 'number' => 7], 'braces.docx');

        $this->assertStringContainsString('>literal {{number}}|{{#x}</w:t>', $docx->word_doc);
    }

    public function test_images_are_replaced(): void
    {
        $this->template('images.docx', ['word/document.xml' => $this->document(
            $this->picture('{{logo}}')                                   // smaller than the placeholder: keeps its 245x245
            .$this->picture('alt {{big}} text', 'anchor', 3000000, 1000000) // 578x192 available: fitted to the height
            .$this->picture('{{big}}')                                   // 295x413 available: fitted to the width
            .$this->picture('just a description')
        )]);

        $docx = $this->execute(['logo' => $this->local(), 'big' => $this->bigImage()], 'images.docx');

        $document = simplexml_load_string($docx->word_doc);
        $document->registerXPathNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');
        $document->registerXPathNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $attributes = fn ($xpath, $name) => array_map(fn ($node) => (string) $node[$name], $document->xpath($xpath));

        $this->assertSame(['', 'alt  text', '', 'just a description'], $attributes('//wp:docPr', 'descr'));
        $this->assertSame(['wrklstId1', 'wrklstId2', 'wrklstId3', 'rId6'], array_map(
            fn ($blip) => (string) $blip->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')['embed'],
            $document->xpath('//a:blip')
        ));
        $this->assertSame(['1270968', '1768980', '1530350', '1530350'], $attributes('//wp:extent', 'cx'));
        $this->assertSame(['1270968', '996024', '861146', '2142490'], $attributes('//wp:extent', 'cy'));
        $this->assertSame($attributes('//wp:extent', 'cx'), $attributes('//a:ext', 'cx'));
        $this->assertSame($attributes('//wp:extent', 'cy'), $attributes('//a:ext', 'cy'));

        $rels = $this->part($this->path($docx), 'word/_rels/document.xml.rels');
        $sizes = [];
        foreach ([1, 2, 3] as $n) {
            $this->assertStringContainsString('<Relationship Id="wrklstId'.$n.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/wrklst_image'.$n.'.jpeg"/>', $rels);
            $info = getimagesizefromstring($this->part($this->path($docx), 'word/media/wrklst_image'.$n.'.jpeg'));
            $this->assertSame('image/jpeg', $info['mime']);
            $sizes[] = $info[0].'x'.$info[1];
        }
        $this->assertSame(['245x245', '341x192', '295x166'], $sizes);
        $this->assertFalse($this->part($this->path($docx), 'word/media/wrklst_image4.jpeg'));
        $this->assertStringContainsString('<Default Extension="jpeg" ContentType="image/jpeg"/>', $this->part($this->path($docx), '[Content_Types].xml'));
    }

    public function test_higher_dpi_stores_more_pixels_at_the_same_size(): void
    {
        $this->template('dpi.docx', ['word/document.xml' => $this->document($this->picture('{{big}}'))]);

        $docx = $this->execute(['big' => $this->bigImage()], 'dpi.docx', 144);

        $info = getimagesizefromstring($this->part($this->path($docx), 'word/media/wrklst_image1.jpeg'));
        $this->assertSame([590, 332], [$info[0], $info[1]]);
        $this->assertStringContainsString('<wp:extent cx="1530350" cy="861146"/>', $docx->word_doc);
    }

    public function test_images_are_loaded_from_urls(): void
    {
        $server = new Process(['php', '-S', '127.0.0.1:8377', '-t', __DIR__.'/../example']);
        $server->start();
        $server->waitUntil(fn () => (bool) @fsockopen('127.0.0.1', 8377));

        try {
            $url = fn ($file) => '[IMG-REPLACE]http://127.0.0.1:8377/'.$file.'[/IMG-REPLACE]';
            $this->template('url.docx', ['word/document.xml' => $this->document($this->picture('{{found}}').$this->picture('{{missing}}'))]);

            $docx = new DocxMustache(['found' => $url('logo.png'), 'missing' => $url('nope.png')], 'url.docx', 'local', 'app/', '?w=1800');
            $docx->execute();
        } finally {
            $server->stop();
        }

        $this->assertNotFalse($this->part($this->path($docx), 'word/media/wrklst_image1.jpeg'));
        // an image that cannot be loaded keeps the placeholder
        $this->assertFalse($this->part($this->path($docx), 'word/media/wrklst_image2.jpeg'));
        $this->assertSame(1, substr_count($docx->word_doc, 'r:embed="wrklstId1"'));
        $this->assertSame(1, substr_count($docx->word_doc, 'r:embed="rId6"'));
        $this->assertStringNotContainsString('wrklstId2', $docx->word_doc.$this->part($this->path($docx), 'word/_rels/document.xml.rels'));
    }

    #[DataProvider('imageTags')]
    public function test_image_tags(string $descr, string $mode, string $source, string $rest, bool $valid): void
    {
        $docx = new class([], 'x.docx') extends DocxMustache
        {
            public function analyse($string)
            {
                return $this->analyseImgUrlString($string);
            }
        };

        $this->assertSame(
            ['mode' => $mode, 'url' => $mode == 'url' ? $source : '', 'path' => $mode == 'path' ? $source : '', 'rest' => $rest, 'valid' => $valid],
            $docx->analyse(str_replace('<storage>', storage_path(), $descr))
            + ['path' => '']
        );
    }

    public static function imageTags(): array
    {
        $url = fn ($url) => '[IMG-REPLACE]'.$url.'[/IMG-REPLACE]';
        $path = fn ($path) => '[LOCAL_IMG_REPLACE]'.$path.'[/LOCAL_IMG_REPLACE]';

        return [
            'no tag'               => ['  plain text ', 'nothing', '', 'plain text', false],
            'http url'             => [$url('http://example.com/a.png'), 'url', 'http://example.com/a.png', '', true],
            'https url, padded'    => ['before '.$url('  https://example.com/a.png?x=1&y=2 ').' after', 'url', 'https://example.com/a.png?x=1&y=2', 'before  after', true],
            'empty url'            => ['x '.$url('').' y', 'nothing', '', 'x  y', false],
            'no scheme'            => [$url('example.com/a.png'), 'url', 'example.com/a.png', '', false],
            'other scheme'         => [$url('ftp://example.com/http.png'), 'url', 'ftp://example.com/http.png', '', false],
            'local file as url'    => [$url('/etc/hosts#http'), 'url', '/etc/hosts#http', '', false],
            'stream wrapper'       => [$url('php://filter/resource=http://x'), 'url', 'php://filter/resource=http://x', '', false],
            'unclosed url'         => ['[IMG-REPLACE]http://example.com/a.png', 'nothing', '', 'http://example.com/a.png', false],
            'stray closing tags'   => ['[/IMG-REPLACE] a [/LOCAL_IMG_REPLACE]', 'nothing', '', 'a', false],
        ];
    }

    public function test_local_image_tags_are_limited_to_the_storage_path(): void
    {
        $docx = new class([], 'x.docx') extends DocxMustache
        {
            public function analyse($string)
            {
                return $this->analyseImgUrlString($string);
            }
        };
        $tag = fn ($path) => $docx->analyse('alt [LOCAL_IMG_REPLACE]'.$path.'[/LOCAL_IMG_REPLACE] text');
        $storage = storage_path();

        $this->assertSame(['mode' => 'path', 'url' => '', 'path' => $storage.'/app/logo.png', 'rest' => 'alt  text', 'valid' => true], $tag($storage.'/app/logo.png'));
        $this->assertSame($storage.'/app//app/logo.png', $tag($storage.'/app/../app/logo.png')['path']);
        $this->assertFalse($tag($storage.'/../../etc/hosts')['valid'] && str_contains($tag($storage.'/../../etc/hosts')['path'], '..'));
        $this->assertFalse($tag('/etc/hosts')['valid']);
        $this->assertFalse($tag('relative/logo.png')['valid']);
        $this->assertFalse($tag($storage.'-other/logo.png')['valid']);
    }

    public function test_document_that_is_no_valid_xml_is_still_written(): void
    {
        $this->template('broken.docx', ['word/document.xml' => '<w:document '.self::NAMESPACES.'><w:body>'.$this->paragraph('{{title}}').'</w:document>']);
        Log::spy();

        $docx = new DocxMustache(['title' => 'T'], 'broken.docx');
        $docx->verbose = true;
        $docx->execute();

        $this->assertStringContainsString('>T</w:t>', $this->part($this->path($docx), 'word/document.xml'));
        Log::shouldHaveReceived('error')->withArgs(fn ($message) => str_starts_with($message, 'Error: Could not load XML file.'))->once();
    }

    public function test_nothing_is_logged_unless_verbose(): void
    {
        $this->template('example_template.docx');
        Log::spy();

        $this->execute([], 'example_template.docx');

        Log::shouldNotHaveReceived('error');
    }

    public function test_missing_template(): void
    {
        $this->expectExceptionMessage('Cannot copy template: app/nope.docx');

        $this->execute([], 'nope.docx');
    }

    public function test_template_that_is_no_docx(): void
    {
        file_put_contents($this->storage.'/app/private/app/fake.docx', 'not a zip');

        $this->expectExceptionMessage('Cannot open zip: ');

        $this->execute([], 'fake.docx');
    }

    public function test_storage_disk_and_prefix(): void
    {
        config(['filesystems.disks.other' => ['driver' => 'local', 'root' => $this->storage.'/other']]);
        mkdir($this->storage.'/other/temp/sub', 0777, true);
        copy(__DIR__.'/../example/example_template.docx', $this->storage.'/other/temp/sub/t.docx');

        $docx = new DocxMustache([], 'sub/t.docx');
        $docx->storageDisk = 'other';
        $docx->storagePathPrefix = 'temp/';
        $docx->execute();

        $this->assertMatchesRegularExpression('#^temp/DocxMustache/[^/]+/$#', $docx->local_path);
        $this->assertSame('t.docx', $docx->template_file_name);
        $this->assertSame($this->storage.'/other/'.$docx->local_path.'t.docx', $this->path($docx));
        $this->assertFileExists($this->path($docx));

        $docx->useStoragePath = true;
        $this->assertFileExists($this->path($docx));
    }

    /** Runs where LibreOffice is installed (the CI has a job for it): a real office suite opens the result. */
    public function test_libreoffice_converts_the_generated_document(): void
    {
        if (! (new \Symfony\Component\Process\ExecutableFinder)->find('soffice')) {
            if (getenv('REQUIRE_LIBREOFFICE')) {
                $this->fail('soffice is not on the PATH');
            }
            $this->markTestSkipped('LibreOffice is not installed');
        }

        $this->template('example_template.docx');
        $docx = $this->execute(['items' => [['name' => 'Someone', 'captions' => '*[[DONOTESCAPE]]*<b>bold</b><br>text', 'img_url' => $this->local()]]], 'example_template.docx');

        $pdf = $docx->saveAsPdf();

        $this->assertFileExists($pdf);
        $this->assertStringStartsWith('%PDF', file_get_contents($pdf));
        $this->assertGreaterThan(5000, filesize($pdf)); // not an empty page: it embeds the logo
    }

    public function test_save_as_pdf(): void
    {
        // a stand-in for LibreOffice that records its arguments and writes the pdf
        mkdir($this->storage.'/bin');
        file_put_contents($this->storage.'/bin/soffice', "#!/bin/sh\necho \"$@\" > \"{$this->storage}/args\"\n[ -e \"{$this->storage}/fail\" ] && exit 3\ntouch \"$6/$(basename \"$4\" .docx).pdf\"\n");
        chmod($this->storage.'/bin/soffice', 0755);
        $path = getenv('PATH');
        putenv('PATH='.$this->storage.'/bin:'.$path);

        try {
            $this->template('example_template.docx');
            $docx = $this->execute([], 'example_template.docx');

            $pdf = $docx->saveAsPdf();

            $this->assertSame(dirname($this->path($docx)).'/example_template.pdf', $pdf);
            $this->assertFileExists($pdf);
            $this->assertSame('--headless --convert-to pdf '.$this->path($docx).' --outdir '.dirname($this->path($docx))."/\n", file_get_contents($this->storage.'/args'));

            touch($this->storage.'/fail');
            $this->expectException(ProcessFailedException::class);
            $docx->saveAsPdf();
        } finally {
            putenv('PATH='.$path);
        }
    }
}
