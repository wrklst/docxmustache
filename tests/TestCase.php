<?php

namespace WrkLst\DocxMustache\Tests;

use Illuminate\Filesystem\Filesystem;
use Orchestra\Testbench\TestCase as Orchestra;
use ZipArchive;

abstract class TestCase extends Orchestra
{
    protected const NAMESPACES = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

    /** storage_path() of the test app, a fresh temp dir per test */
    protected string $storage;

    protected function setUp(): void
    {
        parent::setUp();

        $this->storage = sys_get_temp_dir().'/docxmustache-'.uniqid();
        mkdir($this->storage.'/app/private/app', 0777, true);
        $this->app->useStoragePath($this->storage);
        config(['filesystems.default' => 'local', 'filesystems.disks.local.root' => $this->storage.'/app/private']);

        copy(__DIR__.'/../example/logo.png', $this->storage.'/app/logo.png');
    }

    protected function tearDown(): void
    {
        (new Filesystem)->deleteDirectory($this->storage);

        parent::tearDown();
    }

    /**
     * Put a template on the disk: the example docx with some of its parts replaced.
     *
     * @param  array<string, string>  $parts  content by part name
     */
    protected function template(string $name, array $parts = []): void
    {
        $path = $this->storage.'/app/private/app/'.$name;
        copy(__DIR__.'/../example/example_template.docx', $path);

        $zip = new ZipArchive;
        $zip->open($path);
        foreach ($parts as $part => $content) {
            $zip->addFromString($part, $content);
        }
        $zip->close();
    }

    protected function document(string $body): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n".'<w:document '.self::NAMESPACES.'><w:body>'.$body.'</w:body></w:document>';
    }

    protected function paragraph(string $text): string
    {
        return '<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>'.$text.'</w:t></w:r></w:p>';
    }

    /** A picture as Word stores it, with $descr as its alt text. */
    protected function picture(string $descr, string $position = 'inline', int $cx = 1530350, int $cy = 2142490): string
    {
        return '<w:p><w:r><w:drawing><wp:'.$position.'><wp:extent cx="'.$cx.'" cy="'.$cy.'"/><wp:docPr id="1" name="Picture 1" descr="'.$descr.'"/>'
            .'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            .'<pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>'
            .'<pic:blipFill><a:blip r:embed="rId6"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            .'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'.$cx.'" cy="'.$cy.'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            .'</pic:pic></a:graphicData></a:graphic></wp:'.$position.'></w:drawing></w:r></w:p>';
    }

    /** Content of one part of a docx, false if it has no such part. */
    protected function part(string $docx, string $name): string|false
    {
        $zip = new ZipArchive;
        $this->assertTrue($zip->open($docx));
        $content = $zip->getFromName($name);
        $zip->close();

        return $content;
    }
}
