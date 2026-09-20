<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;

class docxMustacheExampleController extends Controller
{
    public function index(Request $request)
    {
        //copy the example doc file and logo.png into your storage directory or corret these paths
        $local_template_file = 'example_template.docx';
        $local_image = storage_path('app/logo.png');

        //define date to be replaced
        $data = [
            [
                'name'     => 'Someone Other',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[LOCAL_IMG_REPLACE]'.$local_image.'[/LOCAL_IMG_REPLACE]',
            ],
            [
                'name'     => 'Person X',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[LOCAL_IMG_REPLACE]'.$local_image.'[/LOCAL_IMG_REPLACE]',
            ],
            [
                'name'     => 'Person Y',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[LOCAL_IMG_REPLACE]'.$local_image.'[/LOCAL_IMG_REPLACE]',
            ],
        ];

        //call class
        $docx_creation = new \WrkLst\DocxMustache\DocxMustache(['items'=>$data], $local_template_file);

        //optionally change some setting before the class gets executed
        $docx_creation->storageDisk = 'local';
        $docx_creation->storagePathPrefix = 'temp/';

        //execute class
        $docx_creation->execute();

        //return path of generated docx file
        return [
            'docx_file' => $docx_creation->local_path.$docx_creation->template_file_name,
        ];
    }
}
