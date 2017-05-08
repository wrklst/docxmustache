<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;

class docxMustacheExampleController extends Controller
{
    public function index(Request $request)
    {
        //copy the example doc file into your storage directory or corret this path
        $local_template_file = storage_path('example_template.docx');

        //define date to be replaced
        $data = [
            [
                'name'     => 'Someone Other',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[IMG-REPLACE]http://placehold.it/350x150[/IMG-REPLACE]',
            ],
            [
                'name'     => 'Person X',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[IMG-REPLACE]http://placehold.it/350x150[/IMG-REPLACE]',
            ],
            [
                'name'     => 'Person Y',
                'captions' => '*[[DONOTESCAPE]]*<b>something bold</b><br />and so on',
                'img_url'  => '[IMG-REPLACE]http://placehold.it/350x150[/IMG-REPLACE]',
            ],
        ];

        //call class
        $docx_creation = new \WrkLst\DocxMustache\DocxMustache(['items'=>$data], $local_template_file);

        //optionally change some setting before the class gets executed
        $docx_creation->storageDisk = 'local';

        //execute class
        $docx_creation->execute();

        //return path of generated docx file
        return [
            'docx_file' => $docx_creation->local_path.$docx_creation->template_file_name,
        ];
    }
}
