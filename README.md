![DocxMutache Logo](https://github.com/wrklst/docxmustache/raw/master/example/logo.png)
# DocxMustache *for Laravel 6.0.*

[![Software License](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square)](LICENSE) [![Quality Score](https://img.shields.io/scrutinizer/g/wrklst/docxmustache.svg?style=flat-square&b=master)](https://scrutinizer-ci.com/g/wrklst/docxmustache/?branch=master) [![Build Status](https://scrutinizer-ci.com/g/wrklst/docxmustache/badges/build.png?b=master)](https://scrutinizer-ci.com/g/wrklst/docxmustache/build-status/master)
[![StyleCI](https://styleci.io/repos/90483440/shield?branch=master)](https://styleci.io/repos/90483440)

Docx template manipulation class for Laravel 6.0, based on [mustache templating language](https://mustache.github.io). This class is still under heavy development and works more like proof of concept at the moment. Things will change quickly and might break things.

![Template Example in Word](https://github.com/wrklst/docxmustache/raw/master/example/ExampleMustacheTemplate.png)
This package helps you to use docx files with mustache syntax as templates to merge richly styled documents with information from any data source. It can replace text and images and supports basic html styling (bold, itallic, underline).

## Installation
`composer require wrklst/docxmustache`

Please check depencies down below and examples folder for configuration and usage.

## HTML conversion

Current HTML conversion is basic and only supports singular runs of bold, italic and underlined text and no combination of these. It requires all values non html to be escaped with
`htmlspecialchars($value, ENT_COMPAT, 'UTF-8');`
and a prefix of
`*[[DONOTESCAPE]]*`
so the class knows not to escape the html before it is converted to openXML.


## Replacing images

The image needs to be a reachable URL with a image in a supported format. The url value needs to be placed into the alt text description field of the image.
Images will be resampled to the constraints of the placeholder image.
The Image value needs to be formated the with pseudo tags around, such as:
`[IMG-REPLACE]http://placehold.it/350x150[/IMG-REPLACE]`

## Example
Please also checkout the example in the example folder to get a basic understand of how to use this class.

## DOCX to PDF conversion

Conversion to PDF requires `libreoffice-common` to be installed on the server (used for conversion).
Use `sudo apt install libreoffice-common` on your ubuntu/debian based server. Also install ttf-mscorefonts if you need support for Arial font when converting docx documents to pdf `sudo apt-get install ttf-mscorefonts-installer `

## Other Dependencies
The package is dependent on several Laravel specific functions. It could easily be ported to other frameworks or be ported to be framework agnostic. In addition to the Laravel dependency, the page uses the following packages:

* [mustache/mustache](https://packagist.org/packages/mustache/mustache)
* [intervention/image](http://image.intervention.io) (requires adding [provider and alias](http://image.intervention.io/getting_started/installation#laravel) to your app config as well as gd or imagick, [please check the image intervention webpage for details](http://image.intervention.io/getting_started/installation#laravel))

Laravel specific dependencies (only relevant if ported into non Laravel environment):

* [Storage and File class, based on Flysystem](https://flysystem.thephpleague.com) (for local file access, could also be replaced by phps own fopen etc methods)
* [\Symfony\Component\Process\Process](http://symfony.com/doc/current/components/process.html) (only for PDF conversion)

## Contributions
If you would like to contribute something to this package, please feel free to make a pull request and a corresponding issue and we will be happy to review and discuss.

## Why another openXML / docx template solution?
There are some classes out there that help with writing and or changing the content of word documents, some with commercial licenses and some free. For our particular purpose we did not need most of the features many of the libraries out there have – we needed a simple solution that would allow to replace values and images and traverse through data in a easy and straightforward manner.

Many of the other libraries use cloning to repeat a block with some custom templating syntax instead of using a existing template syntax. We use the usual mustache syntax, also to achieve repeating content in as many dimensions as needed.

Other PHP Classes to manipulate openXML word documents:

* [openTBS – Tiny But Strong](http://www.tinybutstrong.com/opentbs.php)
* [PHPWord](https://github.com/PHPOffice/PHPWord)
* [docxtemplater pro](https://modules.docxtemplater.com) (basic opensource / free version / MIT license available as of writing; image replacing is a commercial plugin)
* [docxpresso](http://www.docxpresso.com) (commercial)
* [phpdocx](https://www.phpdocx.com) (commercial)
