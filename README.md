# DocxMustache for Laravel 5.4


Docx template manipulation class for Laravel 5.4, based on [mustache templating language](https://mustache.github.io). This class is still under heavy development and works more like proof of concept at the moment. Things will change quickly and might break things.

![Template Example in Word](https://github.com/wrklst/docxmustache/raw/master/example/ExampleMustacheTemplate.png)
This package helps you to use docx files with mustache syntax as templates for richtly styled reporting templates. It can replace text and images and supports basic html styling (bold, itallic, underline).

## Installation
`composer require wrklst/docxmustache`

Please check depencies down below and examples folder for configuration and usage.

## Why another openXML / docx template solution?
There are plenty of solutions out there, many with commercial licenses and some free. For our particular purpose we did not need most of the features many of the libraries out there have – we simply needed a solution that would allow to replace values and iterate through an array as well as basic image placeholder replacing.

Many of the other libraries use cloning to repeat a block. We use the usual mustache syntax to achieve repeating content in as many dimensions as needed. This introduces some issues, if the template is not setup in a manner that the content can be repeated, but in most cases it is perfectly sufficient, if the template is setup with this in mind.

## HTML conversion

Current HTML conversion is basic and only supports singular runs of bold, itallic and underlined text and no combination of these. It requires all values non html to be escaped with 
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

Conversion to PDF requires `soffice` to be installed on the server (used for conversion).
Use `sudo apt install soffice` on your ubuntu/debian based server. Also install ttf-mscorefonts if you need support for Arial font when converting docx documents to pdf `sudo apt-get install ttf-mscorefonts-installer `

## Other Dependencies
The package is dependent on several Laravel specific functions. It could easily be ported to other frameworks or be ported to be framework agnostic. In addition to the Laravel dependency, the page uses the following packages:

* [mustache/mustache](https://packagist.org/packages/mustache/mustache)
* [chumper/zipper](https://github.com/Chumper/Zipper)
* [intervention/image](http://image.intervention.io) (requires adding [provider and alias](http://image.intervention.io/getting_started/installation#laravel) to your app config as well as gd or imagick, [please check the image intervention webpage for details](http://image.intervention.io/getting_started/installation#laravel))

Laravel specific dependencies (only relevant if ported into non Laravel environment):

* [Storage and File class, based on Flysystem](https://flysystem.thephpleague.com) (for local file access, could also be replaced by phps own fopen etc methods)
* [\Symfony\Component\Process\Process](http://symfony.com/doc/current/components/process.html) (only for PDF conversion)

## Contributions

If you would like to contribute something to this package, please feel free to make a pull request and a corresponding issue and we will be happy to review and discuss.