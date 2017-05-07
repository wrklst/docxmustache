# DocxMoustache for Laravel 5.4


Docx template manipulation class for Laravel 5.4, based on moustache templating language. This class is still under heavy development and works more like proof of concept at the moment. Things will change quickly and might break things.

![Template Example](https://github.com/wrklst/docxmoustache/raw/master/example/ExampleMoustacheTemplate.png)


## Why another openXML / docx template solution?
There are plenty of solutions out there, many with commercial licenses and some free. For our particular purpose we did not need most of the features many of the libraries out there have – we simply needed a solution that would allow to replace values and iterate through an array as well as basic image placeholder replacing.

Many of the other libraries use cloning to repeat a block. We use the usual moustache syntax to achieve repeating content in as many dimensions as needed. This introduces some issues, if the template is not setup in a manner that the content can be repeated, but in most cases it is perfectly sufficient, if the template is setup with this in mind.

## DOCX to PDF conversion

Conversion to PDF requires soffice to be installed on the server (used for conversion).
Use `sudo apt install soffice` on your ubuntu/debian based server. Also install ttf-mscorefonts if you need support for Arial font when converting docx documents to pdf `sudo apt-get install ttf-mscorefonts-installer `


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