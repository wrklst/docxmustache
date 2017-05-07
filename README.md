Laravel 5.4 DocxMoustache
===============

Docx template manipulation class for Laravel 5.4, based on moustache templating language. This class is still under heavy development and works more like proof of concept at the moment. Things will change quickly and might break things.


DOCX to PDF conversion
==

Conversion to PDF requires soffice to be installed on the server (used for conversion).


HTML conversion
==

Current HTML conversion is basic and only supports singular runs of bold, itallic and underlined text and no combination of these. It requires all values non html to be escaped with htmlspecialchars($value, ENT_COMPAT, 'UTF-8') and a prefix of *[[DONOTESCAPE]]* so the class knows not to escape the html before it is converted to openXML.


Replacing images
==

The image needs to be a reachable URL with a image in a supported format. The url value needs to be placed into the alt text description field of the image.
Images will be resampled to the constraints of the placeholder image.
The Image value needs to be formated the with pseudo tags around, such as: [IMG-REPLACE]https://urldtoimg.com/img.jpg[/IMG-REPLACE]
