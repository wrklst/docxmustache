![DocxMutache Logo](https://github.com/wrklst/docxmustache/raw/master/example/logo.png)
# DocxMustache *for Laravel 12.x.*

[![Software License](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square)](LICENSE) [![Quality Score](https://img.shields.io/scrutinizer/g/wrklst/docxmustache.svg?style=flat-square&b=master)](https://scrutinizer-ci.com/g/wrklst/docxmustache/?branch=master) [![Build Status](https://scrutinizer-ci.com/g/wrklst/docxmustache/badges/build.png?b=master)](https://scrutinizer-ci.com/g/wrklst/docxmustache/build-status/master) [![StyleCI](https://styleci.io/repos/90483440/shield?branch=master)](https://styleci.io/repos/90483440)

Docx template manipulation class for Laravel 12.x, based on [mustache templating language](https://mustache.github.io). This class is still under heavy development and works more like proof of concept at the moment. Things will change quickly and might break things.

![Template Example in Word](https://github.com/wrklst/docxmustache/raw/master/example/ExampleMustacheTemplate.png)

With DocxMustache, you can:
- Replace text using Mustache syntax.
- Embed images dynamically.
- Support basic HTML styling (bold, italic, underline).

---

## Installation

Install the package via Composer:

```bash
composer require wrklst/docxmustache
```

Refer to the [Examples](#example) section and the `examples` folder for configuration and usage instructions.

---

## Features

### HTML Conversion

Basic HTML conversion is supported, including:
- Bold (`<b>`)
- Italic (`<i>`)
- Underline (`<u>`)

**Limitations:**
- Does not support combined styling (e.g., bold + italic).
- Non-HTML values must be escaped using:

```php
htmlspecialchars($value, ENT_COMPAT, 'UTF-8');
```

**Special Note:**
To prevent unnecessary escaping of HTML, prefix the value with:

```php
*[[DONOTESCAPE]]*
```

---

### Replacing Images

Dynamic image replacement is supported. Follow these steps:

1. Add the image URL (reachable and supported format) to the image's **alt text description** field in the DOCX template.
2. Use pseudo-tags around the URL, like so:

```text
[IMG-REPLACE]http://placehold.it/350x150[/IMG-REPLACE]
```

**Note:** Images will be resampled to match the constraints of the placeholder image in the template.

---

### DOCX to PDF Conversion

To enable DOCX-to-PDF conversion, install `libreoffice-common` on your server:

```bash
sudo apt install libreoffice-common
```

For Arial font support, install:

```bash
sudo apt-get install ttf-mscorefonts-installer
```

---

## Dependencies

### Required Packages
- [mustache/mustache](https://packagist.org/packages/mustache/mustache)
- [intervention/image](http://image.intervention.io)

For Intervention Image, ensure you:
1. Add the [provider and alias](http://image.intervention.io/getting_started/installation#laravel) to your app config.
2. Install `gd` or `imagick` as required ([details here](http://image.intervention.io/getting_started/installation#laravel)).

### Laravel-Specific Dependencies
If you want to port the package to a non-Laravel environment, consider replacing:
- **Storage and File Classes**: Based on [Flysystem](https://flysystem.thephpleague.com), can be replaced with PHP native file handling.
- **Process Handling**: Uses `\Symfony\Component\Process\Process` for PDF conversion.

---

## Example

Check out the `examples` folder for sample templates and usage.

---

## Contributions

Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a pull request with your changes.
3. Include a corresponding issue for discussion.

We’ll be happy to review and discuss your ideas!

---

## Why Another DOCX Templating Solution?

While there are existing libraries for DOCX manipulation, they often:
- Use proprietary or non-standard templating syntaxes.
- Focus on complex feature sets, making them heavyweight for simple use cases.

DocxMustache was designed to:
- Use the widely adopted Mustache syntax.
- Provide simple and intuitive value and image replacement.
- Support repeating content in multiple dimensions.

### Alternatives
Here are other popular PHP libraries for DOCX manipulation:
- [openTBS – Tiny But Strong](http://www.tinybutstrong.com/opentbs.php)
- [PHPWord](https://github.com/PHPOffice/PHPWord)
- [docxtemplater pro](https://modules.docxtemplater.com) (MIT licensed; image replacement requires a commercial plugin)
- [docxpresso](http://www.docxpresso.com) (commercial)
- [phpdocx](https://www.phpdocx.com) (commercial)
