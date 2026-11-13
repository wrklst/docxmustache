<?php

namespace WrkLst\DocxMustache;

use finfo;
use Intervention\Image\Drivers\Gd\Driver;
use Intervention\Image\Encoders\JpegEncoder;
use Intervention\Image\ImageManager;

class DocImage
{
    /**
     * EMUs (the unit of sizes in a docx) per pixel at 72 dpi: 295px = 1530350 EMUs.
     * See https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/.
     */
    private const EMUS_PER_PIXEL = 5187.627118644067797;

    public function allowedContentTypeImages()
    {
        return [
            'image/gif' => '.gif',
            'image/jpeg' => '.jpeg',
            'image/png' => '.png',
            'image/bmp' => '.bmp',
        ];
    }

    /**
     * Load an image from a url or a local path.
     *
     * @param  string  $url
     * @param  string  $manipulation  appended to the url
     * @return array{data: string, mime: string}|false false if it cannot be loaded or is no supported image
     */
    public function getImageFromUrl($url, $manipulation)
    {
        if (! trim($url) || ! ($handle = @fopen($url.$manipulation, 'rb'))) {
            return false;
        }

        $data = stream_get_contents($handle);
        fclose($handle);

        $mime = (new finfo(FILEINFO_MIME_TYPE))->buffer($data);
        if (! isset($this->allowedContentTypeImages()[$mime])) {
            return false;
        }

        return [
            'data' => $data,
            'mime' => $mime,
        ];
    }

    /**
     * Fit an image into a placeholder of $cx by $cy EMUs and encode it as jpeg.
     *
     * The size in the document is always measured at 72 dpi. A higher $dpi only
     * stores more pixels, so the picture stays sharp on high resolution screens.
     *
     * @param  string  $data  binary image
     * @return array{data: string, width: int, height: int, width_emus: int, height_emus: int}
     */
    public function resampleImage($data, $cx, $cy, $dpi = 72)
    {
        $manager = ImageManager::usingDriver(Driver::class);
        $image = $manager->decodeBinary($data);

        $imgWidth = $image->width();
        $imgHeight = $image->height();
        $availableWidth = (int) ($cx / self::EMUS_PER_PIXEL);
        $availableHeight = (int) ($cy / self::EMUS_PER_PIXEL);

        // fit to the width, or to the height if that makes the image too tall
        $h = (($imgHeight / $imgWidth) * $availableWidth);
        $w = (($imgWidth / $imgHeight) * $h);
        if ($h > $availableHeight) {
            $w = (($imgWidth / $imgHeight) * $availableHeight);
        }

        $image->scaleDown(width: (int) $w);
        $width = $image->width();
        $height = $image->height();

        if ($dpi != 72) {
            $image = $manager->decodeBinary($data)->scaleDown(width: (int) ($w * ($dpi / 72)));
        }

        return [
            'data' => (string) $image->encode(new JpegEncoder(quality: 80)),
            'width' => $width,
            'height' => $height,
            'width_emus' => (int) ($width * self::EMUS_PER_PIXEL),
            'height_emus' => (int) ($height * self::EMUS_PER_PIXEL),
        ];
    }
}
