<?php

namespace WrkLst\DocxMustache;

class DocImage
{
    public function AllowedContentTypeImages()
    {
        return [
            'image/gif'  => '.gif',
            'image/jpeg' => '.jpeg',
            'image/png'  => '.png',
            'image/bmp'  => '.bmp',
        ];
    }

    public function GetImageFromUrl($url, $manipulation)
    {
        $allowed_imgs = $this->AllowedContentTypeImages();

        if (trim($url)) {
            if ($img_file_handle = @fopen($url.$manipulation, 'rb')) {
                $img_data = stream_get_contents($img_file_handle);
                fclose($img_file_handle);
                $fi = new \finfo(FILEINFO_MIME);

                $image_mime = strstr($fi->buffer($img_data), ';', true);
                //dd($image_mime);
                if (isset($allowed_imgs[$image_mime])) {
                    return [
                        'data' => $img_data,
                        'mime' => $image_mime,
                    ];
                }
            }
        }

        return false;
    }

    public function ResampleImage($parent, $imgs, $k, $data, $dpi = 72)
    {
        \Storage::disk($parent->storageDisk)->put($parent->local_path.'word/media/'.$imgs[$k]['img_file_src'], $data);

        //rework img to new size and jpg format
        $img_rework = \Image::make($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_src']))->encode('jpg', 80);
        if ($dpi != 72) {
            $img_rework2 = \Image::make($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_src']))->encode('jpg', 80);
        }

        $imgWidth = $img_rework->width();
        $imgHeight = $img_rework->height();

        //check https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        // for EMUs calculation
        /*
        295px @72 dpi = 1530350 EMUs = Multiplier for 72dpi pixels 5187.627118644067797
        413px @72 dpi = 2142490 EMUs = Multiplier for 72dpi pixels 5187.627118644067797
        */
        $availableWidth = (int) ($imgs[$k]['cx'] / 5187.627118644067797);
        $availableHeight = (int) ($imgs[$k]['cy'] / 5187.627118644067797);

        //height based resize
        $h = (($imgHeight / $imgWidth) * $availableWidth);
        $w = (($imgWidth / $imgHeight) * $h);

        //if height based resize has too large width, do width based resize
        if ($h > $availableHeight) {
            $w = (($imgWidth / $imgHeight) * $availableHeight);
            $h = (($imgHeight / $imgWidth) * $w);
        }

        $h = null;

        //for getting non high dpi measurements, as the document is on 72 dpi.
        //TODO: this should be improved. it does not really need the resampling to identify the new sizes.
        //instead this should just be calculated, as resampling the image is too process instensive.
        $img_rework->resize($w, $h, function($constraint) {
            $constraint->aspectRatio();
            $constraint->upsize();
        });
        $new_height = $img_rework->height();
        $new_width = $img_rework->width();

        if ($dpi != 72) {
            //for storing the image in high dpi, so it has good quality on high dpi screens
            $img_rework2->resize(((int) $w * ($dpi / 72)), $h, function($constraint) { // make high dpi version for actual storage
                $constraint->aspectRatio();
                $constraint->upsize();
            });
            $img_rework2->save($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));

        //set dpi of image to high dpi
            /*$im = new \imagick();
            $im->setResolution($dpi,$dpi);
            $im->readImage($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));
            $im->writeImage($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));
            $im->clear();
            $im->destroy();
            */
        } else {
            $img_rework->save($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));
        }

        $parent->zipper->folder('word/media')->add($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));


        return [
            'height' => $new_height,
            'width'  => $new_width,
            'height_emus' => (int) ($new_height * 5187.627118644067797),
            'width_emus' => (int) ($new_width * 5187.627118644067797),
        ];
    }
}
