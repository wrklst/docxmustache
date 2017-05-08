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

        if ($img_file_handle = fopen($url.$manipulation, 'rb')) {
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

        return false;
    }

    public function ResampleImage($parent, $imgs, $k, $data)
    {
        \Storage::disk($parent->storageDisk)->put($parent->local_path.'word/media/'.$imgs[$k]['img_file_src'], $data);

        //rework img to new size and jpg format
        $img_rework = \Image::make($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_src']));

        $w = $imgs[$k]['width'];
        $h = $imgs[$k]['height'];

        if ($w > $h) {
            $h = null;
        } else {
            $w = null;
        }

        $img_rework->resize($w, $h, function ($constraint) {
            $constraint->aspectRatio();
            $constraint->upsize();
        });

        $new_height = $img_rework->height();
        $new_width = $img_rework->width();
        $img_rework->save($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));

        $parent->zipper->folder('word/media')->add($parent->StoragePath($parent->local_path.'word/media/'.$imgs[$k]['img_file_dest']));

        return [
            'height' => $new_height,
            'width'  => $new_width,
        ];
    }
}
