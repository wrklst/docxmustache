<?php

namespace WrkLst\DocxMoustache;

use Illuminate\Support\Facades\Facade;

/**
 * DocxMoustacheFacade
 *
 */
class DocxMoustacheFacade extends Facade
{
    /**
     * Get the registered name of the component.
     *
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'docxmoustache';
    }
}
