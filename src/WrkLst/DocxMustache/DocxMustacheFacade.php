<?php

namespace WrkLst\DocxMustache;

use Illuminate\Support\Facades\Facade;

/**
 * DocxMustacheFacade.
 */
class DocxMustacheFacade extends Facade
{
    /**
     * Get the registered name of the component.
     *
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'docxmustache';
    }
}
