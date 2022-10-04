<?php

namespace Varacron\SpreadsheetTable\Facades;

use Illuminate\Support\Facades\Facade;

class SpreadsheetTable extends Facade
{
    protected static function getFacadeAccessor()
    {
        return 'SpreadsheetTable';
    }
}
