<?php

namespace Varacron\SpreadsheetTable;

use Illuminate\Support\ServiceProvider;

use Esemve\Hook\Facades\Hook;

class SpreadsheetTableServiceProvider extends ServiceProvider
{
    public function register()
    {
        $this->app->singleton('SpreadsheetTable', function () {
            return new SpreadsheetTable();
        });
    }
}
