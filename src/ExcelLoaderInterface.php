<?php

namespace LeKoala\ExcelImportExport;

/**
 */
interface ExcelLoaderInterface
{
    /**
     * @param string $file
     * @param string $name
     * @return \SilverStripe\Dev\BulkLoader_Result|string
     */
    public function load(string $file, string $name);
}
