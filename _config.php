<?php

use LeKoala\SpreadCompat\SpreadCompat;
use LeKoala\ExcelImportExport\ExcelImportExport;

if (ExcelImportExport::config()->preferred_csv_adapter) {
    SpreadCompat::$preferredCsvAdapter = ExcelImportExport::config()->preferred_csv_adapter;
}
if (ExcelImportExport::config()->preferred_xlsx_adapter) {
    SpreadCompat::$preferredXslxAdapter = ExcelImportExport::config()->preferred_xlsx_adapter;
}
