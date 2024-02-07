<?php

use LeKoala\SpreadCompat\SpreadCompat;
use LeKoala\ExcelImportExport\ExcelImportExport;

SpreadCompat::$preferredCsvAdapter = ExcelImportExport::config()->preferred_csv_adapter ?? 'Native';
SpreadCompat::$preferredXslxAdapter = ExcelImportExport::config()->preferred_xlsx_adapter ?? 'Native';
