<?php

namespace LeKoala\ExcelImportExport;

use Exception;
use Generator;
use LeKoala\SpreadCompat\Common\Options;
use LeKoala\SpreadCompat\SpreadCompat;
use SilverStripe\Forms\Form;
use SilverStripe\Core\ClassInfo;
use SilverStripe\ORM\DataObject;
use SilverStripe\Control\Controller;
use SilverStripe\Core\Config\Config;
use SilverStripe\Control\HTTPResponse;
use PhpOffice\PhpSpreadsheet\IOFactory;
use SilverStripe\Dev\BulkLoader_Result;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use SilverStripe\Core\Config\Configurable;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use SilverStripe\Assets\FileNameFilter;

/**
 * Support class for the module
 *
 * @author Koala
 */
class ExcelImportExport
{
    use Configurable;

    /**
     * Setting this to false improve performance but may lead to skipped cells
     * @var bool
     */
    public static $iterate_only_existing_cells = false;

    /**
     * Useful if importing only one sheet or if computation fails
     * @var bool
     */
    public static $use_old_calculated_value = false;

    /**
     * Some excel sheets need extra processing
     * @var boolean
     */
    public static $process_headers = false;

    /**
     * @var string
     */
    public static $default_tmp_reader = 'Xlsx';

    /**
     * @var integer
     */
    public static $limit_exports = 1000;

    /**
     * Get all db fields for a given dataobject class
     *
     * @param string $class
     * @return array
     */
    public static function allFieldsForClass($class)
    {
        $dataClasses = ClassInfo::dataClassesFor($class);
        $fields      = [];
        $dataObjectSchema = DataObject::getSchema();
        foreach ($dataClasses as $dataClass) {
            $dataFields = $dataObjectSchema->databaseFields($dataClass);
            $fields = array_merge($fields, array_keys($dataFields));
        }
        return array_combine($fields, $fields);
    }

    /**
     * Get fields that should be exported by default for a class
     *
     * @param string $class
     * @return array
     */
    public static function exportFieldsForClass($class)
    {
        $singl = singleton($class);
        if ($singl->hasMethod('exportedFields')) {
            return $singl->exportedFields();
        }
        $exportedFields = Config::inst()->get($class, 'exported_fields');

        if (!$exportedFields) {
            $exportedFields = array_keys(self::allFieldsForClass($class));
        }

        $unexportedFields = Config::inst()->get($class, 'unexported_fields');

        if ($unexportedFields) {
            $exportedFields = array_diff($exportedFields, $unexportedFields);
        }

        $fields = [];
        foreach ($exportedFields as $key => $value) {
            if (is_int($key)) {
                $key = $value;
            }
            $fields[$key] = $value;
        }

        return $fields;
    }

    /**
     * Get fields that can be imported by default for a class
     *
     * @param string $class
     * @return array
     */
    public static function importFieldsForClass($class)
    {
        $singl = singleton($class);
        if ($singl->hasMethod('importedFields')) {
            return $singl->importedFields();
        }
        $importedFields = Config::inst()->get($class, 'imported_fields');

        if (!$importedFields) {
            $importedFields = array_keys(self::allFieldsForClass($class));
        }

        $unimportedFields = Config::inst()->get($class, 'unimported_Fields');

        if ($unimportedFields) {
            $importedFields = array_diff($importedFields, $unimportedFields);
        }
        return array_combine($importedFields, $importedFields);
    }

    /**
     * Output a sample file for a class
     *
     * A custom file can be provided with a custom sampleExcelFile method
     * either as a file path or as a Excel instance
     *
     * @param string $class
     * @return void
     */
    public static function sampleFileForClass($class)
    {
        $ext = self::getDefaultExtension();
        $filter = new FileNameFilter();
        $fileName = $filter->filter("sample-file-for-$class.$ext");
        $spreadsheet = null;

        $sng = singleton($class);

        // Deprecated
        if ($sng->hasMethod('sampleExcelFile')) {
            $spreadsheet = $sng->sampleExcelFile();
            // PHPSpreadsheet is required for this
            $writer = self::getDefaultWriter($spreadsheet);
            self::outputHeaders($fileName);
            $writer->save('php://output');
            exit();
        }

        if ($sng->hasMethod('sampleImportData')) {
            $data = $sng->sampleImportData();
        } else {
            // Simply output the default headers
            $data = [ExcelImportExport::importFieldsForClass($class)];
        }

        if (!is_iterable($data)) {
            throw new Exception("`sampleImportData` must return an iterable");
        }

        $options = new Options();
        $options->creator = "SilverStripe";
        SpreadCompat::output($data, $fileName, $options);
        exit();
    }

    /**
     * @param Controller $controller
     * @return bool
     */
    public static function checkImportForm($controller)
    {
        if (!$controller->showImportForm) {
            return false;
        }
        $modelClass = $controller->getModelClass();
        if (is_array($controller->showImportForm)) {
            /** @var array $valid */
            $valid = $controller->showImportForm;
            if (!in_array($modelClass, $valid)) {
                return false;
            }
        }
        return true;
    }

    /**
     * @param string $handler
     * @param Form $form
     * @param Controller $controller
     * @return HTTPResponse
     */
    public static function useCustomHandler($handler, Form $form, Controller $controller)
    {
        if (!$handler || !method_exists($handler, "load")) {
            $form->sessionMessage("Invalid handler: $handler", 'bad');
            return $controller->redirectBack();
        }
        $file = $_FILES['_CsvFile']['tmp_name'];
        $name = $_FILES['_CsvFile']['name'];
        $inst = new $handler();

        if (!empty($_POST['OnlyUpdateRecords']) && method_exists($handler, 'setOnlyUpdate')) {
            $inst->setOnlyUpdate(true);
        }

        /** @var BulkLoader_Result|string $results  */
        try {
            $results = $inst->load($file, $name);
        } catch (Exception $e) {
            $form->sessionMessage($e->getMessage(), 'bad');
            return $controller->redirectBack();
        }

        $message = '';
        if ($results instanceof BulkLoader_Result) {
            if ($results->CreatedCount()) {
                $message .= _t(
                    'ModelAdmin.IMPORTEDRECORDS',
                    "Imported {count} records.",
                    ['count' => $results->CreatedCount()]
                );
            }
            if ($results->UpdatedCount()) {
                $message .= _t(
                    'ModelAdmin.UPDATEDRECORDS',
                    "Updated {count} records.",
                    ['count' => $results->UpdatedCount()]
                );
            }
            if ($results->DeletedCount()) {
                $message .= _t(
                    'ModelAdmin.DELETEDRECORDS',
                    "Deleted {count} records.",
                    ['count' => $results->DeletedCount()]
                );
            }
            if (!$results->CreatedCount() && !$results->UpdatedCount()) {
                $message .= _t('ModelAdmin.NOIMPORT', "Nothing to import");
            }
        } else {
            // Or we have a simple result
            $message = $results;
        }

        $form->sessionMessage($message, 'good');
        return $controller->redirectBack();
    }

    public static function getDefaultExtension()
    {
        return self::config()->default_extension ?? 'xlsx';
    }

    /**
     * Get default writer for PHPSpreadsheet if installed
     *
     * @param Spreadsheet $spreadsheet
     * @return IWriter
     */
    public static function getDefaultWriter(Spreadsheet $spreadsheet): IWriter
    {
        if (!self::isPhpSpreadsheetAvailable()) {
            throw new Exception("PHPSpreadsheet is not installed");
        }
        $writer = ucfirst(self::getDefaultExtension());
        return IOFactory::createWriter($spreadsheet, $writer);
    }

    /**
     * Output excel headers
     *
     * @param string $fileName
     * @return void
     */
    public static function outputHeaders($fileName)
    {
        $ext = pathinfo($fileName, PATHINFO_EXTENSION);
        switch ($ext) {
            case 'csv':
                header('Content-Type: text/csv');
                break;
            case 'xlsx':
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                break;
            default:
                header('Content-type: application/vnd.ms-excel');
                break;
        }

        header('Content-Disposition: attachment; filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');
        ob_clean();
    }

    /**
     * Generate a default import file with all field name
     * @deprecated
     * @param string $class
     * @return string
     */
    public static function generateDefaultSampleFile($class)
    {
        $opts = [
            'creator' => "SilverStripe"
        ];
        $allFields = ExcelImportExport::importFieldsForClass($class);
        $tmpname = SpreadCompat::getTempFilename();
        SpreadCompat::write([
            $allFields
        ], $tmpname, ...$opts);
        return $tmpname;
    }

    /**
     * Show valid extensions helper (for uploaders)
     *
     * @return string
     */
    public static function getValidExtensionsText()
    {
        return _t(
            'ExcelImportExport.VALIDEXTENSIONS',
            "Allowed extensions: {extensions}",
            array('extensions' => implode(', ', self::getValidExtensions()))
        );
    }

    /**
     * Extracted from PHPSpreadsheet
     *
     * @param string $ext
     * @return string
     */
    public static function getReaderForExtension($ext)
    {
        switch (strtolower($ext)) {
            case 'xlsx': // Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm': // Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return 'Xlsx';
            case 'xls': // Excel (BIFF) Spreadsheet
            case 'xlt': // Excel (BIFF) Template
                return 'Xls';
            case 'ods': // Open/Libre Offic Calc
            case 'ots': // Open/Libre Offic Calc Template
                return 'Ods';
            case 'slk':
                return 'Slk';
            case 'xml': // Excel 2003 SpreadSheetML
                return 'Xml';
            case 'gnumeric':
                return 'Gnumeric';
            case 'htm':
            case 'html':
                return 'Html';
            case 'csv':
                return 'Csv';
            case 'tmp': // Useful when importing uploaded files
                return self::$default_tmp_reader;
            default:
                throw new Exception("Unsupported file type : $ext");
        }
    }

    public static function isPhpSpreadsheetAvailable()
    {
        return class_exists(\PhpOffice\PhpSpreadsheet\Spreadsheet::class);
    }

    /**
     * If you exported separated files, you can merge them in one big file
     * Requires PHPSpreadsheet
     * @param array $files
     * @return Spreadsheet
     */
    public static function mergeExcelFiles($files)
    {
        if (!self::isPhpSpreadsheetAvailable()) {
            throw new Exception("PHPSpreadsheet is not installed");
        }
        $merged = new Spreadsheet;
        $merged->removeSheetByIndex(0);
        foreach ($files as $filename) {
            $remoteExcel = IOFactory::load($filename);
            $merged->addExternalSheet($remoteExcel->getActiveSheet());
        }
        return $merged;
    }

    /**
     * Get valid extensions
     *
     * @return array
     */
    public static function getValidExtensions()
    {
        $v = self::config()->allowed_extensions;
        if (!$v || !is_array($v)) {
            return [];
        }
        return $v;
    }

    /**
     * Save content of an array to a file
     *
     * @param iterable $data
     * @param string $filepath
     * @return void
     */
    public static function arrayToFile($data, $filepath)
    {
        SpreadCompat::write($data, $filepath);
    }

    /**
     * Fast saving to csv
     *
     * @param array $data
     * @param string $filepath
     * @param string $delimiter
     * @param string $enclosure
     * @param string $escapeChar
     */
    public static function arrayToCsv($data, $filepath, $delimiter = ',', $enclosure = '"', $escapeChar = '\\')
    {
        if (is_file($filepath)) {
            unlink($filepath);
        }
        $options = new Options();
        $options->separator = $delimiter;
        $options->enclosure = $enclosure;
        $options->escape = $escapeChar;
        SpreadCompat::write($data, $filepath, $options);
    }

    public static function excelColumnRange(string $lower = 'A', string $upper = 'ZZ'): Generator
    {
        ++$upper;
        for ($i = $lower; $i !== $upper; ++$i) {
            yield $i;
        }
    }

    /**
     * String from column index.
     *
     * @param int $index Column index (1 = A)
     * @param $fallback
     * @return string
     */
    public static function getLetter($index)
    {
        foreach (self::excelColumnRange() as $letter) {
            $index--;
            if ($index <= 0) {
                return $letter;
            }
        }
    }

    /**
     * Convert a file to an array
     *
     * @param string $filepath
     * @param string $delimiter (csv only)
     * @param string $enclosure (csv only)
     * @param string $ext if extension cannot be deducted from filepath (eg temp files)
     * @return array
     */
    public static function fileToArray($filepath, $delimiter = ';', $enclosure = '', $ext = null)
    {
        if ($ext === null) {
            $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        }

        $data = iterator_to_array(
            SpreadCompat::read(
                $filepath,
                extension: $ext,
                separator: $delimiter,
                enclosure: $enclosure
            )
        );
        return $data;
    }

    /**
     * Convert an excel file to an array
     *
     * @param string $filepath
     * @param string $sheetname Load a specific worksheet by name
     * @param true $onlyExisting Avoid reading empty columns
     * @param string $ext if extension cannot be deducted from filepath (eg temp files)
     * @return array
     */
    public static function excelToArray($filepath, $sheetname = null, $onlyExisting = true, $ext = null)
    {
        if (!self::isPhpSpreadsheetAvailable()) {
            throw new Exception("PHPSpreadsheet is not installed");
        }
        if ($ext === null) {
            $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        }
        $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        $readerType = self::getReaderForExtension($ext);
        $reader = IOFactory::createReader($readerType);
        $reader->setReadDataOnly(true);
        if ($sheetname) {
            $reader->setLoadSheetsOnly($sheetname);
        }
        $data = [];
        if ($reader->canRead($filepath)) {
            $excel = $reader->load($filepath);
            if ($onlyExisting) {
                $data = [];
                foreach ($excel->getActiveSheet()->getRowIterator() as $row) {
                    $cellIterator = $row->getCellIterator();
                    if (self::$iterate_only_existing_cells) {
                        $cellIterator->setIterateOnlyExistingCells(true);
                    }
                    $cells = [];
                    foreach ($cellIterator as $cell) {
                        if (self::$use_old_calculated_value) {
                            $cells[] = $cell->getOldCalculatedValue();
                        } else {
                            $cells[] = $cell->getFormattedValue();
                        }
                    }
                    $data[] = $cells;
                }
            } else {
                $data = $excel->getActiveSheet()->toArray(null, true, false, false);
            }
        } else {
            throw new Exception("Cannot read $filepath");
        }
        return $data;
    }

    /**
     * @link https://stackoverflow.com/questions/44304795/how-to-retrieve-date-from-table-cell-using-phpspreadsheet#44304796
     * @param int $v
     * @return string
     */
    public static function convertExcelDate($v)
    {
        if (!is_numeric($v)) {
            return '';
        }
        if (!self::isPhpSpreadsheetAvailable()) {
            throw new Exception("PHPSpreadsheet is not installed");
        }
        return date('Y-m-d', Date::excelToTimestamp($v));
    }

    /**
     * Convert an excel file to an associative array
     *
     * Suppose the first line are the headers of the file
     * Headers are trimmed in case you have crappy whitespace in your files
     *
     * @param string $filepath
     * @param string $sheetname Load a specific worksheet by name
     * @param string $ext if extension cannot be deducted from filepath (eg temp files)
     * @return array
     */
    public static function excelToAssocArray($filepath, $sheetname = null, $ext = null)
    {
        if (!self::isPhpSpreadsheetAvailable()) {
            throw new Exception("PHPSpreadsheet is not installed");
        }
        if ($ext === null) {
            $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        }
        $readerType = self::getReaderForExtension($ext);
        $reader = IOFactory::createReader($readerType);
        $reader->setReadDataOnly(true);
        if ($sheetname) {
            $reader->setLoadSheetsOnly($sheetname);
        }
        $data = [];
        if ($reader->canRead($filepath)) {
            $excel = $reader->load($filepath);
            $data = [];
            $headers = [];
            $headersCount = 0;
            foreach ($excel->getActiveSheet()->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                if (self::$iterate_only_existing_cells) {
                    $cellIterator->setIterateOnlyExistingCells(true);
                }
                $cells = [];
                foreach ($cellIterator as $cell) {
                    if (self::$use_old_calculated_value) {
                        $cells[] = $cell->getOldCalculatedValue();
                    } else {
                        $cells[] = $cell->getFormattedValue();
                    }
                }
                if (empty($headers)) {
                    $headers = $cells;
                    // Some crappy excel file may need this
                    if (self::$process_headers) {
                        $headers = array_map(function ($v) {
                            // Numeric headers are most of the time dates
                            if (is_numeric($v)) {
                                $v =  self::convertExcelDate($v);
                            }
                            // trim does not always work great and headers can contain utf8 stuff
                            return is_string($v) ? preg_replace('/(^\s+)|(\s+$)/us', '', $v) : $v;
                        }, $headers);
                    }
                    $headersCount = count($headers);
                } else {
                    $diff = count($cells) - $headersCount;
                    if ($diff != 0) {
                        if ($diff > 0) {
                            // we have too many cells
                            $cells = array_slice($cells, 0, $headersCount);
                        } else {
                            // we are missing some cells
                            for ($i = 0; $i < abs($diff); $i++) {
                                $cells[] = null;
                            }
                        }
                    }
                    $data[] = array_combine($headers, $cells);
                }
            }
        } else {
            throw new Exception("Cannot read $filepath");
        }
        return $data;
    }
}
