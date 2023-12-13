<?php

namespace LeKoala\ExcelImportExport;

use Exception;
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
        $fields      = array();
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
        $fileName = "sample-file-for-$class.xlsx";
        $spreadsheet = null;

        $sng = singleton($class);
        if ($sng->hasMethod('sampleExcelFile')) {
            $spreadsheet = $sng->sampleExcelFile();

            // We have a file, output directly
            if (is_string($spreadsheet) && is_file($spreadsheet)) {
                self::outputHeaders($fileName);
                readfile($spreadsheet);
                exit();
            }
        }
        if (!$spreadsheet) {
            $spreadsheet = self::generateDefaultSampleFile($class);
        }

        $writer = self::getDefaultWriter($spreadsheet);
        self::outputHeaders($fileName);
        $writer->save('php://output');
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

    public static function getDefaultWriter($spreadsheet)
    {
        return IOFactory::createWriter($spreadsheet, self::config()->default_writer);
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
     *
     * @param string $class
     * @return Spreadsheet
     */
    public static function generateDefaultSampleFile($class)
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getProperties()
            ->setCreator('SilverStripe')
            ->setTitle("Sample file for $class");
        $sheet = $spreadsheet->getActiveSheet();

        $row = 1;
        $col = 1;
        $allFields = ExcelImportExport::importFieldsForClass($class);
        foreach ($allFields as $header) {
            $sheet->setCellValue([$col, $row], $header);
            $col++;
        }
        return $spreadsheet;
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
     * Extracted from PHPSpreadhseet
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
     * @param array $data
     * @param string $filepath
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @return void
     */
    public static function arrayToFile($data, $filepath)
    {
        $spreadsheet = new Spreadsheet;
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->fromArray($data);

        $ext = pathinfo($filepath, PATHINFO_EXTENSION);

        // Writer is the same as read : Csv, Xlsx...
        $writerType = self::getReaderForExtension($ext);
        $writer = IOFactory::createWriter($spreadsheet, $writerType);
        $writer->save($filepath);
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
        $fp = fopen($filepath, 'w');
        // UTF 8 fix
        fprintf($fp, "\xEF\xBB\xBF");
        foreach ($data as $row) {
            fputcsv($fp, $row, $delimiter, $enclosure, $escapeChar);
        }
        return fclose($fp);
    }


    /**
     * @param IReader $reader
     * @return Csv
     */
    protected static function getCsvReader(IReader $reader)
    {
        return $reader;
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
        $readerType = self::getReaderForExtension($ext);
        $reader = IOFactory::createReader($readerType);
        if ($readerType == 'Csv') {
            // @link https://phpspreadsheet.readthedocs.io/en/latest/topics/reading-and-writing-to-file/#setting-csv-options_1
            $reader = self::getCsvReader($reader);
            $reader->setDelimiter($delimiter);
            $reader->setEnclosure($enclosure);
        } else {
            // Does not apply to CSV
            $reader->setReadDataOnly(true);
        }
        $data = array();
        if ($reader->canRead($filepath)) {
            $excel = $reader->load($filepath);
            $data = $excel->getActiveSheet()->toArray(null, true, false, false);
        } else {
            throw new Exception("Cannot read $filepath");
        }
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
        $data = array();
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
        if ($ext === null) {
            $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        }
        $readerType = self::getReaderForExtension($ext);
        $reader = IOFactory::createReader($readerType);
        $reader->setReadDataOnly(true);
        if ($sheetname) {
            $reader->setLoadSheetsOnly($sheetname);
        }
        $data = array();
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
