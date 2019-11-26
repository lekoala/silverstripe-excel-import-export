<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\Core\ClassInfo;
use SilverStripe\ORM\DataObject;
use SilverStripe\Core\Config\Config;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use SilverStripe\Core\Config\Configurable;
use Exception;

/**
 * Support class for the module
 *
 * @author Koala
 */
class ExcelImportExport
{
    use Configurable;

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
        $spreadsheet    = null;

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
            $sheet->setCellValueByColumnAndRow($col, $row, $header);
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
     * Convert a file to an array
     *
     * @param string $filepath
     * @param string $delimiter (csv only)
     * @param string $enclosure (csv only)
     * @return array
     */
    public static function fileToArray($filepath, $delimiter = ';', $enclosure = '')
    {
        $ext = pathinfo($filepath, PATHINFO_EXTENSION);
        $readerType = self::getReaderForExtension($ext);
        $reader = IOFactory::createReader($readerType);
        if ($readerType == 'Csv') {
            /* @var $reader \PhpOffice\PhpSpreadsheet\Writer\Csv */
            // @link https://phpspreadsheet.readthedocs.io/en/latest/topics/reading-and-writing-to-file/#setting-csv-options_1
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
     * @return array
     */
    public static function excelToArray($filepath, $sheetname = null, $onlyExisting = true)
    {
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
                    $cellIterator->setIterateOnlyExistingCells(true);
                    $cells = [];
                    foreach ($cellIterator as $cell) {
                        $cells[] = $cell->getFormattedValue();
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
     * Convert an excel file to an associative array
     *
     * Suppose the first line are the headers of the file
     * Headers are trimmed in case you have crappy whitespace in your files
     *
     * @param string $filepath
     * @param string $sheetname Load a specific worksheet by name
     * @return array
     */
    public static function excelToAssocArray($filepath, $sheetname = null)
    {
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
            $data = [];
            $headers = [];
            $headersCount = 0;
            foreach ($excel->getActiveSheet()->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(true);
                $cells = [];
                foreach ($cellIterator as $cell) {
                    $cells[] = $cell->getFormattedValue();
                }
                if (empty($headers)) {
                    $headers = $cells;
                    $headers = array_map(function ($v) {
                        // trim does not always work great
                        return is_string($v) ? preg_replace('/(^\s+)|(\s+$)/us', '', $v) : $v;
                    }, $headers);
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
