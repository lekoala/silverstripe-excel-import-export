<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Support class for the module
 *
 * @author Koala
 */
class ExcelImportExport
{

    public static function allFieldsForClass($class)
    {
        $dataClasses = ClassInfo::dataClassesFor($class);
        $fields      = array('ID');
        foreach ($dataClasses as $dataClass) {
            $fields = array_merge(
                $fields,
                array_keys(DataObject::database_fields($dataClass))
            );
        }
        return array_combine($fields, $fields);
    }

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
        return array_combine($exportedFields, $exportedFields);
    }

    public static function sampleFileForClass($class)
    {
        $fileName = "sample-file-for-$class.xlsx";
        $excel    = null;

        $sng = singleton($class);
        if ($sng->hasMethod('sampleExcelFile')) {
            $excel = $sng->sampleExcelFile();
            if (is_string($excel) && is_file($excel)) {
                header('Content-type: application/vnd.ms-excel');
                header('Content-Disposition: attachment; filename="' . $fileName . '"');
                readfile($excel);
                exit();
            }
        }
        if (!$excel) {
            $excel = self::generateDefaultSampleFile($class);
        }

        $writer = IOFactory::createWriter($excel, 'Xlsx');

        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $fileName . '"');
        ob_clean();
        $writer->save('php://output');
        exit();
    }

    public static function generateDefaultSampleFile($class)
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        $row = 1;
        $col = 1;
        foreach (ExcelImportExport::allFieldsForClass($class) as $header) {
            $sheet->setCellValueByColumnAndRow($col, $row, $header);
            $col++;
        }
        return $excel;
    }

    public static function getValidExtensionsText()
    {
        return _t(
            'ExcelImportExport.VALIDEXTENSIONS',
            "Allowed extensions: {extensions}",
            array('extensions' => implode(', ', self::getValidExtensions()))
        );
    }

    public static function getValidExtensions()
    {
        return array('csv', 'ods', 'xlsx', 'xls', 'txt');
    }
}
