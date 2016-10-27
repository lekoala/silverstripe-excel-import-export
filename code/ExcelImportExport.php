<?php

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
        $fields      = array();
        foreach ($dataClasses as $dataClass) {
            $fields = array_merge($fields,
                array_keys(DataObject::database_fields($dataClass)));
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
        $excel = new PHPExcel();
        $sheet = $excel->getActiveSheet();

        $row = 1;
        $col = 0;
        foreach (ExcelImportExport::allFieldsForClass($class) as $header) {
            $sheet->setCellValueByColumnAndRow($col, $row, $header);
            $col++;
        }

        $fileName = "sample-file-for-$class.xlsx";

        $writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="'.$fileName.'"');
        ob_clean();
        $writer->save('php://output');
        exit();
    }

    public static function getValidExtensionsText()
    {
        return _t('ExcelImportExport.VALIDEXTENSIONS',
            "Allowed extensions: {extensions}",
            array('extensions' => implode(', ', self::getValidExtensions())));
    }

    public static function getValidExtensions()
    {
        return array('csv', 'ods', 'xlsx', 'xls', 'txt');
    }
}