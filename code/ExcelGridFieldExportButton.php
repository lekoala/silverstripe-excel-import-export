<?php

/**
 * Adds an "Export list" button to the bottom of a {@link GridField}.
 */
class ExcelGridFieldExportButton implements GridField_HTMLProvider, GridField_ActionProvider,
    GridField_URLHandler
{
    /**
     * @var array Map of a property name on the exported objects, with values being the column title in the file.
     * Note that titles are only used when {@link $hasHeader} is set to TRUE.
     */
    protected $exportColumns;

    /**
     * Fragment to write the button to
     */
    protected $targetFragment;

    /**
     * @var boolean
     */
    protected $hasHeader = true;

    /**
     * @var string
     */
    protected $exportType = 'xlsx';

    /**
     * @var string
     */
    protected $exportName = null;

    /**
     * @param string $targetFragment The HTML fragment to write the button into
     * @param array $exportColumns The columns to include in the export
     */
    public function __construct($targetFragment = "after", $exportColumns = null)
    {
        $this->targetFragment = $targetFragment;
        $this->exportColumns  = $exportColumns;
    }

    /**
     * Place the export button in a <p> tag below the field
     */
    public function getHTMLFragments($gridField)
    {
        $button = new GridField_FormAction(
            $gridField, 'excelexport',
            _t('TableListField.XLSEXPORT', 'Export to Excel'), 'excelexport',
            null
        );
        $button->setAttribute('data-icon', 'download-excel');
        $button->addExtraClass('no-ajax action_export');
        $button->setForm($gridField->getForm());
        return array(
            $this->targetFragment => '<p class="grid-excel-button">'.$button->Field().'</p>',
        );
    }

    /**
     * export is an action button
     */
    public function getActions($gridField)
    {
        return array('excelexport');
    }

    public function handleAction(GridField $gridField, $actionName, $arguments,
                                 $data)
    {
        if ($actionName == 'excelexport') {
            return $this->handleExport($gridField);
        }
    }

    /**
     * it is also a URL
     */
    public function getURLHandlers($gridField)
    {
        return array(
            'excelexport' => 'handleExport',
        );
    }

    /**
     * Handle the export, for both the action button and the URL
     */
    public function handleExport($gridField, $request = null)
    {
        $now = Date("Ymd_Hi");

        if ($excel = $this->generateExportFileData($gridField)) {

            $ext      = $this->exportType;
            $name     = $this->exportName;
            $fileName = "$name-$now.$ext";

            switch ($ext) {
                case 'xls':
                    $writer = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
                    break;
                case 'xlsx';
                    $writer = PHPExcel_IOFactory::createWriter($excel,
                            'Excel2007');
                    break;
                default:
                    throw new Exception("$ext is not supported");
            }

            header('Content-type: application/vnd.ms-excel');
            header('Content-Disposition: attachment; filename="'.$fileName.'"');

            $writer->save('php://output');
            exit();
        }
    }

    /**
     * Generate export fields for Excel.
     *
     * @param GridField $gridField
     * @return PHPExcel
     */
    public function generateExportFileData($gridField)
    {
        $class    = $gridField->getModelClass();
        $columns  = ($this->exportColumns) ? $this->exportColumns : ExcelImportExport::exportFieldsForClass($class);
        $fileData = '';

        $singl = singleton($class);

        $singular = $class ? $singl->i18n_singular_name() : '';
        $plural   = $class ? $singl->i18n_plural_name() : '';

        $filter = new FileNameFilter;
        if ($this->exportName) {
            $this->exportName = $filter->filter($this->exportName);
        } else {
            $this->exportName = $filter->filter('export-'.$plural);
        }

        $excel = new PHPExcel();
        $excelProperties = $excel->getProperties();
        $excelProperties->setTitle($this->exportName);
        
        $sheet = $excel->getActiveSheet();
        if ($plural) {
            $sheet->setTitle($plural);
        }

        $row = 1;
        $col = 0;

        if ($this->hasHeader) {
            $headers = array();

            // determine the headers. If a field is callable (e.g. anonymous function) then use the
            // source name as the header instead
            foreach ($columns as $columnSource => $columnHeader) {
                $headers[] = (!is_string($columnHeader) && is_callable($columnHeader))
                        ? $columnSource : $columnHeader;
            }

            foreach ($headers as $header) {
                $sheet->setCellValueByColumnAndRow($col, $row, $header);
                $col++;
            }

            $endcol = PHPExcel_Cell::stringFromColumnIndex($col - 1);
            $sheet->setAutoFilter("A1:{$endcol}1");
            $sheet->getStyle("A1:{$endcol}1")->getFont()->setBold(true);

            $col = 0;
            $row++;
        }

        // Autosize
        $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
        try {
            $cellIterator->setIterateOnlyExistingCells(true);
        } catch (Exception $ex) {
            continue;
        }
        foreach ($cellIterator as $cell) {
            $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
        }

        //Remove GridFieldPaginator as we're going to export the entire list.
        $gridField->getConfig()->removeComponentsByType('GridFieldPaginator');

        $items = $gridField->getManipulatedList();

        // @todo should GridFieldComponents change behaviour based on whether others are available in the config?
        foreach ($gridField->getConfig()->getComponents() as $component) {
            if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                $items = $component->getManipulatedData($gridField, $items);
            }
        }

        foreach ($items->limit(null) as $item) {
            if (!$item->hasMethod('canView') || $item->canView()) {
                foreach ($columns as $columnSource => $columnHeader) {
                    if (!is_string($columnHeader) && is_callable($columnHeader)) {
                        if ($item->hasMethod($columnSource)) {
                            $relObj = $item->{$columnSource}();
                        } else {
                            $relObj = $item->relObject($columnSource);
                        }

                        $value = $columnHeader($relObj);
                    } else {
                        $value = $gridField->getDataFieldValue($item,
                            $columnSource);

                        if ($value === null) {
                            $value = $gridField->getDataFieldValue($item,
                                $columnHeader);
                        }
                    }

                    $value = str_replace(array("\r", "\n"), "\n", $value);

                    $sheet->setCellValueByColumnAndRow($col, $row, $value);
                    $col++;
                }
            }

            if ($item->hasMethod('destroy')) {
                $item->destroy();
            }

            $col = 0;
            $row++;
        }

        return $excel;
    }

    /**
     * @return array
     */
    public function getExportColumns()
    {
        return $this->exportColumns;
    }

    /**
     * @param array
     */
    public function setExportColumns($cols)
    {
        $this->exportColumns = $cols;
        return $this;
    }

    /**
     * @return boolean
     */
    public function getHasHeader()
    {
        return $this->hasHeader;
    }

    /**
     * @param boolean
     */
    public function setHasHeader($bool)
    {
        $this->hasHeader = $bool;
        return $this;
    }

    /**
     * @return string
     */
    public function getExportType()
    {
        return $this->exportType;
    }

    /**
     * @param string xlsx (default) or xls
     */
    public function setExportType($exportType)
    {
        $this->exportType = $exportType;
        return $this;
    }

    /**
     * @return string
     */
    public function getExportName()
    {
        return $this->exportName;
    }

    /**
     * @param string $exportName
     * @return \ExcelGridFieldExportButton
     */
    public function setExportName($exportName)
    {
        $this->exportName = $exportName;
        return $this;
    }
}