<?php

namespace LeKoala\ExcelImportExport;

use PhpOffice\PhpSpreadsheet\IOFactory;
use SilverStripe\Assets\FileNameFilter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use SilverStripe\Forms\GridField\GridField;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use SilverStripe\Forms\GridField\GridFieldPaginator;
use SilverStripe\Forms\GridField\GridField_FormAction;
use SilverStripe\Forms\GridField\GridField_URLHandler;
use SilverStripe\Forms\GridField\GridFieldFilterHeader;
use SilverStripe\Forms\GridField\GridField_HTMLProvider;
use SilverStripe\Forms\GridField\GridFieldSortableHeader;
use SilverStripe\Forms\GridField\GridField_ActionProvider;
use Exception;
use SilverStripe\ORM\DataList;

/**
 * Adds an "Export list" button to the bottom of a {@link GridField}.
 */
class ExcelGridFieldExportButton implements
    GridField_HTMLProvider,
    GridField_ActionProvider,
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
     *
     * @var string
     */
    protected $buttonTitle = null;

    /**
     *
     * @var bool
     */
    protected $checkCanView = true;

    /**
     *
     * @var array
     */
    protected $listFilters = array();

    /**
     *
     * @var callable
     */
    protected $afterExportCallback;

    /**
     * Static instance counter to allow multiple instances to work together
     * @var int
     */
    protected static $instances = 0;

    /**
     * Current instance count
     * @var int
     */
    protected $instance;

    /**
     * @param string $targetFragment The HTML fragment to write the button into
     * @param array $exportColumns The columns to include in the export
     */
    public function __construct($targetFragment = "after", $exportColumns = null)
    {
        $this->targetFragment = $targetFragment;
        $this->exportColumns = $exportColumns;
        self::$instances++;
        $this->instance = self::$instances;
    }

    public function getActionName()
    {
        return 'excelexport_' . $this->instance;
    }

    /**
     * Place the export button in a <p> tag below the field
     */
    public function getHTMLFragments($gridField)
    {
        $title = $this->buttonTitle ? $this->buttonTitle : _t(
            'ExcelImportExport.XLSEXPORT',
            'Export to Excel'
        );

        $name = $this->getActionName();

        $button = new GridField_FormAction(
            $gridField,
            $name,
            $title,
            $name,
            null
        );
        $button->addExtraClass('btn btn-secondary no-ajax font-icon-down-circled action_export');
        $button->setForm($gridField->getForm());

        return array(
            $this->targetFragment => $button->Field()
        );
    }

    /**
     * export is an action button
     */
    public function getActions($gridField)
    {
        return array($this->getActionName());
    }

    public function handleAction(
        GridField $gridField,
        $actionName,
        $arguments,
        $data
    ) {
        if (in_array($actionName, $this->getActions($gridField))) {
            return $this->handleExport($gridField);
        }
    }

    /**
     * it is also a URL
     */
    public function getURLHandlers($gridField)
    {
        return array($this->getActionName() => 'handleExport');
    }

    /**
     * Handle the export, for both the action button and the URL
     */
    public function handleExport($gridField, $request = null)
    {
        $now = Date("Ymd_Hi");

        if ($excel = $this->generateExportFileData($gridField)) {
            $ext = $this->exportType;
            $name = $this->exportName;
            $fileName = "$name-$now.$ext";

            switch ($ext) {
                case 'xls':
                    $writer = IOFactory::createWriter($excel, 'Xls');
                    break;
                case 'xlsx':
                    $writer = IOFactory::createWriter($excel, 'Xlsx');
                    break;
                default:
                    throw new Exception("$ext is not supported");
            }
            if ($this->afterExportCallback) {
                $func = $this->afterExportCallback;
                $func();
            }

            header('Content-type: application/vnd.ms-excel');
            header('Content-Disposition: attachment; filename="' . $fileName . '"');

            $writer->save('php://output');
            exit();
        }
    }

    /**
     * Generate export fields for Excel.
     *
     * @param GridField $gridField
     * @return Spreadsheet
     */
    public function generateExportFileData($gridField)
    {
        $class = $gridField->getModelClass();
        $columns = ($this->exportColumns) ? $this->exportColumns : ExcelImportExport::exportFieldsForClass($class);

        $singl = singleton($class);

        $plural = $class ? $singl->i18n_plural_name() : '';

        $filter = new FileNameFilter;
        if ($this->exportName) {
            $this->exportName = $filter->filter($this->exportName);
        } else {
            $this->exportName = $filter->filter('export-' . $plural);
        }

        $excel = new Spreadsheet();
        $excelProperties = $excel->getProperties();
        $excelProperties->setTitle($this->exportName);

        $sheet = $excel->getActiveSheet();
        if ($plural) {
            $sheet->setTitle($plural);
        }

        $row = 1;
        $col = 1;

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

            $endcol = Coordinate::stringFromColumnIndex($col - 1);
            $sheet->setAutoFilter("A1:{$endcol}1");
            $sheet->getStyle("A1:{$endcol}1")->getFont()->setBold(true);

            $col = 1;
            $row++;
        }

        // Autosize
        $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
        try {
            $cellIterator->setIterateOnlyExistingCells(true);
        } catch (Exception $ex) {
            // Ignore exceptions
        }
        foreach ($cellIterator as $cell) {
            $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
        }

        //Remove GridFieldPaginator as we're going to export the entire list.
        $gridField->getConfig()->removeComponentsByType(GridFieldPaginator::class);

        $items = $gridField->getManipulatedList();

        // @todo should GridFieldComponents change behaviour based on whether others are available in the config?
        foreach ($gridField->getConfig()->getComponents() as $component) {
            if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                $items = $component->getManipulatedData($gridField, $items);
            }
        }

        if ($items instanceof DataList) {
            $list = $items->limit(null);
        }
        if (!empty($this->listFilters)) {
            $list = $list->filter($this->listFilters);
        }

        foreach ($list as $item) {
            if (!$this->checkCanView || !$item->hasMethod('canView') || $item->canView()) {
                foreach ($columns as $columnSource => $columnHeader) {
                    if (!is_string($columnHeader) && is_callable($columnHeader)) {
                        if ($item->hasMethod($columnSource)) {
                            $relObj = $item->{$columnSource}();
                        } else {
                            $relObj = $item->relObject($columnSource);
                        }

                        $value = $columnHeader($relObj);
                    } else {
                        $value = $gridField->getDataFieldValue($item, $columnSource);
                    }

                    $sheet->setCellValueByColumnAndRow($col, $row, $value);
                    $col++;
                }
            }

            if ($item->hasMethod('destroy')) {
                $item->destroy();
            }

            $col = 1;
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
     * @return ExcelGridFieldExportButton
     */
    public function setExportName($exportName)
    {
        $this->exportName = $exportName;
        return $this;
    }

    /**
     * @return string
     */
    public function getButtonTitle()
    {
        return $this->buttonTitle;
    }

    /**
     * @param string $buttonTitle
     * @return ExcelGridFieldExportButton
     */
    public function setButtonTitle($buttonTitle)
    {
        $this->buttonTitle = $buttonTitle;
        return $this;
    }

    /**
     *
     * @return bool
     */
    public function getCheckCanView()
    {
        return $this->checkCanView;
    }

    /**
     *
     * @param bool $checkCanView
     * @return ExcelGridFieldExportButton
     */
    public function setCheckCanView($checkCanView)
    {
        $this->checkCanView = $checkCanView;
        return $this;
    }

    /**
     *
     * @return array
     */
    public function getListFilters()
    {
        return $this->listFilters;
    }

    /**
     *
     * @param array $listFilters
     * @return ExcelGridFieldExportButton
     */
    public function setListFilters($listFilters)
    {
        $this->listFilters = $listFilters;
        return $this;
    }

    /**
     *
     * @return callable
     */
    public function getAfterExportCallback()
    {
        return $this->afterExportCallback;
    }

    /**
     *
     * @param callable $afterExportCallback
     * @return ExcelGridFieldExportButton
     */
    public function setAfterExportCallback(callable $afterExportCallback)
    {
        $this->afterExportCallback = $afterExportCallback;
        return $this;
    }
}
