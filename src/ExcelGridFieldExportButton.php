<?php

namespace LeKoala\ExcelImportExport;

use InvalidArgumentException;
use SilverStripe\Assets\FileNameFilter;
use SilverStripe\Forms\GridField\GridField;
use SilverStripe\Forms\GridField\GridFieldPaginator;
use SilverStripe\Forms\GridField\GridField_FormAction;
use SilverStripe\Forms\GridField\GridField_URLHandler;
use SilverStripe\Forms\GridField\GridFieldFilterHeader;
use SilverStripe\Forms\GridField\GridField_HTMLProvider;
use SilverStripe\Forms\GridField\GridFieldSortableHeader;
use SilverStripe\Forms\GridField\GridField_ActionProvider;
use LeKoala\SpreadCompat\SpreadCompat;
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
     * Map of a property name on the exported objects, with values being the column title in the file.
     * Note that titles are only used when {@link $hasHeader} is set to TRUE.
     */
    protected ?array $exportColumns;

    /**
     * Fragment to write the button to
     */
    protected string $targetFragment;

    protected bool $hasHeader = true;

    protected string $exportType = 'xlsx';

    protected ?string $exportName = null;

    protected ?string $buttonTitle = null;

    protected bool $checkCanView = true;

    protected bool $isLimited = true;

    protected array $listFilters = [];

    /**
     *
     * @var callable
     */
    protected $afterExportCallback;

    protected bool $ignoreFilters = false;

    /**
     * @param string $targetFragment The HTML fragment to write the button into
     * @param array $exportColumns The columns to include in the export
     */
    public function __construct($targetFragment = "after", $exportColumns = null)
    {
        $this->targetFragment = $targetFragment;
        $this->exportColumns = $exportColumns;
    }

    /**
     * @param GridField $gridField
     * @return string
     */
    public function getActionName($gridField)
    {
        $name = strtolower($gridField->getName());
        return 'excelexport_' . $name;
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

        $name = $this->getActionName($gridField);

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
        return array($this->getActionName($gridField));
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
        return array($this->getActionName($gridField) => 'handleExport');
    }

    /**
     * Handle the export, for both the action button and the URL
     */
    public function handleExport($gridField, $request = null)
    {
        $now = Date("Ymd_Hi");

        $this->updateExportName($gridField);

        $data = $this->generateExportFileData($gridField);

        $ext = $this->exportType;
        $name = $this->exportName;
        $fileName = "$name-$now.$ext";

        if ($this->afterExportCallback) {
            $func = $this->afterExportCallback;
            $func();
        }

        $opts = [
            'extension' => $ext,
        ];

        if ($ext != 'csv') {
            $end = ExcelImportExport::getLetter(count($this->getRealExportColumns($gridField)));
            $opts['creator'] = "SilverStripe";
            $opts['autofilter'] = "A1:{$end}1";
        }

        SpreadCompat::output($data, $fileName, ...$opts);
        exit();
    }


    /**
     * @param GridField|\LeKoala\Tabulator\TabulatorGrid $gridField
     */
    protected function updateExportName($gridField)
    {
        $filter = new FileNameFilter;
        if ($this->exportName) {
            $this->exportName = $filter->filter($this->exportName);
        } else {
            $class = $gridField->getModelClass();
            $singl = singleton($class);
            $plural = $class ? $singl->i18n_plural_name() : '';

            $this->exportName = $filter->filter('export-' . $plural);
        }
    }

    /**
     * @param GridField|\LeKoala\Tabulator\TabulatorGrid $gridField
     * @return DataList|ArrayList
     */
    protected function retrieveList($gridField)
    {
        // Remove GridFieldPaginator as we're going to export the entire list.
        $gridField->getConfig()->removeComponentsByType(GridFieldPaginator::class);

        /** @var DataList|ArrayList $items */
        $items = $gridField->getManipulatedList();

        // Keep filters
        if (!$this->ignoreFilters) {
            foreach ($gridField->getConfig()->getComponents() as $component) {
                if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                    $items = $component->getManipulatedData($gridField, $items);
                }
            }
        }

        $list = $items;
        $limit = ExcelImportExport::$limit_exports;
        if ($list instanceof DataList) {
            if ($this->isLimited && $limit > 0) {
                $list = $list->limit($limit);
            }
            if (!empty($this->listFilters)) {
                $list = $list->filter($this->listFilters);
            }
        }

        return $list;
    }

    /**
     * @param GridField|\LeKoala\Tabulator\TabulatorGrid $gridField
     */
    protected function getRealExportColumns($gridField)
    {
        $class = $gridField->getModelClass();
        return ($this->exportColumns) ? $this->exportColumns : ExcelImportExport::exportFieldsForClass($class);
    }

    /**
     * Generate export fields for Excel.
     *
     * @param GridField|\LeKoala\Tabulator\TabulatorGrid $gridField
     */
    public function generateExportFileData($gridField): iterable
    {
        $columns = $this->getRealExportColumns($gridField);

        if ($this->hasHeader) {
            $headers = [];

            // determine the headers. If a field is callable (e.g. anonymous function) then use the
            // source name as the header instead
            foreach ($columns as $columnSource => $columnHeader) {
                $headers[] = (!is_string($columnHeader) && is_callable($columnHeader))
                    ? $columnSource : $columnHeader;
            }

            yield $headers;
        }

        $list = $this->retrieveList($gridField);

        if (!$list) {
            return;
        }

        $exportFormat = ExcelImportExport::config()->export_format;

        foreach ($list as $item) {
            // This can be really slow for large exports depending on how canView is implemented
            if ($this->checkCanView) {
                $canView = true;
                if ($item->hasMethod('canView') && !$item->canView()) {
                    $canView = false;
                }
                if (!$canView) {
                    continue;
                }
            }

            $dataRow = [];

            // Loop and transforms records as needed
            foreach ($columns as $columnSource => $columnHeader) {
                if (!is_string($columnHeader) && is_callable($columnHeader)) {
                    if ($item->hasMethod($columnSource)) {
                        $relObj = $item->{$columnSource}();
                    } else {
                        $relObj = $item->relObject($columnSource);
                    }

                    $value = $columnHeader($relObj);
                } else {
                    if (is_string($columnSource)) {
                        // It can be a method
                        if (strpos($columnSource, '(') !== false) {
                            $matches = [];
                            preg_match('/([a-zA-Z]*)\((.*)\)/', $columnSource, $matches);
                            $func = $matches[1];
                            $params = explode(",", $matches[2]);
                            // Support only one param for now
                            $value = $item->$func($params[0]);
                        } else {
                            if (array_key_exists($columnSource, $exportFormat)) {
                                $format = $exportFormat[$columnSource];
                                $value = $item->dbObject($columnSource)->$format();
                            } else {
                                $value = $gridField->getDataFieldValue($item, $columnSource);
                            }
                        }
                    } else {
                        // We can also use a simple dot notation
                        $parts = explode(".", $columnHeader);
                        if (count($parts) == 1) {
                            $value = $item->$columnHeader;
                        } else {
                            $value = $item->relObject($parts[0]);
                            if ($value) {
                                $relObjField = $parts[1];
                                $value = $value->$relObjField;
                            }
                        }
                    }
                }

                $dataRow[] = $value;
            }

            if ($item->hasMethod('destroy')) {
                $item->destroy();
            }

            yield $dataRow;
        }
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
     * @param string xlsx (default), xls or csv
     */
    public function setExportType($exportType)
    {
        if (!in_array($exportType, ['xls', 'xlsx', 'csv'])) {
            throw new InvalidArgumentException("Export type must be one of : xls, xlsx, csv");
        }
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

    /**
     * Get the value of isLimited
     */
    public function getIsLimited(): bool
    {
        return $this->isLimited;
    }

    /**
     * Set the value of isLimited
     *
     * @param bool $isLimited
     */
    public function setIsLimited(bool $isLimited)
    {
        $this->isLimited = $isLimited;
        return $this;
    }

    /**
     * Get the value of ignoreFilters
     */
    public function getIgnoreFilters(): bool
    {
        return $this->ignoreFilters;
    }

    /**
     * Set the value of ignoreFilters
     *
     * @param bool $ignoreFilters
     */
    public function setIgnoreFilters(bool $ignoreFilters): self
    {
        $this->ignoreFilters = $ignoreFilters;
        return $this;
    }
}
