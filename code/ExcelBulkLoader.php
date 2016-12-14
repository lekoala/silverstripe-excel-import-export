<?php

/**
 * Use PHPExcel to expand BulkLoader file format support
 *
 * @author Koala
 */
class ExcelBulkLoader extends BulkLoader
{
    /**
     * Delimiter character (Default: comma).
     *
     * @var string
     */
    public $delimiter = ',';

    /**
     * Enclosure character (Default: doublequote)
     *
     * @var string
     */
    public $enclosure = '"';

    /**
     * Identifies if csv the has a header row.
     *
     * @var boolean
     */
    public $hasHeaderRow = true;

    /**
     *
     * @var array
     */
    protected $uploadFile = null;

    /**
     *
     * @var DataObject
     */
    protected $singleton = null;

    /**
     * @var array
     */
    protected $db = array();

    /**
     * Type of file if not able to determine through uploaded file
     * 
     * @var string
     */
    protected $fileType = 'xlsx';

    /**
     * @inheritDoc
     */
    public function preview($filepath)
    {
        return $this->processAll($filepath, true);
    }

    /**
     * Load the given file via {@link self::processAll()} and {@link self::processRecord()}.
     * Optionally truncates (clear) the table before it imports.
     *
     * @return BulkLoader_Result See {@link self::processAll()}
     */
    public function load($filepath)
    {
        // A small hack to allow model admin import form to work properly
        if (!is_array($filepath) && isset($_FILES['_CsvFile'])) {
            $filepath = $_FILES['_CsvFile'];
        }
        if (is_array($filepath)) {
            $this->uploadFile = $filepath;
            $filepath         = $filepath['tmp_name'];
        }

        increase_time_limit_to();
        increase_memory_limit_to('512M');

        //get all instances of the to be imported data object
        if ($this->deleteExistingRecords) {
            DataObject::get($this->objectClass)->removeAll();
        }

        return $this->processAll($filepath);
    }

    protected function getReaderType($ext)
    {
        $readerType = null;
        switch (strtolower($ext)) {
            case 'xlsx':   //	Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm':   //	Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx':   //	Excel (OfficeOpenXML) Template
            case 'xltm':   //	Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                $readerType = 'Excel2007';
                break;
            case 'xls':    //	Excel (BIFF) Spreadsheet
            case 'xlt':    //	Excel (BIFF) Template
                $readerType = 'Excel5';
                break;
            case 'ods':    //	Open/Libre Offic Calc
            case 'ots':    //	Open/Libre Offic Calc Template
                $readerType = 'OOCalc';
                break;
            case 'slk':
                $readerType = 'SYLK';
                break;
            case 'xml':    //	Excel 2003 SpreadSheetML
                $readerType = 'Excel2003XML';
                break;
            case 'gnumeric':
                $readerType = 'Gnumeric';
                break;
            case 'htm':
            case 'html':
                $readerType = 'HTML';
                break;
            case 'csv':
                $readerType = 'CSV';
                break;
            default:
                throw new Exception("Unsupported extension: $ext");
        }

        return $readerType;
    }

    protected function getUploadFileExtension()
    {
        if ($this->uploadFile) {
            return pathinfo($this->uploadFile['name'], PATHINFO_EXTENSION);
        }
        return $this->fileType;
    }

    protected function mergeRowWithHeaders($row, $headers, $headersCount = null)
    {
        if ($headersCount === null) {
            $headersCount = count($headers);
        }
        $row = array_slice($row, 0, $headersCount);
        $row = array_combine($headers, $row);
        return $row;
    }

    /**
     * @param string $filepath
     * @param boolean $preview
     */
    protected function processAll($filepath, $preview = false)
    {
        $results = new BulkLoader_Result();
        $ext     = $this->getUploadFileExtension();

        $readerType = $this->getReaderType($ext);
        $reader     = PHPExcel_IOFactory::createReader($readerType);
        $reader->setReadDataOnly(true);
        if ($readerType == 'CSV') {
            /* @var $reader PHPExcel_Reader_CSV */
            $reader->setDelimiter($this->delimiter);
            $reader->setEnclosure($this->enclosure);
        }
        $data = array();
        if ($reader->canRead($filepath)) {
            $excel = $reader->load($filepath);
            $data  = $excel->getActiveSheet()->toArray(null, true, false, false);
        } else {
            throw new Exception("Cannot read $filepath");
        }

        $headers = array();

        if ($this->hasHeaderRow) {
            $headers = array_shift($data);
            $headers = array_map('trim', $headers);
            $headers = array_filter($headers);
        }

        $headersCount = count($headers);

        $this->db = Config::inst()->get($this->objectClass, 'db');

        foreach ($data as $row) {
            $row = $this->mergeRowWithHeaders($row, $headers, $headersCount);
            $id  = $this->processRecord($row, $this->columnMap, $results,
                $preview);
        }

        return $results;
    }

    /**
     * @todo Better messages for relation checks and duplicate detection
     * Note that columnMap isn't used.
     *
     * @param array $record
     * @param array $columnMap
     * @param BulkLoader_Result $results
     * @param boolean $preview
     *
     * @return int
     */
    protected function processRecord($record, $columnMap, &$results,
                                     $preview = false, $makeRelations = false)
    {
        $class           = $this->objectClass;
        $this->singleton = singleton($class);

        // find existing object, or create new one
        $existingObj = $this->findExistingObject($record, $columnMap);
        $obj         = ($existingObj) ? $existingObj : new $class();

        // first run: find/create any relations and store them on the object
        // we can't combine runs, as other columns might rely on the relation being present
        if ($makeRelations) {
            $relations = array();
            foreach ($record as $fieldName => $val) {
                // don't bother querying of value is not set
                if ($this->isNullValue($val)) continue;

                // checking for existing relations
                if (isset($this->relationCallbacks[$fieldName])) {
                    // trigger custom search method for finding a relation based on the given value
                    // and write it back to the relation (or create a new object)
                    $relationName = $this->relationCallbacks[$fieldName]['relationname'];
                    if ($this->hasMethod($this->relationCallbacks[$fieldName]['callback'])) {
                        $relationObj = $this->{$this->relationCallbacks[$fieldName]['callback']}($obj,
                            $val, $record);
                    } elseif ($obj->hasMethod($this->relationCallbacks[$fieldName]['callback'])) {
                        $relationObj = $obj->{$this->relationCallbacks[$fieldName]['callback']}($val,
                            $record);
                    }
                    if (!$relationObj || !$relationObj->exists()) {
                        $relationClass = $obj->hasOneComponent($relationName);
                        $relationObj   = new $relationClass();
                        //write if we aren't previewing
                        if (!$preview) $relationObj->write();
                    }
                    $obj->{"{$relationName}ID"} = $relationObj->ID;
                    //write if we are not previewing
                    if (!$preview) {
                        $obj->write();
                        $obj->flushCache(); // avoid relation caching confusion
                    }
                } elseif (strpos($fieldName, '.') !== false) {
                    // we have a relation column with dot notation
                    list($relationName, $columnName) = explode('.', $fieldName);
                    // always gives us an component (either empty or existing)
                    $relationObj                = $obj->getComponent($relationName);
                    if (!$preview) $relationObj->write();
                    $obj->{"{$relationName}ID"} = $relationObj->ID;

                    //write if we are not previewing
                    if (!$preview) {
                        $obj->write();
                        $obj->flushCache(); // avoid relation caching confusion
                    }
                }
            }
        }


        // second run: save data

        $db = $this->db;

        foreach ($record as $fieldName => $val) {
            // break out of the loop if we are previewing
            if ($preview) {
                break;
            }

            // look up the mapping to see if this needs to map to callback
            $mapped = $this->columnMap && isset($this->columnMap[$fieldName]);

            if ($mapped && strpos($this->columnMap[$fieldName], '->') === 0) {
                $funcName = substr($this->columnMap[$fieldName], 2);

                $this->$funcName($obj, $val, $record);
            } else if ($obj->hasMethod("import{$fieldName}")) {
                $obj->{"import{$fieldName}"}($val, $record);
            } else {
                // Basic value mapping if needed

                if (isset($db[$fieldName])) {
                    switch ($db[$fieldName]) {
                        case 'Boolean':
                            if ($val == 'yes') {
                                $val = true;
                            } else if ($val == 'no') {
                                $val = false;
                            }
                    }
                }

                $obj->update(array($fieldName => $val));
            }
        }

        // write record
        $id = ($preview) ? 0 : $obj->write();

        // @todo better message support
        $message = '';

        // save to results
        if ($existingObj) {
            $results->addUpdated($obj, $message);
        } else {
            $results->addCreated($obj, $message);
        }

        $objID = $obj->ID;

        $obj->destroy();

        // memory usage
        unset($existingObj);
        unset($obj);

        return $objID;
    }

    /**
     * Find an existing objects based on one or more uniqueness columns
     * specified via {@link self::$duplicateChecks}.
     *
     * @param array $record CSV data column
     *
     * @return mixed
     */
    public function findExistingObject($record)
    {
        $SNG_objectClass = $this->singleton;

        // checking for existing records (only if not already found)
        foreach ($this->duplicateChecks as $fieldName => $duplicateCheck) {
            if (is_string($duplicateCheck)) {

                // Skip current duplicate check if field value is empty
                if (empty($record[$duplicateCheck])) continue;

                $existingRecord = DataObject::get($this->objectClass)
                    ->filter($duplicateCheck, $record[$duplicateCheck])
                    ->first();

                if ($existingRecord) return $existingRecord;
            } elseif (is_array($duplicateCheck) && isset($duplicateCheck['callback'])) {
                if ($this->hasMethod($duplicateCheck['callback'])) {
                    $existingRecord = $this->{$duplicateCheck['callback']}($record[$fieldName],
                        $record);
                } elseif ($SNG_objectClass->hasMethod($duplicateCheck['callback'])) {
                    $existingRecord = $SNG_objectClass->{$duplicateCheck['callback']}($record[$fieldName],
                        $record);
                } else {
                    user_error("CsvBulkLoader::processRecord():"
                        ." {$duplicateCheck['callback']} not found on importer or object class.",
                        E_USER_ERROR);
                }

                if ($existingRecord) {
                    return $existingRecord;
                }
            } else {
                user_error('CsvBulkLoader::processRecord(): Wrong format for $duplicateChecks',
                    E_USER_ERROR);
            }
        }

        return false;
    }

    /**
     * Determine whether any loaded files should be parsed with a
     * header-row (otherwise we rely on {@link self::$columnMap}.
     *
     * @return boolean
     */
    public function hasHeaderRow()
    {
        return ($this->hasHeaderRow || isset($this->columnMap));
    }

    public function setFileType($fileType)
    {
        $this->fileType = $fileType;
    }

    public function getFileType()
    {
        return $this->fileType;
    }
}