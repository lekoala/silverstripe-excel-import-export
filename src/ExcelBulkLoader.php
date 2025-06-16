<?php

namespace LeKoala\ExcelImportExport;

use Exception;
use LeKoala\SpreadCompat\SpreadCompat;
use SilverStripe\Dev\BulkLoader;
use SilverStripe\ORM\DataObject;
use SilverStripe\Core\Environment;
use SilverStripe\Dev\BulkLoader_Result;
use SilverStripe\Control\HTTPResponse_Exception;
use SilverStripe\Core\ClassInfo;
use SilverStripe\ORM\DB;

/**
 * @author Koala
 */
class ExcelBulkLoader extends BulkLoader
{
    private bool $useTransaction = false;

    /**
     * Delimiter character
     * We use auto detection for csv because we can't ask the user what he is using
     *
     * @var string
     */
    public $delimiter = 'auto';

    /**
     * Enclosure character (Default: doublequote)
     *
     * @var string
     */
    public $enclosure = '"';

    /**
     * Identifies if the file has a header row.
     *
     * @var boolean
     */
    public $hasHeaderRow = true;

    /**
     * @var array<string,string>
     */
    public $duplicateChecks = [
        'ID' => 'ID',
    ];

    /**
     * The uploaded file infos
     * @var array<mixed>
     */
    protected $uploadFile = null;

    /**
     *
     * @var DataObject
     */
    protected $singleton = null;

    /**
     * @var array<mixed>
     */
    protected $db = [];

    /**
     * Type of file if not able to determine through uploaded file
     *
     * @var string
     */
    protected $fileType = 'xlsx';

    /**
     * @return BulkLoader_Result
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
            $filepath = $filepath['tmp_name'];
        }
        if (is_string($filepath)) {
            $ext = pathinfo($filepath, PATHINFO_EXTENSION);
            if ($ext == 'csv' || $ext == 'xlsx') {
                $this->fileType = $ext;
            }
        }

        // upload is resource intensive
        Environment::increaseTimeLimitTo(3600);
        Environment::increaseMemoryLimitTo('512M');

        if ($this->useTransaction) {
            DB::get_conn()->transactionStart();
        }

        try {
            //get all instances of the to be imported data object
            if ($this->deleteExistingRecords) {
                if ($this->getCheckPermissions()) {
                    // We need to check each record, in case there's some fancy conditional logic in the canDelete method.
                    // If we can't delete even a single record, we should bail because otherwise the result would not be
                    // what the user expects.
                    /** @var DataObject $record */
                    foreach (DataObject::get($this->objectClass) as $record) {
                        if (!$record->canDelete()) {
                            $type = $record->i18n_singular_name();
                            throw new HTTPResponse_Exception(
                                _t(__CLASS__ . '.CANNOT_DELETE', "Not allowed to delete '{type}' records", ["type" => $type]),
                                403
                            );
                        }
                    }
                }
                DataObject::get($this->objectClass)->removeAll();
            }

            $result = $this->processAll($filepath);

            if ($this->useTransaction) {
                DB::get_conn()->transactionEnd();
            }
        } catch (Exception $e) {
            if ($this->useTransaction) {
                DB::get_conn()->transactionRollback();
            }
            $code = $e->getCode() ?: 500;
            throw new HTTPResponse_Exception($e->getMessage(), $code);
        }
        return $result;
    }

    /**
     * @return string
     */
    protected function getUploadFileExtension()
    {
        if ($this->uploadFile) {
            return pathinfo($this->uploadFile['name'], PATHINFO_EXTENSION);
        }
        return $this->fileType;
    }

    /**
     * Merge a row with its headers
     *
     * @param array $row
     * @param array $headers
     * @param int $headersCount (optional) Limit to a specifc number of headers
     * @return array
     */
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
        $this->extend('onBeforeProcessAll', $filepath, $preview);

        $results = new BulkLoader_Result();
        $ext = $this->getUploadFileExtension();

        if (!is_readable($filepath)) {
            throw new Exception("Cannot read $filepath");
        }

        $opts = [
            'separator' => $this->delimiter,
            'enclosure' => $this->enclosure,
            'extension' => $ext,
        ];
        if ($this->hasHeaderRow) {
            $opts['assoc'] = true;
        }

        $data = SpreadCompat::read($filepath, ...$opts);

        $objectClass = $this->objectClass;
        $objectConfig = $objectClass::config();
        $this->db = $objectConfig->db;
        $this->singleton = singleton($objectClass);

        foreach ($data as $row) {
            $this->processRecord(
                $row,
                $this->columnMap,
                $results,
                $preview
            );
        }

        $this->extend('onAfterProcessAll', $result, $preview);

        return $results;
    }

    /**
     *
     * @param array $record
     * @param array $columnMap
     * @param BulkLoader_Result $results
     * @param boolean $preview
     *
     * @return int
     */
    protected function processRecord(
        $record,
        $columnMap,
        &$results,
        $preview = false,
        $makeRelations = false
    ) {
        // find existing object, or create new one
        $existingObj = $this->findExistingObject($record, $columnMap);
        $alreadyExists = (bool) $existingObj;

        // If we can't edit the existing object, bail early.
        if ($this->getCheckPermissions() && !$preview && $alreadyExists && !$existingObj->canEdit()) {
            $type = $existingObj->i18n_singular_name();
            throw new HTTPResponse_Exception(
                _t(BulkLoader::class . '.CANNOT_EDIT', "Not allowed to edit '{type}' records", ["type" => $type]),
                403
            );
        }

        $class = $record['ClassName'] ?? $this->objectClass;
        $obj = $existingObj ? $existingObj : new $class();

        // If we can't create a new record, bail out early.
        if ($this->getCheckPermissions() && !$preview && !$alreadyExists && !$obj->canCreate()) {
            $type = $obj->i18n_singular_name();
            throw new HTTPResponse_Exception(
                _t(BulkLoader::class . '.CANNOT_CREATE', "Not allowed to create '{type}' records", ["type" => $type]),
                403
            );
        }

        // first run: find/create any relations and store them on the object
        // we can't combine runs, as other columns might rely on the relation being present
        if ($makeRelations) {
            foreach ($record as $fieldName => $val) {
                // don't bother querying of value is not set
                if ($this->isNullValue($val)) {
                    continue;
                }

                // checking for existing relations
                if (isset($this->relationCallbacks[$fieldName])) {
                    // trigger custom search method for finding a relation based on the given value
                    // and write it back to the relation (or create a new object)
                    $relationName = $this->relationCallbacks[$fieldName]['relationname'];
                    if ($this->hasMethod($this->relationCallbacks[$fieldName]['callback'])) {
                        $relationObj = $this->{$this->relationCallbacks[$fieldName]['callback']}(
                            $obj,
                            $val,
                            $record
                        );
                    } elseif ($obj->hasMethod($this->relationCallbacks[$fieldName]['callback'])) {
                        $relationObj = $obj->{$this->relationCallbacks[$fieldName]['callback']}(
                            $val,
                            $record
                        );
                    }
                    if (!$relationObj || !$relationObj->exists()) {
                        $relationClass = $obj->hasOneComponent($relationName);
                        $relationObj = new $relationClass();
                        //write if we aren't previewing
                        if (!$preview) {
                            $relationObj->write();
                        }
                    }
                    $obj->{"{$relationName}ID"} = $relationObj->ID;
                    //write if we are not previewing
                    if (!$preview) {
                        $obj->write();
                        $obj->flushCache(); // avoid relation caching confusion
                    }
                } elseif (strpos($fieldName, '.') !== false) {
                    // we have a relation column with dot notation
                    list($relationName) = explode('.', $fieldName);
                    // always gives us an component (either empty or existing)
                    $relationObj = $obj->getComponent($relationName);
                    if (!$preview) {
                        $relationObj->write();
                    }
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

            // Do not update ID if any exist
            if ($fieldName == 'ID' && $obj->ID) {
                continue;
            }

            // look up the mapping to see if this needs to map to callback
            $mapping = ($columnMap && isset($columnMap[$fieldName])) ? $columnMap[$fieldName]
                : null;

            // Mapping that starts with -> map to a method
            if ($mapping && strpos($mapping, '->') === 0) {
                $funcName = substr($mapping, 2);

                $this->$funcName($obj, $val, $record);
            } elseif ($obj->hasMethod("import{$fieldName}")) {
                // Try to call import_myFieldName
                $obj->{"import{$fieldName}"}($val, $record);
            } else {
                // Map column to field
                $usedName = $mapping ? $mapping : $fieldName;

                // Basic value mapping based on datatype if needed
                if (isset($db[$usedName])) {
                    switch ($db[$usedName]) {
                        case 'Boolean':
                            if ((string) $val == 'yes') {
                                $val = true;
                            } elseif ((string) $val == 'no') {
                                $val = false;
                            }
                    }
                }

                $obj->update(array($usedName => $val));
            }
        }

        // write record
        if (!$preview) {
            $obj->write();
        }

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
     * @param array $columnMap
     *
     * @return DataObject|false
     */
    public function findExistingObject($record, $columnMap = [])
    {
        $SNG_objectClass = $this->singleton;

        // checking for existing records (only if not already found)
        foreach ($this->duplicateChecks as $fieldName => $duplicateCheck) {
            $existingRecord = null;
            if (is_string($duplicateCheck)) {
                // Skip current duplicate check if field value is empty
                if (empty($record[$duplicateCheck])) {
                    continue;
                }

                $dbFieldValue = $record[$duplicateCheck];

                // Even if $record['ClassName'] is a subclass, this will work
                $existingRecord = DataObject::get($this->objectClass)
                    ->filter($duplicateCheck, $dbFieldValue)
                    ->first();

                if ($existingRecord) {
                    return $existingRecord;
                }
            } elseif (is_array($duplicateCheck) && isset($duplicateCheck['callback'])) {
                if ($this->hasMethod($duplicateCheck['callback'])) {
                    $existingRecord = $this->{$duplicateCheck['callback']}(
                        $record[$fieldName],
                        $record
                    );
                } elseif ($SNG_objectClass->hasMethod($duplicateCheck['callback'])) {
                    $existingRecord = $SNG_objectClass->{$duplicateCheck['callback']}(
                        $record[$fieldName],
                        $record
                    );
                } else {
                    throw new \RuntimeException(
                        "ExcelBulkLoader::processRecord():"
                            . " {$duplicateCheck['callback']} not found on importer or object class."
                    );
                }

                if ($existingRecord) {
                    return $existingRecord;
                }
            } else {
                throw new \InvalidArgumentException(
                    'ExcelBulkLoader::processRecord(): Wrong format for $duplicateChecks'
                );
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

    /**
     * Set file type as import
     *
     * @param string $fileType
     * @return void
     */
    public function setFileType($fileType)
    {
        $this->fileType = $fileType;
    }

    /**
     * Get file type (default is xlsx)
     *
     * @return string
     */
    public function getFileType()
    {
        return $this->fileType;
    }

    /**
     * If true, will wrap everything in a transaction
     */
    public function getUseTransaction(): bool
    {
        return $this->useTransaction;
    }

    /**
     * Determines if everything will be wrapped in a transaction
     */
    public function setCheckPermissions(bool $value): ExcelBulkLoader
    {
        $this->useTransaction = $value;
        return $this;
    }
}
