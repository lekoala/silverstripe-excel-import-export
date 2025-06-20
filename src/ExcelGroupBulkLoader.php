<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\ORM\DataObject;
use SilverStripe\Security\Group;
use SilverStripe\Security\Permission;

/**
 * The same as GroupBulkLoader but with ExcelBulkLoader as base class
 *
 * @author Koala
 */
class ExcelGroupBulkLoader extends ExcelBulkLoader
{
    /**
     * @var array<string,string>
     */
    public $duplicateChecks = [
        'Code' => 'Code',
    ];

    /**
     * @param class-string $objectClass
     */
    public function __construct($objectClass = null)
    {
        if (!$objectClass) {
            $objectClass = Group::class;
        }
        parent::__construct($objectClass);
    }

    /**
     * @param array<string,mixed> $record
     * @param array<string,string> $columnMap
     * @param mixed $results
     * @param boolean $preview
     * @param boolean $makeRelations
     * @return int
     */
    protected function processRecord(
        $record,
        $columnMap,
        &$results,
        $preview = false,
        $makeRelations = false
    ) {
        // We match by 'Code', the ID property is confusing the importer
        if (isset($record['ID'])) {
            unset($record['ID']);
        }

        $objID = parent::processRecord($record, $columnMap, $results, $preview);

        $group = Group::get_by_id($objID);
        // set group hierarchies - we need to do this after all records
        // are imported to avoid missing "early" references to parents
        // which are imported later on in the CSV file.
        if (isset($record['ParentCode']) && $record['ParentCode']) {
            $parentGroup = DataObject::get_one(Group::class, array(
                '"Group"."Code"' => $record['ParentCode']
            ));
            if ($parentGroup) {
                $group->ParentID = $parentGroup->ID;
                $group->write();
            }
        }

        // set permission codes - these are all additive, meaning
        // existing permissions arent cleared.
        if (isset($record['PermissionCodes']) && $record['PermissionCodes']) {
            foreach (explode(',', $record['PermissionCodes']) as $code) {
                $p = DataObject::get_one(Permission::class, array(
                    '"Permission"."Code"' => $code,
                    '"Permission"."GroupID"' => $group->ID
                ));
                if (!$p) {
                    $p = new Permission(array('Code' => $code));
                    $p->write();
                }
                $group->Permissions()->add($p);
            }
        }

        return $objID;
    }
}
