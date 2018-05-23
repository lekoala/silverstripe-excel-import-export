<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\ORM\DataObject;
use SilverStripe\Security\Group;
use LeKoala\ExcelImportExport\ExcelBulkLoader;
use SilverStripe\Security\Member;

/**
 * Imports member records, and checks/updates duplicates based on their
 * 'Email' property.
 */
class ExcelMemberBulkLoader extends ExcelBulkLoader
{
    /**
     * Array of {@link Group} records. Import into a specific group.
     * Is overruled by any "Groups" columns in the import.
     *
     * @var array
     */
    protected $groups = array();

    public $duplicateChecks = array(
        'Email' => 'Email',
    );

    public function __construct($objectClass = null)
    {
        if (!$objectClass) {
            $objectClass = Member::class;
        }

        parent::__construct($objectClass);
    }

    public function processRecord(
        $record,
        $columnMap,
        &$results,
        $preview = false
    ) {
        $objID = parent::processRecord($record, $columnMap, $results, $preview);

        $_cache_groupByCode = array();

        // Add to predefined groups
        /** @var Member $member */
        $member = DataObject::get_by_id($this->objectClass, $objID);
        foreach ($this->groups as $group) {
            // TODO This isnt the most memory effective way to add members to a group
            $member->Groups()->add($group);
        }

        // Add to groups defined in CSV
        if (isset($record['Groups']) && $record['Groups']) {
            $groupCodes = explode(',', $record['Groups']);
            foreach ($groupCodes as $groupCode) {
                $groupCode = Convert::raw2url($groupCode);
                if (!isset($_cache_groupByCode[$groupCode])) {
                    $group = Group::get()->filter('Code', $groupCode)->first();
                    if (!$group) {
                        $group = new Group();
                        $group->Code = $groupCode;
                        $group->Title = $groupCode;
                        $group->write();
                    }
                    $member->Groups()->add($group);
                    $_cache_groupByCode[$groupCode] = $group;
                }
            }
        }

        $member->destroy();
        unset($member);

        return $objID;
    }

    /**
     * @param array $groups
     */
    public function setGroups($groups)
    {
        $this->groups = $groups;
    }

    /**
     * @return array
     */
    public function getGroups()
    {
        return $this->groups;
    }
}
