<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\Core\Convert;
use SilverStripe\ORM\DataObject;
use SilverStripe\Security\Group;
use SilverStripe\Security\Member;

/**
 * Imports member records, and checks/updates duplicates based on their
 * 'Email' property.
 */
class ExcelMemberBulkLoader extends ExcelBulkLoader
{
    /**
     * Import into a specific group.
     * Is overruled by any "Groups" columns in the import.
     *
     * @var array<Group>
     */
    protected $groups = [];

    /**
     * @var array<string,string>
     */
    public $duplicateChecks = array(
        'Email' => 'Email',
    );

    /**
     * @param class-string $objectClass
     */
    public function __construct($objectClass = null)
    {
        if (!$objectClass) {
            $objectClass = Member::class;
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
        $objID = parent::processRecord($record, $columnMap, $results, $preview);

        $_cache_groupByCode = [];

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
     * @param array<Group> $groups
     * @return void
     */
    public function setGroups($groups)
    {
        $this->groups = $groups;
    }

    /**
     * @return array<Group>
     */
    public function getGroups()
    {
        return $this->groups;
    }
}
