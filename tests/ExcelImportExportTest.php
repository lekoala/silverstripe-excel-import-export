<?php

namespace LeKoala\ExcelImportExport\Test;

use SilverStripe\Security\Group;
use SilverStripe\Security\Member;
use SilverStripe\Dev\SapphireTest;
use LeKoala\ExcelImportExport\ExcelImportExport;
use LeKoala\ExcelImportExport\ExcelMemberBulkLoader;

/**
 * Tests for ExcelImportExport
 */
class ExcelImportExportTest extends SapphireTest
{
    /**
     * Defines the fixture file to use for this test class
     * @var string
     */
    protected static $fixture_file = 'ExcelImportExportTest.yml';

    public function testGetAllFields()
    {
        $fields = ExcelImportExport::allFieldsForClass(Member::class);
        $this->assertNotEmpty($fields);
    }

    public function testExportedFields()
    {
        $fields = ExcelImportExport::exportFieldsForClass(Member::class);
        $this->assertNotEmpty($fields);
        $this->assertNotContains('Password', $fields);
    }

    public function testCanImportMembers()
    {
        $count = Member::get()->count();
        $membersCount = Group::get()->filter('Code', 'Administrators')->first()->Members()->count();

        // ; separator is properly detected thanks to auto separator feature
        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members.csv');

        $this->assertEquals(1, $result->CreatedCount());

        $newCount = Member::get()->count();
        $newMembersCount = Group::get()->filter('Code', 'Administrators')->first()->Members()->count();

        $this->assertEquals($count + 1, $newCount);
        $this->assertEquals($membersCount + 1, $newMembersCount, "Groups are not updated");
    }
}
