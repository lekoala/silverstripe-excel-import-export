<?php

namespace LeKoala\ExcelImportExport\Test;

use LeKoala\ExcelImportExport\ExcelGridFieldExportButton;
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

    public function testGetAllFields(): void
    {
        $fields = ExcelImportExport::allFieldsForClass(Member::class);
        $this->assertNotEmpty($fields);
    }

    public function testExportedFields(): void
    {
        $fields = ExcelImportExport::exportFieldsForClass(Member::class);
        $this->assertNotEmpty($fields);
        $this->assertNotContains('Password', $fields);
    }

    public function testCanImportMembers(): void
    {
        $count = Member::get()->count();
        /** @var Group $firstGroup */
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $membersCount = $firstGroup->Members()->count();

        // ; separator is properly detected thanks to auto separator feature
        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members.csv');

        $this->assertEquals(1, $result->CreatedCount());

        $newCount = Member::get()->count();
        /** @var Group $firstGroup */
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $newMembersCount = $firstGroup->Members()->count();

        $this->assertEquals($count + 1, $newCount);
        $this->assertEquals($membersCount + 1, $newMembersCount, "Groups are not updated");

        // format is handled according to file extension
        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members.xlsx');

        $this->assertEquals(1, $result->CreatedCount());
        $this->assertEquals(0, $result->UpdatedCount());

        $newCount = Member::get()->count();
        /** @var Group $firstGroup */
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $newMembersCount = $firstGroup->Members()->count();

        $this->assertEquals($count + 2, $newCount);
        $this->assertEquals($membersCount + 2, $newMembersCount, "Groups are not updated");

        // Loading again does nothing new
        $result = $loader->load(__DIR__ . '/data/members.xlsx');

        $this->assertEquals(0, $result->CreatedCount());
        $this->assertEquals(1, $result->UpdatedCount());

        $newCount = Member::get()->count();
        /** @var Group $firstGroup */
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $newMembersCount = $firstGroup->Members()->count();

        $this->assertEquals($count + 2, $newCount);
        $this->assertEquals($membersCount + 2, $newMembersCount, "Groups are not updated");
    }

    public function testSanitize(): void
    {
        $dangerousInput = '=1+2";=1+2';

        $actual = ExcelGridFieldExportButton::sanitizeValue($dangerousInput);
        $expected = "\t" . $dangerousInput;
        $this->assertEquals($expected, $actual);
    }
}
