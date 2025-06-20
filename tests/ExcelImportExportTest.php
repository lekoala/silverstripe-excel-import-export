<?php

namespace LeKoala\ExcelImportExport\Test;

use Exception;
use LeKoala\ExcelImportExport\ExcelGridFieldExportButton;
use SilverStripe\Security\Group;
use SilverStripe\Security\Member;
use SilverStripe\Dev\SapphireTest;
use LeKoala\ExcelImportExport\ExcelImportExport;
use LeKoala\ExcelImportExport\ExcelMemberBulkLoader;
use LeKoala\ExcelImportExport\Test\Mocks\TestExcelMember;

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
    protected static $extra_dataobjects = [
        TestExcelMember::class,
    ];

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
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $membersCount = $firstGroup->Members()->count();

        // ; separator is properly detected thanks to auto separator feature
        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members.csv');

        $this->assertEquals(1, $result->CreatedCount());

        $newCount = Member::get()->count();
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
        $firstGroup = Group::get()->filter('Code', 'Administrators')->first();
        $newMembersCount = $firstGroup->Members()->count();

        $this->assertEquals($count + 2, $newCount);
        $this->assertEquals($membersCount + 2, $newMembersCount, "Groups are not updated");

        // Loading again does nothing new
        $result = $loader->load(__DIR__ . '/data/members.xlsx');

        $this->assertEquals(0, $result->CreatedCount());
        $this->assertEquals(1, $result->UpdatedCount());

        $newCount = Member::get()->count();
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

    public function testImportMultipleClasses(): void
    {
        $beforeCount = Member::get()->count();

        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members-class.csv');

        $totalCount = Member::get()->count();
        $this->assertEquals(2, $result->CreatedCount());
        $this->assertEquals($beforeCount, $totalCount - 2);

        // Is a subclass AND a base class
        $member = TestExcelMember::get()->filter('Email', 'excel@silverstripe.org')->first();
        $this->assertNotEmpty($member);
        $member = Member::get()->filter('Email', 'excel@silverstripe.org')->first();
        $this->assertNotEmpty($member);

        // Not a subclass
        $member = TestExcelMember::get()->filter('Email', 'regular@silverstripe.org')->first();
        $this->assertEmpty($member);

        // Checking duplicate still works even with custom class
        $loader = new ExcelMemberBulkLoader();
        $result = $loader->load(__DIR__ . '/data/members-class.csv');

        $this->assertEquals(0, $result->CreatedCount());
        $this->assertEquals(2, $result->UpdatedCount());

        $newTotalCount = Member::get()->count();
        $this->assertEquals($newTotalCount, $totalCount);
    }

    public function testImportMigrateClasses(): void
    {
        // For members, duplicate is done on emails
        $duplicateEmail = 'migrateme@silverstripe.org';

        $baseMember = new Member();
        $baseMember->Email = $duplicateEmail;
        $baseMember->write();

        $loader = new ExcelMemberBulkLoader();
        $result = $loader->processData([
            [
                'ID' => $baseMember->ID,
                'ClassName' => TestExcelMember::class,
                'Email' => $duplicateEmail,
            ]
        ]);

        $this->assertEquals(0, $result->CreatedCount());
        $this->assertEquals(1, $result->UpdatedCount());

        // Check that our existing member has the new class
        $member = TestExcelMember::get()->filter('Email', $duplicateEmail)->first();
        $this->assertNotEmpty($member);
        $this->assertEquals($baseMember->ID, $member->ID);
    }

    public function testTransaction()
    {
        $beforeCount = Member::get()->count();

        $loader = new ExcelMemberBulkLoader();
        $loader->setUseTransaction(true);

        try {
            $result = $loader->load(__DIR__ . '/data/members-invalid.csv');
        } catch (Exception $e) {
            // expected to fail
        }

        // First row was not created because it was rolled back
        $member = Member::get()->filter('Email', 'valid@silverstripe.org')->first();
        $this->assertEmpty($member);

        $newTotalCount = Member::get()->count();
        $this->assertEquals($newTotalCount, $beforeCount);


        $loader = new ExcelMemberBulkLoader();
        $loader->setUseTransaction(false);

        try {
            $result = $loader->load(__DIR__ . '/data/members-invalid.csv');
        } catch (Exception $e) {
            // expected to fail
        }

        // First row was created because it was not rolled back
        $member = Member::get()->filter('Email', 'valid@silverstripe.org')->first();
        $this->assertNotEmpty($member);

        $newTotalCount = Member::get()->count();
        $this->assertEquals($newTotalCount - 1, $beforeCount);
    }
}
