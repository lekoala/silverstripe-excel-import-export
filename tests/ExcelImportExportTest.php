<?php

namespace LeKoala\ExcelImportExport\Test;

use SilverStripe\Dev\SapphireTest;
use LeKoala\ExcelImportExport\ExcelImportExport;
use SilverStripe\Security\Member;

/**
 * Tests for ExcelImportExport
 */
class ExcelImportExportTest extends SapphireTest
{
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
}
