<?php

namespace LeKoala\ExcelImportExport\Test\Mocks;

use SilverStripe\Dev\TestOnly;
use SilverStripe\Security\Member;
use Exception;

class TestExcelMember extends Member implements TestOnly
{
    public function onBeforeWrite()
    {
        parent::onBeforeWrite();

        // For older ss versions that do not validate emails
        if ($this->Email && !filter_var($this->Email, FILTER_VALIDATE_EMAIL)) {
            throw new Exception("Email is not valid");
        }
    }
}
