SilverStripe Excel Import Export module
==================
Replace all csv import/export functionnalities by Excel. Excel support is provided
by PHPExcel.

These changes apply automatically to SecurityAdmin and ModelAdmin through extension.

Side functionnalities:

- It will enable bulk manager if class exists on the GridField
- Import specs are replaced by a sample file that is easier to use for clients

Configure exported fields
==================

All fields are exported by default (not just summaryFields that are useless by themselves)

If you want to restrict the fields, you can either:

- Implement a exportedFields method on your model that should return an array of fields
- Define a exported_fields config field on your model that will restrict the list to these fields
- Define a unexported_fields config field on your model that will blacklist these fields from being exported

Other modules out there
==================

[Silverstripe Excel Export](https://github.com/firebrandhq/excel-export)
Focus more on the export side of things and support RestfulServer module

[Silverstripe PHPExcel](https://github.com/axyr/silverstripe-phpexcel)
If you just need the export button

[SilverStripe Import/Export](https://github.com/burnbright/silverstripe-importexport)
With a very nice user column mapping

Compatibility
==================
Tested with 3.x

Maintainer
==================
LeKoala - thomas@lekoala.be