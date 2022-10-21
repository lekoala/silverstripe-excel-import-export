# SilverStripe Excel Import Export module

![Build Status](https://github.com/lekoala/silverstripe-debugbar/actions/workflows/ci.yml/badge.svg)
[![scrutinizer](https://scrutinizer-ci.com/g/lekoala/silverstripe-debugbar/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/lekoala/silverstripe-debugbar/)
[![Code coverage](https://codecov.io/gh/lekoala/silverstripe-debugbar/branch/master/graph/badge.svg)](https://codecov.io/gh/lekoala/silverstripe-debugbar)

## Intro

Replace all csv import/export functionalities by Excel.
Excel support is provided by PHPSpreadsheet.

These changes apply automatically to SecurityAdmin and ModelAdmin through extension.

To make import easier, import specs are replaced by a sample file that is ready to use for the user.

## Configure exported fields

All fields are exported by default (not just summary fields that are useless by themselves)

If you want to restrict the fields, you can either:

- Implement a exportedFields method on your model that should return an array of fields
- Define a exported_fields config field on your model that will restrict the list to these fields
- Define a unexported_fields config field on your model that will blacklist these fields from being exported

## Migrate your old code

Previous version was using PHPExcel. This version use PHPSpreadhsheet. Any code using PHPExcel should
be migrated.
You can find the following guide helpful: https://phpspreadsheet.readthedocs.io/en/develop/topics/migration-from-PHPExcel/

## Todo

- More tests and refactoring
- Handle large exports with ease

## Compatibility

Tested with ^4.6 and up

## Maintainer

LeKoala - thomas@lekoala.be
