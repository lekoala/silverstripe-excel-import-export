<?php

/**
 * Import members from excel
 *
 * @author Koala
 */
class ExcelMemberImportForm extends MemberImportForm
{

    public function __construct($controller, $name, $fields = null,
                                $actions = null, $validator = null)
    {
        if (!$fields) {
            $helpHtml = _t(
                'ExcelMemberImportForm.Help1',
                '<p><a href="{link}">Download sample file</a></p>',
                array('link' => $controller->Link('downloadsample/Member'))
            );
            $helpHtml .= _t(
                'ExcelMemberImportForm.Help2',
                '<ul>'
                .'<li>Existing users are matched by their unique <em>Email</em> property, and updated with any new values from '
                .'the imported file.</li>'
                .'<li>Groups can be assigned by the <em>Groups</em> column. Groups are identified by their <em>Code</em> property, '
                .'multiple groups can be separated by comma. Existing group memberships are not cleared.</li>'
                .'</ul>'
            );

            $importer   = new MemberCsvBulkLoader();
            $importSpec = $importer->getImportSpec();
            $helpHtml   = sprintf($helpHtml,
                implode(', ', array_keys($importSpec['fields'])));

            $fields    = new FieldList(
                new LiteralField('Help', $helpHtml),
                $fileField = new FileField(
                'File',
                _t(
                    'ExcelMemberImportForm.FileFieldLabel',
                    'File <small>(Allowed extensions: *.csv, *.xls, *.xlsx, *.ods, *.txt)</small>'
                )
                )
            );
            $fileField->getValidator()->setAllowedExtensions(ExcelImportExport::getValidExtensions());
        }

        if (!$actions) {
            $action  = new FormAction('doImport',
                _t('ExcelMemberImportForm.BtnImport', 'Import from file'));
            $action->addExtraClass('ss-ui-button');
            $actions = new FieldList($action);
        }

        if (!$validator) $validator = new RequiredFields('File');

        parent::__construct($controller, $name, $fields, $actions, $validator);

        $this->addExtraClass('cms');
        $this->addExtraClass('import-form');
    }

    public function doImport($data, $form)
    {
        $loader = new ExcelMemberBulkLoader();

        // optionally set group relation
        if ($this->group) $loader->setGroups(array($this->group));

        // load file
        $result = $loader->load($data['File']);

        // result message
        $msgArr   = array();
        if ($result->CreatedCount())
                $msgArr[] = _t(
                'ExcelMemberImportForm.ResultCreated',
                'Created {count} members',
                array('count' => $result->CreatedCount())
            );
        if ($result->UpdatedCount())
                $msgArr[] = _t(
                'ExcelMemberImportForm.ResultUpdated',
                'Updated {count} members',
                array('count' => $result->UpdatedCount())
            );
        if ($result->DeletedCount())
                $msgArr[] = _t(
                'ExcelMemberImportForm.ResultDeleted', 'Deleted %d members',
                array('count' => $result->DeletedCount())
            );
        $msg      = ($msgArr) ? implode(',', $msgArr) : _t('ExcelMemberImportForm.ResultNone',
                'No changes');

        $this->sessionMessage($msg, 'good');

        return $this->getController()->redirectBack();
    }
}