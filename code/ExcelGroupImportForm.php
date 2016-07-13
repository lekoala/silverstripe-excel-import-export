<?php

/**
 * Improved group import form
 *
 * @author Koala
 */
class ExcelGroupImportForm extends GroupImportForm
{

    public function __construct($controller, $name, $fields = null,
                                $actions = null, $validator = null)
    {
        if (!$fields) {
            $helpHtml = _t(
                'ExcelGroupImportForm.Help1',
                '<a href="{link}">Download sample file</a></small></p>',
                array('link' => $controller->Link('downloadsample/Group'))
            );
            $helpHtml .= _t(
                'ExcelGroupImportForm.Help2',
                '<ul>'
                .'<li>Existing groups are matched by their unique <em>Code</em> value, and updated with any new values from the '
                .'imported file</li>'
                .'<li>Group hierarchies can be created by using a <em>ParentCode</em> column.</li>'
                .'<li>Permission codes can be assigned by the <em>PermissionCode</em> column. Existing permission codes are not '
                .'cleared.</li>'
                .'</ul>'
                .'</div>'
            );

            $importer   = new GroupCsvBulkLoader();
            $importSpec = $importer->getImportSpec();
            $helpHtml   = sprintf($helpHtml,
                implode(', ', array_keys($importSpec['fields'])));

            $fields    = new FieldList(
                new LiteralField('Help', $helpHtml),
                $fileField = new FileField(
                'File',
                _t(
                    'ExcelGroupImportForm.FileFieldLabel',
                    'File <small>(Allowed extensions: *.csv, *.xls, *.xlsx, *.ods, *.txt)</small>'
                )
                )
            );
            $fileField->getValidator()->setAllowedExtensions(array('csv'));
        }

        if (!$actions) {
            $action  = new FormAction('doImport',
                _t('ExcelGroupImportForm.BtnImport', 'Import from file'));
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
        $loader = new ExcelGroupBulkLoader();

        // load file
        $result = $loader->load($data['File']['tmp_name']);

        // result message
        $msgArr   = array();
        if ($result->CreatedCount())
                $msgArr[] = _t(
                'ExcelGroupImportForm.ResultCreated', 'Created {count} groups',
                array('count' => $result->CreatedCount())
            );
        if ($result->UpdatedCount())
                $msgArr[] = _t(
                'ExcelGroupImportForm.ResultUpdated', 'Updated %d groups',
                array('count' => $result->UpdatedCount())
            );
        if ($result->DeletedCount())
                $msgArr[] = _t(
                'ExcelGroupImportForm.ResultDeleted', 'Deleted %d groups',
                array('count' => $result->DeletedCount())
            );
        $msg      = ($msgArr) ? implode(',', $msgArr) : _t('ExcelGroupImportForm.ResultNone',
                'No changes');

        $this->sessionMessage($msg, 'good');

        return $this->getController()->redirectBack();
    }
}