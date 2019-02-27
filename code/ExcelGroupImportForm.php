<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\Forms\FieldList;
use SilverStripe\Forms\FileField;
use SilverStripe\Forms\FormAction;
use SilverStripe\Forms\LiteralField;
use SilverStripe\Admin\GroupImportForm;

/**
 * Improved group import form
 *
 * @author Koala
 */
class ExcelGroupImportForm extends GroupImportForm
{
    public function __construct(
        $controller,
        $name,
        $fields = null,
        $actions = null,
        $validator = null
    ) {
        $downloadSampleLink = $controller->Link('downloadsample/Group');
        $downloadSample = '<a href="'.$downloadSampleLink.'" class="no-ajax" target="_blank">' . _t(
            'ExcelImportExport.DownloadSample',
            'Download sample file'
        ) . '</a>';

        $helpHtml = '';
        $helpHtml .= '<ul>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.GroupImportHelp1',
            'For supported columns, please check the sample file.'
        ) . ' ' . $downloadSample . '</li>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.GroupImportHelp2',
            'Existing groups are matched by their unique <em>Code</em> value, and updated with any new values from the imported file.'
        ) . '</li>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.GroupImportHelp3',
            'Group hierarchies can be created by using a <em>ParentCode</em> column.'
        ) . '</li>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.GroupImportHelp4',
            'Permission codes can be assigned by the <em>PermissionCode</em> column. Existing permission codes are not cleared.'
        ) . '</li>';
        $helpHtml .= '</ul>';

        $file = new FileField('File');
        $csvDescription = ExcelImportExport::getValidExtensionsText();
        $file->setDescription($csvDescription);
        $file->getValidator()->setAllowedExtensions(ExcelImportExport::getValidExtensions());

        $fields    = new FieldList(
            new LiteralField('Help', $helpHtml),
            $file
        );

        $action  = new FormAction(
            'doImport',
            _t('ExcelImportExport.BtnImport', 'Import from file')
        );
        $action->addExtraClass('btn-primary');
        $actions = new FieldList($action);

        parent::__construct($controller, $name, $fields, $actions, $validator);
    }

    public function doImport($data, $form)
    {
        $loader = new ExcelGroupBulkLoader();

        // load file
        $result = $loader->load($data['File']['tmp_name']);

        // result message
        $msgArr   = array();
        if ($result->CreatedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultCreated',
                'Created {count} groups',
                array('count' => $result->CreatedCount())
            );
        }
        if ($result->UpdatedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultUpdated',
                'Updated %d groups',
                array('count' => $result->UpdatedCount())
            );
        }
        if ($result->DeletedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultDeleted',
                'Deleted %d groups',
                array('count' => $result->DeletedCount())
            );
        }
        $msg = ($msgArr) ? implode(',', $msgArr) : _t(
            'ExcelImportExport.ResultNone',
            'No changes'
        );

        $this->sessionMessage($msg, 'good');

        return $this->getController()->redirectBack();
    }
}
