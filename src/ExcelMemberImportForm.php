<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\Forms\FieldList;
use SilverStripe\Forms\FileField;
use SilverStripe\Forms\FormAction;
use SilverStripe\Forms\LiteralField;
use SilverStripe\Admin\MemberImportForm;

/**
 * Import members from excel
 *
 * @author Koala
 */
class ExcelMemberImportForm extends MemberImportForm
{
    public function __construct(
        $controller,
        $name,
        $fields = null,
        $actions = null,
        $validator = null
    ) {
        $downloadSampleLink = $controller->Link('downloadsample/Member');
        $downloadSample = '<a href="' . $downloadSampleLink . '" class="no-ajax" target="_blank">' . _t(
            'ExcelImportExport.DownloadSample',
            'Download sample file'
        ) . '</a>';

        $helpHtml = '';
        $helpHtml .= '<ul>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.MemberImportHelp1',
            'For supported columns, please check the sample file.'
        ) . ' ' . $downloadSample . '</li>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.MemberImportHelp2',
            'Existing users are matched by their unique <em>Email</em> property, and updated with any new values from the imported file.'
        ) . '</li>';
        $helpHtml .= '<li>' . _t(
            'ExcelImportExport.MemberImportHelp3',
            'Groups can be assigned by the <em>Groups</em> column. Groups are identified by their <em>Code</em> property, multiple groups can be separated by comma. Existing group memberships are not cleared.'
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
        $loader = new ExcelMemberBulkLoader();

        // optionally set group relation
        if ($this->group) {
            $loader->setGroups(array($this->group));
        }

        // load file
        $result = $loader->load($data['File']);

        // result message
        $msgArr   = array();
        if ($result->CreatedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultCreated',
                'Created {count} members',
                array('count' => $result->CreatedCount())
            );
        }
        if ($result->UpdatedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultUpdated',
                'Updated {count} members',
                array('count' => $result->UpdatedCount())
            );
        }
        if ($result->DeletedCount()) {
            $msgArr[] = _t(
                'ExcelImportExport.ResultDeleted',
                'Deleted %d members',
                array('count' => $result->DeletedCount())
            );
        }
        $msg      = ($msgArr) ? implode(',', $msgArr) : _t(
            'ExcelImportExport.ResultNone',
            'No changes'
        );

        $this->sessionMessage($msg, 'good');

        return $this->getController()->redirectBack();
    }
}
