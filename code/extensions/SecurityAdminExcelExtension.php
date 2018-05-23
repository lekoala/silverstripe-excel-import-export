<?php

namespace LeKoala\ExcelImportExport\Extensions;

use SilverStripe\Forms\Form;
use SilverStripe\Core\Extension;
use SilverStripe\Security\Group;
use SilverStripe\Security\Member;
use SilverStripe\View\Requirements;
use SilverStripe\Forms\LiteralField;
use SilverStripe\Security\Permission;
use SilverStripe\Forms\GridField\GridField;
use SilverStripe\Core\Config\Config_ForClass;
use LeKoala\ExcelImportExport\ExcelImportExport;
use LeKoala\ExcelImportExport\ExcelGroupImportForm;
use LeKoala\ExcelImportExport\ExcelMemberImportForm;
use SilverStripe\Forms\GridField\GridFieldButtonRow;
use SilverStripe\Forms\GridField\GridFieldDetailForm;
use SilverStripe\Forms\GridField\GridFieldExportButton;
use SilverStripe\Forms\GridField\GridFieldImportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldExportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldImportButton;

/**
 * Extends {@link SecurityAdmin}. to bind new forms and features
 *
 * @author Koala
 */
class SecurityAdminExcelExtension extends Extension
{
    private static $allowed_actions = array(
        'excelmemberimport',
        'ExcelMemberImportForm',
        'excelgroupimport',
        'ExcelGroupImportForm',
        'downloadsample',
    );

    /**
     * @param Form $form
     * @return GridField
     */
    protected function getMembersGridField($form)
    {
        return $form->Fields()->dataFieldByName('Members');
    }

    /**
     * @param Form $form
     * @return GridField
     */
    protected function getGroupsGridField($form)
    {
        return $form->Fields()->dataFieldByName('Groups');
    }

    public function updateEditForm(Form $form)
    {
        /* @var $owner SecurityAdmin */
        $owner       = $this->owner;
        /* @var $config Config_ForClass */
        $classConfig = $owner->config();

        $members = $this->getMembersGridField($form);
        $membersConfig = $members->getConfig();
        $groups = $this->getGroupsGridField($form);
        $groupsConfig = $groups->getConfig();

        if ($classConfig->allow_import && Permission::check('ADMIN')) {
            $membersConfig->removeComponentsByType(GridFieldImportButton::class);
            $ExcelGridFieldImportButton = new ExcelGridFieldImportButton('buttons-before-left');
            $ExcelGridFieldImportButton->setImportIframe($this->owner->Link('excelmemberimport'));
            $ExcelGridFieldImportButton->setModalTitle(_t('ExcelImportExport.IMPORTMEMBERSFROMFILE', 'Import members from a file'));
            $membersConfig->addComponent($ExcelGridFieldImportButton);

            $groupsConfig->removeComponentsByType(GridFieldImportButton::class);
            $ExcelGridFieldImportButton = new ExcelGridFieldImportButton('buttons-before-left');
            $ExcelGridFieldImportButton->setImportIframe($this->owner->Link('excelgroupimport'));
            $ExcelGridFieldImportButton->setModalTitle(_t('ExcelImportExport.IMPORTGROUPSFROMFILE', 'Import groups from a file'));
            $groupsConfig->addComponent($ExcelGridFieldImportButton);
        }

        // Export features
        $this->addExportButton($members, $classConfig);
        $this->addExportButton($groups, $classConfig);
    }

    public function downloadsample()
    {
        $class = $this->owner->getRequest()->param('ID');
        switch ($class) {
            case 'Group':
                $class = Group::class;
                break;
            case 'Member':
                $class = Member::class;
                break;
        }
        ExcelImportExport::sampleFileForClass($class);
    }

    protected function requireAdminAssets()
    {
        Requirements::clear();
        Requirements::javascript('silverstripe/admin: client/dist/js/vendor.js');
        Requirements::javascript('silverstripe/admin: client/dist/js/MemberImportForm.js');
        Requirements::css('silverstripe/admin: client/dist/styles/bundle.css');
    }

    protected function addExportButton(
        GridField $gridfield,
        Config_ForClass $classConfig
    ) {
        if (!$gridfield) {
            return;
        }

        $config = $gridfield->getConfig();
        $class  = $gridfield->getModelClass();

        // Better export
        if ($classConfig->export_csv) {
            /* @var $export GridFieldExportButton */
            $export = $config->getComponentByType(GridFieldExportButton::class);
            $export->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
        } else {
            $config->removeComponentsByType(GridFieldExportButton::class);
        }
        if ($classConfig->export_excel) {
            $ExcelGridFieldExportButton = new ExcelGridFieldExportButton('buttons-before-left');
            $config->addComponent($ExcelGridFieldExportButton);
        }
    }

    /**
     * @see SecurityAdmin_MemberImportForm
     *
     * @return Form
     */
    public function ExcelMemberImportForm()
    {
        if (!Permission::check('ADMIN')) {
            return false;
        }

        $group = $this->owner->currentPage();
        $form  = new ExcelMemberImportForm(
            $this->owner,
            'ExcelMemberImportForm'
        );
        $form->setGroup($group);

        return $form;
    }

    public function excelmemberimport()
    {
        $this->requireAdminAssets();

        return $this->owner->renderWith(
            'BlankPage',
            array('Form' => $this->ExcelMemberImportForm()->forTemplate(), 'Content' => ' ')
        );
    }

    /**
     * @see SecurityAdmin_GroupImportForm
     *
     * @return Form
     */
    public function ExcelGroupImportForm()
    {
        if (!Permission::check('ADMIN')) {
            return false;
        }

        $form = new ExcelGroupImportForm(
            $this->owner,
            'ExcelGroupImportForm'
        );

        return $form;
    }

    public function excelgroupimport()
    {
        $this->requireAdminAssets();

        return $this->owner->renderWith(
            'BlankPage',
            array('Form' => $this->ExcelGroupImportForm()->forTemplate(), 'Content' => ' ' )
        );
    }
}
