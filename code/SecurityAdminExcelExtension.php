<?php

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

    function updateEditForm(Form $form)
    {
        /* @var $owner SecurityAdmin */
        $owner       = $this->owner;
        /* @var $config Config_ForClass */
        $classConfig = $owner->config();

        // Import features
        $form->Fields()->removeByName('MemberImportFormIframe');
        $form->Fields()->removeByName('GroupImportFormIframe');

        if ($classConfig->allow_import && Permission::check('ADMIN')) {
            $field = new LiteralField(
                'MemberImportFormIframe',
                sprintf(
                    '<iframe src="%s" id="MemberImportFormIframe" width="100%%" height="250px" frameBorder="0">'
                    .'</iframe>', $this->owner->Link('excelmemberimport')
                )
            );

            $form->Fields()->addFieldToTab('Root.Users', $field);

            $field = new LiteralField(
                'GroupImportFormIframe',
                sprintf(
                    '<iframe src="%s" id="GroupImportFormIframe" width="100%%" height="250px" frameBorder="0">'
                    .'</iframe>', $this->owner->Link('excelgroupimport')
                )
            );

            $form->Fields()->addFieldToTab('Root.Groups', $field);
        }

        // Export features

        $members = $form->Fields()->dataFieldByName('Members');
        $this->setupGridField($members, $classConfig);

        /* @var $groups GridField */
        $groups     = $form->Fields()->dataFieldByName('Groups');
        $this->setupGridField($groups, $classConfig);
        /* @var $detailForm GridFieldDetailForm */
        $detailForm = $groups->getConfig()->getComponentByType('GridFieldDetailForm');
        $detailForm->setItemEditFormCallback(function($form) use($classConfig) {
            $members = $form->Fields()->dataFieldByName('Members');
            if (!$members) {
                return;
            }
            $config = $members->getConfig();
            if ($classConfig->export_csv) {
                /* @var $export GridFieldExportButton */
                $export = $config->getComponentByType('GridFieldExportButton');
                $export->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
            } else {
                $config->removeComponentsByType('GridFieldExportButton');
            }
            if ($classConfig->export_excel) {
                $config->addComponent(new ExcelGridFieldExportButton('buttons-after-left'));
            }
        });
    }

    public function downloadsample()
    {
        $class = $this->owner->getRequest()->param('ID');
        ExcelImportExport::sampleFileForClass($class);
    }

    protected function setupGridField(GridField $gridfield,
                                      Config_ForClass $classConfig)
    {
        if (!$gridfield) {
            return;
        }

        $config = $gridfield->getConfig();
        $class  = $gridfield->getModelClass();

        // More item per page
        $paginator = $config->getComponentByType('GridFieldPaginator');
        $paginator->setItemsPerPage(50);

        // Bulk manage
        if ($classConfig->bulk_manage && class_exists('GridFieldBulkManager')) {
            $already = $config->getComponentByType('GridFieldBulkManager');
            if (!$already) {
                $config->addComponent($bulkManager = new GridFieldBulkManager());
                $bulkManager->removeBulkAction('unLink');
            }
        }

        // Better export
        if ($classConfig->export_csv) {
            /* @var $export GridFieldExportButton */
            $export = $config->getComponentByType('GridFieldExportButton');
            $export->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
        } else {
            $config->removeComponentsByType('GridFieldExportButton');
        }
        if ($classConfig->export_excel) {
            if ($class == 'Group') {
                $config->addComponent(new GridFieldButtonRow('after'));
            }
            $config->addComponent(new ExcelGridFieldExportButton('buttons-after-left'));
        }
    }

    /**
     * @see SecurityAdmin_MemberImportForm
     *
     * @return Form
     */
    public function ExcelMemberImportForm()
    {
        if (!Permission::check('ADMIN')) return false;

        $group = $this->owner->currentPage();
        $form  = new ExcelMemberImportForm(
            $this->owner, 'ExcelMemberImportForm'
        );
        $form->setGroup($group);

        return $form;
    }

    protected function importFormRequirements()
    {
        Requirements::clear();
        Requirements::javascript(THIRDPARTY_DIR.'/jquery/jquery.js');
        Requirements::css(FRAMEWORK_ADMIN_DIR.'/css/screen.css');
    }

    public function excelmemberimport()
    {
        $this->importFormRequirements();

        return $this->owner->renderWith('BlankPage',
                array(
                'Form' => $this->ExcelMemberImportForm()->forTemplate(),
                'Content' => ' '
        ));
    }

    /**
     * @see SecurityAdmin_GroupImportForm
     *
     * @return Form
     */
    public function ExcelGroupImportForm()
    {
        if (!Permission::check('ADMIN')) return false;

        $form = new ExcelGroupImportForm(
            $this->owner, 'ExcelGroupImportForm'
        );

        return $form;
    }

    public function excelgroupimport()
    {
        $this->importFormRequirements();

        return $this->owner->renderWith('BlankPage',
                array(
                'Form' => $this->ExcelGroupImportForm()->forTemplate(),
                'Content' => ' '
        ));
    }
}