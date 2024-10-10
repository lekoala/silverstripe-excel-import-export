<?php

namespace LeKoala\ExcelImportExport\Extensions;

use SilverStripe\Forms\Form;
use SilverStripe\Core\Extension;
use SilverStripe\Admin\ModelAdmin;
use SilverStripe\Forms\CheckboxField;
use SilverStripe\Forms\OptionsetField;
use LeKoala\ExcelImportExport\ExcelBulkLoader;
use LeKoala\ExcelImportExport\ExcelImportExport;
use SilverStripe\Forms\GridField\GridFieldConfig;
use SilverStripe\Forms\GridField\GridFieldExportButton;
use SilverStripe\Forms\GridField\GridFieldImportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldExportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldImportButton;
use LeKoala\ExcelImportExport\ExcelGroupBulkLoader;
use LeKoala\ExcelImportExport\ExcelMemberBulkLoader;
use SilverStripe\Admin\SecurityAdmin;

/**
 * Extends {@link ModelAdmin}. to bind new forms and features
 *
 * @author Koala
 */
class ModelAdminExcelExtension extends Extension
{
    /**
     * @var array<string>
     */
    private static $allowed_actions = array(
        'downloadsample'
    );

    /**
     * @return void
     */
    public function onAfterInit()
    {
        $this->updateModelImporters();
    }

    /**
     * Replace default model import with a predefined value
     *
     * @return void
     */
    protected function updateModelImporters()
    {
        /** @var ModelAdmin $owner */
        $owner = $this->owner;
        //@phpstan-ignore-next-line
        $config = $this->owner::config();

        // Overwrite model imports
        $importerClasses = $config->get('model_importers');

        if ($owner instanceof SecurityAdmin) {
            $importerClasses = [
                "users" => ExcelMemberBulkLoader::class,
                "SilverStripe\Security\Member" => ExcelMemberBulkLoader::class,
                "groups" => ExcelGroupBulkLoader::class,
                "SilverStripe\Security\Group" => ExcelGroupBulkLoader::class,
            ];
            $config->set('model_importers', $importerClasses);
        } elseif (is_null($importerClasses)) {
            $models = $owner->getManagedModels();
            foreach (array_keys($models) as $modelName) {
                $importerClasses[$modelName] = ExcelBulkLoader::class;
            }

            $config->set('model_importers', $importerClasses);
        }
    }

    /**
     * @return void
     */
    public function downloadsample()
    {
        /** @var \SilverStripe\Admin\ModelAdmin $owner */
        $owner = $this->owner;
        $importer = $owner->getRequest()->getVar('importer');
        if ($importer && class_exists($importer) && method_exists($importer, 'getSampleFile')) {
            $importer::getSampleFile();
        } else {
            ExcelImportExport::sampleFileForClass($owner->getModelClass());
        }
    }

    /**
     * @param GridFieldConfig $config
     * @return void
     */
    public function updateGridFieldConfig(GridFieldConfig $config)
    {
        /** @var ModelAdmin $owner */
        $owner = $this->owner;
        $class = $owner->getModelClass();
        $classConfig = $owner->config();

        // Add/remove csv export. Replace with our own implementation if necessary.
        if ($classConfig->export_csv) {
            $GridFieldExportButton = $this->getGridFieldExportButton($config);
            if ($GridFieldExportButton) {
                if ($classConfig->use_framework_csv) {
                    $GridFieldExportButton->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
                } else {
                    $config->removeComponentsByType(GridFieldExportButton::class);
                    $ExcelGridFieldExportButton = new ExcelGridFieldExportButton('buttons-before-left');
                    $ExcelGridFieldExportButton->setExportType('csv');
                    $config->addComponent($ExcelGridFieldExportButton);
                }
            }
        } else {
            $config->removeComponentsByType(GridFieldExportButton::class);
        }

        // Add/remove csv export (add by default)
        if ($classConfig->export_excel) {
            $ExcelGridFieldExportButton = new ExcelGridFieldExportButton('buttons-before-left');
            $config->addComponent($ExcelGridFieldExportButton);
        }

        // Rename import button
        $config->removeComponentsByType(GridFieldImportButton::class);
        if ((is_bool($owner->showImportForm)
                && $owner->showImportForm)
            ||
            (is_array($owner->showImportForm)
                && in_array($class, $owner->showImportForm))
        ) {
            $importForm = $owner->ImportForm();
            if ($importForm) {
                $ExcelGridFieldImportButton = new ExcelGridFieldImportButton('buttons-before-left');
                $ExcelGridFieldImportButton->setImportForm($importForm);
                $ExcelGridFieldImportButton->setModalTitle(_t('ExcelImportExport.IMPORTFROMFILE', 'Import from a file'));
                $config->addComponent($ExcelGridFieldImportButton);
            }
        }
    }

    /**
     * @param \SilverStripe\Forms\GridField\GridFieldConfig $config
     * @return GridFieldExportButton|null
     */
    protected function getGridFieldExportButton($config)
    {
        /** @var GridFieldExportButton|null $comp */
        $comp = $config->getComponentByType(GridFieldExportButton::class);
        if ($comp) {
            return $comp;
        }
        return null;
    }

    /**
     * @param Form $form
     * @return void
     */
    public function updateImportForm(Form $form)
    {
        /** @var ModelAdmin $owner */
        $owner = $this->owner;
        $class = $owner->getModelClass();
        $classConfig = $owner->config();

        $modelSNG = singleton($class);
        $modelConfig = $modelSNG->config();
        $modelName = $modelSNG->i18n_singular_name();

        $fields = $form->Fields();

        // We can implement a custom handler
        $importHandlers = [];
        $htmlDesc = '';
        $useDefaultSample = true;
        if ($modelSNG->hasMethod('listImportHandlers')) {
            $importHandlers = array_merge([
                'default' => _t('ExcelImportExport.DefaultHandler', 'Default import handler'),
            ], $modelSNG->listImportHandlers());

            $supportOnlyUpdate = [];
            foreach ($importHandlers as $class => $label) {
                if (!class_exists($class)) {
                    continue;
                }
                if (method_exists($class, 'setOnlyUpdate')) {
                    $supportOnlyUpdate[] = $class;
                }
                if (method_exists($class, 'getImportDescription')) {
                    $htmlDesc .= '<div class="js-import-desc" data-name="' . $class . '" hidden>' . $class::getImportDescription() . '</div>';
                }
                if (method_exists($class, 'getSampleFileLink')) {
                    $useDefaultSample = false;
                    $htmlDesc .= '<div class="js-import-desc" data-name="' . $class . '" hidden>' . $class::getSampleFileLink() . '</div>';
                }
            }
        }

        /** @var \SilverStripe\Forms\FileField|null $file */
        $file = $fields->dataFieldByName('_CsvFile');
        if ($file) {
            $csvDescription = ExcelImportExport::getValidExtensionsText();
            if ($useDefaultSample) {
                $downloadSample = ExcelImportExport::createDownloadSampleLink();
                $csvDescription .= '. ' . $downloadSample;
            }
            $file->setDescription($csvDescription);
            $file->getValidator()->setAllowedExtensions(ExcelImportExport::getValidExtensions());
        }

        // Hide by default, but allow on per class basis (opt-in)
        if ($classConfig->hide_replace_data && !$modelConfig->show_replace_data) {
            // This is way too dangerous and customers don't understand what this is most of the time
            $fields->removeByName("EmptyBeforeImport");
        }
        // If you cannot delete, you cannot empty
        if (!$modelSNG->canDelete()) {
            $fields->removeByName('EmptyBeforeImport');
        }

        // We moved the specs into a nice to use download sample button
        $fields->removeByName("SpecFor{$modelName}");

        if (!empty($importHandlers)) {
            $form->Fields()->push($OnlyUpdateRecords = new CheckboxField("OnlyUpdateRecords", _t('ExcelImportExport.OnlyUpdateRecords', "Only update records")));
            $OnlyUpdateRecords->setAttribute("data-handlers", implode(",", $supportOnlyUpdate));

            $form->Fields()->push($ImportHandler = new OptionsetField("ImportHandler", _t('ExcelImportExport.PleaseSelectImportHandler', "Please select the import handler"), $importHandlers));
            // Simply check of this is supported or not for the given handler (if not, disable it)
            $js = <<<JS
var desc=document.querySelectorAll('.js-import-desc');var cb=document.querySelector('#OnlyUpdateRecords');var accepted=cb.dataset.handlers.split(',');var item=([...this.querySelectorAll('input')].filter((input) => input.checked)[0]); cb.disabled=(item && accepted.includes(item.value)) ? '': 'disabled';desc.forEach((el)=>el.hidden=!item||el.dataset.name!=item.value);;
JS;
            $ImportHandler->setAttribute("onclick", $js);
            if ($htmlDesc) {
                $ImportHandler->setDescription($htmlDesc); // Description is an HTMLFragment
            }
        }

        $actions = $form->Actions();

        // Update import button
        $import = $actions->dataFieldByName('action_import');
        if ($import) {
            $import->setTitle(_t(
                'ExcelImportExport.ImportExcel',
                "Import from Excel"
            ));
            $import->removeExtraClass('btn-outline-secondary');
            $import->addExtraClass('btn-primary');
        }
    }
}
