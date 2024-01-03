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
    private static $allowed_actions = array(
        'downloadsample'
    );

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
        $config = $this->owner->config();

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

    public function downloadsample()
    {
        ExcelImportExport::sampleFileForClass($this->owner->modelClass);
    }

    public function updateGridFieldConfig(GridFieldConfig $config)
    {
        /** @var ModelAdmin $owner */
        $owner = $this->owner;
        $class = $owner->modelClass;
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
        if (is_bool($this->owner->showImportForm) && $this->owner->showImportForm || is_array($this->owner->showImportForm) && in_array($class, $this->owner->showImportForm)) {
            $ExcelGridFieldImportButton = new ExcelGridFieldImportButton('buttons-before-left');
            $ExcelGridFieldImportButton->setImportForm($this->owner->ImportForm());
            $ExcelGridFieldImportButton->setModalTitle(_t('ExcelImportExport.IMPORTFROMFILE', 'Import from a file'));
            $config->addComponent($ExcelGridFieldImportButton);
        }
    }

    /**
     * @param GridFieldConfig $config
     * @return GridFieldExportButton
     */
    protected function getGridFieldExportButton($config)
    {
        return $config->getComponentByType(GridFieldExportButton::class);
    }

    public function updateImportForm(Form $form)
    {
        /** @var ModelAdmin $owner */
        $owner = $this->owner;
        $class = $owner->modelClass;
        $classConfig = $owner->config();

        $modelSNG = singleton($class);
        $modelConfig = $modelSNG->config();
        $modelName = $modelSNG->i18n_singular_name();

        $fields = $form->Fields();

        $downloadSampleLink = $owner->Link(str_replace('\\', '-', $class) . '/downloadsample');
        $downloadSample = '<a href="' . $downloadSampleLink . '" class="no-ajax" target="_blank">' . _t(
            'ExcelImportExport.DownloadSample',
            'Download sample file'
        ) . '</a>';

        $file = $fields->dataFieldByName('_CsvFile');
        if ($file) {
            $csvDescription = ExcelImportExport::getValidExtensionsText();
            $csvDescription .= '. ' . $downloadSample;
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

        // We can implement a custom handler
        $importHandlers = [];
        if ($modelSNG->hasMethod('listImportHandlers')) {
            $importHandlers = array_merge([
                'default' => _t('ExcelImportExport.DefaultHandler', 'Default import handler'),
            ], $modelSNG->listImportHandlers());

            $supportOnlyUpdate = [];
            foreach ($importHandlers as $class => $label) {
                if (class_exists($class) && method_exists($class, 'setOnlyUpdate')) {
                    $supportOnlyUpdate[] = $class;
                }
            }

            $form->Fields()->push($OnlyUpdateRecords = new CheckboxField("OnlyUpdateRecords", _t('ExcelImportExport.OnlyUpdateRecords', "Only update records")));
            $OnlyUpdateRecords->setAttribute("data-handlers", implode(",", $supportOnlyUpdate));
        }
        if (!empty($importHandlers)) {
            $form->Fields()->push($ImportHandler = new OptionsetField("ImportHandler", _t('ExcelImportExport.PleaseSelectImportHandler', "Please select the import handler"), $importHandlers));
            // Simply check of this is supported or not for the given handler (if not, disable it)
            $js = <<<JS
var cb=document.querySelector('#OnlyUpdateRecords');var accepted=cb.dataset.handlers.split(',');var item=([...this.querySelectorAll('input')].filter((input) => input.checked)[0]); cb.disabled=(item && accepted.includes(item.value)) ? '': 'disabled';
JS;
            $ImportHandler->setAttribute("onclick", $js);
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
