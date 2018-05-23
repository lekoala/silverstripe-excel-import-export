<?php

namespace LeKoala\ExcelImportExport\Extensions;

use SilverStripe\Forms\Form;
use SilverStripe\Core\Extension;
use SilverStripe\Core\Config\Config;
use SilverStripe\Forms\LiteralField;
use LeKoala\ExcelImportExport\ExcelBulkLoader;
use LeKoala\ExcelImportExport\ExcelImportExport;
use SilverStripe\Forms\GridField\GridFieldConfig;
use SilverStripe\Forms\GridField\GridFieldExportButton;
use SilverStripe\Forms\GridField\GridFieldImportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldExportButton;
use LeKoala\ExcelImportExport\ExcelGridFieldImportButton;

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
        /* @var $owner ModelAdmin */
        $owner = $this->owner;
        $class = $owner->modelClass;
        $config = $this->owner->config();

        // Overwrite model imports
        $importerClasses = $config->get('model_importers');

        if (is_null($importerClasses)) {
            $models = $owner->getManagedModels();
            foreach ($models as $modelName => $options) {
                $importerClasses[$modelName] = ExcelBulkLoader::class;
            }

            $config->set('model_importers', $importerClasses);
        }
    }

    public function downloadsample()
    {
        ExcelImportExport::sampleFileForClass($this->owner->modelClass);
    }

    /**
     * Helper method to return a gridfield that your ide loves
     *
     * @param Form $form
     * @param string $sanitisedClass
     * @return GridField
     */
    protected function getGridFieldForClass($form, $sanitisedClass)
    {
        return $form->Fields()->dataFieldByName($sanitisedClass);
    }

    public function updateEditForm(Form $form)
    {
        /* @var $owner ModelAdmin */
        $owner       = $this->owner;
        $class       = $owner->modelClass;
        $sanitisedClass = str_replace('\\', '-', $class);
        $classConfig = $owner->config();

        $gridfield = $this->getGridFieldForClass($form, $sanitisedClass);
        if (!$gridfield) {
            return;
        }
        $config = $gridfield->getConfig();

        // Handle export buttons
        if ($classConfig->export_csv) {
            /* @var $export GridFieldExportButton */
            $GridFieldExportButton = $config->getComponentByType(GridFieldExportButton::class);
            $GridFieldExportButton->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
        } else {
            $config->removeComponentsByType(GridFieldExportButton::class);
        }
        if ($classConfig->export_excel) {
            $ExcelGridFieldExportButton = new ExcelGridFieldExportButton('buttons-before-left');
            $config->addComponent($ExcelGridFieldExportButton);
        }

        // Rename import button
        $config->removeComponentsByType(GridFieldImportButton::class);
        if ($this->owner->showImportForm) {
            $ExcelGridFieldImportButton = new ExcelGridFieldImportButton('buttons-before-left');
            $ExcelGridFieldImportButton->setImportForm($this->owner->ImportForm());
            $ExcelGridFieldImportButton->setModalTitle(_t('ExcelImportExport.IMPORTFROMFILE', 'Import from a file'));
            $config->addComponent($ExcelGridFieldImportButton);
        }
    }

    public function updateImportForm(Form $form)
    {
        /* @var $owner ModelAdmin */
        $owner = $this->owner;
        $class = $owner->modelClass;
        $modelSNG  = singleton($class);
        $modelName = $modelSNG->i18n_singular_name();

        $fields = $form->Fields();

        $downloadSampleLink = $owner->Link($class.'/downloadsample');
        $downloadSample = '<a href="'.$downloadSampleLink.'" class="no-ajax" target="_blank">' . _t(
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

        // We moved the specs into a nice to use download sample button
        $fields->removeByName("SpecFor{$modelName}");

        // If you cannot delete, you cannot empty
        if (!$modelSNG->canDelete()) {
            $fields->removeByName('EmptyBeforeImport');
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
