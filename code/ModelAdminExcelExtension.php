<?php

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

    public function downloadsample()
    {
        ExcelImportExport::sampleFileForClass($this->owner->modelClass);
    }

    function updateEditForm(Form $form)
    {
        /* @var $owner ModelAdmin */
        $owner       = $this->owner;
        $class       = $owner->modelClass;
        $classConfig = $owner->config();

        $gridfield = $form->Fields()->dataFieldByName($class);
        if (!$gridfield) {
            return;
        }
        /* @var $config GridFieldConfig */
        $config = $gridfield->getConfig();

        // Bulk manage
        if ($classConfig->bulk_manage && class_exists('GridFieldBulkManager')) {
            $already = $config->getComponentByType('GridFieldBulkManager');
            if (!$already) {
                $config->addComponent($bulkManager = new GridFieldBulkManager());
                $bulkManager->removeBulkAction('unLink');
            }
        }

        if ($classConfig->export_csv) {
            /* @var $export GridFieldExportButton */
            $export = $config->getComponentByType('GridFieldExportButton');
            $export->setExportColumns(ExcelImportExport::exportFieldsForClass($class));
        } else {
            $config->removeComponentsByType('GridFieldExportButton');
        }
        if ($classConfig->export_excel) {
            $config->addComponent(new ExcelGridFieldExportButton('buttons-before-left'));
        }
    }

    function updateImportForm(Form $form)
    {
        /* @var $owner ModelAdmin */
        $owner = $this->owner;
        $class = $owner->modelClass;

        // Overwrite model imports 
        $importerClasses = $owner->stat('model_importers');

        if (is_null($importerClasses)) {
            $models = $owner->getManagedModels();
            foreach ($models as $modelName => $options) {
                $importerClasses[$modelName] = 'ExcelBulkLoader';
            }

            $owner->set_stat('model_importers', $importerClasses);
        }

        $modelSNG  = singleton($class);
        $modelName = $modelSNG->i18n_singular_name();

        $fields = $form->Fields();

        $content = _t('ModelAdminExcelExtension.DownloadSample',
            '<div class="field"><a href="{link}">Download sample file</a></div>',
            array('link' => $owner->Link($class.'/downloadsample')));

        $file = $fields->dataFieldByName('_CsvFile');
        if ($file) {
            $file->setDescription(ExcelImportExport::getValidExtensionsText());
            $file->getValidator()->setAllowedExtensions(ExcelImportExport::getValidExtensions());
        }

        $fields->removeByName("SpecFor{$modelName}");
        $fields->insertAfter('EmptyBeforeImport',
            new LiteralField("SampleFor{$modelName}", $content));

        if (!$modelSNG->canDelete()) {
            $fields->removeByName('EmptyBeforeImport');
        }

        $actions = $form->Actions();

        $import = $actions->dataFieldByName('action_import');
        if ($import) {
            $import->setTitle(_t('ModelAdminExcelExtension.ImportExcel',
                    "Import from Excel"));
        }
    }
}