<?php

namespace LeKoala\ExcelImportExport;

use SilverStripe\Forms\Form;
use SilverStripe\View\SSViewer;
use SilverStripe\Forms\GridField\GridField;
use SilverStripe\Forms\GridField\GridField_FormAction;
use SilverStripe\Forms\GridField\GridFieldImportButton;
use SilverStripe\Model\ArrayData;

/**
 * Adds an "Import" action
 */
class ExcelGridFieldImportButton extends GridFieldImportButton
{
    /**
     * We need to override the getHTMLFragments to allow changing the title
     *
     * @param GridField $gridField
     * @return array<int|string,mixed>
     */
    public function getHTMLFragments($gridField)
    {
        $modalID = $gridField->ID() . '_ImportModal';

        // Check for form message prior to rendering form (which clears session messages)
        /** @var Form|null $form */
        $form = $this->getImportForm();
        $hasMessage = $form && $form->getMessage();

        // Render modal
        $template = SSViewer::get_templates_by_class(GridFieldImportButton::class, '_Modal');
        $viewer = new ArrayData([
            'ImportModalTitle' => $this->getModalTitle(),
            'ImportModalID' => $modalID,
            'ImportIframe' => $this->getImportIframe(),
            'ImportForm' => $this->getImportForm(),
        ]);
        $modal = $viewer->renderWith($template)->forTemplate();

        // Build action button
        $button = new GridField_FormAction(
            $gridField,
            'import',
            _t('ExcelImportExport.XLSIMPORT', 'Import'),
            'import',
            []
        );
        $button
            ->addExtraClass('btn btn-secondary font-icon-upload btn--icon-large action_import')
            ->setForm($gridField->getForm())
            ->setAttribute('data-toggle', 'modal')
            ->setAttribute('aria-controls', $modalID)
            ->setAttribute('data-target', "#{$modalID}")
            ->setAttribute('data-modal', $modal);

        // If form has a message, trigger it to automatically open
        if ($hasMessage) {
            $button->setAttribute('data-state', 'open');
        }

        return [
            $this->targetFragment => $button->Field()
        ];
    }
}
