<?php

/**
 * (c) Daniel Richardt
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */

namespace dpmR\Bundle\TranslationExtractorBundle\Command;


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\DependencyInjection\ParameterBag\ContainerBagInterface;
use Symfony\Component\Translation\Translator;
use Symfony\Contracts\Translation\TranslatorInterface;

class ExportCommand extends Command
{

    /**
     * @var string
     */
    protected static $defaultName = 'dpmR:export-translations';

    /**
     * @var Translator
     */
    protected $translator;

    /**
     * @var ContainerBagInterface
     */
    protected $container;

    public function __construct(
        TranslatorInterface $translator,
        ContainerBagInterface $container,
        ?string $name = null
    )
    {
        $this->translator = $translator;
        $this->container = $container;
        parent::__construct($name);
    }

    /**
     * @param InputInterface $input
     * @param OutputInterface $output
     * @return int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function execute(
        InputInterface $input,
        OutputInterface $output
    ): int
    {
        $spreadsheet = new Spreadsheet();

        $spreadsheet->removeSheetByIndex($spreadsheet->getActiveSheetIndex());

        $activeSheet = 0;
        foreach ($this->container->get('php_translation.locales') as $locale) {
            $sheet = $spreadsheet->createSheet($activeSheet);
            $sheet->setTitle($locale);

            $i = 1;
            foreach (['Domain', 'Original', 'Translation'] as $header) {
                $sheet->setCellValueByColumnAndRow($i, 1, $header);
                $i++;
            }

            $catalogue = $this->translator->getCatalogue($locale);
            $domains = $catalogue->getDomains();

            $domainMessages = $catalogue->all();

            $fallbackCatalogue = $catalogue->getFallbackCatalogue();
            if (!empty($fallbackCatalogue)) {
                $fallbackMessages = $fallbackCatalogue->all();

                foreach ($domains as $domain) {
                    foreach ($fallbackMessages[$domain] as $key => $fallbackMessage) {
                        if (!isset($domainMessages[$domain][$key])) {
                            $domainMessages[$domain][$key] = '';
                        }
                    }
                }
            }

            $row = 2;
            foreach ($domains as $domain) {
                foreach ($domainMessages[$domain] as $key => $domainMessage) {
                    $sheet->setCellValueByColumnAndRow(1, $row, $domain);
                    $sheet->setCellValueByColumnAndRow(2, $row, $key);
                    $sheet->setCellValueByColumnAndRow(3, $row, $domainMessage);

                    $row++;
                }
            }

            $activeSheet++;
        }

        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            $spreadsheet->setActiveSheetIndex($spreadsheet->getIndex($worksheet));

            $sheet = $spreadsheet->getActiveSheet();
            $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(true);

            foreach ($cellIterator as $cell) {
                $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
            }
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($this->container->get('kernel.project_dir') . '/var/Translations.xlsx');

        return 0;
    }
}
