<?php


namespace NodaSoft\PhpXlsTemplating;


use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Settings;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ProcessSpreadsheet
{
    /**
     * Массив с заменами для страниц, ключи массива совпадают с именами страниц XLS-файла
     * Если пустой ключ, то берется активная страница
     *
     * @var SheetTplData[]
     */
    private $tplData;

    /**
     * @var ProcessOptions
     */
    private $options;

    /**
     * Объект PHPExcel
     *
     * @var Spreadsheet
     */
    private $spreadsheet;

    public function __construct()
    {
        $this->options = new ProcessOptions();
    }

    /**
     * Устанавливает массив со значениями шаблонов для всех листов
     *
     * @param SheetTplData[] $tplData
     * @return self
     */
    public function setTplData(array $tplData): self
    {
        $this->tplData = $tplData;

        return $this;
    }

    /**
     * @param ProcessOptions $options
     * @return self
     */
    public function setOptions(ProcessOptions $options): self
    {
        $this->options = $options;

        return $this;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @return Spreadsheet
     * @throws Exception
     */
    public function run(Spreadsheet $spreadsheet): Spreadsheet
    {
        // Локаль для формул
        Settings::setLocale('en');
        $this->spreadsheet = $spreadsheet;
        $this->replaceWorksheets();

        return $this->spreadsheet;
    }

    /**
     * Сопоставляет массивы шаблонов с листами и производит замену на листах
     *
     * @throws Exception
     */
    private function replaceWorksheets(): void
    {
        $activeWorkSheet = clone $this->spreadsheet->getActiveSheet();
        foreach ($this->tplData as $worksheetTplData) {
            if (!empty($worksheetTplData->sheetName)) {
                $worksheet = $this->spreadsheet->getSheetByName($worksheetTplData->sheetName);
                if ($worksheet === null) {
                    $newWorksheet = clone $activeWorkSheet;
                    $sheetTitle = $activeWorkSheet->getTitle();
                    $newWorksheet->setTitle("{$sheetTitle}_{$worksheetTplData->sheetName}");
                    if ($activeWorkSheet->getParent() !== null) {
                        $worksheet = $activeWorkSheet->getParent()->addSheet($newWorksheet);
                    }
                }
            } else {
                $worksheet = $this->spreadsheet->getActiveSheet();
            }

            $this->processWorksheet($worksheet, $worksheetTplData);
        }
    }

    /**
     * Производит замену шаблонов на листе
     *
     * @param Worksheet $worksheet
     * @param SheetTplData $tplData
     * @throws Exception
     */
    private function processWorksheet(Worksheet $worksheet, SheetTplData $tplData): void
    {
        $pw = new ProcessWorksheet();
        $pw->setWorksheet($worksheet)->setTplData($tplData)->setOptions($this->options)->run();
    }
}