<?php


namespace NodaSoft\PhpXlsTemplating;


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Pdf;

class Templating
{
    /**
     * @var string
     */
    private $tplFileName = '';

    /**
     * @var string
     */
    private $resultFileName = '';

    /**
     * @var ProcessOptions
     */
    private $options;

    /**
     * @var bool
     */
    private $isSendToBrowser = false;

    /**
     * @var bool
     */
    private $isSendAsAttach = false;

    /**
     * @param SheetTplData[] $tplData
     * @param string $writerType
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function run(array $tplData, string $writerType = 'Xlsx'): void
    {
        $process = new ProcessSpreadsheet();
        $process->setTplData($tplData);
        if ($this->options !== null) {
            $process->setOptions($this->options);
        }
        $spreadsheet = $this->loadSpreadsheet();
        $spreadsheet = $process->run($spreadsheet);
        $this->save($spreadsheet, $writerType);
    }

    /**
     * Загружает файл шаблона
     *
     * @return Spreadsheet
     */
    private function loadSpreadsheet(): Spreadsheet
    {
        $spreadsheet = IOFactory::load($this->tplFileName);
        $spreadsheet
            ->getProperties()
            ->setCreator('User User')
            ->setLastModifiedBy('User User')
            ->setTitle('Document')
            ->setSubject('Document')
            ->setDescription('')
            ->setKeywords('')
            ->setCategory('');

        return $spreadsheet;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string $writerType
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function save(Spreadsheet $spreadsheet, string $writerType): void
    {
        $writer = IOFactory::createWriter($spreadsheet, $writerType);

        $fileDir = dirname($this->resultFileName);
        $outputFileName = basename($this->resultFileName);
        if (in_array($writerType,['Tcpdf', 'Mpdf',  'Dompdf'], true)) {
            // Не вычислять формулы при загрузке документа
            $writer->setPreCalculateFormulas(false);
            if ($this->isSendToBrowser) {
                $this->sendPdfHeader($outputFileName);
            }
        } else {
            if ($this->isSendToBrowser) {
                $this->sendExcelHeader($outputFileName);
            }
        }

        if ($this->isSendToBrowser) {
            $writer->save('php://output');
        } else {
            $writer->save("{$fileDir}/{$outputFileName}");
        }
    }

    /**
     * Отправляет в браузер хидер с данными о PDF-файле
     *
     * @param string $fileName
     */
    private function sendPdfHeader(string $fileName): void
    {
        $disposition = $this->isSendAsAttach ? 'attachment' : 'inline';
        header('Cache-Control: no-cache, must-revalidate');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Cache-Control: private', false);
        header('Content-type: application/pdf');
        header('Content-Disposition: ' . $disposition . '; filename="' . $fileName . '"');
    }

    /**
     * Отправляет в браузер хидер с данными о XLS-файле
     *
     * @param string $fileName
     */
    private function sendExcelHeader(string $fileName): void
    {
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');
    }

    /**
     * @param string $tplFileName
     * @return Templating
     */
    public function setTplFileName(string $tplFileName): Templating
    {
        $this->tplFileName = $tplFileName;

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
     * @param string $resultFileName
     * @return Templating
     */
    public function setResultFileName(string $resultFileName): Templating
    {
        $this->resultFileName = $resultFileName;

        return $this;
    }

    /**
     * @param bool $isSendToBrowser
     * @return Templating
     */
    public function setIsSendToBrowser(bool $isSendToBrowser): Templating
    {
        $this->isSendToBrowser = $isSendToBrowser;

        return $this;
    }

    /**
     * @param bool $isSendAsAttach
     * @return Templating
     */
    public function setIsSendAsAttach(bool $isSendAsAttach): Templating
    {
        $this->isSendAsAttach = $isSendAsAttach;

        return $this;
    }
}