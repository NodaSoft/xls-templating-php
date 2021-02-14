<?php


namespace NodaSoft\PhpXlsTemplating;


use NodaSoft\PhpXlsTemplating\Helpers\Image;
use NodaSoft\PhpXlsTemplating\Helpers\Merge;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ProcessWorksheet
{
    /**
     * Максимальное количество колонок, которое поддерживает генерация
     */
    private const COLUMNS_LIMIT = 255;

    /**
     * Текущая страница
     *
     * @var Worksheet
     */
    private $worksheet;

    /**
     * Текущий массив с заменами
     *
     * @var SheetTplData
     */
    private $tplData;

    /**
     * @var ProcessOptions
     */
    private $options;

    /**
     * Количество строк на текущей странице
     *
     * @var int
     */
    private $maxRowIndex;

    /**
     * Количество колонок на текущей странице
     *
     * @var int
     */
    private $maxColumnIndex;

    /**
     * @var Tags
     */
    private $tags;

    private $rowIndex = 0;

    private $columnIndex = 0;

    /**
     * @var ProcessCell
     */
    private $cellProcessor;

    public function __construct()
    {
    }

    /**
     * @throws Exception
     */
    public function run(): void
    {
        $this->tags = new Tags($this->options->openTag, $this->options->closeTag);
        $this->cellProcessor = new ProcessCell($this->tags, $this->tplData);

        $this->initMax();
        for ($this->rowIndex = 1; $this->rowIndex <= $this->maxRowIndex; ++$this->rowIndex) {
            for ($this->columnIndex = 0; $this->columnIndex <= $this->maxColumnIndex; ++$this->columnIndex) {
                $this->processCell();
            }
        }
        //			if ($this->isAutoHeightRow) {
        //				foreach($this->curWorksheet->getRowDimensions() as $rd) {
        //					$rd->setRowHeight(-1);
        //				}
        //			}
    }

    /**
     * Получает максимальные номера строк и столбцов на листе
     *
     * @throws Exception
     */
    protected function initMax(): void
    {
        $this->maxRowIndex = $this->worksheet->getHighestDataRow();
        $highestColumn = $this->worksheet->getHighestDataColumn();
        $highestColumnPos = Coordinate::columnIndexFromString($highestColumn);
        $this->maxColumnIndex = min(self::COLUMNS_LIMIT, $highestColumnPos);
    }

    /**
     * Обработка содержимого ячейки
     *
     * @throws Exception
     */
    private function processCell(): void
    {
        $cell = $this->getCell($this->columnIndex, $this->rowIndex);
        if ($cell === null) {
            return;
        }

        if (($this->columnIndex === 1) && $this->tags->matchAll($cell->getValue(), $matchedOperators)) {
            foreach ($matchedOperators[1] as $var) {
                if (isset($this->tplData->data[$var]) && is_array($this->tplData->data[$var])) {
                    $this->processArrayVar($this->tplData->data[$var]);
                }
            }
        }

        $res = $this->cellProcessor->process($cell);
        if ($res->isHideRow) {
            $this->hideRow($this->rowIndex);
            // Пропускаем все колонки до конца строки
            $this->columnIndex = $this->maxColumnIndex;

            return;
        }
        $this->setHeight($cell, $res->cellHeightLines);
        $this->insertImage($cell, $res->imageFilePath, $res->imageOptions);
    }

    /**
     * Проверяет наличие флага изменения высота в ячейке и если надо, изменяет высоту строки в соответствии с текстом
     *
     * @param Cell $cell
     * @param int|null $cellHeightLines
     */
    private function setHeight(Cell $cell, ?int $cellHeightLines): void
    {
        // Если флаг по высоте задан в ячейке
        if ($cellHeightLines !== null) {
            $this->worksheet->getRowDimension($cell->getRow())->setRowHeight($cellHeightLines);
            $this->worksheet->getStyle($cell->getCoordinate())->getAlignment()->setWrapText(true);
        }
    }

    /**
     * Скрываем строку
     *
     * @param $rowIndex
     */
    private function hideRow($rowIndex): void
    {
        //                // Скрываем текущую строку
        //                if ($this->useExcel2007Format) {
        //                    $this->worksheet->removeRow($rowIndex, 1);
        //                    $this->curHighestRow--;
        //                    $this->rowIndex--;
        //                } else {
        $this->worksheet->getRowDimension($rowIndex)->setVisible(false);
        //                }
    }

    /**
     * @param array $rowsArray Набор переменных для подстановки в строки при размножении
     * @throws Exception
     */
    private function processArrayVar(array $rowsArray): void
    {
        // Кол-во строк в одной группе, которую нужно мультиплицировать
        $mergedRowsCount = Merge::getMergedRowsCount($this->columnIndex, $this->rowIndex, $this->worksheet);
        // Если ячейка с указанием на мультиплицирование строки объединяет несколько строк
        if ($mergedRowsCount > 1) {
            // мультиплицирование нескольких строк шаблона
            $this->replaceAndMultiplyMergedRows($rowsArray, $mergedRowsCount);
        } else {
            // мультиплицирование одной строки шаблона
            $this->replaceAndMultiplyRow($rowsArray);
        }

        $this->columnIndex = $this->maxColumnIndex; // все колонки до конца были обработаны в вызванных функциях
    }

    /**
     * Производит замену шаблонов в строке и мультиплицирование строки
     *
     * @param array $rowsArray Набор переменных для подстановки в строки при размножении
     * @throws Exception
     */
    private function replaceAndMultiplyRow(array $rowsArray): void
    {
        // Получаем строку для копирования
        $sourceRowCells = $this->getRowCellsArray($this->columnIndex, $this->rowIndex);

        // Получаем объединенные ячейки в текущей строке
        $horizontalMerge = Merge::getHorizontalMergeInRow($this->rowIndex, $this->worksheet);

        $rowsCount = count($rowsArray);
        foreach ($rowsArray as $index => $rowVars) {
            $rowIndex = $this->rowIndex + $index;
            if ($index < ($rowsCount - 1)) { // Не последняя строка
                $this->insertRow($rowIndex, $rowIndex + 1, $sourceRowCells);
            }

            // Объединяем ячейки в новой строке по образцу первой строки
            Merge::setHorizontal($horizontalMerge, $rowIndex, $this->worksheet);

            // Перебираем ячейки строки
            for ($colIndex = 0; $colIndex <= $this->maxColumnIndex; ++$colIndex) {
                $cell = $this->getCell($colIndex, $rowIndex);
                if ($cell === null) {
                    continue;
                }
                $res = $this->cellProcessor->process($cell, $rowVars);
                if ($res->isHideRow) {
                    $this->hideRow($this->rowIndex);
                    // Пропускаем все колонки до конца строки
                    continue 2;
                }
                $this->setHeight($cell, $res->cellHeightLines);
                $this->insertImage($cell, $res->imageFilePath, $res->imageOptions);
            }
        }
        // Увеличивает текущий номер строки на количество добавленных строк,
        // т.к. в них уже ничего заменять не надо
        $this->rowIndex += count($rowsArray) - 1;
        $this->maxRowIndex += count($rowsArray) - 1;
    }

    /**
     * Вставляет новую строку под $rowIndex
     *
     * @param int $sourceRowIndex
     * @param int $targetRowIndex
     * @param array $rowCells Ячейки строки
     * @throws Exception
     */
    private function insertRow(int $sourceRowIndex, int $targetRowIndex, array $rowCells): void
    {
        // Создаем новую пустую строку под текущей
        if ($this->options->insertInsteadOfCopy) {
            $this->worksheet->insertNewRowBefore($targetRowIndex);
        } else {
            // Вставляем в неё скопированную первоначальную
            $this->worksheet->fromArray($rowCells, null, 'A' . $targetRowIndex, false);
            if ($this->options->copyStyles) {
                $this->duplicateRowStyle($sourceRowIndex, $targetRowIndex);
            }
        }
    }

    /**
     * Производит мультиплицирование групп по несколько строк и замену шаблонов в них
     *
     * @param array $rowsArray Набор переменных для подстановки в строки при размножении
     * @param int $mergedRowsCount Кол-во строк в одной группе, которую нужно мультиплицировать
     * @throws Exception
     */
    private function replaceAndMultiplyMergedRows(array $rowsArray, int $mergedRowsCount): void
    {
        // Сбор шаблонов строк для копирования
        $templateRows = [];
        for ($i = 0; $i < $mergedRowsCount; $i++) {
            $templateRows[] = $this->getRowCellsArray($this->columnIndex, $this->rowIndex + $i);
        }

        // Цикл создания всех необходимых строк - последовательно по одной строке из группы
        $insertedReal = 0;
        $rowsArrayCount = count($rowsArray);
        foreach ($rowsArray as $_) { // Перебираем значения для подстановки из массива данных
            foreach ($templateRows as $rowCells) {
                $rowIndex = $this->rowIndex + $insertedReal;
                // Вставка с вертикальным смещением в одну группу строк
                $insertTargetRowIndex = $rowIndex + $mergedRowsCount;

                // Т.к. в файле уже присутствует 1 группа строк для вставки, уменьшаем число добавляемых групп на 1
                if ($insertedReal < ($mergedRowsCount * ($rowsArrayCount - 1))) {
                    $this->insertRow($rowIndex, $insertTargetRowIndex, $rowCells);
                    $insertedReal++;
                }
            }
        }

        // Цикл замены шаблонов в созданных строках и попутное объединение ячеек
        $rowsCounter = 0;
        foreach ($rowsArray as $replacementRowKey => $rowVars) {
            foreach ($templateRows as $templateRowKey => $rowCells) {
                $rowIndex = $this->rowIndex + $rowsCounter;
                Merge::doBothWays($this->rowIndex + $templateRowKey, $rowIndex, $mergedRowsCount * $replacementRowKey,
                    $this->worksheet);

                for ($colIndex = 0; $colIndex <= $this->maxColumnIndex; ++$colIndex) {
                    $cell = $this->getCell($colIndex, $rowIndex);
                    if ($cell === null) {
                        return;
                    }
                    $res = $this->cellProcessor->process($cell, $rowVars);
                    if ($res->isHideRow) {
                        $this->hideRow($this->rowIndex);
                        // Пропускаем все колонки до конца строки
                        $colIndex = $this->maxColumnIndex;
                        continue;
                    }
                    $this->setHeight($cell, $res->cellHeightLines);
                    $this->insertImage($cell, $res->imageFilePath, $res->imageOptions);
                }
                $rowsCounter++;
            }
        }
        // Увеличивает текущий номер строки на количество добавленных строк,
        // т.к. в них уже ничего заменять не надо
        $this->rowIndex += $insertedReal + $mergedRowsCount - 1;
        $this->maxRowIndex += $insertedReal;
    }

    /**
     * Копирует стили строки в другую строку
     *
     * @param int $sourceRowIndex Номер копируемой строки
     * @param int $targetRowIndex Номер строки, которая получит копируемые стили
     */
    private function duplicateRowStyle(int $sourceRowIndex, int $targetRowIndex): void
    {
        for ($pColumn = 0; $pColumn <= $this->maxColumnIndex; ++$pColumn) {
            $cellAddressCurrent = Coordinate::stringFromColumnIndex($pColumn) . $sourceRowIndex;
            $cellAddressNew = Coordinate::stringFromColumnIndex($pColumn) . $targetRowIndex;
            $this->worksheet->duplicateStyle($this->worksheet->getStyle($cellAddressCurrent),
                "{$cellAddressNew}:{$cellAddressNew}");
        }
    }

    /**
     * Возврашает массив со значениями ячеек строки, начиная с ячейки с заданными координатами
     *
     * @param int $colIndex Порядкой номер столбца начальной ячейки
     * @param int $rowIndex Порядковый номер строки начальной ячейки
     * @return array
     */
    private function getRowCellsArray(int $colIndex, int $rowIndex): array
    {
        $columnLetterStartRow = Coordinate::stringFromColumnIndex($colIndex);
        $coordinateStartRow = $columnLetterStartRow . $rowIndex;
        $columnLetterEndRow = Coordinate::stringFromColumnIndex($this->maxColumnIndex);
        $coordinateEndRow = $columnLetterEndRow . $rowIndex;

        return $this->worksheet->rangeToArray("{$coordinateStartRow}:{$coordinateEndRow}");
    }

    /**
     * Проверяет наличие картинки в ячейке, если есть, вставляет её
     *
     * @param Cell $cell
     * @param string|null $imageFilePath
     * @param string|null $imageOptions
     * @throws Exception
     */
    private function insertImage(Cell $cell, ?string $imageFilePath, ?string $imageOptions): void
    {
        if (!empty($imageFilePath)) {
            Image::insertInCell($cell, $imageFilePath, $imageOptions, $this->worksheet);
        }
    }

    /**
     * Скрывает столбцы
     */
    //    private function hideColumns()
    //    {
    //        foreach (array_keys($this->columnsToHide) as $column) {
    //            $this->processWorksheet->getColumnDimensionByColumn($column)->setVisible(false);
    //        }
    //    }

    /**
     * @param int $colIndex Порядкой номер столбца
     * @param int $rowIndex Порядковый номер строки
     * @return Cell|null
     * @return Cell|null
     */
    private function getCell(int $colIndex, int $rowIndex): ?Cell
    {
        return $this->worksheet->getCellByColumnAndRow($colIndex, $rowIndex);
    }

    //////////////////////////////////////////////////////////
    /// //////////////////////////////////////////////////////////
    /// //////////////////////////////////////////////////////////

    /**
     * @param ProcessOptions $options
     * @return ProcessWorksheet
     */
    public function setOptions(ProcessOptions $options): self
    {
        $this->options = $options;

        return $this;
    }

    /**
     * @param Worksheet $worksheet
     * @return ProcessWorksheet
     */
    public function setWorksheet(Worksheet $worksheet): self
    {
        $this->worksheet = $worksheet;

        return $this;
    }

    /**
     * @param SheetTplData $tplData
     * @return ProcessWorksheet
     */
    public function setTplData(SheetTplData $tplData): self
    {
        $this->tplData = $tplData;

        return $this;
    }
}