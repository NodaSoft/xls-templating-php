<?php


namespace NodaSoft\PhpXlsTemplating;


use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception;

/**
 * One cell processing
 */
class ProcessCell
{
    /**
     * @var Tags
     */
    private $tags;

    /**
     * @var SheetTplData
     */
    private $tplData;

    public function __construct(Tags $tags, SheetTplData $tplData)
    {
        $this->tags = $tags;
        $this->tplData = $tplData;
    }

    /**
     * Process one call - replace all variable and process all tqgs in the cell
     * @param Cell $cell
     * @param array $extTplData
     * @return ProcessCellResult
     * @throws Exception
     */
    public function process(Cell $cell, array $extTplData = []): ProcessCellResult
    {
        $res = new ProcessCellResult();
        if ($this->checkFormula($cell)) {
            return $res; // была формула, дальнейшую обработку пропускаем
        }
        $cellValue = $cell->getValue();
        $res->isHideRow = $this->checkHideRowConditions($cellValue, $extTplData);
        if ($res->isHideRow) {
            $cell->setValue($cellValue);
            return $res; // строку надо скрыть, дальнейшую обработку пропускаем
        }
        // Если найден какой-то шаблон
        if (!$this->tags->matchAll($cellValue, $matches)) {
            return $res;
        }

        $cellValue = $this->checkVarConditions($cellValue, $extTplData);

        [$res->imageFilePath, $res->imageOptions] = $this->checkImages($cellValue, $extTplData);

        $res->cellHeightLines = $this->checkHeight($cellValue);

        // Проверяем, есть ли флаг принудительной установки строкового формата
        $setString = $this->tags->getSetString($cellValue);

        $cellValue = $this->replaceVars($cellValue, $matches[1], $extTplData);

        // Сохраняем значение ячейки
        if ($setString) { // Если надо, принудительно ставим текстовый формат ячейке
            $cell->setValueExplicit($cellValue, DataType::TYPE_STRING);
        } else {
            $cell->setValue($cellValue);
        }

        return $res;
    }

    /**
     * @param string|null $cellValue
     * @param array $vars
     * @param array $extTplData
     * @return string|null
     */
    private function replaceVars(?string $cellValue, array $vars, array $extTplData = []): ?string
    {
        foreach ($vars as $var) {
            $newValue = '';
            if (isset($extTplData[$var])) { // Ищем шаблоны для строки
                $newValue = $extTplData[$var];
            } elseif (isset($this->tplData->data[$var])) { // Ищем общие шаблоны
                $newValue = $this->tplData->data[$var];
            }
            if (is_array($newValue)) {
                $newValue = '';
            }
            $cellValue = $this->simpleReplace($cellValue, $var, $newValue);
        }

        return $cellValue;
    }

    /**
     * @param string|null $cellValue
     * @param string $var
     * @param $newValue
     * @return string|null
     */
    private function simpleReplace(?string $cellValue, string $var, $newValue): ?string
    {
        // Если найдена переменная для замены
        if (!is_null($newValue)) {
            // Заменяем переменную
            $cellValue = $this->tags->replaceVar($var, $newValue, $cellValue);
        }

        return $cellValue;
    }

    /**
     * Проверяет наличие условий видимости переменных и обрабатывает их
     *
     * @param string|null $cellValue
     * @param array $extTplData
     * @return string|null
     */
    private function checkVarConditions(?string $cellValue, array $extTplData = []): ?string
    {
        // Если нет условий в строке
        if (!$this->tags->matchAllConditions($cellValue, $matchesConditions)) {
            return $cellValue;
        }
        foreach ($matchesConditions[1] as $conditionVar) {
            if (!empty($this->tplData->data[$conditionVar])) {
                $cellValue = $this->tags->enableCondition($cellValue, $conditionVar);
            } elseif (!empty($extTplData[$conditionVar])) {
                $cellValue = $this->tags->enableCondition($cellValue, $conditionVar);
            } else {
                $cellValue = $this->tags->disableCondition($cellValue, $conditionVar);
            }
        }

        return $cellValue;
    }

    /**
     * Проверяет наличие формул в ячейке
     *
     * @param Cell $cell
     * @return bool
     * @throws Exception
     */
    private function checkFormula(Cell $cell): bool
    {
        // Если формулы не найдены
        if (!$this->tags->matchFormula($cell->getValue(), $matchesFormulas)) {
            return false;
        }
        $formula = $this->tags->getFormula($cell->getValue());
        if ($formula !== '') {
            // Подставляем в ячейку
            $cell->setValueExplicit($formula, DataType::TYPE_FORMULA);

            return true;
        }

        return false;
    }

    /**
     * Проверяет наличие флага изменения высота в ячейке и если надо, изменяет высоту строки в соответствии с текстом
     *
     * @param string|null $cellValue
     * @return int|null
     */
    private function checkHeight(?string &$cellValue): ?int
    {
        // Если флаг по высоте задан в ячейке
        if ($this->tags->matchHeightTag($cellValue, $matchesCheckFlag)) {
            $lineCount = substr_count($cellValue, "\n") + 1;
            $lineCount = empty($lineCount) ? 1 : $lineCount;
            $lineOneHeight = 14;
            if (!empty($matchesCheckFlag[1][0])) {
                $lineOneHeight = (int)$matchesCheckFlag[1][0];
            }
            $cellValue = $this->tags->clearHeightFlag($cellValue);

            return $lineCount * $lineOneHeight;
        }

        return null;
    }

    /**
     * Проверяет наличие условий в текущей ячейке
     *
     * @param string|null $cellValue
     * @param array $extTplData
     * @return bool
     */
    private function checkHideRowConditions(?string &$cellValue, array $extTplData = []): bool
    {
        // Если есть условие видимости строки
        if ($this->tags->matchShowRowCondition($cellValue, $matchedVariable)) {
            $cellValue = $this->tags->clearShowRowCondition($cellValue);
            $var = $matchedVariable[1];

            // Если переменная в условии пустая, т.е. условие видимости не выполняется
            return empty($this->tplData->data[$var]) && empty($extTplData[$var]);
        }

        return false;
    }

    /**
     * Проверяет наличие картинки в ячейке, если есть, вставляет её
     *
     * @param string|null $cellValue
     * @param array $extTplData
     * @return array
     */
    private function checkImages(?string &$cellValue, array $extTplData = []): array
    {
        $matchesImages = null;
        if ($this->tags->matchImages($cellValue, $matchesImages)) {
            $cellValue = $this->tags->clearImageTags($cellValue);
            if (is_array($matchesImages) && !empty($matchesImages[2][0])) {
                $options = $matchesImages[1][0];
                $imageVarNameTpl = $matchesImages[2][0]; // {SUPER_IMAGE}
                $this->tags->matchAll($imageVarNameTpl, $mathesImg);
                $imageVar = $mathesImg[1][0];
                if (!empty($this->tplData->data[$imageVar])) {
                    $imageFilePath = $this->tplData->data[$imageVar];
                } else if (!empty($extTplData[$imageVar])) {
                    $imageFilePath = $extTplData[$imageVar];
                } else {
                    $imageFilePath = null;
                }
                if (!empty($imageFilePath)) {
                    return [$imageFilePath, $options];
                }
            }
        }

        return [null, null];
    }
}