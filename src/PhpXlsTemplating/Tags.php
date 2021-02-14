<?php


namespace NodaSoft\PhpXlsTemplating;


/**
 * Обеспечивает работу с шаблонами операторов и переменных
 */
class Tags
{
    /**
     * Символ, выделяющий имя переменной в шаблоне
     *
     * @var string
     */
    private $openTag;

    /**
     * Символ, завершающий имя переменной в шаблоне
     *
     * @var string
     */
    private $closeTag;

    public function __construct(string $openTag, string $closeTag)
    {
        $this->openTag = $openTag;
        $this->closeTag = $closeTag;
    }

    /// //////////////////////////////
    /// IMG
    /// //////////////////////////////

    /**
     * Ищет картинки в ячейке
     *
     * @param string|null $cellValue Содержимое ячейки
     * @param array|null $matches Массив с результатами поиска
     * @return boolean
     */
    public function matchImages(?string $cellValue, ?array &$matches): bool
    {
        return preg_match_all("/{$this->openTag}IMG(.*?){$this->closeTag}(.+?){$this->openTag}\/IMG{$this->closeTag}/i",
            $cellValue, $matches);
    }

    /**
     * Удаляет тег картинки
     *
     * @param string|null $cellValue
     * @return string
     */
    public function clearImageTags(?string $cellValue): ?string
    {
        return preg_replace("/{$this->openTag}IMG.*?{$this->closeTag}(.+?){$this->openTag}\/IMG{$this->closeTag}/i", '',
            $cellValue);
    }

    /// //////////////////////////////
    /// CONDITION IFSHOWROW
    /// //////////////////////////////

    /**
     * Ищет условие видимости строки в ячейке
     *
     * @param string|null $cellValue Содержимое ячейки
     * @param array|null $matches Найденные условия
     * @return boolean
     */
    public function matchShowRowCondition(?string $cellValue, ?array &$matches): bool
    {
        return preg_match("/{$this->openTag}IFSHOWROW (.+?){$this->closeTag}/i", $cellValue, $matches);
    }

    /**
     * Ищет условие видимости строки в ячейке
     *
     * @param string $cellValue Содержимое ячейки
     * @return string
     */
    public function clearShowRowCondition(string $cellValue): ?string
    {
        return preg_replace("/{$this->openTag}IFSHOWROW (.+?){$this->closeTag}/i", '', $cellValue);
    }

    /// //////////////////////////////
    /// CHECK_HEIGHT
    /// //////////////////////////////

    /**
     * Ищет флаг проверки высоты в тексте ячейки
     *
     * @param string|null $cellValue
     * @param array|null $matches
     * @return false|int
     */
    public function matchHeightTag(?string $cellValue, ?array &$matches)
    {
        return preg_match_all("/{$this->openTag}CHECK_HEIGHT=([0-9]+){$this->closeTag}/i", $cellValue, $matches);
    }

    /**
     * Удаляет флаг проверки высоты из текста ячейки
     *
     * @param string|null $cellValue
     * @return string
     */
    public function clearHeightFlag(?string $cellValue): ?string
    {
        return preg_replace("/{$this->openTag}CHECK_HEIGHT=([0-9]+){$this->closeTag}/i", '', $cellValue);
    }

    /// //////////////////////////////
    /// CONDITION IF
    /// //////////////////////////////

    /**
     * Находит все условия в ячейке
     *
     * @param string|null $cellValue Содержимое ячейки
     * @param array|null $matches Найденные условия
     * @return boolean
     */
    public function matchAllConditions(?string $cellValue, ?array &$matches): bool
    {
        return preg_match_all("/{$this->openTag}IF (.+?){$this->closeTag}.+?{$this->openTag}\/IF{$this->closeTag}/i",
            $cellValue, $matches);
    }

    /**
     * Разрешает условие, сохраняет его содержимое, возвращает содержимое ячейки без условия
     *
     * @param string|null $cellValue Содержимое ячейки
     * @param string $conditionVar Имя переменной в условии
     * @return string
     */
    public function enableCondition(?string $cellValue, string $conditionVar): ?string
    {
        $conditionsElse = "/{$this->openTag}IF {$conditionVar}{$this->closeTag}(.*?)({$this->openTag}\/IF{$this->closeTag}|{$this->openTag}ELSE{$this->closeTag}(.*?){$this->openTag}\/IF{$this->closeTag})/i";
        if (preg_match($conditionsElse, $cellValue)) {
            return preg_replace($conditionsElse, '$1', $cellValue);
        }

        return preg_replace("/{$this->openTag}IF {$conditionVar}{$this->closeTag}(.+?){$this->openTag}\/IF{$this->closeTag}/i",
            '$1', $cellValue);
    }

    /**
     * Запрещает условие, удалет его содержимое
     *
     * @param string|null $cellValue Содержимое ячейки
     * @param string|int|bool|null $conditionValue Имя переменной в условии
     * @return string
     */
    public function disableCondition(?string $cellValue, $conditionValue): ?string
    {
        $conditionsElse = "/{$this->openTag}IF {$conditionValue}{$this->closeTag}(.*?)({$this->openTag}\/IF{$this->closeTag}|{$this->openTag}ELSE{$this->closeTag}(.*?){$this->openTag}\/IF{$this->closeTag})/i";
        if (preg_match($conditionsElse, $cellValue)) {
            return preg_replace($conditionsElse, '$3', $cellValue);
        }

        return preg_replace("/{$this->openTag}IF {$conditionValue}{$this->closeTag}(.+?){$this->openTag}\/IF{$this->closeTag}/i",
            '', $cellValue);
    }

    /// //////////////////////////////
    /// FORMULA
    /// //////////////////////////////

    /**
     * Получает значение формулы
     *
     * @param string|null $cellValue
     * @return string
     */
    public function getFormula(?string $cellValue): ?string
    {
        return preg_replace("/{$this->openTag}FORMULA{$this->closeTag}(.+?){$this->openTag}\/FORMULA{$this->closeTag}/i",
            '$1', $cellValue);
    }

    /**
     * Ищем формулы в ячейке
     *
     * @param string|null $cellValue Значение ячейки
     * @param array|null $matches Найденные результаты
     * @return bool
     */
    public function matchFormula(?string $cellValue, ?array &$matches): bool
    {
        return preg_match_all("/{$this->openTag}FORMULA{$this->closeTag}.+?{$this->openTag}\/FORMULA{$this->closeTag}/i",
            $cellValue, $matches);
    }

    /// //////////////////////////////
    /// COMMON
    /// //////////////////////////////

    /**
     * Заменяет шаблон его значением
     *
     * @param string $templateName Имя переменной шаблона
     * @param string|int|float|null $replacementValue На что менять
     * @param string|null $cellValue Содержимое ячейки, в котором производить замену
     * @return string Текст, полученный после замены
     */
    public function replaceVar(string $templateName, $replacementValue, ?string $cellValue): ?string
    {
        $templateName = str_replace('/', '\/', $templateName);

        return preg_replace("/{$this->openTag}{$templateName}{$this->closeTag}/i", $replacementValue, $cellValue);
    }

    /**
     * Выбирает все шаблоны из текста, возвращает 0,\ если ничего не найдено или количество совпадений
     *
     * @param string|null $cellValue
     * @param array|null $matches
     * @return int
     */
    public function matchAll(?string $cellValue, ?array &$matches): int
    {
        return preg_match_all("/{$this->openTag}(.+?){$this->closeTag}/i", $cellValue, $matches);
    }

    /**
     * Проверяет, есть флаг принудительной установки текстового формата ячейке
     *
     * @param string|null $cellValue Содержимое ячейки
     * @return boolean
     */
    public function getSetString(?string &$cellValue): bool
    {
        if (preg_match("/{$this->openTag}SETSTRING{$this->closeTag}/i", $cellValue)) {
            $cellValue = preg_replace("/{$this->openTag}SETSTRING{$this->closeTag}/i", '', $cellValue);

            return true;
        }

        return false;
    }
}