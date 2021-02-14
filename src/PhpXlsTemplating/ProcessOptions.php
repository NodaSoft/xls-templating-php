<?php


namespace NodaSoft\PhpXlsTemplating;


class ProcessOptions
{
    /**
     * Использовать вставку новой строки при её размножении вместо копирования (вставка работает более медленно)
     *
     * @var bool
     */
    public $insertInsteadOfCopy = true;

    /**
     * Копировать стили строки при её копировании.
     *
     * @var bool
     */
    public $copyStyles = false;

    /**
     * Символ, выделяющий имя переменной в шаблоне
     *
     * @var string
     */
    public $openTag = '{';

    /**
     * Символ, закрывающий имя переменной в шаблоне
     *
     * @var string
     */
    public $closeTag = '}';
}