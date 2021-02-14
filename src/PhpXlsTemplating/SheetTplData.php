<?php


namespace NodaSoft\PhpXlsTemplating;


class SheetTplData
{
    /**
     * @var string
     */
    public $sheetName = '';

    /**
     * @var array
     */
    public $data = [];

    public function __construct(string $sheetName, array $tplData)
    {
        $this->data = $tplData;
        $this->sheetName = $sheetName;
    }
}