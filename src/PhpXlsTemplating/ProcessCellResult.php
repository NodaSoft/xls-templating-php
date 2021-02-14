<?php


namespace NodaSoft\PhpXlsTemplating;


class ProcessCellResult
{
    /**
     * @var bool
     */
    public $isHideRow = false;

    /**
     * @var ?int
     */
    public $cellHeightLines;

    /**
     * @var ?string
     */
    public $imageFilePath;

    /**
     * @var ?string
     */
    public $imageOptions;
}