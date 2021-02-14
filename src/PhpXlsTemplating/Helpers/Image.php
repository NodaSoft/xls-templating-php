<?php


namespace NodaSoft\PhpXlsTemplating\Helpers;


use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Image
{
    /**
     * Вставляет объект изображения в ячейку
     *
     * @param Cell $cell
     * @param string $imageFilePath Путь в файлу картинки
     * @param string $options Строка с опциями картинки
     * @param Worksheet $worksheet
     * @throws Exception
     */
    public static function insertInCell(Cell $cell, string $imageFilePath, string $options, Worksheet $worksheet): void
    {
        $objDrawing = new Drawing();
        $objDrawing
            ->setCoordinates($cell->getCoordinate())
            ->setWorksheet($worksheet)
            ->setPath($imageFilePath);
        self::setImageOptions($objDrawing, $options);
    }

    /**
     * Задает параметры объекта изображения
     *
     * @param Drawing $objDrawing Объект изображения
     * @param string $options Строка с опциями картинки
     */
    private static function setImageOptions(Drawing $objDrawing, string $options): void
    {
        if (empty($options)) {
            return;
        }
        $optionsArray = explode(';', $options);
        for ($i = 1, $iMax = count($optionsArray); $i < $iMax; $i++) {
            switch ($i) {
                case 1:
                    $objDrawing->setName($optionsArray[$i]);
                    break;
                case 2:
                    $objDrawing->setDescription($optionsArray[$i]);
                    break;
                case 3:
                    if (!empty($optionsArray[$i])) {
                        $objDrawing->setHeight($optionsArray[$i]);
                    }
                    break;
                case 4:
                    if (!empty($optionsArray[$i])) {
                        $objDrawing->setWidth($optionsArray[$i]);
                    }
                    break;
                case 5:
                    $objDrawing->setRotation($optionsArray[$i]);
                    break;
                case 6:
                    $objDrawing->setOffsetX($optionsArray[$i]);
                    break;
                case 7:
                    $objDrawing->setOffsetY($optionsArray[$i]);
                    break;
                case 8:
                    $objDrawing->setResizeProportional(($optionsArray[$i] === '1'));
                    break;
                case 9:
                    $objDrawing->getShadow()->setVisible(($optionsArray[$i] === '1'));
                    break;
                case 10:
                    $objDrawing->getShadow()->setDirection($optionsArray[$i]);
                    break;
            }
        }
    }
}