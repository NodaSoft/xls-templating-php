<?php


namespace NodaSoft\PhpXlsTemplating\Helpers;


use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Merge
{

    /**
     * Объединение ячеек в указанной строке по горизонтали
     *
     * @param array $mergeCellsCoordinates координаты ячеек
     * @param int $rowIndex номер строки
     * @param Worksheet $worksheet
     * @throws Exception
     */
    public static function setHorizontal(array $mergeCellsCoordinates, int $rowIndex, Worksheet $worksheet): void
    {
        foreach ($mergeCellsCoordinates as $key => $value) {
            $mergeStartCell = Coordinate::stringFromColumnIndex($value[0]) . $rowIndex;
            $mergeStopCell = Coordinate::stringFromColumnIndex($value[1]) . $rowIndex;
            $worksheet->mergeCells("$mergeStartCell:$mergeStopCell");
        }
    }

    /**
     * Объединение ячеек в указанной строке по вертикали
     *
     * @param array $mergeCellsInRow сведения о вертикальном объединении ячеек в эталонной строке
     * @param int $offset сдвиг координат ячеек по вертикали
     * @param Worksheet $worksheet
     * @throws Exception
     */
    public static function doVerticalWay(array $mergeCellsInRow, int $offset, Worksheet $worksheet): void
    {
        foreach ($mergeCellsInRow as $key => $value) {
            $mergeStartCell = Coordinate::stringFromColumnIndex($value[0]) . ($value[1] + $offset);
            $mergeStopCell = Coordinate::stringFromColumnIndex($value[2]) . ($value[3] + $offset);
            $worksheet->mergeCells("$mergeStartCell:$mergeStopCell");
        }
    }

    /**
     * Объединение ячеек по горизонтали и по вертикали
     *
     * @param int $fromRowIndex строка-образец
     * @param int $toRowIndex строка для объединения по горизонтали
     * @param int $verticalOffset сдвиг координат ячеек по вертикали
     * @param Worksheet $worksheet
     * @throws Exception
     */
    public static function doBothWays(int $fromRowIndex, int $toRowIndex, int $verticalOffset, Worksheet $worksheet): void
    {
        // Запоминаем ячейки, объединенные по горизонтали в строке-образце
        $mergeCellsInRow = self::getHorizontalMergeInRow($fromRowIndex, $worksheet);
        // Объединяем ячейки в новой строке по взятому образцу
        self::setHorizontal($mergeCellsInRow, $toRowIndex, $worksheet);
        // То же с объединением по вертикали
        $mergeColsInRow = self::getVerticalMergeInRow($fromRowIndex, $worksheet);
        self::doVerticalWay($mergeColsInRow, $verticalOffset, $worksheet);
    }

    /**
     * Возвращает полную информацию о вертикальном объединении ячеек в указанной строке
     *
     * @param int $rowIndex
     * @param Worksheet $worksheet
     * @return array
     * @throws Exception
     */
    public static function getVerticalMergeInRow(int $rowIndex, Worksheet $worksheet): array
    {
        $aMergeCells = $worksheet->getMergeCells();
        $mergeColCellsInRow = [];
        foreach ($aMergeCells as $value) {
            $coors = explode(':', $value);
            $start = Coordinate::coordinateFromString($coors[0]);
            $stop = Coordinate::coordinateFromString($coors[1]);
            if ((int)$start[1] === $rowIndex && (int)$stop[1] !== $rowIndex) {
                $mergeColCellsInRow[] = [
                    Coordinate::columnIndexFromString($start[0]), // номер колонки ячейки 1
                    (int)$start[1], // номер строки ячейки 1
                    Coordinate::columnIndexFromString($stop[0]), // номер колонки ячейки 2
                    (int)$stop[1], // номер строки ячейки 2
                ];
            }
        }

        return $mergeColCellsInRow;
    }

    /**
     * Возвращает информацию о горизонтальном объединении ячеек в строке
     *
     * @param int $rowIndex
     * @param Worksheet $worksheet
     * @return array
     * @throws Exception
     */
    public static function getHorizontalMergeInRow(int $rowIndex, Worksheet $worksheet): array
    {
        $aMergeCells = $worksheet->getMergeCells();
        $mergeCellsInRow = [];
        foreach ($aMergeCells as $value) {
            $coors = explode(':', $value);
            $start = Coordinate::coordinateFromString($coors[0]);
            $stop = Coordinate::coordinateFromString($coors[1]);
            if ((int)$start[1] === $rowIndex && (int)$stop[1] === $rowIndex) { // нужная ячейка
                $mergeCellsInRow[] = [
                    Coordinate::columnIndexFromString($start[0]),
                    Coordinate::columnIndexFromString($stop[0]),
                ];
            }
        }

        return $mergeCellsInRow;
    }

    /**
     * Возвращает кол-во строк, объединенных с указанной ячейкой
     *
     * @param int $colIndex Порядкой номер столбца
     * @param int $rowIndex Порядковый номер строки
     * @param Worksheet $worksheet
     * @return int
     * @throws Exception
     */
    public static function getMergedRowsCount(int $colIndex, int $rowIndex, Worksheet $worksheet): int
    {
        $rowsCount = 1;
        foreach ($worksheet->getMergeCells() as $mergedCellCoordinates) {
            $coors = explode(':', $mergedCellCoordinates);
            $start = Coordinate::coordinateFromString($coors[0]);
            $stop = Coordinate::coordinateFromString($coors[1]);
            $startColIndex = Coordinate::columnIndexFromString($start[0]);
            //            $stopColIndex = Coordinate::columnIndexFromString($stop[0]) - 1;

            if ((int)$start[1] === $rowIndex // нужная строка
                && $startColIndex === $colIndex // нужная колонка
                //                && $stopColIndex === $colIndex
            ) {
                $rowsCount = (int)$stop[1] - (int)$start[1] + 1;
                break;
            }
        }

        return $rowsCount;
    }
}