<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Utils
{
    protected $cols;
    protected $rows;

    public static function appendSheet(Worksheet &$src, Worksheet &$dst)
    {
        $srcColMax = $src->getHighestColumn();
        $srcRowMax = $src->getHighestRow();

        $dstRow = $dst->getHighestRow() + 1;
        $dstCol = 'A';


        // Merge Cells
        foreach ($src->getMergeCells() as $mergeCell) {
            $mc = explode(':', $mergeCell);
            $col_s = preg_replace('/[0-9]*/', '', $mc[0]);
            $col_e = preg_replace('/[0-9]*/', '', $mc[1]);
            $row_s = ((int) preg_replace('/[A-Z]*/', '', $mc[0])) - 1;
            $row_e = ((int) preg_replace('/[A-Z]*/', '', $mc[1])) - 1;

            if (0 <= $row_s && $row_s < $srcRowMax) {
                $merge = $col_s. (string) ($dstRow + $row_s) . ':'. $col_e . (string) ($dstRow + $row_e);
                $dst->mergeCells($merge);
            }
        }

        // Copy data
        $data = $src->rangeToArray('A1:' . $srcColMax.$srcRowMax);
        $dst->fromArray($data, null, $dstCol . $dstRow);

        $colMax = Coordinate::columnIndexFromString($srcColMax);
        $rowMax = $srcRowMax;

        // Copy style
        for ($col = 1; $col <= $colMax; ++$col) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            for ($row = 1; $row <= $rowMax; ++$row) {
                $cellCordStart = $colLetter . $row;
                $cellCordEnd = $colLetter . ($row + $dstRow - 1);
//                echo 'Copy Style from '.$cellCordStart. ' -> '.$cellCordEnd.PHP_EOL;
                $style = $src->getStyle($cellCordStart);
                $dst->duplicateStyle($style, $cellCordEnd);
            }
        }

        // Copy row height
        // Cols sollte identisch wie erste seite sein
        for ($row = 1; $row <= $rowMax; ++$row) {
            $dim = $src->getRowDimension($row, true);
            $dstDim = $dst->getRowDimension($row + $dstRow - 1, true);

            $dstDim->setCollapsed($dim->getCollapsed());
            $dstDim->setOutlineLevel($dim->getOutlineLevel());
            $dstDim->setRowHeight($dim->getRowHeight());
            $dstDim->setRowIndex($dim->getRowIndex());
            $dstDim->setVisible($dim->getVisible());
            $dstDim->setZeroHeight($dim->getZeroHeight());
        }


        // Copy images
        $drawings = $src->getDrawingCollection();
        if (count($drawings) > 0) {
            foreach ($drawings as $drawing) {
                $coords = $drawing->getCoordinates();
                $coords = self::parseCoord($coords);
                $coords = $coords[0].($coords[1] + $dstRow);
                if ($drawing instanceof Drawing) {
                    $drawingCopy = clone $drawing;
                    $drawingCopy->setCoordinates($coords);
                    $drawingCopy->setWorksheet($dst, true);
                }
                if ($drawing instanceof MemoryDrawing) {
                    $drawing->setCoordinates($coords);
                    $drawing->setWorksheet($dst, true);
                }
            }
        }
    }

    public static function copyRows(Worksheet $srcSheet, string $srcFrom, string $srcTo, Worksheet $dstSheet, string $dstFrom)
    {
        $dstFromRow = $dstSheet->getHighestRow();
        $dstFromCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($dstSheet->getHighestColumn());

        $srcToRow = $srcSheet->getHighestRow();
        $srcToCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($srcSheet->getHighestColumn());

        $dstFromPos = ['row' => $dstFromRow, 'col' => $dstFromCol];

        $srcFromPos = ['row' => 1, 'col' => 1];
        $srcToPos = ['row' => $srcToRow, 'col' => $srcToCol];

        self::copyStyle($srcSheet, $dstSheet, $dstFromPos, $srcFromPos, $srcToPos);

        self::mergeCells($srcSheet, $dstFromPos, $dstSheet);

        self::copyContent($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom);
    }

    public static function copyStyle(Worksheet $srcSheet, Worksheet $dstSheet, $dstFromPos, $srcFromPos, $srcToPos)
    {
        for ($col = $srcFromPos['col']; $col <= $srcToPos['col']; ++$col) {
            $colindex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
            for ($row = $srcFromPos['row']; $row <= $srcToPos['row']; ++$row) {
                $style = $srcSheet->getStyleByColumnAndRow($col, $row);
                $dstSheet->duplicateStyle($style, $colindex . ($row + $dstFromPos['row']));
            }
        }
    }

    public static function copyContent(Worksheet $srcSheet, $srcFrom, $srcTo, Worksheet $dstSheet, $dstFrom)
    {
        $cellValues = $srcSheet->rangeToArray($srcFrom . ':' . $srcTo);
        $dstSheet->fromArray($cellValues, null, $dstFrom);
    }

    public static function mergeCells(Worksheet $srcSheet, $dstFromPos, Worksheet $dstSheet)
    {
        $arr = $srcSheet->getMergeCells();
        $a = [];
        foreach ($arr as $key => $val) {
            $keynew = explode(':', $key);
            $keya = $keynew[0];
            $keyb = $keynew[1];

            // get Row
            preg_match_all('!\d+!', $keya, $newcoler);
            $newcoler = $newcoler[0][0];
            $col = str_replace($newcoler, '', $keya);

            $newcol = $newcoler + $dstFromPos['row'];
            $newcola = $col.$newcol;

            unset($newcoler);

            preg_match_all('!\d+!', $keyb, $newcoler);
            $newcoler = $newcoler[0][0];
            $col = str_replace($newcoler, '', $keyb);

            $newcol = (int) $newcoler + (int) $dstFromPos['row'];
            $newcolb = $col.$newcol;

            $valnew = $newcola.':'.$newcolb;

            $a[$valnew] = $valnew;
        }

        $b = $dstSheet->getMergeCells();
        $c = array_merge($a, $b);
        $dstSheet->setMergeCells($c);
    }

    public static function parseCoord($coord) {
        $matches = [];
        preg_match('/[A-Za-z]+/', $coord, $matches);
        $letters = $matches[0];

        $matches = [];
        preg_match('/\d+/', $coord, $matches);
        $digits = $matches[0];

        return [$letters, $digits];
    }
}
