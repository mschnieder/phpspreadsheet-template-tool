<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Drawing;
use PhpOffice\PhpSpreadsheet\Shared\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Table
{
    /**
     * @param Worksheet $worksheet
     * @param string    $name
     * @param int       $h         horizontal index
     * @param int       $v         vertical index
     *
     * @return int
     */
    public static function countTableRows(&$worksheet, $name, $h, $v)
    {
        if (strpos($name, '[') === false) {
            return -1;
        }
        $count = 1;
        ++$v;
        while ($h < 10000) {
            ++$count;
            if (!$worksheet->getCellByColumnAndRow($h, $v)->getValue() == '') {
                return $count;
            }
            ++$v;
        }
        return $count;
    }

    /**
     * @param Worksheet $worksheet
     * @param array     $celldata
     * @param array     $data
     *
     * @return Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function fill(&$worksheet, $celldata, $data)
    {
        return self::setValues($worksheet, $celldata, $data);
    }

    /**
     * @param Worksheet $worksheet
     * @param array     $celldata
     * @param array     $data
     *
     * @return Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function setValues(&$worksheet, $celldata, $data)
    {
        $order = [];
        foreach ($celldata as $tableSet) {
            for ($i = $tableSet['begin']; $i <= $tableSet['end']; ++$i) {
                $order[] = $tableSet['col'].$i;
            }
        }

        $dataIndex = -1;
        foreach ($order as $cellCoord) {
            ++$dataIndex;
            if (!isset($data[$dataIndex])) {
                break;
            }
            $value = $data[$dataIndex];
            if (gettype($value) == 'resource') {
                self::addImageToCell($worksheet, $cellCoord, $value);
            } else {
                $worksheet->getCell($cellCoord)->setValue($value);
            }
        }

        return $worksheet;
    }

    /**
     * @param Worksheet $worksheet
     * @param resource  $img
     * @param int       $h horizontal index
     * @param int       $v vertical index
     * @param int       $width
     * @param int       $height
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function addImage(&$worksheet, $img, $h, $v, $width = 163, $height = 30)
    {
        //  Add the In-Memory image to a worksheet
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName('In-Memory Drawing 2');
        $drawing->setCoordinates($worksheet->getCellByColumnAndRow($h, $v)->getCoordinate());
        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(
            \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_PNG
        );
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
        $drawing->setWidth($width);
        $drawing->setHeight($height);

        $drawing->setWorksheet($worksheet);
        $worksheet->getCellByColumnAndRow($h, $v)->setValue('');
    }

    public static function addImageToCell(Worksheet &$worksheet, $cellCoord, $img, $width = 163, $height = 30)
    {
//        $cell = $worksheet->getCell($cellCoord);
        $a = $worksheet->getMergeCells();

        $width2 = 0;
        $height2 = 0;
        foreach($a as $key => $val) {
            $widthrow = 0;
            $lastcol = 0;
            if(substr($key, 0, strlen($cellCoord)) == $cellCoord) {
                $tmp = Coordinate::extractAllCellReferencesInRange($key);
                foreach($tmp as $key => $coord) {
                    $cords = $worksheet->getCell($coord);
                    if($widthrow == 0) {
                        $widthrow = $cords->getRow();
                    }
                    $aktuelcol = $cords->getColumn();
                    if($lastcol === 0 || $lastcol === $aktuelcol) {
                        $po = $cords->getRow();
                        $tmpheight = $worksheet->getRowDimension($po)->getRowHeight();
                        if($tmpheight == -1) {
                            $font = $worksheet->getParent()->getDefaultStyle()->getFont();
                            $pointRowHeight = Font::getDefaultRowHeightByFont($font);
                            $tmpheight = $pointRowHeight;
                        } else {
                            $tmpheight = Drawing::pointsToPixels($tmpheight);
                        }
                        $height2 += $tmpheight;
                        $lastcol = $aktuelcol;
                    }
                    $co = $cords->getColumn();
                    if ($widthrow == $cords->getRow())
                        $width2 += $worksheet->getColumnDimension($co)->getWidth();
                }
            }
        }

        $width = $width2 * 7.8138144947537;
        $height = $height2;
//        $height = $height2 * 7.8138144947537;

        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
        $drawing->setName('In-Memory Drawing 2');
        $drawing->setCoordinates($cellCoord);

        $drawing->setImageResource($img);
        $drawing->setRenderingFunction(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_PNG);
        $drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
        $drawing->setWidth($width);
        $drawing->setWidthAndHeight($width, $height);

        if($height < Drawing::pointsToPixels($drawing->getHeight())) {
            $drawing->setHeight($height);
        }


        $drawing->setWorksheet($worksheet);
    }
}
