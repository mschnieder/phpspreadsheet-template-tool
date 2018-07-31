<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * @author bloep
 */

class TemplateParser
{
    const INDEX_ONEPAGER = 0;
    const INDEX_TWOPAGER = 1;
    const INDEX_STARTPAGE = 2;
    const INDEX_MULTIPAGER = 3;
    const INDEX_ENDPAGE = 4;

    const ONEPAGER = 'onepager';
    const TWOPAGER = 'twopager';
    const MULTIPAGER = 'multipager';

    const NAME_ONEPAGER = 'einseitig';
    const NAME_TWOPAGER = 'zweiseitig';
    const NAME_STARTPAGE = 'ersteseite';
    const NAME_MULTIPAGER = 'mittlere';
    const NAME_ENDPAGE = 'letzte';

    const TPL_NORMAL = 0;
    const TPL_ONEPAGEER_ONLY = 1;
    const TPL_NO_TWOPAGER = 2;
    const TPL_NO_MULTIPAGER = 3;

    /** @var array */
    private $variablesTable = [];

    /** @var Spreadsheet */
    private $spreadsheet;

    /** @var Worksheet */
    private $worksheet;

    /** @var array */
    private $variables = [];

    /** @var array */
    private $breakPoints = [];

    /** @var int */
    private $selectedIndex;

    /** @var int */
    private $additionalPages = 0;

    /** @var int */
    private $spreadsheetType;

    /** @var string */
    private $path;

    public function __construct($path)
    {
        $this->path = $path;
        $this->spreadsheet = IOFactory::load($path);
        $this->detectTemplateStructure();
    }

    public function getBreakPoints()
    {
        return $this->breakPoints;
    }

    public function getTemplateMode($rows)
    {
        $breakPoints = $this->getBreakPoints();
        $breakPoints = reset($breakPoints);

        if ($rows <= $breakPoints[self::ONEPAGER]) {
            $this->selectedIndex = self::INDEX_ONEPAGER;
            $this->worksheet = $this->spreadsheet->getSheet(self::INDEX_ONEPAGER);
            return self::ONEPAGER;
        }
        if (($this->hasWorksheetType(self::TWOPAGER) || $this->hasWorksheetType(self::MULTIPAGER)) === false) {
            throw new Exception('Table is too large for the given template and twopager/multipager doesn\'t exists');
        }
        if ($rows <= $breakPoints[self::TWOPAGER]) {
            $this->selectedIndex = self::INDEX_TWOPAGER;
            $this->worksheet = $this->spreadsheet->getSheet(self::INDEX_TWOPAGER);
            return self::TWOPAGER;
        }
        if ($this->hasWorksheetType(self::MULTIPAGER) === false) {
            throw new Exception('Table is too large for the given template and multipager doesn\'t exists');
        }
        $this->selectedIndex = self::INDEX_MULTIPAGER;
        $this->worksheet = $this->spreadsheet->getSheet(self::INDEX_STARTPAGE);
        return self::MULTIPAGER;
    }

    public function parseTemplate()
    {
        $sheetCount = $this->spreadsheet->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            $worksheet = $this->spreadsheet->getSheet($i);
            $this->findVariables('', $worksheet);
        }

        // Jetzt sind alle worksheets eingelesen und die variablen und größen geparsed.

        // If tables exist
        if (count($this->variablesTable) > 0) {
            $values = reset($this->variablesTable);
            foreach ($values as $key => $val) {
                $tableVar = $val['variable'];
                $parsedVar = self::getVariableName($tableVar);
                $this->breakPoints[$parsedVar] = $this->getTableBreakpoints($tableVar);
            }
        }

        // jezt gibt es pro variable die table breakpoints
    }

    public function findVariables($variable, &$worksheet, $vOffset = 1, $hOffset = 1)
    {
        if (!$worksheet) {
            $worksheet = $this->worksheet;
        }
        $title = $worksheet->getTitle();
        if (!isset($this->variables[$title])) {
            $this->variables[$title] = [];
        }
        if (!isset($this->variablesTable[$title])) {
            $this->variablesTable[$title] = [];
        }

        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

        for ($v = $vOffset; $v <= $highestRow; ++$v) {
            for ($h = $hOffset; $h <= $highestColumnIndex; ++$h) {
                $inhalt = $worksheet->getCellByColumnAndRow($h, $v)->getValue();
                if ($variable != '') {
                    if (substr($inhalt, 0, 4) == '§§' && strpos($inhalt, ']§§') === false && strpos($inhalt, ']END§§') === false) {
                        if (strpos($inhalt, $variable) !== false) {
                            $erg = [];

                            $ext = $this->findVariables($variable, $worksheet, $v + 1, 1);
                            if ($ext != false) {
                                $erg[] = $ext;
                                $erg[] = ['h' => $h, 'v' => $v];
                            } else {
                                $erg = ['h' => $h, 'v' => $v];
                            }
                            return $erg;
                        }
                    } elseif (strpos($inhalt, '[') !== false && strpos($inhalt, '§§') !== false && strpos($inhalt, ']END§§') === false) {
                        if (strpos($inhalt, $variable) !== false) {
                            $ext = $this->findVariables($variable, $worksheet, $v + 1, 1);
                            $erg = [];
                            if ($ext != false) {
                                $erg[] = $ext;
                                $erg[] = ['h' => $h, 'v' => $v];
                            } else {
                                $erg = ['h' => $h, 'v' => $v];
                            }
                            return $erg;
                        }
                    }
                } else {
                    if (substr($inhalt, 0, 4) == '§§' && strpos($inhalt, ']§§') === false && strpos($inhalt, ']END§§') === false) {
                        $this->addVariable($title, $inhalt, $h, $v);
                    } elseif (strpos($inhalt, '[') !== false && strpos($inhalt, '§§') !== false && strpos($inhalt, ']END§§') === false) {
                        $this->addVariableTable($worksheet, $inhalt, $h, $v);
                    }
                }
            }
        }
        return false;
    }

    protected function getTableSize($variable, $worksheetIndex)
    {
        $worksheet = $this->spreadsheet->getSheet($worksheetIndex);

        $pos = $this->findVariables($variable, $worksheet);

        if (!is_array($pos) || count($pos) == 0) {
            return 0;
        }

        $size = 0;

        if (isset($pos['h'])) {
            $size = Table::countTableRows($worksheet, $variable, $pos['h'], $pos['v']);
        } else {
            foreach ($pos as $p) {
                $size += Table::countTableRows($worksheet, $variable, $p['h'], $p['v']);
            }
        }
        return $size;
    }

    /**
     * @param Worksheet $worksheet
     * @param string    $variable
     * @param int       $h
     * @param int       $v
     */
    protected function addVariableTable(&$worksheet, $variable, $h, $v)
    {
        $worksheetName = $worksheet->getTitle();

        $countV = $v + 1;
        $verticals = [];
        $verticals[] = $v;
        while ($countV < 1000) {
            if ($worksheet->getCellByColumnAndRow($h, $countV)->getValue() == '') {
                $verticals[] = $countV;
            } else {
                break;
            }
            ++$countV;
        }
        $verticals[] = $countV;

        if (is_array($this->variablesTable[$worksheetName])) {
            foreach ($this->variablesTable[$worksheetName] as $key => $arr) {
                if ($variable == $arr['variable']) {
                    foreach ($verticals as $k => $v1) {
                        $this->variablesTable[$worksheetName][$key]['v'][] = $v1;
                    }
                    return;
                }
            }
        }

        $this->variablesTable[$worksheetName][] = [
            'variable' => $variable,
            'h' => $h,
            'v' => $verticals,
            'tablesize' => Table::countTableRows($worksheet, $variable, $h, $v),
        ];
    }

    protected function addVariable($worksheetName, $variable, $h, $v)
    {
        $this->variables[$worksheetName][] = [
            'variable' => $variable,
            'h' => $h,
            'v' => $v,
        ];
    }

    public static function getVariableName($uncleanvariable)
    {
        return explode('[', str_replace('§§', '', $uncleanvariable))[0];
    }

    public static function getColName($uncleanvariable)
    {
        $colname = explode('[', str_replace('§§', '', $uncleanvariable));
        $colname = str_replace(']END', '', $colname[1]);
        $colname = str_replace(']', '', $colname);
        return $colname;
    }

    private function getTableBreakpoints($variable)
    {
        $d = [
            'onepager' => $this->getTableSize($variable, self::INDEX_ONEPAGER),
        ];
        if ($this->hasWorksheetType(self::TWOPAGER)) {
            $d['twopager'] = $this->getTableSize($variable, self::INDEX_TWOPAGER);
        }
        if ($this->hasWorksheetType(self::MULTIPAGER)) {
            $d['multipager'] = $this->getTableSize($variable, self::INDEX_MULTIPAGER);
        }

        return $d;
    }

    public function hasTable()
    {
        foreach ($this->variablesTable as $sheetName => $value) {
            if (count($value) > 0) {
                return true;
            }
        }
        return false;
    }

    public function getSheetByType($createtype)
    {
        $index = self::getIndexByTypeName($createtype);
        return $this->spreadsheet->getSheet($index);
    }

    public function getVariablesByType($createtype)
    {
        $index = self::getIndexByTypeName($createtype);
        $sheet = $this->spreadsheet->getSheet($index);
        return $this->variables[$sheet->getTitle()];
    }

    public function getVariablesTableByType($createtype)
    {
        $index = self::getIndexByTypeName($createtype);
        $sheet = $this->spreadsheet->getSheet($index);
        return $this->variablesTable[$sheet->getTitle()];
    }

    public static function getIndexByTypeName($type)
    {
        if ($type === self::ONEPAGER) {
            return self::INDEX_ONEPAGER;
        }
        if ($type === self::TWOPAGER) {
            return self::INDEX_TWOPAGER;
        }
        if ($type === self::MULTIPAGER) {
            return self::INDEX_STARTPAGE;
        }
    }

    public static function getTypeNameByIndex($index)
    {
        if ($index === self::INDEX_ONEPAGER) {
            return self::ONEPAGER;
        }
        if ($index === self::INDEX_TWOPAGER) {
            return self::TWOPAGER;
        }
        if ($index === self::INDEX_STARTPAGE || $index === self::INDEX_MULTIPAGER || $index === self::INDEX_ENDPAGE) {
            return self::MULTIPAGER;
        }
    }

    public function appendNeededSheets($tableSize)
    {
        $breakPoints = $this->getBreakPoints();
        $breakPoints = reset($breakPoints);

        $neededSize = $tableSize - $breakPoints[self::TWOPAGER];
        $middleSize = $breakPoints[self::MULTIPAGER];

        $neededSheets = ceil($neededSize / $middleSize);

        $this->additionalPages = $neededSheets;

        $middleSheet = $this->spreadsheet->getSheet(self::INDEX_MULTIPAGER);
        $endSheet = $this->spreadsheet->getSheet(self::INDEX_ENDPAGE);
        for ($i = 0; $i < $neededSheets; ++$i) {
            Utils::appendSheet($middleSheet, $this->worksheet);
        }
        // Append last page
        Utils::appendSheet($endSheet, $this->worksheet);

        $title = $this->worksheet->getTitle();

        $this->variables[$title] = [];
        $this->variablesTable[$title] = [];

        $this->findVariables('', $this->worksheet);
    }

    public function createNewTemplate($tableSize, $worksheetName = null)
    {
        $createtype = $this->getTemplateMode($tableSize);

        if ($createtype === self::MULTIPAGER) {
            $this->appendNeededSheets($tableSize);
        }

        $curIndex = 0;
        while ($this->spreadsheet->getSheetCount() > 1) {
            if ($curIndex === $this->spreadsheet->getIndex($this->worksheet)) {
                ++$curIndex;
            }

            $sheet = $this->spreadsheet->getSheet($curIndex);
            $title = $sheet->getTitle();
            unset($this->variablesTable[$title]);
            unset($this->variables[$title]);
            $this->spreadsheet->removeSheetByIndex($curIndex);
        }

        $this->spreadsheet->setActiveSheetIndex(0);
        $this->worksheet = $this->spreadsheet->getActiveSheet();

        if ($worksheetName) {
            $title = $this->worksheet->getTitle();
            $this->variables[$worksheetName] = $this->variables[$title];
            $this->variablesTable[$worksheetName] = $this->variablesTable[$title];
            unset($this->variablesTable[$title]);
            unset($this->variables[$title]);
            $this->worksheet->setTitle($worksheetName);
        }

        $this->spreadsheet->garbageCollect();
        $this->worksheet->garbageCollect();
    }

    public function getPreparedWorksheet()
    {
        return $this->worksheet;
    }

    public function getPreparedSpreadsheet()
    {
        return $this->spreadsheet;
    }

    public function getCacheKey()
    {
        return TemplateCache::getCacheKey(basename($this->path), self::getTypeNameByIndex($this->selectedIndex), $this->additionalPages);
    }

    public function getPath()
    {
        return $this->path;
    }

    public function getSelectedMode()
    {
        return self::getTypeNameByIndex($this->selectedIndex);
    }

    public function getTotalRows()
    {
        $breakPoints = $this->getBreakPoints();
        $breakPoints = reset($breakPoints);

        return $breakPoints[self::TWOPAGER] + ($breakPoints[self::MULTIPAGER] * $this->additionalPages);
    }

    public function setLogo($path, $header, $position = HeaderFooter::IMAGE_HEADER_LEFT, $width = 90)
    {
        $drawing = new HeaderFooterDrawing();
        $drawing->setName('Logo');
        $drawing->setPath($path);
        $drawing->setWidth($width);
        $this->worksheet->getHeaderFooter()->addImage($drawing, $position);
        $this->worksheet->getHeaderFooter()->setFirstHeader($header);
        $this->worksheet->getHeaderFooter()->setEvenHeader($header);
        $this->worksheet->getHeaderFooter()->setOddHeader($header);
    }

    private function detectTemplateStructure()
    {
        $sheetNames = $this->spreadsheet->getSheetNames();

        $onepagerExists = false;
        $twopagerExists = false;
        $startpageExits = false;
        $multipageExists = false;
        $endpageExists = false;

        foreach ($sheetNames as $name) {
            if ($name === self::NAME_ONEPAGER) {
                $onepagerExists = true;
            }
            if ($name === self::NAME_TWOPAGER) {
                $twopagerExists = true;
            }
            if ($name === self::NAME_STARTPAGE) {
                $startpageExits = true;
            }
            if ($name === self::NAME_MULTIPAGER) {
                $multipageExists = true;
            }
            if ($name === self::NAME_ENDPAGE) {
                $endpageExists = true;
            }
        }

        $multipageComplete = $startpageExits && $multipageExists && $endpageExists;

        if ($onepagerExists && !$twopagerExists && !$multipageComplete) {
            $this->spreadsheetType = self::TPL_ONEPAGEER_ONLY;
        }
        if ($onepagerExists && !$twopagerExists && $multipageComplete) {
            $this->spreadsheetType = self::TPL_NO_TWOPAGER;
        }
        if ($onepagerExists && $twopagerExists && !$multipageComplete) {
            $this->spreadsheetType = self::TPL_NO_MULTIPAGER;
        }
        if ($onepagerExists && $twopagerExists && $multipageComplete) {
            $this->spreadsheetType = self::TPL_NORMAL;
        }
        return $this->spreadsheetType;
    }

    public function hasWorksheetType($type)
    {
        $sheetType = $this->spreadsheetType;
        if ($type === self::ONEPAGER) {
            return true; // Ist aktuell immer dabei
        }
        if ($type === self::TWOPAGER) {
            return $sheetType === self::TPL_NORMAL || $sheetType === self::TPL_NO_MULTIPAGER;
        }
        if ($type === self::MULTIPAGER) {
            return $sheetType === self::TPL_NORMAL || $sheetType === self::TPL_NO_TWOPAGER;
        }
        throw new \InvalidArgumentException('"'.$type.'" is not a valid option');
    }
}
