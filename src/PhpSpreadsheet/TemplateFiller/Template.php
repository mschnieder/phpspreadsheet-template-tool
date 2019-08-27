<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter;

class Template
{
    const INDEX_ONEPAGER = 0;
    const INDEX_TWOPAGER = 1;
    const INDEX_MULTIPAGER = 3;

    const ONEPAGER = 'onepager';
    const TWOPAGER = 'twopager';
    const MULTIPAGER = 'multipager';

    /** @var \PhpOffice\PhpSpreadsheet\Spreadsheet */
    protected $spreadsheet;

    /** @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet */
    protected $worksheet;

    /** @var string */
    protected $path;

    /** @var array */
    protected $variablesTable;

    /** @var array */
    protected $pagetablesize;

    /** @var array */
    protected $variables;

    /** @var array */
    protected $data;

    /** @var TemplateParser */
    protected $templateParser;

    /** @var TemplateCache */
    protected $templateCache;

    /** @var string */
    protected $worksheetName;

    /** @var array */
    protected $logo;

    /** @var array */
    private $headerFooter;

    /** @var bool|string */
    private $probeausdruck = false;

    public function __construct()
    {
        $this->path = '';
        $this->variables = [];
        $this->variablesTable = [];
    }

    /**
     * Set the Filename for the Template.
     *
     * @param $filename
     * @param $path
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function setTemplate($filename, $path)
    {
        $this->path = $path.'/'.$filename;
        $this->templateCache = new TemplateCache();
    }

    public function setWorksheetName($name)
    {
        $this->worksheetName = $name;
    }

    public function setData($d)
    {
        $this->data = $d;

        if ($this->hasTable($d)) {
            list($maxRows, $maxRowsVar) = $this->getMaxTableEntries($d);
        } else {
            $maxRows = 0;
            $maxRowsVar = '';
        }

        $cachedTemplateKey = $this->templateCache->getCacheTemplateKey(basename($this->path), $maxRows, $this->probeausdruck);

        if ($cachedTemplateKey && $this->templateCache->exists($cachedTemplateKey) && !$this->templateCache->isInvalid($cachedTemplateKey, $this->path)) {
            $this->templateParser = $this->templateCache->loadFromCache($cachedTemplateKey);
            if (!$this->templateParser) {
                throw new Exception('Fehler beim Lesen der Cache-Datei');
            }
        } else {
            $this->templateParser = new TemplateParser($this->path);
            $this->templateParser->parseTemplate();
            $this->templateParser->createNewTemplate($maxRows, $this->worksheetName);

            if ($this->logo) {
                call_user_func_array([$this->templateParser, 'setLogo'], $this->logo);
            }
            if ($this->probeausdruck) {
                $this->templateParser->setProbeausdruck($this->probeausdruck);
            }

            $this->templateCache->store($this->templateParser);
        }

        $createtype = self::ONEPAGER;

        $this->variables = $this->templateParser->getVariablesByType($createtype);
        $this->variablesTable = $this->templateParser->getVariablesTableByType($createtype);
        $this->headerFooter = $this->templateParser->getHeaderFooterByType($createtype);
        $this->pagetablesize = 0;
        if ($this->hasTable($d)) {
            $breakPoints = $this->templateParser->getBreakPoints();
            if ($maxRows > 0) {
                $this->pagetablesize = $breakPoints[$maxRowsVar];
            }
        }

        $this->worksheet = $this->templateParser->getPreparedWorksheet();
        $this->spreadsheet = $this->templateParser->getPreparedSpreadsheet();
        $this->_setData($d);
    }

    public function _setData($d)
    {
        if ($this->templateParser->hasTable()) {
            foreach ($this->variablesTable as $key => $celldata) {
                $data = $d[TemplateParser::getVariableName($key)];
                $tableKey = TemplateParser::getVariableTableKey($key);
                $data = array_map(function($v) use ($tableKey) {
                    if (is_object($v)) {
                        $v = (array) $v;
                    }
                    return $v[$tableKey] ?? '';
                }, $data);
                Table::fill($this->worksheet, $celldata, $data);
            }
        }
        $this->fillData($d);
        $this->fillHeaderFooter($d);
        $this->writeVariables();
    }

    private function fillData($d)
    {
        foreach ($this->variables as $coord => $val) {
            $coord = str_replace('-', '', $coord);
            $value = $val['raw'];
            foreach($val['vars'] as $i => $varName) {
                if (!isset($d[$varName])) {
                    continue;
                }
                $cellValue = $d[$varName];
                if (gettype($cellValue) == 'resource') {
                    Table::addImageToCell($this->worksheet, $coord, $cellValue);
                } else {
                    $value = str_replace($val['matches'][$i], $cellValue, $value);
                    $this->worksheet->getCell($coord)->setValue($value);
                }
            }
        }
    }

    private function fillHeaderFooter(array $d) {
        if (is_array($this->headerFooter) && count($this->headerFooter) > 0) {
            $worksheetHF = $this->worksheet->getHeaderFooter();
            foreach ($this->headerFooter as $key => $value) {
                $raw = $value['raw'];
                foreach ($value['vars'] as $i => $varName) {
                    $raw = str_replace($value['matches'][$i], $d[$varName], $raw);
                }
                if ($key === 'headerFirst') {
                    $worksheetHF->setFirstHeader($raw);
                    continue;
                }
                if ($key === 'footerFirst') {
                    $worksheetHF->setFirstFooter($raw);
                    continue;
                }
                if ($key === 'headerEven') {
                    $worksheetHF->setEvenHeader($raw);
                    continue;
                }
                if ($key === 'footerEvent') {
                    $worksheetHF->setEvenFooter($raw);
                    continue;
                }
                if ($key === 'headerOdd') {
                    $worksheetHF->setOddHeader($raw);
                    continue;
                }
                if ($key === 'footerOdd') {
                    $worksheetHF->setOddFooter($raw);
                    continue;
                }
            }
        }
    }

    public function setLogo($path, $header, $position = HeaderFooter::IMAGE_HEADER_LEFT, $width = 90)
    {
        $this->logo = func_get_args();
    }

    public function setProbeausdruck($path = false)
    {
        if (!file_exists($path)) {
            throw new Exception('phpspreadsheet: file not found `'.$path.'`');
        }
        $this->probeausdruck = $path;
    }

    protected function writeVariables()
    {
        foreach ($this->variables as $key => $val) {
            if (isset($val['value'])) {
                $this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($val['value']);
            }
        }
    }

    private function cleanup()
    {
        if ($this->worksheetName) {
            $this->worksheet->setTitle($this->worksheetName);
        }

        if (is_array($this->variables) && count($this->variables) > 0) {
            foreach ($this->variables as $coord => $variable) {
                $cell = $this->worksheet->getCell(str_replace('-', '', $coord));
                $cell->setValue(str_replace($variable['raw'], '', $cell->getValue()));
            }
        }

        if (is_array($this->variablesTable) && count($this->variablesTable) > 0) {
            foreach ($this->variablesTable as $variable) {
                foreach ($variable as $tableSet) {
                    $cell = $this->worksheet->getCell($tableSet['col'].$tableSet['begin']);
                    if (strpos($cell->getValue(), '§§') === 0) {
                        $cell->setValue('');
                    }
                    $cell = $this->worksheet->getCell($tableSet['col'].$tableSet['end']);
                    if (strpos($cell->getValue(), '§§') === 0) {
                        $cell->setValue('');
                    }
                }
            }
        }
    }

    public function lock($password = null)
    {
        if ($password) {
            $randomPW = $password;
        } else {
            $randomPW = bin2hex(openssl_random_pseudo_bytes(64));
        }

        $this->spreadsheet->getSecurity()->setLockRevision(true);
        $this->spreadsheet->getSecurity()->setLockStructure(true);
        $this->spreadsheet->getSecurity()->setLockWindows(true);
        $this->spreadsheet->getSecurity()->setWorkbookPassword($randomPW);
        $this->spreadsheet->getSecurity()->setRevisionsPassword($randomPW);
        $sheet = $this->spreadsheet->getActiveSheet();
        $proctection = $sheet->getProtection();

        $proctection->setSheet(true);
        $proctection->setPassword($randomPW);
        $proctection->setAutoFilter(true);
        $proctection->setDeleteColumns(true);
        $proctection->setDeleteRows(true);
        $proctection->setFormatCells(true);
        $proctection->setFormatColumns(true);
        $proctection->setFormatRows(true);
        $proctection->setInsertColumns(true);
        $proctection->setInsertHyperlinks(true);
        $proctection->setInsertRows(true);
        $proctection->setObjects(true);
        $proctection->setPivotTables(true);
        $proctection->setScenarios(true);
        $proctection->setSelectLockedCells(true);
        $proctection->setSelectUnlockedCells(true);
        $proctection->setSort(true);

        $sheet->setPrintGridlines(false);
        $sheet->setShowGridlines(false);
    }

    public function save($filename, $path = '')
    {
        $this->cleanup();

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        $writer->setIncludeCharts(true);
        $writer->setPreCalculateFormulas(false);
        $writer->save($path.$filename);
    }

    public function sendToBrowser($filename)
    {
        $this->cleanup();

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        $writer->setIncludeCharts(true);
        $writer->setPreCalculateFormulas(false);
        $writer->save('php://output');
        exit();
    }

    private function hasTable($data)
    {
        foreach ($data as $varname => $cellData) {
            if (is_array($cellData)) {
                return true;
            }
        }
        return false;
    }

    private function getMaxTableEntries($d)
    {
        $maxRowsName = '';
        $maxRows = 0;
        foreach ($d as $varName => $value) {
            if (is_array($value)) {
                $rows = count($value);
                if ($maxRows < $rows) {
                    $maxRows = $rows;
                    $maxRowsName = $varName;
                }
            }
        }
        return [$maxRows, $maxRowsName];
    }
}
