<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Document\Security;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;
use Psr\SimpleCache\CacheInterface;

Class Template {
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

    public function __construct() {
		$this->path = '';
		$this->variables = [];
		$this->variablesTable = [];
	}

	/**
	 * @param $filename
	 * @param $path
	 * @throws \PhpOffice\PhpSpreadsheet\Exception
	 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
	 * Set the Filename for the Template
	 */
	public function setTemplate($filename, $path) {
        $this->path = $path."/".$filename;
        $this->templateCache = new TemplateCache();
	}

	public function setWorksheetName($name) {
	    $this->worksheetName = $name;
    }

	public function setData($d){
	    $this->data = $d;

	    if ($this->hasTable($d)) {
	        list($maxRows, $maxRowsVar) = $this->getMaxTableEntries($d);

	        $cachedTemplateKey = $this->templateCache->getCacheTemplateKey(basename($this->path), $maxRows);

	        if($cachedTemplateKey && $this->templateCache->exists($cachedTemplateKey) && !$this->templateCache->isInvalid($cachedTemplateKey, $this->path)) {
	            $this->templateParser = $this->templateCache->loadFromCache($cachedTemplateKey);
	            if (!$this->templateParser) {
	                throw new Exception('Fehler beim lesen der Cache-Datei');
                }
            } else {
	            $this->templateParser = new TemplateParser($this->path);
	            $this->templateParser->parseTemplate();
	            $this->templateParser->createNewTemplate($maxRows, $this->worksheetName);

	            if($this->logo) {
	                call_user_func_array([$this->templateParser, 'setLogo'], $this->logo);
                }

	            $this->templateCache->store($this->templateParser);
            }

            $createtype = self::ONEPAGER;

            $this->variables = $this->templateParser->getVariablesByType($createtype);
            $this->variablesTable = $this->templateParser->getVariablesTableByType($createtype);
            $breakPoints = $this->templateParser->getBreakPoints();
            $this->pagetablesize = $breakPoints[$maxRowsVar];

            $this->worksheet = $this->templateParser->getPreparedWorksheet();
            $this->spreadsheet = $this->templateParser->getPreparedSpreadsheet();
        } else {
	        throw new Exception('not working now');
        }
	    $this->_setData($d);
    }

	public function _setData($d) {
		if($this->templateParser->hasTable()) {
			foreach($this->variablesTable as $key => $celldata) {
				$celldata['variable_blank'] = TemplateParser::getColName($celldata['variable']);
                Table::fill($this->worksheet, $celldata, $d[TemplateParser::getVariableName($celldata['variable'])]);
			}
            $this->fillData($d);
		} else {
		    throw new Exception('noch nicht fertig');
//			$this->findVariables();
//			$this->fillData($d);
		}
		$this->writeVariables();
    }

	private function fillData($d) {
		foreach($this->variables as $key => $val) {
			if(strpos($val['variable'], "[") === false && !is_array($val['v'])) {
                $varname = TemplateParser::getVariableName($val['variable']);
                if (!isset($d[$varname])) {
                	continue;
                }
                if(gettype($d[$varname]) == 'resource') {
                    Table::addImage($this->worksheet, $d[$varname], $val['h'], $val['v'], 163, 500, 30);
                } else {
					$this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($d[$varname]);
                }
			}
		}
	}

	public function setLogo($path, $header, $position = HeaderFooter::IMAGE_HEADER_LEFT, $width = 90) {
	    $this->logo = func_get_args();
    }

    public function setProbeausdruck()
    {
    	//TODO mit spreadsheet noch herausfinden
    }

    protected function writeVariables()
    {
		foreach ($this->variables as $key => $val) {
			if (isset($val['value']))  {
                $this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($val['value']);
            }
		}
	}

	private function cleanup() {
		if(is_array($this->variables) && count($this->variables) > 0) {
            foreach($this->variables as $variable) {
                $cell = $this->worksheet->getCellByColumnAndRow($variable['h'], $variable['v']);
                if(strpos($cell->getValue(), '§§') === 0) {
                    $cell->setValue('');
                }
            }
        }

        if(is_array($this->variablesTable) && count($this->variablesTable) > 0) {
            foreach($this->variablesTable as $variable) {
                $cell = $this->worksheet->getCellByColumnAndRow($variable['h'], $variable['v'][0]);
                if(strpos($cell->getValue(), '§§') === 0) {
                    $cell->setValue('');
                }
                $cell = $this->worksheet->getCellByColumnAndRow($variable['h'], $variable['v'][sizeof($variable['v']) -1]);
                if(strpos($cell->getValue(), '§§') === 0) {
                    $cell->setValue('');
                }
            }
        }
	}

	public function lock($password = null) {
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

	public function save($filename, $path = '') {
		$this->cleanup();

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);
        $writer->setIncludeCharts(true);
        $writer->setPreCalculateFormulas(false);
		$writer->save($path.$filename);
	}

	public function sendToBrowser($filename) {
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
        foreach($data as $varname => $cellData) {
            if(is_array($cellData) && count($cellData) > 0) {
                return true;
            }
        }
        return false;
    }

    private function getMaxTableEntries($d)
    {
        $maxRowsName = '';
        $maxRows = 0;
        foreach($d as $varName => $value) {
            if(is_array($value)) {
                $rows = count($value);
                if($maxRows < $rows) {
                    $maxRows = $rows;
                    $maxRowsName = $varName;
                }
            }
        }
        return [$maxRows, $maxRowsName];
    }
}
