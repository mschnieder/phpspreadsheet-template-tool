<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Document\Security;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Protection;

Class Template {
    /** @var \PhpOffice\PhpSpreadsheet\Spreadsheet */
	protected $spreadsheet;

    /** @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet */
	protected $worksheet;
	protected $path;
	
	protected $finalSpreadsheet;
	protected $finalWorksheet;
	protected $finalRowPos;

	protected $variables;
	protected $variablesTable;

	public function __construct() {
		$this->pfad = '';
		$this->variables = [];
		$this->variablesTable = '';

		$this->finalSpreadsheet =  new \PhpOffice\PhpSpreadsheet\Spreadsheet();
		$this->finalWorksheet = $this->finalSpreadsheet->getActiveSheet();
		$this->finalRowPos = 1;
	}

	/**
	 * @param $filename
	 * @param $path
	 * @throws \PhpOffice\PhpSpreadsheet\Exception
	 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
	 * Set the Filename for the Template
	 */
	public function setTemplate($filename, $path) {
		$this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path."/".$filename);
		$this->worksheet = $this->spreadsheet->getSheet(0);
	}

	public function setData($d) {
		$this->findVariables();

		// if Tables existing
		if($this->variablesTable != '') {
			$createtype = '';
			foreach($this->variablesTable as $key => $val) {
				$createtype = $this->checkType(count($d[$this->getVariableName($val['variable'])]), $val['variable']);

				// Take the Template by Data size
				switch($createtype) {
					case 'onepager':
						$this->worksheet = $this->spreadsheet->getSheet(0);
						$this->variables = [];
						$this->variablesTable = [];
						$this->findVariables();
						break;
					case 'twopager':
						$this->worksheet = $this->spreadsheet->getSheet(1);
						$this->variables = [];
						$this->variablesTable = [];
						$this->findVariables();
						break;
					case 'multipager':
						$this->worksheet = $this->spreadsheet->getSheet(2);
						$this->variables = [];
						$this->variablesTable = [];
						$this->findVariables();
						break;
				}
				break;
			}

			// Fill Table in a worksheet by createtype
			$first = true;
			foreach($this->variablesTable as $key => $celldata) {
				$celldata['variable_blank'] = $this->getColName($celldata['variable']);
				switch($createtype) {
					case 'onepager':
						$this->worksheet = $this->spreadsheet->getSheet(0);

						table::fill($this->worksheet, $celldata, $d[$this->getVariableName($celldata['variable'])]);
						break;
					case 'twopager':
						$this->worksheet = $this->spreadsheet->getSheet(1);

						table::fill($this->worksheet, $celldata, $d[$this->getVariableName($celldata['variable'])]);
						break;
					case 'multipager':
						$this->worksheet = $this->spreadsheet->getSheet(2);
						if($first == true)
							$this->copySheetAtLeast($this->spreadsheet->getSheet(3));

						$first = false;
//						$this->findVariables();

//						die("MULTIPAGER");
						break;
				}
			}
            $this->fillData($d);
		} else {
			$this->findVariables();
			$this->fillData($d);
		}
//		$this->writeVariables();
	}

	private function copySheetAtLeast($copysheet) {
		$highestRow = $copysheet->getHighestRow();
		$highestColumn = $copysheet->getHighestColumn();

		Utils::copyRows($copysheet, 'A1', 'AU37', $this->worksheet, 'A39', 'AU76');
	}

	public function copyRows($sheet,$srcRow,$dstRow,$height,$width) {

		for ($row = 0; $row < $height; $row++) {
			for ($col = 0; $col < $width; $col++) {
				$cell = $sheet->getCellByColumnAndRow($col, $srcRow + $row);
				$style = $sheet->getStyleByColumnAndRow($col, $srcRow + $row);
				$dstCell = $sheet->getCellByColumnAndRow($col, ($dstRow + $row))->getValue();
				$sheet->setCellValue($dstCell, $cell->getValue());
				$sheet->duplicateStyle($style, $dstCell);
			}

			$h = $sheet->getRowDimension($srcRow + $row)->getRowHeight();
			$sheet->getRowDimension($dstRow + $row)->setRowHeight($h);
		}

		foreach ($sheet->getMergeCells() as $mergeCell) {
			$mc = explode(":", $mergeCell);
			$col_s = preg_replace("/[0-9]*/", "", $mc[0]);
			$col_e = preg_replace("/[0-9]*/", "", $mc[1]);
			$row_s = ((int)preg_replace("/[A-Z]*/", "", $mc[0])) - $srcRow;
			$row_e = ((int)preg_replace("/[A-Z]*/", "", $mc[1])) - $srcRow;

			if (0 <= $row_s && $row_s < $height) {
				$merge = $col_s . (string)($dstRow + $row_s) . ":" . $col_e . (string)($dstRow + $row_e);
				$sheet->mergeCells($merge);
			}
		}
	}


	private function fillData($d) {
		foreach($this->variables as $key => $val) {
			if(strpos($val['variable'], "[") === false && !is_array($val['v'])) {
                if(gettype($d[$this->getVariableName($val['variable'])]) == 'resource') {
                    Table::addImage($this->worksheet, $d[$this->getVariableName($val['variable'])], $val['h'], $val['v'], 163, 500, 30);
                } else {
                    $this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($d[$this->getVariableName($val['variable'])]);
                }
			}
		}
	}

	public function findVariables($variable = '', $worksheet = '', $vOffset = 1, $hOffset = 1) {
		if($worksheet == '')
			$worksheet = $this->worksheet;

		$highestRow = $worksheet->getHighestRow();
		$highestColumn = $worksheet->getHighestColumn();
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

		for($v = $vOffset; $v <= $highestRow; $v++) {
			for($h = $hOffset; $h <= $highestColumnIndex; $h++) {
				if($variable != '') {
					$inhalt = $worksheet->getCellByColumnAndRow($h, $v)->getValue();
					if (substr($inhalt, 0, 4) == '§§' && strpos($inhalt, ']§§') === false && strpos($inhalt, ']END§§') === false) {
						if(strpos($inhalt, $variable) !== false) {
							$erg = '';

							$ext = $this->findVariables($variable, $worksheet, $v + 1, 1);
							if ($ext != false) {
								$erg[] = $ext;
								$erg[] = ['h' => $h, 'v' => $v];
							} else {
								$erg = ['h' => $h, 'v' => $v];
							}
							return $erg;
						}
					} else if (strpos($inhalt, '[') !== false && strpos($inhalt, '§§') !== false && strpos($inhalt, ']END§§') === false) {
						if(strpos($inhalt, $variable) !== false) {
							$erg = '';
							$ext = $this->findVariables($variable, $worksheet, $v + 1, 1);
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
					$inhalt = $worksheet->getCellByColumnAndRow($h, $v)->getValue();
					if (substr($inhalt, 0, 4) == '§§' && strpos($inhalt, ']§§') === false && strpos($inhalt, ']END§§') === false) {
						$this->addVariable($inhalt, $h, $v);
					} else if (strpos($inhalt, '[') !== false && strpos($inhalt, '§§') !== false && strpos($inhalt, ']END§§') === false) {
						$this->addVariableTable($inhalt, $h, $v);
					}
				}
			}
		}
		return false;
	}

	protected function checkType($dataCount, $variable) {
		$this->pagetablesize['onepager'] = $this->getTableSize($variable, 0);
		$this->pagetablesize['twopager'] = $this->getTableSize($variable, 1);
		$this->pagetablesize['multipager'] = $this->getTableSize($variable, 3);


		if($dataCount <= $this->pagetablesize['onepager']) {
			return 'onepager';
		} else if($dataCount <= $this->pagetablesize['twopager']) {
			return 'twopager';
		} else {
			return 'multipager';
		}
	}
	
	protected function getTableSize($variable, $worksheet) {
		$pos = $this->findVariables($variable, $this->spreadsheet->getSheet($worksheet));
		$size = 0;

		if(isset($pos['h'])) {
			$size = table::countTableRows($this->spreadsheet->getSheet($worksheet), $variable, $pos['h'], $pos['v']);
		} else {
			foreach($pos as $p) {
				$size += table::countTableRows($this->spreadsheet->getSheet($worksheet), $variable, $p['h'], $p['v']);
			}
		}
		return $size;
	}

	protected function addVariableTable($variable, $h, $v) {
		$countV = $v + 1;
		$verticals = [];
		$verticals[] = $v;
		while($countV < 1000) {
			if($this->worksheet->getCellByColumnAndRow($h, $countV)->getValue() == '') {
				$verticals[] = $countV;
			} else {
				break 1;
			}
			$countV++;
		}
		$verticals[] = $countV;

		if(is_array($this->variablesTable))
		foreach($this->variablesTable as $key => $arr) {
			if($variable == $arr['variable']) {
				foreach($verticals as $k => $v)
					$this->variablesTable[$key]['v'][] = $v;
				
				return;
			}
		}

		$var = ['variable' => $variable,
				'h' => $h,
				'v' => $verticals,
				'tablesize' => table::countTableRows($this->worksheet, $variable, $h, $v)];

		$this->variablesTable[] = $var;
	}

	protected function addVariable($variable, $h, $v) {
		$var = ['variable' => $variable,
				'h' => $h,
				'v' => $v];

		$this->variables[] = $var;
	}

	protected function getVariableName($uncleanvariable) {
		$colname = explode('[', str_replace('§§', '', $uncleanvariable));
		$vname = $colname[0];
		return $vname;
	}

	protected function getColName($uncleanvariable) {
		$colname = explode('[', str_replace('§§', '', $uncleanvariable));
		$colname = str_replace(']END', '', $colname[1]);
		$colname = str_replace("]", "", $colname);
		return $colname;
	}

    public function setProbeausdruck()
    {
    	//TODO mit spreadsheet noch herausfinden
    }

    protected function writeVariables() {
		foreach($this->variables as $key => $val) {
			if(isset($val['value']))
			$this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($val['value']);
		}
	}

	private function _lastChanges() {
        foreach($this->variables as $variable) {
            $cell = $this->worksheet->getCellByColumnAndRow($variable['h'], $variable['v']);
            if(strpos($cell->getValue(), '§§') === 0) {
                $cell->setValue('');
            }
        }

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


//		$class = \PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf::class;
//		\PhpOffice\PhpSpreadsheet\IOFactory::registerWriter('Pdf', $class);
//		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Pdf');
//		$writer->save($path.'test.pdf');
//


        $sheetCount = $this->spreadsheet->getSheetCount();
        for($i=$sheetCount-1;$i>=0;$i--) {
            if($this->spreadsheet->getActiveSheetIndex() != $i) {
                $this->spreadsheet->removeSheetByIndex($i);
            } else {
                $this->spreadsheet->getActiveSheet()->setTitle('Quittierungsbeleg');
            }
        }

        $this->spreadsheet->getSecurity()->setLockRevision(true);
        $this->spreadsheet->getSecurity()->setLockStructure(true);
        $this->spreadsheet->getSecurity()->setLockWindows(true);
        $this->spreadsheet->getSecurity()->setWorkbookPassword('MEMORIA_UNLOCK_PASSWORD');
        $this->spreadsheet->getSecurity()->setRevisionsPassword('MEMORIA_UNLOCK_PASSWORD');
        $this->spreadsheet->getActiveSheet()->getProtection()->setSheet(true);
        $this->spreadsheet->getActiveSheet()->getProtection()->setPassword('MEMORIA_UNLOCK_PASSWORD');
	}

	public function save($filename, $path = '', $filetype = 'Xls') {
		$this->_lastChanges();
		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xls');
		$writer->save($path.$filename);
	}

	public function sendToBrowser($filename) {
        $this->_lastChanges();

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xls');
        $writer->save('php://output');
        exit();
	}
}
?>
