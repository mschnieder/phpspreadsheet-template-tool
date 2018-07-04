<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

Class Template {
	protected $spreadsheet;
	protected $worksheet;
	protected $path;
	protected $maxcols;
	protected $maxrows;

	protected $variables;

	public function __construct() {
		parent::__construct();
		$this->pfad = '';
		$this->maxcols = 47;
		$this->maxrows = 50;
		$this->variables = [];
	}

	/**
	 * @param $filename
	 * @param $path
	 * @throws \PhpOffice\PhpSpreadsheet\Exception
	 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
	 * Set the Filename for the Template
	 */
	public function setTemplate($filename, $path) {
		$this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path.$filename);
		$this->worksheet = $this->spreadsheet->getActiveSheet();
	}

	public function findVariables() {
		for($v = 1; $v <= $this->maxrows; $v++) {
			for($h = 1; $h <= $this->maxcols; $h++) {
				$inhalt = $this->worksheet->getCellByColumnAndRow($h, $v)->getValue();
				if(substr($inhalt, 0, 4) == '§§') {
					$this->addVariable($inhalt, $h, $v);
				}
			}
		}
	}

	protected function addVariable($variable, $h, $v) {
		$this->variables[] = ['variable' => $variable, 'h' => $h, 'v' => $v];
	}

	public function setVariables($d) {
		echo "<pre>";
//		print_r($d);
		foreach($this->variables as $key => $val) {
			$tmp = str_replace('§', '', $val['variable']);

			$tmp = explode('[', $tmp);
			$tmp = $tmp[0];

			if(!is_array($d[$tmp])) {
				$this->variables[$key]['value'] = $d[$tmp];
			} else {
//				print_r($d);
				$h = $val['h'];
				$v = $val['v'];
				$a = $this->worksheet->getCellByColumnAndRow($h, $v);
				unset($this->variables[$key]);

				$colname = explode('[', str_replace('§§', '', $val['variable']));
				$vname = $colname[0];
				$colname = str_replace(']', '', $colname[1]);

				$anzahl = sizeof($d[$vname]);

				foreach($d[$vname] as $key => $o) {
					$stopit = false;
					if(strpos($a->getValue(), ']END') !== false) {
						unset($this->variables[$a->getValue()]);
						$stopit = true;
					}

					$a->setValue($o->$colname);

					if($stopit == true)
						break 2;

					$v++;
					$a = $this->worksheet->getCellByColumnAndRow($h, $v);
				}


//				while($a->getValue() != '') {
//					$a->setValue($d['variable'])
//					$v++;
//					if($v > 50) break;
//				}
			}

		}
		$this->writeVariables();
	}

	protected function writeVariables() {
		foreach($this->variables as $key => $val) {
			$this->worksheet->getCellByColumnAndRow($val['h'], $val['v'])->setValue($val['value']);
		}
	}

	public function setData($data) {

	}

	public function save() {
		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xls');
		$writer->save('writeit.xls');
	}
}
?>