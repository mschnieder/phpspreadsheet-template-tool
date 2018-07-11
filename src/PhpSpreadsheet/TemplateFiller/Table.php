<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

Class Table {
	protected $cols;
	protected $rows;

	public function __construct() {

	}

	static function countTableRows($worksheet, $name, $h, $v) {
		if(strpos($name, "[") === false) return -1;
		$count = 1;
		$v++;
		while($h < 10000) {
			if ($worksheet->getCellByColumnAndRow($h, $v)->getValue() == '') {
				$count++;
			} else {
				$count++;
				return $count;
			}
			$v++;
		}
		return $count;
	}

	static function fill(&$worksheet, $celldata, $data) {
		$a = new Table();

		$a->setValues($worksheet, $celldata, $data);
	}
	
	protected function setValues(&$worksheet, $celldata, $data) {
		$selectedCell = $worksheet->getCellByColumnAndRow($celldata['h'], $celldata['v'][0]);
		$colname = $celldata['variable_blank'];

		$datapos = 0;
		$v = 0;

		foreach($data as $key => $o) {
			if(isset($o->$colname) == false && isset($o[$colname])) {
				$o = (object) $o;
			}
			if(gettype($o->$colname) == 'resource') {
				$this->addImage($worksheet, $o->$colname, $celldata['h'], $celldata['v'][$v]);
			} else {
				$selectedCell->setValue($o->$colname);
			}
			$datapos++;

			if(isset($celldata['v'][$v + 1]) === false)
				break 1;
			$v++;
			$selectedCell = $worksheet->getCellByColumnAndRow($celldata['h'], $celldata['v'][$v]);
		}
		return $worksheet;
	}


	protected function addImage(&$worksheet, $img, $h, $v) {
//  Add the In-Memory image to a worksheet
		$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing();
		$drawing->setName('In-Memory image 1');
		$drawing->setDescription('In-Memory image 1');
		$drawing->setCoordinates($worksheet->getCellByColumnAndRow($h, $v)->getCoordinate());
		$drawing->setImageResource($img);
		$drawing->setRenderingFunction(
			\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::RENDERING_PNG
		);
		$drawing->setMimeType(\PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT);
		$drawing->setWidth(163);
		$drawing->setOffsetX(200);
		$drawing->setWorksheet($worksheet);
		$worksheet->getCellByColumnAndRow($h, $v)->setValue('');
	}
}
?>