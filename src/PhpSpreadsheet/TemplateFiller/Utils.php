<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;


use \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use \PhpOffice\PhpSpreadsheet\Style as style;

Class Utils
{
	protected $cols;
	protected $rows;

	public function __construct() {

	}

	public static function copyRows(Worksheet $srcSheet, string $srcFrom, string $srcTo, Worksheet $dstSheet, string $dstFrom, string $dstTo) {
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

	public static function copyStyle(Worksheet $srcSheet, Worksheet $dstSheet, $dstFromPos, $srcFromPos, $srcToPos) {

		for($col = $srcFromPos['col']; $col <= $srcToPos['col']; $col++) {
			$colindex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col);
			for($row = $srcFromPos['row']; $row <= $srcFromPos['row']; $row++) {
				$style = $srcSheet->getStyleByColumnAndRow($col, $row);
				$dstSheet->duplicateStyle($style, $colindex.($row + $dstFromPos['row']));
			}
		}
//		$dstSheet->duplicateConditionalStyle($style, 'A39');
//		$a = [];
//		self::copyAlignment($srcStyle->getAlignment()->getStyleArray($a));

//		die("HALT!");
	}


	public static function copyContent(Worksheet $srcSheet, $srcFrom, $srcTo, Worksheet $dstSheet, $dstFrom) {
		$cellValues = $srcSheet->rangeToArray($srcFrom.':'.$srcTo);
		$dstSheet->fromArray($cellValues, NULL, $dstFrom);
	}

	public static function copyAlignment($srcstyle) { // }, $srcFrom, $srcTo, $dstSheet, $dstFrom) {
		print_r($srcstyle);
	}

	public static function mergeCells(Worksheet $srcSheet, $dstFromPos, Worksheet $dstSheet) {
		$arr = $srcSheet->getMergeCells();
		$a = [];
		foreach($arr as $key => $val) {
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


			$newcol = intval($newcoler) + intval($dstFromPos['row']);
			$newcolb = $col.$newcol;

			$valnew = $newcola.':'.$newcolb;


			$a[$valnew] = $valnew;
		}

		$b = $dstSheet->getMergeCells();
		$c = array_merge($a, $b);
		$dstSheet->setMergeCells($c);
	}
}
