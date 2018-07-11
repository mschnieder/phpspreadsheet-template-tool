<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;


use \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as wrks;
use \PhpOffice\PhpSpreadsheet\Style as style;

Class Style extends
{
	protected $cols;
	protected $rows;

	public function __construct() {

	}

	public static function copyRows(wrks $srcSheet, $srcFrom, $srcTo, wrks $dstSheet, $dstFrom, $dstTo) {
		self::copyContent($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom);

		self::copyStyle($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom, $dstTo);
	}

	public static function copyStyle(wrks $srcSheet, $srcFrom, $srcTo, wrks $dstSheet, $dstFrom, $dstTo) {
		$a = $srcSheet->;

		print_r($a);
//		$a = [];
//		self::copyAlignment($srcStyle->getAlignment()->getStyleArray($a));

		die("HALT!");
	}


	public static function copyContent($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom) {
		$cellValues = $srcSheet->rangeToArray($srcFrom.':'.$srcTo);
		$dstSheet->fromArray($cellValues, NULL, $dstFrom);
	}

	public static function copyAlignment($srcstyle) { // }, $srcFrom, $srcTo, $dstSheet, $dstFrom) {

		print_r($srcstyle);
	}
}