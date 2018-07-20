<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;


use \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Style extends \PhpOffice\PhpSpreadsheet\Style\Style
{
	protected $cols;
	protected $rows;

	public static function copyRows(Worksheet $srcSheet, string $srcFrom, string $srcTo, Worksheet $dstSheet, string $dstFrom, string $dstTo) {
		self::copyContent($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom);
		self::copyStyle($srcSheet, $srcFrom, $srcTo, $dstSheet, $dstFrom, $dstTo);
	}

	public static function copyStyle(Worksheet $srcSheet, string $srcFrom, string $srcTo, Worksheet $dstSheet, string $dstFrom, string $dstTo) {
//		$a = $srcSheet->;

//		print_r($a);
//		$a = [];
//		self::copyAlignment($srcStyle->getAlignment()->getStyleArray($a));

		die("HALT!");
	}


	public static function copyContent(Worksheet $srcSheet, string $srcFrom, string $srcTo, Worksheet $dstSheet, string $dstFrom) {
		$cellValues = $srcSheet->rangeToArray($srcFrom.':'.$srcTo);
		$dstSheet->fromArray($cellValues, NULL, $dstFrom);
	}

	public static function copyAlignment($srcstyle) { // }, $srcFrom, $srcTo, $dstSheet, $dstFrom) {
		print_r($srcstyle);
	}
}
