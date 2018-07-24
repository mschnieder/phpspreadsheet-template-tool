<?php
ini_set('display_errors', 1);
error_reporting(E_ALL);

$start = microtime(true);

use PhpOffice\PhpSpreadsheet\TemplateFiller\Template;
use PhpOffice\PhpSpreadsheet\TemplateFiller\Cache\SimpleCache;

require_once '../vendor/autoload.php';

$entry = (object) [
    'von' => '12:00',
    'bis' => '12:30',
    'bis_zeit' => '12:30',
    'aufnahmestammdatenid' => '1',
    'betreuungsart' => 'Einzel',
    'betreuer' => 'Ein Mitarbeiter',
    'stunden' => '0.5',
    'stundensatz' => 25.0,
    'summeeuro' => 12.5,
    'unterschriftklient'=> '|Unterschrift|',
    'datum' => '01.09.2018'
];

$tabelle = [];

// <= 20
// <= 55


for($i=0;$i<85;$i++)
    $tabelle[] = $entry;

$data = [
    'klientenname' => 'Max Mustermann',
    'kvnr' => 'KVNR',
    'klientennr' => 1000,
    'mitarbeiter' => 'Ein Mitarbeiter',
    'abrechnungsmonat' => 'September 2018',
    'azua' => $tabelle,
    'datum' => date('d.m.Y'),
    'unterschriftmitarbeiter' => '|Unterschrift|',
    'fahrzeit' => 100,
    'fahrzeitstundensatz' => 100,
    'fahrzeitsumme' => 100,
    'einzel' => 100,
    'einzelstundensatz' => 100,
    'einzelsumme' => 100,
    'gruppen' => 100,
    'gruppenstundensatz' => 100,
    'gruppensumme' => 100,
    'gesamtsumme' => 100,
    'gesamtstunden' => 100,
];


$cache = new SimpleCache();
$cache->setCacheDir(__DIR__.DIRECTORY_SEPARATOR.'cache'.DIRECTORY_SEPARATOR);
//$cache->clear();

\PhpOffice\PhpSpreadsheet\TemplateFiller\TemplateCache::setCache($cache);

$template = new Template();
$template->setTemplate('test_file.xlsx', __DIR__);
$template->setLogo('./test_logo.png', '&L&G&CTestausdruck');
$template->setWorksheetName('Quittierungsbeleg');
$template->setData($data);


$template->save('output.xlsx', __DIR__.DIRECTORY_SEPARATOR);

$ende = microtime(true);

echo $ende - $start.PHP_EOL;
