<?php
ini_set('display_errors', 1);
error_reporting(E_ALL);

use PhpOffice\PhpSpreadsheet\TemplateFiller\Template;
use PhpOffice\PhpSpreadsheet\TemplateFiller\Cache\SimpleCache;
use PhpOffice\PhpSpreadsheet\TemplateFiller\TemplateCache;

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
    'unterschriftklient'=> imagecreatefrompng(__DIR__.'/test_logo.png'),
    'datum' => '01.09.2018'
];

$data = [
    'klientenname' => 'Max Mustermann',
    'kvnr' => 'KVNR',
    'klientennr' => 1000,
    'mitarbeiter' => 'Ein Mitarbeiter',
    'abrechnungsmonat' => 'September 2018',
    'azua' => [],
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

for($i=0;$i<100;$i++)
    $data['azua'][] = $entry;


$cache = new SimpleCache();
$cache->setCacheDir(__DIR__.DIRECTORY_SEPARATOR.'cache'.DIRECTORY_SEPARATOR);
//$cache->clear();

TemplateCache::setCache($cache);

//TemplateCache::warmup(__DIR__.'/test_file.xlsx', '&L&G&CTestausdruck', './test_logo.png', $entry, $data,'azua',  0, 200);

$template = new Template();
$template->setTemplate('test_file.xlsx', __DIR__);
$template->setLogo('./test_logo.png', '&L&G&CTestausdruck');
$template->setWorksheetName('Quittierungsbeleg');
$template->setData($data);

$template->save('output.xlsx', __DIR__.DIRECTORY_SEPARATOR);
