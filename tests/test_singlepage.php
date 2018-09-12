<?php
ini_set('display_errors', 1);
error_reporting(E_ALL);

use PhpOffice\PhpSpreadsheet\TemplateFiller\Template;
use PhpOffice\PhpSpreadsheet\TemplateFiller\Cache\SimpleCache;

require_once '../vendor/autoload.php';

$data = [
    'asdf' => '|asdf|',
    'test' => '|Test|',
    'zeitraum_von' => '01.01.2018',
    'zeitraum_bis' => '31.12.2018',
    'tab' => [

    ],
    'jahr' => 2020,
];
$entry = [
    'lol' => 'LOL',
    'entry' => 'ENTRY',
    'blub' => 'BLUB',
];

for($i=0;$i<11;++$i) {
    $data['tab'][] = $entry;
}


$cache = new SimpleCache();
$cache->setCacheDir(__DIR__.DIRECTORY_SEPARATOR.'cache'.DIRECTORY_SEPARATOR);
//$cache->clear();

//TemplateCache::setCache($cache);

//TemplateCache::warmup(__DIR__.'/test_file.xlsx', '&L&G&CTestausdruck', './test_logo.png', $entry, $data,'azua',  0, 200);

$template = new Template();
$template->setTemplate('test_singlepage.xlsx', __DIR__);
$template->setWorksheetName('SinglePage');
$template->setData($data);

$template->save('output.xlsx', __DIR__.DIRECTORY_SEPARATOR);
