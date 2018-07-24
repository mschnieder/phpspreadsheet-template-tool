<?php
namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use Psr\SimpleCache\CacheInterface;

/**
 * @author bloep
 */
class TemplateCache {
    const CACHE_METADATA = 'metadata';

    /** @var CacheInterface */
    private static $cacheClass;

    /** @var array */
    private $meta;

    public static function setCache(CacheInterface $cacheClass) {
        self::$cacheClass = $cacheClass;
    }

    public function getCacheTemplateKey($filename, $maxRows)
    {
        $meta = $this->getTemplateMeta();
        if(self::$cacheClass && isset($meta[$filename])) {
            $fileCache = $meta[$filename];
            $breakPoints = $fileCache['breakpoints'];

            $additionalPages = 0;
            if($breakPoints[TemplateParser::ONEPAGER] > $maxRows) {
                $selectedType = TemplateParser::ONEPAGER;
            } elseif ($breakPoints[TemplateParser::TWOPAGER] > $maxRows) {
                $selectedType = TemplateParser::TWOPAGER;
            } else {
                $selectedType = TemplateParser::MULTIPAGER;
                $neededRows = $maxRows - $breakPoints[TemplateParser::TWOPAGER];
                $additionalPages = ceil($neededRows / $breakPoints[TemplateParser::MULTIPAGER]);
            }

            return self::getCacheKey($filename, $selectedType, $additionalPages);
        }
        return null;
    }

    public function isInvalid($cachedTemplate, $path)
    {
        if (self::$cacheClass && ($cachedTemplate != null)) {
            $meta = $this->getTemplateMeta();
            $filename = basename($path);
            $templateTimestamp = @filemtime($path);
            if(
                $meta[$filename]['timestamp'] >= $templateTimestamp ||
                $meta[$filename]['cachefiles'][$cachedTemplate] >= $templateTimestamp
            ) {
                return false;
            }
        }
        return true;
    }

    private function getTemplateMeta()
    {
        if (self::$cacheClass) {
            if($this->meta) {
                return $this->meta;
            }

            if (self::$cacheClass->has(self::CACHE_METADATA)) {
                return self::$cacheClass->get(self::CACHE_METADATA);
            }
        }
        return null;
    }

    public static function getCacheKey($filename, $type, $additionalPages = 0) {
        return $filename.'_'.$type.'_'.$additionalPages;
    }

    public function loadFromCache($cachedTemplateKey)
    {
        if(self::$cacheClass) {
            return self::$cacheClass->get($cachedTemplateKey);
        }
        return null;
    }

    /**
     * @param TemplateParser $templateParser
     *
     * @throws \Psr\SimpleCache\InvalidArgumentException
     */
    public function store($templateParser)
    {
        if(self::$cacheClass) {
            $path = $templateParser->getPath();
            $filename = basename($path);
            $type = $templateParser->getSelectedMode();

            $cacheKey = $templateParser->getCacheKey();
            self::$cacheClass->set($cacheKey, $templateParser);

            $meta = $this->getTemplateMeta();

            if(!isset($meta[$filename])) {
                $breakpoints = $templateParser->getBreakPoints();
                $breakpoints = reset($breakpoints);

                $meta[$filename] = [];
                $meta[$filename]['timestamp'] = @filemtime($path);
                $meta[$filename]['breakpoints'] = $breakpoints;
                $meta[$filename]['cachefiles'] = [];
            } else if($meta[$filename]['timestamp'] < @filemtime($path)) {
                $breakpoints = $templateParser->getBreakPoints();
                $breakpoints = reset($breakpoints);

                $meta[$filename] = [];
                $meta[$filename]['timestamp'] = @filemtime($path);
                $meta[$filename]['breakpoints'] = $breakpoints;
                $meta[$filename]['cachefiles'] = [];
            }
            $meta[$filename]['cachefiles'][$cacheKey] = time();

            self::$cacheClass->set(self::CACHE_METADATA, $meta);
        }
/*
        $meta = [
            'test_file.xlsx' => [
                'timestamp' => time(),
                'breakpoints' => 'Normale Breakpoints',
                'cachefiles' => [
                    'test_file.xlsx_multipager_1' => 'timestamp',
                    'test_file.xlsx_multipager_1' => 'timestamp',
                    'test_file.xlsx_multipager_1' => 'timestamp',
                    'test_file.xlsx_multipager_1' => 'timestamp',
                    'test_file.xlsx_multipager_1' => 'timestamp',
                ]
            ]
        ];
*/
    }
}
