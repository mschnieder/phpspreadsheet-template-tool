<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller;

use PhpOffice\PhpSpreadsheet\Exception;
use Psr\SimpleCache\CacheInterface;

/**
 * @author bloep
 */
class TemplateCache
{
    const CACHE_METADATA = 'metadata';

    /** @var CacheInterface */
    private static $cacheClass;

    /** @var array */
    private $meta;

    public static function setCache(CacheInterface $cacheClass)
    {
        self::$cacheClass = $cacheClass;
    }

    public function getCacheTemplateKey($filename, $maxRows)
    {
        $meta = $this->getTemplateMeta();
        if (self::$cacheClass && isset($meta[$filename])) {
            $fileCache = $meta[$filename];
            $breakPoints = $fileCache['breakpoints'];

            if ($breakPoints[TemplateParser::ONEPAGER] >= $maxRows) {
                return self::getCacheKey($filename, TemplateParser::ONEPAGER, 0);
            }

            if (!isset($breakPoints[TemplateParser::TWOPAGER]) && !isset($breakPoints[TemplateParser::MULTIPAGER])) {
                throw new Exception('Table is too large for the given template and twopager doesn\'t exists');
            }

            if (isset($breakPoints[TemplateParser::TWOPAGER]) && $breakPoints[TemplateParser::TWOPAGER] >= $maxRows) {
                return self::getCacheKey($filename, TemplateParser::TWOPAGER, 0);
            }

            if (!isset($breakPoints[TemplateParser::MULTIPAGER])) {
                throw new Exception('Table is too large for the given template and multipager doesn\'t exists');
            }
            if(isset($breakPoints[TemplateParser::TWOPAGER])) {
                $neededRows = $maxRows - $breakPoints[TemplateParser::TWOPAGER];
            } else {
                $neededRows = $maxRows - ($breakPoints[TemplateParser::NAME_STARTPAGE] + $breakPoints[TemplateParser::NAME_ENDPAGE]);
            }
            $additionalPages = max(0, ceil($neededRows / $breakPoints[TemplateParser::MULTIPAGER]));
            return self::getCacheKey($filename, TemplateParser::MULTIPAGER, $additionalPages);
        }
        return null;
    }

    public function isInvalid($cachedTemplate, $path)
    {
        if (self::$cacheClass && ($cachedTemplate != null)) {
            $meta = $this->getTemplateMeta();
            $filename = basename($path);
            $templateTimestamp = @filemtime($path);
            if (
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
            if ($this->meta) {
                return $this->meta;
            }

            if (self::$cacheClass->has(self::CACHE_METADATA)) {
                $this->meta = self::$cacheClass->get(self::CACHE_METADATA);
                return $this->meta;
            }
        }
        return null;
    }

    public static function getCacheKey($filename, $type, $additionalPages = 0)
    {
        return $filename.'_'.$type.'_'.$additionalPages;
    }

    public function loadFromCache($cachedTemplateKey)
    {
        if (self::$cacheClass) {
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
        if (self::$cacheClass) {
            $path = $templateParser->getPath();
            $filename = basename($path);

            $cacheKey = $templateParser->getCacheKey();
            $templateParser->garbageCollect();
            self::$cacheClass->set($cacheKey, $templateParser);

            $meta = $this->getTemplateMeta();
            if (!$meta) {
                $meta = [];
            }

            if (!isset($meta[$filename])) {
                $breakpoints = $templateParser->getBreakPoints();
                $breakpoints = reset($breakpoints);

                $meta[$filename] = [];
                $meta[$filename]['timestamp'] = @filemtime($path);
                $meta[$filename]['breakpoints'] = $breakpoints;
                $meta[$filename]['cachefiles'] = [];
            } elseif ($meta[$filename]['timestamp'] < @filemtime($path)) {
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
    }

    public function exists($cachedTemplateKey)
    {
        if (self::$cacheClass) {
            if (self::$cacheClass->has($cachedTemplateKey)) {
                return true;
            }
        }
        return false;
    }

    public static function warmup($template, $header, $logo, $testentry, $data, $tablekey, $from = 1, $to = 20)
    {
        ini_set('memory_limit', -1);
        $dir = dirname($template);
        $file = basename($template);

        if (!is_array($tablekey)) {
            $tablekey = [$tablekey];
        }

        foreach ($tablekey as $key) {
            $data[$key] = [];
        }
        echo 'Generating template for rows from '.$from.' to '.$to.PHP_EOL.PHP_EOL;
        $lastFile = '';
        for ($rows = 0; $rows <= $to; ++$rows) {
            foreach ($tablekey as $key) {
                $data[$key][] = $testentry;
            }

            $cache = new self();
            $cachedTemplateKey = $cache->getCacheTemplateKey($file, $rows);
            if ($cachedTemplateKey) {
                if ($lastFile != $cachedTemplateKey) {
                    echo 'Checking '.$cachedTemplateKey.PHP_EOL;
                }
                $lastFile = $cachedTemplateKey;
                if ($cache->exists($cachedTemplateKey)) {
                    if (!$cache->isInvalid($cachedTemplateKey, $template)) {
                        continue;
                    }
                    echo 'Invalid, Regenerating '.$cachedTemplateKey.PHP_EOL.PHP_EOL;
                } else {
                    echo 'Generating '.$cachedTemplateKey.PHP_EOL.PHP_EOL;
                }
            }

            // Trigger generator
            $doc = new Template();
            $doc->setTemplate($file, $dir);
            if($logo && $header) {
                $doc->setLogo($logo, $header);
            }
            $doc->setWorksheetName('Quittierungsbeleg');

            $doc->setData($data);
            unset($doc);
        }
    }
}
