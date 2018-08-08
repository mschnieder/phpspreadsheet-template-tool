<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller\Cache;

use PhpOffice\PhpSpreadsheet\Exception;
use Psr\SimpleCache\CacheInterface;

/**
 * @author bloep
 */
class SimpleCache implements CacheInterface
{
    private $cacheDir;
    private $useIgbinary = false;

    public function setCacheDir($dir)
    {
        if (file_exists($dir)) {
            if (!is_readable($dir)) {
                throw new Exception('Directory '.$dir. ' is not readable');
            }
            if (!is_writable($dir)) {
                throw new Exception('Directory '.$dir. ' is not writeable');
            }
        } else {
            if (mkdir($dir, 0777, true) === false) {
                throw new Exception('Directory '.dirname($dir). ' is not writeable');
            }
        }

        if (file_exists($dir) && is_readable($dir) && is_writable($dir)) {
            $this->cacheDir = $dir;
            return true;
        }
        
        if(function_exists('igbinary_serialize') && function_exists('igbinary_unserialize')) {
            $this->useIgbinary = true;
        }

        throw new Exception('Directory '.$dir. ' not found and cannot created');
    }

    private function getFilePath($key)
    {
        return $this->cacheDir.$key.'.cache';
    }

    /**
     * {@inheritdoc}
     */
    public function get($key, $default = null)
    {
        $file = $this->getFilePath($key);
        if (file_exists($file)) {
            if($this->useIgbinary) {
                return igbinary_unserialize(file_get_contents($file));
            } else {
                return unserialize(file_get_contents($file));
            }
        }
        return $default;
    }

    /**
     * {@inheritdoc}
     */
    public function set($key, $value, $ttl = null)
    {
        $file = $this->getFilePath($key);
        if($this->useIgbinary) {
            return file_put_contents($file, igbinary_serialize($value));
        } else {
            return file_put_contents($file, serialize($value));
        }
    }

    /**
     * {@inheritdoc}
     */
    public function delete($key)
    {
        $file = $this->getFilePath($key);
        if (file_exists($file)) {
            return unlink($file);
        }
        return true;
    }

    /**
     * {@inheritdoc}
     */
    public function clear()
    {
        $it = new \DirectoryIterator($this->cacheDir);

        /** @var \SplFileInfo $item */
        foreach ($it as $item) {
            if ($item->isFile() && $item->isReadable()) {
                if (unlink($item->getRealPath()) === false) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * {@inheritdoc}
     */
    public function getMultiple($keys, $default = null)
    {
        $results = [];
        foreach ($keys as $key) {
            $results[$key] = $this->get($key, $default);
        }
        return $results;
    }

    /**
     * {@inheritdoc}
     */
    public function setMultiple($values, $ttl = null)
    {
        foreach ($values as $key => $value) {
            if ($this->set($key, $value) === false) {
                return false;
            }
        }
        return true;
    }

    /**
     * {@inheritdoc}
     */
    public function deleteMultiple($keys)
    {
        foreach ($keys as $key) {
            if ($this->delete($key) === false) {
                return false;
            }
        }
        return true;
    }

    /**
     * {@inheritdoc}
     */
    public function has($key)
    {
        $file = $this->getFilePath($key);
        return file_exists($file);
    }
}
