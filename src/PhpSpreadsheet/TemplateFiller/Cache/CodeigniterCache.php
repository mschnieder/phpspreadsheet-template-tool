<?php

namespace PhpOffice\PhpSpreadsheet\TemplateFiller\Cache;

/**
 * @author bloep
 */
class CodeigniterCache implements \Psr\SimpleCache\CacheInterface
{
    /** @var \CI_Controller */
    private $ci;

    /** @var \CI_Cache */
    private $cache;

    /** @var string */
    private $_cache_path;

    public function __construct(\CI_Controller $controller)
    {
        $this->ci = $controller;
        $this->ci->load->driver('cache', ['adapter' => 'apc', 'backup' => 'file']);
        $this->cache = $this->ci->cache;

        $path = $this->ci->config->item('cache_path');
        $this->_cache_path = ($path === '') ? APPPATH.'cache/' : $path;
    }

    /**
     * {@inheritdoc}
     */
    public function get($key, $default = null)
    {
//        echo '[CACHE] GET '.$key.PHP_EOL;
        if ($r = $this->cache->get($key)) {
            return $r;
        }
        return $default;
    }

    /**
     * {@inheritdoc}
     */
    public function set($key, $value, $ttl = null)
    {
//        echo '[CACHE] SET '.$key.PHP_EOL;
        return $this->cache->save($key, $value, $ttl);
    }

    /**
     * {@inheritdoc}
     */
    public function delete($key)
    {
//        echo '[CACHE] DELETE '.$key.PHP_EOL;
        return $this->cache->delete($key);
    }

    /**
     * {@inheritdoc}
     */
    public function clear()
    {
//        echo '[CACHE] CLEAN'.PHP_EOL;
        return $this->cache->clean();
    }

    /**
     * {@inheritdoc}
     */
    public function getMultiple($keys, $default = null)
    {
        // TODO: Implement getMultiple() method.
    }

    /**
     * {@inheritdoc}
     */
    public function setMultiple($values, $ttl = null)
    {
        // TODO: Implement setMultiple() method.
    }

    /**
     * {@inheritdoc}
     */
    public function deleteMultiple($keys)
    {
        // TODO: Implement deleteMultiple() method.
    }

    /**
     * {@inheritdoc}
     */
    public function has($key)
    {
//        echo '[CACHE] HAS '.$key.PHP_EOL;
        return file_exists($this->_cache_path.$key);
    }
}
