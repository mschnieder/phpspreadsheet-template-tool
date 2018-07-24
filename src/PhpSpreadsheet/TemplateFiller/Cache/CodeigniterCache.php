<?php
/**
 * @author bloep
 */
class CodeigniterCache implements Psr\SimpleCache\CacheInterface {

    /** @var CI_Controller */
    private $ci;

    /** @var CI_Cache */
    private $cache;
    
    public function __construct(CI_Controller $controller)
    {
        $this->ci = $controller;
    }

    /**
     * {@inheritdoc}
     */
    public function get($key, $default = null)
    {
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
        return $this->cache->save($key, $value, $ttl);
    }

    /**
     * {@inheritdoc}
     */
    public function delete($key)
    {
        return $this->cache->delete($key);
    }

    /**
     * {@inheritdoc}
     */
    public function clear()
    {
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
        return $this->cache->get_metadata($key) != null;
    }
}
