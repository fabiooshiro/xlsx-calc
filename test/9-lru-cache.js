"use strict";
const LRUCache = require('../src/LRUCache.js');
const assert = require('assert');

describe('LRU cache', () => {
    it('should return null if missing from cache', () => {
        const cache = new LRUCache()
        assert.equal(cache.get('key'), null);
    });

    it('should cache results', () => {
        const cache = new LRUCache()
        cache.set('key', 'value')
        assert.equal(cache.get('key'), 'value');
    });

    it('should remove least recently used if at capacity', () => {
        const cache = new LRUCache(2)
        cache.set('key1', 'value1')
        cache.set('key2', 'value2')
        cache.set('key3', 'value3')

        // assert keys
        assert.equal(cache.get('key1'), null);
        assert.equal(cache.get('key2'), 'value2');
        assert.equal(cache.get('key3'), 'value3');
    });


    it('should update cache when accessed', () => {
        const cache = new LRUCache(2)
        cache.set('key1', 'value1')
        cache.set('key2', 'value2')
        // accessing key 1 to update recently used
        cache.get('key1')

        cache.set('key3', 'value3')

        assert.equal(cache.get('key1'), 'value1');
        assert.equal(cache.get('key2'), null);
        assert.equal(cache.get('key3'), 'value3');
    });

    it('should empty cache when cleared', () => {
        const cache = new LRUCache()
        cache.set('key1', 'value1')
        cache.set('key2', 'value2')
        cache.set('key3', 'value3')

        cache.clear();

        assert.equal(cache.get('key1'), null);
        assert.equal(cache.get('key2'), null);
        assert.equal(cache.get('key3'), null);
    });
})
