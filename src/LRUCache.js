class LRUCache {
    constructor(capacity = 500) {
        this.cache = new Map();
        this.capacity = capacity;
    }

    clear() {
        this.cache = new Map();
    }

    get(key) {
        if (!this.cache.has(key)) return null;

        let val = this.cache.get(key);

        this.cache.delete(key);
        this.cache.set(key, val);

        return val;
    }

    set(key, value) {
        this.cache.delete(key);

        if (this.cache.size === this.capacity) {
            this.cache.delete(this.cache.keys().next().value);
            this.cache.set(key, value);
        } else {
            this.cache.set(key, value);
        }
    }
}

module.exports = LRUCache;
