// Standalone replacement for the Claude-artifact-only `window.storage` API.
// Same shape (get/set/delete/list, all async, same return objects) but
// backed by the browser's own localStorage, so it works on any normal
// website with no server or account needed. Data lives only in this
// browser, on this device.

const PREFIX = 'showbasen:';

function read(key) {
  const raw = localStorage.getItem(PREFIX + key);
  return raw === null ? null : raw;
}

window.storage = {
  async get(key) {
    const value = read(key);
    if (value === null) return null;
    return { key, value, shared: false };
  },

  async set(key, value) {
    try {
      localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: false };
    } catch (e) {
      // Most likely quota exceeded — surface it the same way window.storage would.
      return null;
    }
  },

  async delete(key) {
    const existed = read(key) !== null;
    localStorage.removeItem(PREFIX + key);
    return { key, deleted: existed, shared: false };
  },

  async list(prefix = '') {
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(PREFIX)) {
        const bare = k.slice(PREFIX.length);
        if (bare.startsWith(prefix)) keys.push(bare);
      }
    }
    return { keys, prefix, shared: false };
  },
};
