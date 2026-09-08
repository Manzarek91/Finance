// Ricrea l'API window.storage (get/set/delete/list) usata dall'artifact
// originale, appoggiandosi a localStorage del browser. Essendo un portale
// a uso personale, il parametro "shared" viene ignorato: tutto resta
// esclusivamente nel browser in cui apri il sito.

const PREFIX = "libro-mastro:";

function fullKey(key) {
  return PREFIX + key;
}

window.storage = {
  async get(key) {
    try {
      const raw = localStorage.getItem(fullKey(key));
      if (raw === null) return null;
      return { key, value: raw, shared: false };
    } catch (e) {
      throw new Error("Errore di lettura da localStorage: " + e.message);
    }
  },

  async set(key, value) {
    try {
      localStorage.setItem(fullKey(key), value);
      return { key, value, shared: false };
    } catch (e) {
      throw new Error("Errore di scrittura su localStorage: " + e.message);
    }
  },

  async delete(key) {
    try {
      const existed = localStorage.getItem(fullKey(key)) !== null;
      localStorage.removeItem(fullKey(key));
      return { key, deleted: existed, shared: false };
    } catch (e) {
      throw new Error("Errore di cancellazione da localStorage: " + e.message);
    }
  },

  async list(prefix = "") {
    try {
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(PREFIX + prefix)) {
          keys.push(k.slice(PREFIX.length));
        }
      }
      return { keys, prefix, shared: false };
    } catch (e) {
      throw new Error("Errore di lettura da localStorage: " + e.message);
    }
  },
};
