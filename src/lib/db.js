// Tiny IndexedDB helper: 'pending' (screenshot upload queue) and 'keys' (encryption key).
const NAME = 'snap-to-excel';
let dbPromise;

function open() {
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open(NAME, 2);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('pending')) db.createObjectStore('pending', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('keys')) db.createObjectStore('keys');
      if (!db.objectStoreNames.contains('handles')) db.createObjectStore('handles');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

async function run(store, mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, mode);
    const req = fn(tx.objectStore(store));
    tx.oncomplete = () => resolve(req ? req.result : undefined);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

export const dbPut = (store, value, key) => run(store, 'readwrite', (o) => o.put(value, key));
export const dbGet = (store, key) => run(store, 'readonly', (o) => o.get(key));
export const dbAll = (store) => run(store, 'readonly', (o) => o.getAll());
export const dbDel = (store, key) => run(store, 'readwrite', (o) => o.delete(key));
export const dbClear = (store) => run(store, 'readwrite', (o) => o.clear());
