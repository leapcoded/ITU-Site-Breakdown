// Database operations module
const DB_NAME = 'TERRA_DB';
const DB_VERSION = 2; // bump DB version to create new store for alerts/rules
export const STORES = { NOTES: 'notes', IMAGES: 'images', FILES: 'files', RULES: 'rules' };

export let db = null;

export function initDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onerror = (e) => reject('DB Error: ' + e.target.errorCode);
        request.onsuccess = (e) => {
            db = e.target.result;
            resolve(db);
        };
        request.onupgradeneeded = (e) => {
            db = e.target.result;
            Object.values(STORES).forEach(store => {
                if (!db.objectStoreNames.contains(store)) {
                    db.createObjectStore(store, { keyPath: 'id', autoIncrement: true });
                }
            });
        };
    });
}

// Ensure the DB connection is open; init if needed
export async function ensureDB() {
    if (db) return db;
    return await initDB();
}

export function dbSave(store, obj) {
    return new Promise(async (resolve, reject) => {
        try {
            await ensureDB();
            const tx = db.transaction([store], 'readwrite');
            const storeObj = tx.objectStore(store);
            const req = (obj && obj.id) ? storeObj.put(obj) : storeObj.add(obj);
            req.onsuccess = () => resolve(req.result);
            req.onerror = (e) => reject(e);
        } catch (e) {
            // If DB connection was closing, try to re-init once and retry the operation
            if (e && e.name === 'InvalidStateError') {
                try {
                    await initDB();
                    const tx = db.transaction([store], 'readwrite');
                    const storeObj = tx.objectStore(store);
                    const req = (obj && obj.id) ? storeObj.put(obj) : storeObj.add(obj);
                    req.onsuccess = () => resolve(req.result);
                    req.onerror = (err) => reject(err);
                } catch (e2) { reject(e2); }
            } else {
                reject(e);
            }
        }
    });
}

export function dbLoadAll(store) {
    return new Promise(async (resolve, reject) => {
        try {
            await ensureDB();
            const tx = db.transaction(store, 'readonly');
            const storeObj = tx.objectStore(store);
            const req = storeObj.getAll();
            req.onsuccess = () => resolve(req.result);
            req.onerror = (e) => reject(e);
        } catch (e) {
            if (e && e.name === 'InvalidStateError') {
                try { await initDB(); const tx = db.transaction(store, 'readonly'); const storeObj = tx.objectStore(store); const req = storeObj.getAll(); req.onsuccess = () => resolve(req.result); req.onerror = (err) => reject(err); } catch (e2) { reject(e2); }
            } else reject(e);
        }
    });
}

export function dbDelete(store, id) {
    return new Promise(async (resolve, reject) => {
        try {
            await ensureDB();
            const tx = db.transaction(store, 'readwrite');
            const storeObj = tx.objectStore(store);
            const req = storeObj.delete(id);
            req.onsuccess = () => resolve();
            req.onerror = (e) => reject(e);
        } catch (e) {
            if (e && e.name === 'InvalidStateError') {
                try { await initDB(); const tx = db.transaction(store, 'readwrite'); const storeObj = tx.objectStore(store); const req = storeObj.delete(id); req.onsuccess = () => resolve(); req.onerror = (err) => reject(err); } catch (e2) { reject(e2); }
            } else reject(e);
        }
    });
}

export function dbClear(store) {
    return new Promise(async (resolve, reject) => {
        try {
            await ensureDB();
            const tx = db.transaction(store, 'readwrite');
            const storeObj = tx.objectStore(store);
            const req = storeObj.clear();
            req.onsuccess = () => resolve();
            req.onerror = (e) => reject(e);
        } catch (e) {
            if (e && e.name === 'InvalidStateError') {
                try { await initDB(); const tx = db.transaction(store, 'readwrite'); const storeObj = tx.objectStore(store); const req = storeObj.clear(); req.onsuccess = () => resolve(); req.onerror = (err) => reject(err); } catch (e2) { reject(e2); }
            } else reject(e);
        }
    });
}

// Remove null/undefined or non-object records from a store. Useful as a migration/cleanup step.
export function cleanNullishRecords(store) {
    return new Promise(async (resolve, reject) => {
        try {
            await ensureDB();
            const tx = db.transaction([store], 'readwrite');
            const storeObj = tx.objectStore(store);
            const req = storeObj.openCursor();
            req.onsuccess = (e) => {
                const cursor = e.target.result;
                if (!cursor) {
                    resolve();
                    return;
                }
                try {
                    const val = cursor.value;
                    // delete entries that are null/undefined or not plain objects
                    if (val == null || typeof val !== 'object') {
                        const del = cursor.delete();
                        del.onsuccess = () => { console.debug('db: removed nullish record in', store); };
                        del.onerror = () => { console.warn('db: failed to delete nullish record in', store); };
                    }
                } catch (err) {
                    // if inspecting the value throws, attempt to delete defensively
                    try { cursor.delete(); } catch (e) { /* ignore */ }
                }
                cursor.continue();
            };
            req.onerror = (e) => reject(e.target ? e.target.error : e);
        } catch (e) {
            if (e && e.name === 'InvalidStateError') {
                try { await initDB(); return cleanNullishRecords(store).then(resolve).catch(reject); } catch (e2) { reject(e2); }
            } else reject(e);
        }
    });
}
