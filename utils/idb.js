import { IDB_META } from "./constants.js";
import { blobToDataUrl } from "./filename.js";

let dbPromise = null;

function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("IDB request failed"));
  });
}

function transactionComplete(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("IDB transaction failed"));
    tx.onabort = () => reject(tx.error || new Error("IDB transaction aborted"));
  });
}

export async function openDatabase() {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(IDB_META.name, IDB_META.version);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(IDB_META.stores.images)) {
        db.createObjectStore(IDB_META.stores.images, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(IDB_META.stores.previews)) {
        db.createObjectStore(IDB_META.stores.previews, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(IDB_META.stores.exportTasks)) {
        db.createObjectStore(IDB_META.stores.exportTasks, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Failed to open IDB"));
  });

  return dbPromise;
}

export async function idbPut(storeName, value) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).put(value);
  await transactionComplete(tx);
}

export async function idbGet(storeName, key) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, "readonly");
  const request = tx.objectStore(storeName).get(key);
  const value = await requestToPromise(request);
  await transactionComplete(tx);
  return value;
}

export async function idbDelete(storeName, key) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).delete(key);
  await transactionComplete(tx);
}

export async function idbClear(storeName) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, "readwrite");
  tx.objectStore(storeName).clear();
  await transactionComplete(tx);
}

export async function idbGetAll(storeName) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, "readonly");
  const request = tx.objectStore(storeName).getAll();
  const rows = await requestToPromise(request);
  await transactionComplete(tx);
  return rows;
}

export function buildImageRecordId(documentUrl, token) {
  return `${documentUrl}_${token}`;
}

export async function getImageDataMap(documentUrl, tokens) {
  const map = new Map();
  const uniqueTokens = [...new Set(tokens || [])];
  for (const token of uniqueTokens) {
    const id = buildImageRecordId(documentUrl, token);
    const row = await idbGet(IDB_META.stores.images, id);
    if (!row) {
      continue;
    }
    if (row.dataUrl && typeof row.dataUrl === "string") {
      map.set(token, row.dataUrl);
      continue;
    }
    if (row.blob instanceof Blob) {
      const dataUrl = await blobToDataUrl(row.blob);
      map.set(token, dataUrl);
      continue;
    }
    if (row.base64Data && typeof row.base64Data === "string") {
      const mimeType = row.mimeType || "application/octet-stream";
      map.set(token, `data:${mimeType};base64,${row.base64Data}`);
    }
  }
  return map;
}

export async function removeImagesByTokens(documentUrl, tokens) {
  const uniqueTokens = [...new Set(tokens || [])];
  for (const token of uniqueTokens) {
    const id = buildImageRecordId(documentUrl, token);
    await idbDelete(IDB_META.stores.images, id);
  }
}
