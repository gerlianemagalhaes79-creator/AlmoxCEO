import { initializeApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { initializeFirestore, memoryLocalCache } from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';

// Clean up any corrupted IndexedDB Firestore databases from previous sessions
if (typeof window !== 'undefined' && window.indexedDB) {
  try {
    if (typeof window.indexedDB.databases === 'function') {
      window.indexedDB.databases().then((databases) => {
        databases.forEach((dbInfo) => {
          if (dbInfo.name && dbInfo.name.includes('firestore')) {
            try {
              window.indexedDB.deleteDatabase(dbInfo.name);
            } catch (_) {}
          }
        });
      }).catch(() => {});
    }
  } catch (_) {}
}

const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(app, {
  localCache: memoryLocalCache()
}, (firebaseConfig as any).firestoreDatabaseId);

export const auth = getAuth(app);

