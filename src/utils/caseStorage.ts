import type {
  AiSectionsData,
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  ScrutinySectionKey,
} from '../types/memo';

const DB_NAME = 'memo_engine_case_db';
const DB_VERSION = 1;
const STORE_NAME = 'case_state';
const ACTIVE_CASE_KEY = 'active_case';

export interface PersistedCaseState {
  id: string;
  updatedAt: string;
  extractedDocs: ExtractedDocumentsState;
  manualInputs: ManualInputsData;
  aiSections: AiSectionsData;
  overrides: Record<string, string>;
  manualEditedSections: Partial<Record<ScrutinySectionKey, boolean>>;
  docUploadedFlags?: Partial<Record<DocumentKey, { uploaded: boolean; fileName: string }>>;
}

function openCaseDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB unavailable'));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Failed to open case storage.'));
  });
}

export async function saveCentralCaseState(
  state: Omit<PersistedCaseState, 'id' | 'updatedAt'>,
  docMetas?: Record<DocumentKey, DocumentUploadMeta>
): Promise<void> {
  const docUploadedFlags: PersistedCaseState['docUploadedFlags'] = {};
  if (docMetas) {
    for (const [k, meta] of Object.entries(docMetas) as Array<[DocumentKey, DocumentUploadMeta]>) {
      docUploadedFlags[k] = {
        uploaded: Boolean(meta.uploaded),
        fileName: meta.fileName || '',
      };
    }
  }

  const record: PersistedCaseState = {
    id: ACTIVE_CASE_KEY,
    updatedAt: new Date().toISOString(),
    extractedDocs: state.extractedDocs,
    manualInputs: state.manualInputs,
    aiSections: state.aiSections,
    overrides: state.overrides,
    manualEditedSections: state.manualEditedSections,
    docUploadedFlags,
  };

  try {
    const db = await openCaseDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(record);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error('Failed to save case data.'));
    });
    db.close();
  } catch {
    try {
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem(DB_NAME, JSON.stringify(record));
      }
    } catch {
      // Ignore storage quota / blocked environment errors
    }
  }
}

export async function loadCentralCaseState(): Promise<PersistedCaseState | null> {
  try {
    const db = await openCaseDb();
    const record = await new Promise<PersistedCaseState | null>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(ACTIVE_CASE_KEY);
      req.onsuccess = () => resolve((req.result as PersistedCaseState | undefined) ?? null);
      req.onerror = () => reject(req.error || new Error('Failed to load case data.'));
    });
    db.close();
    if (record) return record;
  } catch {
    // Fallback to localStorage below
  }

  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(DB_NAME);
      if (raw) return JSON.parse(raw) as PersistedCaseState;
    }
  } catch {
    // Ignore parse/storage errors
  }

  return null;
}
