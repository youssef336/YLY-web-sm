import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { Member } from '@/domain/entities/member';
import type { TechnicalEvaluation } from '@/domain/entities/technical-evaluation';
import type { FieldVisit } from '@/domain/entities/field-visit';
import type { Meeting } from '@/domain/entities/meeting';
import type { CategoryScores } from '@/domain/entities/category-scores';

/**
 * IndexedDB schema. Every evaluator's device keeps its own private copy of
 * the data (offline-first); no bytes ever leave the browser.
 */
export interface MemberEvalDb extends DBSchema {
  members: {
    key: string;
    value: Member;
  };
  technical: {
    key: string; // memberId (one row per member)
    value: TechnicalEvaluation;
  };
  categoryScores: {
    key: string; // memberId (one row per member)
    value: CategoryScores;
  };
  fieldVisits: {
    key: string;
    value: FieldVisit;
    indexes: { 'by-member': string };
  };
  meetings: {
    key: string;
    value: Meeting;
    indexes: { 'by-member': string };
  };
}

const DB_NAME = 'member-eval-db';
const DB_VERSION = 2;

let dbPromise: Promise<IDBPDatabase<MemberEvalDb>> | null = null;

/** Opens (once) and caches the database connection. */
export function getDb(): Promise<IDBPDatabase<MemberEvalDb>> {
  if (!dbPromise) {
    dbPromise = openDB<MemberEvalDb>(DB_NAME, DB_VERSION, {
      upgrade(db, oldVersion) {
        if (!db.objectStoreNames.contains('members')) {
          db.createObjectStore('members', { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains('technical')) {
          db.createObjectStore('technical', { keyPath: 'memberId' });
        }
        if (!db.objectStoreNames.contains('categoryScores')) {
          db.createObjectStore('categoryScores', { keyPath: 'memberId' });
        }
        if (!db.objectStoreNames.contains('fieldVisits')) {
          const store = db.createObjectStore('fieldVisits', { keyPath: 'id' });
          store.createIndex('by-member', 'memberId');
        }
        if (!db.objectStoreNames.contains('meetings')) {
          const store = db.createObjectStore('meetings', { keyPath: 'id' });
          store.createIndex('by-member', 'memberId');
        }
        // v1 shipped a `tasks` store that no longer exists; drop it on upgrade.
        if (oldVersion < 2 && (Array.from(db.objectStoreNames) as string[]).includes('tasks')) {
          db.deleteObjectStore('tasks' as 'fieldVisits');
        }
      },
    });
  }
  return dbPromise;
}