export type FieldVisitScore = 0 | 1;

/**
 * A field visit entry.
 *
 * In the SMMEMBER template each of the 15 "FIELD VISITS" columns (D..R) is one
 * shared event: its label lives in the header row (row 4) and every member's
 * 0/1 attendance sits in their row under that column. The `slot` (0-14) is the
 * assigned column; the same slot is reused whenever the same event (name +
 * date) is entered for another member, keeping headers and scores aligned.
 */
export interface FieldVisit {
  id: string;
  memberId: string;
  /** Location / event name. */
  name: string;
  /** Event date (ISO YYYY-MM-DD). */
  date: string;
  /** 1 = complete, 0 = incomplete. */
  score: FieldVisitScore;
  /** Assigned template column index (0-14 -> D..R). */
  slot: number;
  createdAt: Date;
}

/** Label injected into the column header cell (row 4) for a field visit. */
export function fieldVisitHeaderLabel(visit: { name: string; date: string }): string {
  return `${visit.name.trim()} - ${visit.date.trim()}`;
}