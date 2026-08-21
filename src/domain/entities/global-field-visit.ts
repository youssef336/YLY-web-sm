/**
 * A global field visit event — independent of any specific member.
 * Created on the Home page; members then reference it by ID when scoring.
 */
export interface GlobalFieldVisit {
  id: string;
  /** Location / event name shown in the Excel header cell. */
  name: string;
  /** Event date (ISO YYYY-MM-DD). */
  date: string;
  createdAt: Date;
}

/** Label injected into the column header cell (row 3) for a field visit. */
export function fieldVisitHeaderLabel(visit: { name: string; date: string }): string {
  return `${visit.name.trim()} - ${visit.date.trim()}`;
}