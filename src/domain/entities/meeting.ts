export type MeetingScore = 0 | 1;

/**
 * A meeting entry.
 *
 * In the SMMEMBER template each of the 15 "MEETINGS" columns (U..AI) is one
 * shared event: its label lives in the header row (row 4) and every member's
 * 0/1 attendance sits in their row under that column. The `slot` (0-14) is the
 * assigned column; the same slot is reused whenever the same date is entered
 * for another member.
 */
export interface Meeting {
  id: string;
  memberId: string;
  /** Meeting date (ISO YYYY-MM-DD). */
  date: string;
  /** 1 = attended, 0 = absent. */
  score: MeetingScore;
  /** Assigned template column index (0-14 -> U..AI). */
  slot: number;
  createdAt: Date;
}

/** Label injected into the column header cell (row 4) for a meeting. */
export function meetingHeaderLabel(meeting: { date: string }): string {
  return meeting.date.trim();
}