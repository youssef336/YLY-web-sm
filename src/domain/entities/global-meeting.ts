/**
 * A global meeting event — independent of any specific member.
 * Created on the Home page; members then reference it by ID when scoring.
 */
export interface GlobalMeeting {
  id: string;
  /** Meeting date (ISO YYYY-MM-DD), also used as the Excel header label. */
  date: string;
  createdAt: Date;
}

/** Label injected into the column header cell (row 3) for a meeting. */
export function meetingHeaderLabel(meeting: { date: string }): string {
  return meeting.date.trim();
}