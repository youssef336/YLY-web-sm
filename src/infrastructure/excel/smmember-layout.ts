import { APP_CONFIG } from '@/infrastructure/config';

/**
 * Contract describing exactly where to write data inside the real SMMEMBER
 * template (sheet "Member Evaluation"). Discovered by inspecting the actual
 * file: one member per row from row 5 to row 304 (300 registry rows).
 *
 *   A   = Member name
 *   B   = Technical Evaluation /50            (data entry)
 *   C   = Field Visits Entered               (FORMULA — counts labels in row 3;
 *                                             NEVER written to)
 *   D..R (15) = Field Visit scores (0/1); their "Location/Event - Date" labels
 *               live in the HEADER row 3 (D3..R3)
 *   S   = Field Visits Total /20              (template formula, untouched)
 *   T   = Meetings Entered                   (FORMULA — counts labels in row 3;
 *                                             NEVER written to)
 *   U..AI (15)= Meeting scores (0/1); their "Date" labels live in HEADER
 *               row 3 (U3..AI3)
 *   AJ  = Meetings Total /10                  (template formula, untouched)
 *   AK, AL, AM = Interaction / Respect Hierarchy / Bonus (/10 each, direct)
 *   AN  = Total /110, AO = %, AP = Grade      (template formulas, untouched)
 *
 * Template layout (verified against the real file):
 *   row 2 = dark blue section headers (static text)
 *   row 3 = LIGHT YELLOW header row — empty cells where Field Visit and Meeting
 *           labels are injected (D3..R3, U3..AI3)
 *   row 4 = DARK YELLOW numbering row — static column indexes 1..15, NEVER
 *           overwritten
 *   row 5..304 = one member per row
 *
 * Leaving the S/AJ/AN/AO/AP formulas intact means Excel recomputes them on
 * open from the injected values — matching the live UI scores.
 */
export const SMMEMBER_TEMPLATE = {
  sheetName: 'Member Evaluation',
  filePath: APP_CONFIG.TEMPLATE_PATH,
  downloadFileName: APP_CONFIG.EXPORT_FILENAME,
  /**
   * Row reserved for the Field Visit / Meeting column labels. This is the
   * empty light-yellow row ABOVE the dark-yellow numbering row (row 4). The
   * numbering row must never be written to.
   */
  headerRow: 3,
  firstDataRow: 5,
  lastDataRow: 304,
  columns: {
    name: 'A',
    technical: 'B',
    visitsCount: 'C',
    visitsHeaderStart: 'D',
    visitsHeaderEnd: 'R',
    visitsStart: 'D',
    visitsEnd: 'R',
    meetingsCount: 'T',
    meetingsHeaderStart: 'U',
    meetingsHeaderEnd: 'AI',
    meetingsStart: 'U',
    meetingsEnd: 'AI',
    interaction: 'AK',
    respectHierarchy: 'AL',
    bonus: 'AM',
  },
} as const;

export const MAX_FIELD_VISITS = 15;
export const MAX_MEETINGS = 15;
export const MAX_REGISTRY_ROWS = SMMEMBER_TEMPLATE.lastDataRow - SMMEMBER_TEMPLATE.firstDataRow + 1;

export function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}