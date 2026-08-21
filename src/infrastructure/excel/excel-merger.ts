import ExcelJS from 'exceljs';
import {
  SMMEMBER_TEMPLATE,
  MAX_FIELD_VISITS,
  MAX_MEETINGS,
  MAX_REGISTRY_ROWS,
  columnLetterToIndex,
} from './smmember-layout';

const { columns, headerRow, firstDataRow, lastDataRow } = SMMEMBER_TEMPLATE;
const nameCol = columnLetterToIndex(columns.name);
const technicalCol = columnLetterToIndex(columns.technical);
const visitsStartCol = columnLetterToIndex(columns.visitsStart);
const meetingsStartCol = columnLetterToIndex(columns.meetingsStart);
const interactionCol = columnLetterToIndex(columns.interaction);
const respectHierarchyCol = columnLetterToIndex(columns.respectHierarchy);
const bonusCol = columnLetterToIndex(columns.bonus);

export interface OfficialEvent {
  name?: string;
  date: string;
}

function cellToNumber(value: ExcelJS.CellValue): number {
  if (value == null) return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const n = Number(value);
    return Number.isFinite(n) ? n : 0;
  }
  if (typeof value === 'object' && 'result' in value) {
    const r = (value as { result?: unknown }).result;
    return typeof r === 'number' ? r : 0;
  }
  return 0;
}

function cellToString(value: ExcelJS.CellValue): string | null {
  if (value == null) return null;
  if (typeof value === 'string') return value.trim() || null;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'object' && 'result' in value) {
    const r = (value as { result?: unknown }).result;
    return r != null ? String(r) : null;
  }
  return null;
}

/**
 * Extract the date portion from a header label.
 * Visit headers are "Name - Date", meeting headers are just "Date".
 */
function extractDateFromHeader(header: string | null): string | null {
  if (!header) return null;
  const dashIndex = header.lastIndexOf(' - ');
  const datePart = dashIndex >= 0 ? header.substring(dashIndex + 3).trim() : header.trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(datePart) ? datePart : null;
}

/**
 * Read Row 3 of a workbook and return a slot→date map for visits and meetings.
 */
function readHeaderDateMap(
  workbook: ExcelJS.Workbook,
): { visits: Map<number, string>; meetings: Map<number, string> } {
  const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
  const visits = new Map<number, string>();
  const meetings = new Map<number, string>();
  if (!sheet) return { visits, meetings };

  const row = sheet.getRow(headerRow);
  for (let i = 0; i < MAX_FIELD_VISITS; i++) {
    const date = extractDateFromHeader(cellToString(row.getCell(visitsStartCol + i).value));
    if (date) visits.set(i, date);
  }
  for (let i = 0; i < MAX_MEETINGS; i++) {
    const date = extractDateFromHeader(cellToString(row.getCell(meetingsStartCol + i).value));
    if (date) meetings.set(i, date);
  }
  return { visits, meetings };
}

/**
 * Read all member data rows from a workbook. Returns raw row data with scores
 * in their original slot positions (for date-based remapping later).
 */
function extractRawRows(workbook: ExcelJS.Workbook): {
  name: string;
  technical: number;
  visits: (number | null)[];
  meetings: (number | null)[];
  interaction: number;
  respectHierarchy: number;
  bonus: number;
}[] {
  const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
  if (!sheet) return [];

  const rows: {
    name: string;
    technical: number;
    visits: (number | null)[];
    meetings: (number | null)[];
    interaction: number;
    respectHierarchy: number;
    bonus: number;
  }[] = [];

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const row = sheet.getRow(r);
    const name = cellToString(row.getCell(nameCol).value);
    if (!name) break;

    const visits: (number | null)[] = [];
    for (let i = 0; i < MAX_FIELD_VISITS; i++) {
      const v = row.getCell(visitsStartCol + i).value;
      visits.push(v != null ? cellToNumber(v) : null);
    }

    const meetings: (number | null)[] = [];
    for (let i = 0; i < MAX_MEETINGS; i++) {
      const v = row.getCell(meetingsStartCol + i).value;
      meetings.push(v != null ? cellToNumber(v) : null);
    }

    rows.push({
      name,
      technical: cellToNumber(row.getCell(technicalCol).value),
      visits,
      meetings,
      interaction: cellToNumber(row.getCell(interactionCol).value),
      respectHierarchy: cellToNumber(row.getCell(respectHierarchyCol).value),
      bonus: cellToNumber(row.getCell(bonusCol).value),
    });
  }

  return rows;
}

/**
 * Merge multiple uploaded .xlsx files into a single master template.
 *
 * The leader defines "official" events (field visits + meetings) which become
 * the master Row 3 headers. Uploaded files are mapped to the master via DATE
 * matching — sub-leader event names are ignored, only dates matter.
 */
export async function mergeExcelFiles(
  files: File[],
  officialVisits: OfficialEvent[],
  officialMeetings: OfficialEvent[],
): Promise<Uint8Array> {
  // 1. Load pristine master template
  const masterUrl = `${new URL(SMMEMBER_TEMPLATE.filePath, window.location.origin).href}?v=${Date.now()}`;
  const masterRes = await fetch(masterUrl, {
    cache: 'no-store',
    headers: {
      'Cache-Control': 'no-cache, no-store, must-revalidate',
      'Pragma': 'no-cache',
      'Expires': '0',
    },
  });
  if (!masterRes.ok) throw new Error(`Master template not found (HTTP ${masterRes.status})`);

  const masterWb = new ExcelJS.Workbook();
  await masterWb.xlsx.load(await masterRes.arrayBuffer());
  const masterSheet = masterWb.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
  if (!masterSheet) throw new Error(`Sheet "${SMMEMBER_TEMPLATE.sheetName}" not found in template`);

  // 2. Write official headers to Row 3 of the master
  const masterHeaderRow = masterSheet.getRow(headerRow);
  const masterVisitDateToSlot = new Map<string, number>();
  const masterMeetingDateToSlot = new Map<string, number>();

  for (let i = 0; i < officialVisits.length && i < MAX_FIELD_VISITS; i++) {
    const ev = officialVisits[i];
    const label = ev.name ? `${ev.name} - ${ev.date}` : ev.date;
    masterHeaderRow.getCell(visitsStartCol + i).value = label;
    masterVisitDateToSlot.set(ev.date, i);
  }

  for (let i = 0; i < officialMeetings.length && i < MAX_MEETINGS; i++) {
    const ev = officialMeetings[i];
    masterHeaderRow.getCell(meetingsStartCol + i).value = ev.date;
    masterMeetingDateToSlot.set(ev.date, i);
  }

  // 3. Process uploaded files — date-based mapping
  const allRemappedRows: {
    name: string;
    technical: number;
    visits: (number | null)[];
    meetings: (number | null)[];
    interaction: number;
    respectHierarchy: number;
    bonus: number;
  }[] = [];

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    // Read the source file's Row 3 to get date→slot mapping
    const sourceHeaders = readHeaderDateMap(wb);
    // Read all member rows
    const sourceRows = extractRawRows(wb);

    for (const src of sourceRows) {
      // Remap visit scores: source slot → source date → master slot
      const remappedVisits: (number | null)[] = Array(MAX_FIELD_VISITS).fill(null);
      for (const [srcSlot, date] of sourceHeaders.visits) {
        const masterSlot = masterVisitDateToSlot.get(date);
        if (masterSlot != null && src.visits[srcSlot] != null) {
          remappedVisits[masterSlot] = src.visits[srcSlot];
        }
      }

      // Remap meeting scores: source slot → source date → master slot
      const remappedMeetings: (number | null)[] = Array(MAX_MEETINGS).fill(null);
      for (const [srcSlot, date] of sourceHeaders.meetings) {
        const masterSlot = masterMeetingDateToSlot.get(date);
        if (masterSlot != null && src.meetings[srcSlot] != null) {
          remappedMeetings[masterSlot] = src.meetings[srcSlot];
        }
      }

      allRemappedRows.push({
        name: src.name,
        technical: src.technical,
        visits: remappedVisits,
        meetings: remappedMeetings,
        interaction: src.interaction,
        respectHierarchy: src.respectHierarchy,
        bonus: src.bonus,
      });
    }
  }

  if (allRemappedRows.length === 0) {
    throw new Error('No member data found in the uploaded files.');
  }

  if (allRemappedRows.length > MAX_REGISTRY_ROWS) {
    throw new Error(
      `Too many members (${allRemappedRows.length}). The template supports ${MAX_REGISTRY_ROWS} rows maximum.`,
    );
  }

  // 4. Write remapped rows to master template
  for (let i = 0; i < allRemappedRows.length; i++) {
    const data = allRemappedRows[i];
    const row = masterSheet.getRow(firstDataRow + i);

    row.getCell(nameCol).value = data.name;
    row.getCell(technicalCol).value = data.technical;

    for (let s = 0; s < MAX_FIELD_VISITS; s++) {
      if (data.visits[s] != null) row.getCell(visitsStartCol + s).value = data.visits[s];
    }

    for (let s = 0; s < MAX_MEETINGS; s++) {
      if (data.meetings[s] != null) row.getCell(meetingsStartCol + s).value = data.meetings[s];
    }

    row.getCell(interactionCol).value = data.interaction;
    row.getCell(respectHierarchyCol).value = data.respectHierarchy;
    row.getCell(bonusCol).value = data.bonus;
  }

  // 5. Strip cached formula results so Excel recalculates on open
  masterSheet.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell && typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
        delete (cell.value as unknown as Record<string, unknown>).result;
      }
    });
  });

  const buffer = await masterWb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}
