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

interface ExtractedMemberRow {
  name: string | null;
  technical: number;
  visits: (number | null)[];
  meetings: (number | null)[];
  interaction: number;
  respectHierarchy: number;
  bonus: number;
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

interface ExtractedHeaders {
  visits: (string | null)[];
  meetings: (string | null)[];
}

function extractHeadersFromWorkbook(workbook: ExcelJS.Workbook): ExtractedHeaders {
  const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
  const empty: ExtractedHeaders = { visits: Array(MAX_FIELD_VISITS).fill(null), meetings: Array(MAX_MEETINGS).fill(null) };
  if (!sheet) return empty;

  const row = sheet.getRow(headerRow);
  const visits: (string | null)[] = [];
  for (let i = 0; i < MAX_FIELD_VISITS; i++) {
    visits.push(cellToString(row.getCell(visitsStartCol + i).value));
  }
  const meetings: (string | null)[] = [];
  for (let i = 0; i < MAX_MEETINGS; i++) {
    meetings.push(cellToString(row.getCell(meetingsStartCol + i).value));
  }
  return { visits, meetings };
}

function extractRowsFromWorkbook(workbook: ExcelJS.Workbook): ExtractedMemberRow[] {
  const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
  if (!sheet) return [];

  const rows: ExtractedMemberRow[] = [];

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

export async function mergeExcelFiles(files: File[]): Promise<Uint8Array> {
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

  const allRows: ExtractedMemberRow[] = [];
  const mergedVisitHeaders: (string | null)[] = Array(MAX_FIELD_VISITS).fill(null);
  const mergedMeetingHeaders: (string | null)[] = Array(MAX_MEETINGS).fill(null);

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    // Extract and merge Row 3 headers — first non-empty value per slot wins.
    const headers = extractHeadersFromWorkbook(wb);
    for (let i = 0; i < MAX_FIELD_VISITS; i++) {
      if (!mergedVisitHeaders[i] && headers.visits[i]) mergedVisitHeaders[i] = headers.visits[i];
    }
    for (let i = 0; i < MAX_MEETINGS; i++) {
      if (!mergedMeetingHeaders[i] && headers.meetings[i]) mergedMeetingHeaders[i] = headers.meetings[i];
    }

    allRows.push(...extractRowsFromWorkbook(wb));
  }

  if (allRows.length === 0) {
    throw new Error('No member data found in the uploaded files.');
  }

  if (allRows.length > MAX_REGISTRY_ROWS) {
    throw new Error(
      `Too many members (${allRows.length}). The template supports ${MAX_REGISTRY_ROWS} rows maximum.`,
    );
  }

  for (let i = 0; i < allRows.length; i++) {
    const data = allRows[i];
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

  // Write merged headers to Row 3 — only non-empty slots.
  const headerRowCells = masterSheet.getRow(headerRow);
  for (let i = 0; i < MAX_FIELD_VISITS; i++) {
    if (mergedVisitHeaders[i]) headerRowCells.getCell(visitsStartCol + i).value = mergedVisitHeaders[i];
  }
  for (let i = 0; i < MAX_MEETINGS; i++) {
    if (mergedMeetingHeaders[i]) headerRowCells.getCell(meetingsStartCol + i).value = mergedMeetingHeaders[i];
  }

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