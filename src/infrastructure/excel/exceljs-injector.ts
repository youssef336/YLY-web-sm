import ExcelJS from 'exceljs';
import type { ExcelGenerator } from '@/application/ports/excel-generator.port';
import type { MemberProfile } from '@/domain/entities/member-profile';
import type { GlobalFieldVisit } from '@/domain/entities/global-field-visit';
import type { GlobalMeeting } from '@/domain/entities/global-meeting';
import { fieldVisitHeaderLabel } from '@/domain/entities/global-field-visit';
import { meetingHeaderLabel } from '@/domain/entities/global-meeting';
import {
  SMMEMBER_TEMPLATE,
  MAX_FIELD_VISITS,
  MAX_MEETINGS,
  columnLetterToIndex,
} from './smmember-layout';

const { columns, headerRow } = SMMEMBER_TEMPLATE;
const nameCol = columnLetterToIndex(columns.name);
const technicalCol = columnLetterToIndex(columns.technical);
const visitsStartCol = columnLetterToIndex(columns.visitsStart);
const meetingsStartCol = columnLetterToIndex(columns.meetingsStart);
const interactionCol = columnLetterToIndex(columns.interaction);
const respectHierarchyCol = columnLetterToIndex(columns.respectHierarchy);
const bonusCol = columnLetterToIndex(columns.bonus);

/**
 * Injects local member data into the REAL SMMEMBER template using exceljs,
 * entirely in the browser (offline-first — no server involved).
 *
 * Cell injection only, by exact coordinates:
 *   - Field Visit / Meeting labels are written to the HEADER row (row 3 — the
 *     empty light-yellow row ABOVE the dark-yellow numbering row 4), one label
 *     per shared event column (D..R for visits, U..AI for meetings).
 *   - Each member's 0/1 scores are written to that member's row, under the
 *     same column (slot) the label was written to.
 *   - Interaction / Respect Hierarchy / Bonus are written directly to AK / AL /
 *     AM.
 *   - ONLY data-entry cells are written (A, B, D..R, U..AI, AK/AL/AM). The
 *     template's FORMULA columns — C, T, S, AJ, AN, AO, AP — are never
 *     touched, so Excel recalculates them on open from the injected values.
 */
export class ExceljsInjector implements ExcelGenerator {
  async generateAll(input: {
    profiles: MemberProfile[];
    globalFieldVisits: GlobalFieldVisit[];
    globalMeetings: GlobalMeeting[];
  }): Promise<Uint8Array> {
    const { profiles, globalFieldVisits, globalMeetings } = input;

    const templateUrl = new URL(SMMEMBER_TEMPLATE.filePath, window.location.origin).href;
    let templateBytes: ArrayBuffer;
    try {
      const res = await fetch(templateUrl, { cache: 'no-cache' });
      if (!res.ok) throw new Error(`template not found (HTTP ${res.status})`);
      templateBytes = await res.arrayBuffer();
    } catch (cause) {
      throw new Error(
        `SMMEMBER template not bundled. Copy "SMMEMBER .xlsx" from the "exel need" folder into public/templates/ and rebuild.`,
        { cause },
      );
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBytes);
    const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
    if (!sheet) throw new Error(`Sheet "${SMMEMBER_TEMPLATE.sheetName}" not found in template`);

    // exceljs 4.x loads formula cells with cached `result` values from the
    // template (e.g. { formula: '=IF(...)', result: 15 }). If we leave those
    // intact, the stale cached numbers survive into the output and Excel
    // displays them instead of recalculating.  Stripping `result` from every
    // formula cell forces Excel to evaluate natively on open.
    this.clearCachedFormulaResults(sheet);

    this.resetRegistry(sheet);

    // Header labels are global (one event column shared by all members), so
    // derive the slot -> label map by joining member entries with global events.
    const visitHeaders = this.buildSlotLabels(profiles, 'fieldVisits', globalFieldVisits);
    const meetingHeaders = this.buildSlotLabels(profiles, 'meetings', globalMeetings);
    this.writeHeaders(sheet, visitHeaders, visitsStartCol, MAX_FIELD_VISITS);
    this.writeHeaders(sheet, meetingHeaders, meetingsStartCol, MAX_MEETINGS);

    const sorted = [...profiles].sort((a, b) => {
      if (b.summary.total !== a.summary.total) return b.summary.total - a.summary.total;
      return a.member.name.localeCompare(b.member.name);
    });

    sorted.forEach((profile, index) =>
      this.writeMember(sheet, SMMEMBER_TEMPLATE.firstDataRow + index, profile),
    );

    const buffer = await workbook.xlsx.writeBuffer();
    return new Uint8Array(buffer as ArrayBuffer);
  }

  /**
   * Exceljs loads formula cells as `{ formula: '...', result: <cached> }`.
   * The cached result is from the TEMPLATE's last save and is wrong for the
   * new data we are about to inject.  By deleting `result` from every cell
   * that has a `formula` property, we force Excel to re-evaluate all formulas
   * natively when the user opens the downloaded file.
   *
   * We scan the ENTIRE sheet (not just the data rows) because formula cells
   * can appear anywhere — including rows/columns we do not touch.
   */
  private clearCachedFormulaResults(sheet: ExcelJS.Worksheet): void {
    sheet.eachRow((row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell && typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
          delete (cell.value as unknown as Record<string, unknown>).result;
        }
      });
    });
  }

  private resetRegistry(sheet: ExcelJS.Worksheet): void {
    const { firstDataRow, lastDataRow } = SMMEMBER_TEMPLATE;

    // CRITICAL: Use `null` — NOT empty strings (`''`). Excel's COUNTA treats
    // "" as non-empty, which would make COUNTA($D$3:$R$3) always return 15
    // regardless of how many headers were actually written. Setting cells to
    // null makes them truly empty so COUNTA ignores them.

    // Clear header cells (row 3) for visit and meeting labels.
    for (let i = 0; i < MAX_FIELD_VISITS; i++) sheet.getRow(headerRow).getCell(visitsStartCol + i).value = null;
    for (let i = 0; i < MAX_MEETINGS; i++) sheet.getRow(headerRow).getCell(meetingsStartCol + i).value = null;

    // Clear only the DATA-ENTRY cells. The template's formula columns (C, T,
    // S/AJ/AN/AO/AP) are deliberately NOT touched.
    for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      row.getCell(nameCol).value = null;
      row.getCell(technicalCol).value = 0;
      row.getCell(interactionCol).value = 0;
      row.getCell(respectHierarchyCol).value = 0;
      row.getCell(bonusCol).value = 0;
      for (let i = 0; i < MAX_FIELD_VISITS; i++) row.getCell(visitsStartCol + i).value = null;
      for (let i = 0; i < MAX_MEETINGS; i++) row.getCell(meetingsStartCol + i).value = null;
    }
  }

  private buildSlotLabels(
    profiles: MemberProfile[],
    kind: 'fieldVisits' | 'meetings',
    globalEvents: GlobalFieldVisit[] | GlobalMeeting[],
  ): Map<number, string> {
    const bySlot = new Map<number, string>();
    const globalMap = new Map(globalEvents.map((e) => [e.id, e]));

    for (const profile of profiles) {
      for (const entry of profile[kind]) {
        const globalEvent = globalMap.get(entry.globalEventId);
        if (!globalEvent) continue;
        const label =
          kind === 'fieldVisits'
            ? fieldVisitHeaderLabel(globalEvent as GlobalFieldVisit)
            : meetingHeaderLabel(globalEvent as GlobalMeeting);
        bySlot.set(entry.slot, label);
      }
    }
    return bySlot;
  }

  private writeHeaders(
    sheet: ExcelJS.Worksheet,
    labelsBySlot: Map<number, string>,
    startCol: number,
    maxSlots: number,
  ): void {
    const headerRowCells = sheet.getRow(headerRow);
    for (let slot = 0; slot < maxSlots; slot++) {
      const label = labelsBySlot.get(slot);
      headerRowCells.getCell(startCol + slot).value = label ?? null;
    }
  }

  private writeMember(sheet: ExcelJS.Worksheet, rowNumber: number, profile: MemberProfile): void {
    const row = sheet.getRow(rowNumber);
    const technical = profile.technical?.score ?? 0;

    row.getCell(nameCol).value = profile.member.name;
    row.getCell(technicalCol).value = technical;

    for (const visit of profile.fieldVisits) {
      if (visit.slot >= MAX_FIELD_VISITS) continue;
      row.getCell(visitsStartCol + visit.slot).value = visit.score;
    }

    for (const meeting of profile.meetings) {
      if (meeting.slot >= MAX_MEETINGS) continue;
      row.getCell(meetingsStartCol + meeting.slot).value = meeting.score;
    }

    row.getCell(interactionCol).value = profile.scores?.interaction ?? 0;
    row.getCell(respectHierarchyCol).value = profile.scores?.respectHierarchy ?? 0;
    row.getCell(bonusCol).value = profile.scores?.bonus ?? 0;
  }
}