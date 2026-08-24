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
 * Template fidelity rules:
 *   - Loads the ORIGINAL .xlsx from public/excel_need/ via workbook.xlsx.load().
 *   - NEVER creates a new workbook or worksheet — modifies the existing one.
 *   - ONLY writes to cells that have actual data to inject.
 *   - NEVER touches formula columns (C, T, S, AJ, AN, AO, AP).
 *   - NEVER clears cells — the fresh template has no stale data to clear.
 *   - Strips cached formula results so Excel recalculates on open.
 */
export class ExceljsInjector implements ExcelGenerator {
  async generateAll(input: {
    profiles: MemberProfile[];
    globalFieldVisits: GlobalFieldVisit[];
    globalMeetings: GlobalMeeting[];
  }): Promise<Uint8Array> {
    const { profiles, globalFieldVisits, globalMeetings } = input;

    const templateUrl = `${SMMEMBER_TEMPLATE.filePath}?v=${Date.now()}`;
    let templateBytes: ArrayBuffer;
    try {
      const res = await fetch(templateUrl, {
        cache: 'no-store',
        headers: {
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache',
          'Expires': '0',
        },
      });
      if (!res.ok) throw new Error(`template not found (HTTP ${res.status})`);
      templateBytes = await res.arrayBuffer();
    } catch (cause) {
      throw new Error(
        `SMMEMBER template not bundled. Place "SMMEMBER.xlsx" at public/excel_need/SMMEMBER.xlsx and rebuild.`,
        { cause },
      );
    }

    // Load the ORIGINAL template into exceljs — never create a new workbook.
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBytes);
    const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
    if (!sheet) throw new Error(`Sheet "${SMMEMBER_TEMPLATE.sheetName}" not found in template`);

    // Strip cached formula results so Excel recalculates on open.
    this.clearCachedFormulaResults(sheet);

    // Build slot → label maps from global events.
    const visitHeaders = this.buildSlotLabels(profiles, 'fieldVisits', globalFieldVisits);
    const meetingHeaders = this.buildSlotLabels(profiles, 'meetings', globalMeetings);

    // Write header labels to row 3 — ONLY slots that have a label.
    this.writeHeaders(sheet, visitHeaders, visitsStartCol, MAX_FIELD_VISITS);
    this.writeHeaders(sheet, meetingHeaders, meetingsStartCol, MAX_MEETINGS);

    // Write member data — sorted by total descending.
    const sorted = [...profiles].sort((a, b) => {
      if (b.summary.total !== a.summary.total) return b.summary.total - a.summary.total;
      return a.member.name.localeCompare(b.member.name);
    });

    sorted.forEach((profile, index) =>
      this.writeMember(sheet, SMMEMBER_TEMPLATE.firstDataRow + index, profile),
    );

    // Export the modified workbook — all formatting, merges, styles preserved.
    const buffer = await workbook.xlsx.writeBuffer();
    return new Uint8Array(buffer as ArrayBuffer);
  }

  /**
   * Strip cached `result` from every formula cell so Excel recalculates on open.
   * exceljs loads formulas as { formula, result } — the `result` is stale.
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

  /**
   * Write header labels to row 3. ONLY writes to slots that have a label —
   * empty slots are left untouched so COUNTA sees them as truly empty.
   */
  private writeHeaders(
    sheet: ExcelJS.Worksheet,
    labelsBySlot: Map<number, string>,
    startCol: number,
    maxSlots: number,
  ): void {
    const headerRowCells = sheet.getRow(headerRow);
    for (let slot = 0; slot < maxSlots; slot++) {
      const label = labelsBySlot.get(slot);
      if (!label) continue; // skip — leave cell untouched
      headerRowCells.getCell(startCol + slot).value = label;
    }
  }

  /**
   * Write a single member's data to their row. ONLY writes to cells that have
   * actual data — untouched cells preserve the original template formatting.
   */
  private writeMember(sheet: ExcelJS.Worksheet, rowNumber: number, profile: MemberProfile): void {
    const row = sheet.getRow(rowNumber);

    // Name (column A)
    row.getCell(nameCol).value = profile.member.name;

    // Technical score (column B)
    const technical = profile.technical?.score ?? 0;
    row.getCell(technicalCol).value = technical;

    // Field visit scores (columns D..R) — only write slots with data
    for (const visit of profile.fieldVisits) {
      if (visit.slot >= MAX_FIELD_VISITS) continue;
      row.getCell(visitsStartCol + visit.slot).value = visit.score;
    }

    // Meeting scores (columns U..AI) — only write slots with data
    for (const meeting of profile.meetings) {
      if (meeting.slot >= MAX_MEETINGS) continue;
      row.getCell(meetingsStartCol + meeting.slot).value = meeting.score;
    }

    // Category scores (columns AK, AL, AM)
    row.getCell(interactionCol).value = profile.scores?.interaction ?? 0;
    row.getCell(respectHierarchyCol).value = profile.scores?.respectHierarchy ?? 0;
    row.getCell(bonusCol).value = profile.scores?.bonus ?? 0;
  }
}
