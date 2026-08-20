import ExcelJS from 'exceljs';
import type { ExcelGenerator } from '@/application/ports/excel-generator.port';
import type { MemberProfile } from '@/domain/entities/member-profile';
import type { FieldVisit } from '@/domain/entities/field-visit';
import { fieldVisitHeaderLabel } from '@/domain/entities/field-visit';
import type { Meeting } from '@/domain/entities/meeting';
import { meetingHeaderLabel } from '@/domain/entities/meeting';
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
 *     A visit's header is "Name - Date"; a meeting's header is just the date.
 *     The template's static column numbers in row 4 are never touched.
 *   - Each member's 0/1 scores are written to that member's row, under the
 *     same column (slot) the label was written to.
 *   - Interaction / Respect Hierarchy / Bonus are written directly to AK / AL /
 *     AM.
 *   - ONLY data-entry cells are written (A, B, D..R, U..AI, AK/AL/AM). The
 *     template's FORMULA columns — C (Field Visits Entered), T (Meetings
 *     Entered), S, AJ, AN, AO, AP — are never touched, so Excel recalculates
 *     them on open from the injected values. No styles, colors or static text
 *     are modified.
 */
export class ExceljsInjector implements ExcelGenerator {
  async generateAll(profiles: MemberProfile[]): Promise<Uint8Array> {
    const templateUrl = new URL(SMMEMBER_TEMPLATE.filePath, window.location.origin).href;
    let templateBytes: ArrayBuffer;
    try {
      const res = await fetch(templateUrl, { cache: 'no-cache' });
      if (!res.ok) throw new Error(`template not found (HTTP ${res.status})`);
      templateBytes = await res.arrayBuffer();
    } catch (cause) {
      throw new Error(
        `SMMEMBER template is not bundled. Copy "${SMMEMBER_TEMPLATE.filePath}" into public/templates/ and rebuild.`,
        { cause },
      );
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBytes);
    const sheet = workbook.getWorksheet(SMMEMBER_TEMPLATE.sheetName);
    if (!sheet) throw new Error(`Sheet "${SMMEMBER_TEMPLATE.sheetName}" not found in template`);

    this.resetRegistry(sheet);

    // Header labels are global (one event column shared by all members), so
    // derive the slot -> label map from every member's entries.
    const visitHeaders = this.buildSlotLabels(profiles, 'fieldVisits');
    const meetingHeaders = this.buildSlotLabels(profiles, 'meetings');
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

  private resetRegistry(sheet: ExcelJS.Worksheet): void {
    const { firstDataRow, lastDataRow } = SMMEMBER_TEMPLATE;
    // Clear header cells (row 3) for visit and meeting labels.
    for (let i = 0; i < MAX_FIELD_VISITS; i++) sheet.getRow(headerRow).getCell(visitsStartCol + i).value = '';
    for (let i = 0; i < MAX_MEETINGS; i++) sheet.getRow(headerRow).getCell(meetingsStartCol + i).value = '';

    // Clear only the DATA-ENTRY cells. The template's formula columns (C = Field
    // Visits Entered, T = Meetings Entered, S/AJ/AN/AO/AP) are deliberately NOT
    // touched, so their formulas survive and Excel recalculates them on open.
    for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      row.getCell(nameCol).value = '';
      row.getCell(technicalCol).value = 0;
      row.getCell(interactionCol).value = 0;
      row.getCell(respectHierarchyCol).value = 0;
      row.getCell(bonusCol).value = 0;
      for (let i = 0; i < MAX_FIELD_VISITS; i++) row.getCell(visitsStartCol + i).value = '';
      for (let i = 0; i < MAX_MEETINGS; i++) row.getCell(meetingsStartCol + i).value = '';
    }
  }

  private buildSlotLabels(
    profiles: MemberProfile[],
    kind: 'fieldVisits' | 'meetings',
  ): Map<number, string> {
    const bySlot = new Map<number, string>();
    for (const profile of profiles) {
      for (const entry of profile[kind]) {
        const label =
          kind === 'fieldVisits'
            ? fieldVisitHeaderLabel(entry as FieldVisit)
            : meetingHeaderLabel(entry as Meeting);
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
      headerRowCells.getCell(startCol + slot).value = label ?? '';
    }
  }

  private writeMember(sheet: ExcelJS.Worksheet, rowNumber: number, profile: MemberProfile): void {
    const row = sheet.getRow(rowNumber);
    const technical = profile.technical?.score ?? 0;

    // Strictly cell-level writes to DATA-ENTRY columns only:
    //   A = name, B = technical, D..R = visit scores, U..AI = meeting scores,
    //   AK/AL/AM = Interaction / Respect Hierarchy / Bonus.
    // The formula columns (C, T, S, AJ, AN, AO, AP) are never written to.
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