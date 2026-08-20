import type { MemberProfile } from '@/domain/entities/member-profile';

/**
 * Excel generation port (output boundary).
 *
 * The use case asks for a compiled workbook (the real SMMEMBER template with
 * the member data injected into its registry rows). It knows nothing about
 * exceljs, the template file, or cell coordinates. Swap exceljs for SheetJS /
 * a server-side report service without touching business logic.
 */
export interface ExcelGenerator {
  /** Compiles every member's data into the SMMEMBER template. Returns the .xlsx bytes. */
  generateAll(profiles: MemberProfile[]): Promise<Uint8Array>;
}