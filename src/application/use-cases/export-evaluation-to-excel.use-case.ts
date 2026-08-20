import type { LocalMemberRepository } from '@/application/ports/member-repository.port';
import type { ExcelGenerator } from '@/application/ports/excel-generator.port';
import type { MemberProfile } from '@/domain/entities/member-profile';

/**
 * Compiles ALL members' evaluation data into the SMMEMBER Excel template
 * (one member per registry row, ranked by total score descending, matching
 * the template's own AS sort helper).
 *
 * Mapping contract (template -> domain):
 *   - Column A  <-> Member name
 *   - Column B  <-> Technical Evaluation /50
 *   - Columns D..R (15) <-> Field Visit scores (0/1)
 *   - Columns U..AI (15)<-> Meeting scores (0/1)
 *   - Columns AK, AL, AM <-> aggregated Tasks score (Interaction / Respect / Bonus)
 *
 * The template's formulas compute S (Field Visits /20), AJ (Meetings /10),
 * AN (Total /110), AO (%) and AP (Grade) when the file is opened, so this
 * injector only writes the raw data cells. Returns the .xlsx bytes.
 */
export class ExportEvaluationToExcelUseCase {
  constructor(
    private readonly repository: LocalMemberRepository,
    private readonly excel: ExcelGenerator,
  ) {}

  async execute(): Promise<Uint8Array> {
    const members = await this.repository.findAll();

    const profiles = await Promise.all(
      members.map(async (member): Promise<MemberProfile> => {
        const profile = await this.repository.loadProfile(member.id);
        if (!profile) throw new Error(`Missing profile for member ${member.id}`);
        return profile;
      }),
    );

    return this.excel.generateAll(profiles);
  }
}