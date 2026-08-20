import type { LocalMemberRepository } from '@/application/ports/member-repository.port';
import type { TechnicalScoreInput } from '@/application/dto/member.dto';
import { TechnicalScoreInputSchema } from '@/application/dto/member.dto';
import { MemberNotFoundError } from './errors';
import { parseOrThrow } from './validation';

export class UpdateTechnicalScoreUseCase {
  constructor(private readonly repository: LocalMemberRepository) {}

  async execute(memberId: string, input: TechnicalScoreInput): Promise<void> {
    const { score } = parseOrThrow(TechnicalScoreInputSchema, input);
    const member = await this.repository.findById(memberId);
    if (!member) throw new MemberNotFoundError(memberId);
    await this.repository.upsertTechnical(memberId, score);
  }
}