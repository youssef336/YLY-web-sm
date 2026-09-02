import type { LocalMemberRepository } from '@/application/ports/member-repository.port';
import type { GlobalFieldVisitInput } from '@/application/dto/member.dto';
import { GlobalFieldVisitInputSchema } from '@/application/dto/member.dto';
import { parseOrThrow } from './validation';

export class UpdateGlobalFieldVisitUseCase {
  constructor(private readonly repository: LocalMemberRepository) {}

  async execute(id: string, input: GlobalFieldVisitInput): Promise<void> {
    const data = parseOrThrow(GlobalFieldVisitInputSchema, input);
    await this.repository.updateGlobalFieldVisit(id, {
      name: data.name,
      date: data.date,
      shift: data.shift,
    });
  }
}
