import type { LocalMemberRepository } from '@/application/ports/member-repository.port';
import type { GlobalMeetingInput } from '@/application/dto/member.dto';
import { GlobalMeetingInputSchema } from '@/application/dto/member.dto';
import { parseOrThrow } from './validation';

export class UpdateGlobalMeetingUseCase {
  constructor(private readonly repository: LocalMemberRepository) {}

  async execute(id: string, input: GlobalMeetingInput): Promise<void> {
    const data = parseOrThrow(GlobalMeetingInputSchema, input);
    await this.repository.updateGlobalMeeting(id, {
      name: data.name,
      date: data.date,
    });
  }
}
