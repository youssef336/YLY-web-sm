import type { Member } from '@/domain/entities/member';
import type { TechnicalEvaluation } from '@/domain/entities/technical-evaluation';
import type { FieldVisit, FieldVisitScore } from '@/domain/entities/field-visit';
import type { Meeting, MeetingScore } from '@/domain/entities/meeting';
import type { CategoryScores } from '@/domain/entities/category-scores';
import type { MemberProfile } from '@/domain/entities/member-profile';

/**
 * Local repository port (output boundary).
 *
 * The domain/application layers only know this interface. The IndexedDB
 * adapter (web/PWA) in the infrastructure layer implements it. Because the
 * application is fully offline-first, this port is intentionally async so the
 * same use cases can drive SQLite/Hive/Isar adapters on mobile later.
 *
 * Field Visits and Meetings are shared events in the template: each occupies a
 * column (slot 0-14) whose label lives in the header row. The repository is
 * responsible for slot assignment (reuse the slot of an entry with the same
 * label; otherwise allocate the next free slot), so callers never think about
 * columns.
 */
export interface LocalMemberRepository {
  createMember(member: Member): Promise<Member>;
  findById(id: string): Promise<Member | null>;
  findAll(): Promise<Member[]>;
  updateMemberName(id: string, name: string): Promise<Member>;
  deleteMember(id: string): Promise<void>;

  getTechnical(memberId: string): Promise<TechnicalEvaluation | null>;
  upsertTechnical(memberId: string, score: number): Promise<TechnicalEvaluation>;

  getCategoryScores(memberId: string): Promise<CategoryScores | null>;
  upsertCategoryScores(memberId: string, input: { interaction: number; respectHierarchy: number; bonus: number }): Promise<CategoryScores>;

  addFieldVisit(input: {
    memberId: string;
    id: string;
    name: string;
    date: string;
    score: FieldVisitScore;
  }): Promise<FieldVisit>;
  updateFieldVisit(id: string, input: { name: string; date: string; score: FieldVisitScore }): Promise<FieldVisit>;
  removeFieldVisit(id: string): Promise<void>;

  addMeeting(input: { memberId: string; id: string; date: string; score: MeetingScore }): Promise<Meeting>;
  updateMeeting(id: string, input: { date: string; score: MeetingScore }): Promise<Meeting>;
  removeMeeting(id: string): Promise<void>;

  loadProfile(memberId: string): Promise<MemberProfile | null>;
}