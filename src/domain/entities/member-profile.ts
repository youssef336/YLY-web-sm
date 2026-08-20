import type { Member } from './member';
import type { TechnicalEvaluation } from './technical-evaluation';
import type { FieldVisit } from './field-visit';
import type { Meeting } from './meeting';
import type { CategoryScores } from './category-scores';
import type { ScoreSummary } from '../score/scoring';

/**
 * Aggregate root for a single member's full evaluation state.
 * Repositories return this object; use cases operate on it.
 */
export interface MemberProfile {
  member: Member;
  technical: TechnicalEvaluation | null;
  fieldVisits: FieldVisit[];
  meetings: Meeting[];
  scores: CategoryScores | null;
  summary: ScoreSummary;
}