import type { MemberProfile } from '@/domain/entities/member-profile';

/**
 * Presenter (interface adapter): maps a domain MemberProfile aggregate to the
 * shape consumed by the Evaluation Profile screen. Scores always come from
 * the domain scoring rules so the on-screen total matches the Excel export.
 */
export interface ProfileView {
  member: {
    id: string;
    name: string;
    createdAt: string;
  };
  technical: number;
  fieldVisits: Array<{ id: string; name: string; date: string; score: number }>;
  meetings: Array<{ id: string; date: string; score: number }>;
  scores: {
    interaction: number;
    respectHierarchy: number;
    bonus: number;
  };
  summary: {
    technical: number;
    fieldVisits: number;
    meetings: number;
    interaction: number;
    respectHierarchy: number;
    bonus: number;
    total: number;
    percentage: number;
    grade: string;
  };
}

export function profileToView(profile: MemberProfile): ProfileView {
  return {
    member: {
      id: profile.member.id,
      name: profile.member.name,
      createdAt: profile.member.createdAt.toISOString(),
    },
    technical: profile.technical?.score ?? 0,
    fieldVisits: profile.fieldVisits.map((fv) => ({
      id: fv.id,
      name: fv.name,
      date: fv.date,
      score: fv.score,
    })),
    meetings: profile.meetings.map((m) => ({
      id: m.id,
      date: m.date,
      score: m.score,
    })),
    scores: {
      interaction: profile.scores?.interaction ?? 0,
      respectHierarchy: profile.scores?.respectHierarchy ?? 0,
      bonus: profile.scores?.bonus ?? 0,
    },
    summary: profile.summary,
  };
}