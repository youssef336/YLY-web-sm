/**
 * Core scoring business rules, aligned 1:1 with the formulas already present
 * in the SMMEMBER template sheet "Member Evaluation".
 *
 * Template layout (each row = one member):
 *   B        = Technical Evaluation ............... /50  (raw input)
 *   D..R (15)= Field Visit scores (0/1)          /20   -> S = SUM(D:R)/count * 20
 *   U..AI (15)= Meeting scores (0/1)             /10   -> AJ = SUM(U:AI)/count * 10
 *   AK, AL, AM = Interaction, Respect Hierarchy, Bonus (/10 each, direct input)
 *   AN = Total /110 = B + S + AJ + AK + AL + AM
 *   AO = %  = AN/110*100
 *   AP = Grade (A/B/C/D/F)
 *
 * 50 + 20 + 10 + 10 + 10 + 10 = 110.
 *
 * Because both the live UI total and the Excel export compute through this
 * single function, the number shown on screen always matches what the
 * template formulas produce when the file is opened in Excel.
 */
export const SCORING_LIMITS = {
  MAX_TECHNICAL: 50,
  MAX_FIELD_VISITS: 20,
  MAX_MEETINGS: 10,
  MAX_INTERACTION: 10,
  MAX_RESPECT_HIERARCHY: 10,
  MAX_BONUS: 10,
  TOTAL: 110,
} as const;

export interface ScoreSummary {
  technical: number; // /50
  fieldVisits: number; // /20 (normalized: avg x 20)
  meetings: number; // /10 (normalized: avg x 10)
  interaction: number; // /10
  respectHierarchy: number; // /10
  bonus: number; // /10
  total: number; // /110
  percentage: number; // 0-100
  grade: string;
}

export const GRADE_BANDS = [
  { min: 90, grade: 'A' },
  { min: 80, grade: 'B' },
  { min: 70, grade: 'C' },
  { min: 60, grade: 'D' },
  { min: 0, grade: 'F' },
] as const;

export function clampScore(score: number, min: number, max: number): number {
  if (!Number.isFinite(score)) return min;
  return Math.min(max, Math.max(min, score));
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

function round1(value: number): number {
  return Math.round(value * 10) / 10;
}

export interface NormalizedCategory {
  /** Number of entries recorded. */
  count: number;
  /** Normalized score: avg(entry scores) * cap, rounded to 2 decimals. */
  score: number;
}

/**
 * Mirrors the template's `S` and `AJ` formulas:
 *   S  = IFERROR(SUM(D:R)/C*20, 0)
 *   AJ = IFERROR(SUM(U:AI)/T*10, 0)
 */
export function normalizeCategory(entries: Array<{ score: number }>, cap: number): NormalizedCategory {
  if (entries.length === 0) return { count: 0, score: 0 };
  const sum = entries.reduce((acc, entry) => acc + clampScore(entry.score, 0, 1), 0);
  const avg = sum / entries.length;
  return { count: entries.length, score: Math.min(cap, round2(avg * cap)) };
}

export function percentageOf(total: number): number {
  return clampScore(round1((total / SCORING_LIMITS.TOTAL) * 100), 0, 100);
}

export function gradeOf(percentage: number): string {
  for (const band of GRADE_BANDS) {
    if (percentage >= band.min) return band.grade;
  }
  return GRADE_BANDS[GRADE_BANDS.length - 1].grade;
}

export function calculateScoreSummary(input: {
  technical: number;
  fieldVisits: Array<{ score: number }>;
  meetings: Array<{ score: number }>;
  interaction: number;
  respectHierarchy: number;
  bonus: number;
}): ScoreSummary {
  const technical = clampScore(input.technical, 0, SCORING_LIMITS.MAX_TECHNICAL);
  const fieldVisits = normalizeCategory(input.fieldVisits, SCORING_LIMITS.MAX_FIELD_VISITS);
  const meetings = normalizeCategory(input.meetings, SCORING_LIMITS.MAX_MEETINGS);
  const interaction = clampScore(input.interaction, 0, SCORING_LIMITS.MAX_INTERACTION);
  const respectHierarchy = clampScore(input.respectHierarchy, 0, SCORING_LIMITS.MAX_RESPECT_HIERARCHY);
  const bonus = clampScore(input.bonus, 0, SCORING_LIMITS.MAX_BONUS);

  const total = round2(technical + fieldVisits.score + meetings.score + interaction + respectHierarchy + bonus);
  const percentage = percentageOf(total);
  const grade = gradeOf(percentage);

  return {
    technical,
    fieldVisits: fieldVisits.score,
    meetings: meetings.score,
    interaction,
    respectHierarchy,
    bonus,
    total,
    percentage,
    grade,
  };
}

export function emptyScoreSummary(): ScoreSummary {
  return calculateScoreSummary({
    technical: 0,
    fieldVisits: [],
    meetings: [],
    interaction: 0,
    respectHierarchy: 0,
    bonus: 0,
  });
}