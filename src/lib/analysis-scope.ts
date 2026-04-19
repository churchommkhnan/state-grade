import type { StudentRecord, StudentStatus } from "@/types/student";
import {
  PASS_THRESHOLD,
  hasMidtermKeyword,
  hasQuizKeyword,
  hasTheoryKeyword,
  parseMaxPointsFromHeader,
} from "@/lib/grade-utils";

export type AnalysisScope = "full" | "year_work" | "midterm" | "quiz" | "theory";

/** Weighted % from stored component scores (each value is 0–100 of that component’s max, see parser). */
export function weightedAverageFromSubjectEntries(
  subjects: Record<string, number>,
  keys: string[]
): number {
  let earned = 0;
  let possible = 0;
  for (const k of keys) {
    const v = subjects[k];
    if (v === undefined || v === null || Number.isNaN(v)) continue;
    const maxPts = parseMaxPointsFromHeader(k);
    if (maxPts !== null && maxPts > 0) {
      earned += (v / 100) * maxPts;
      possible += maxPts;
    } else {
      earned += v;
      possible += 100;
    }
  }
  if (possible <= 0) return 0;
  return Math.max(0, Math.min(100, (earned / possible) * 100));
}

export function collectAllSubjectKeys(students: StudentRecord[]): string[] {
  const set = new Set<string>();
  for (const s of students) {
    for (const k of Object.keys(s.subjects)) set.add(k);
  }
  return [...set];
}

/**
 * Which merged column labels belong to this analysis scope.
 * `year_work` = quiz + midterm (continuous assessment block), excluding theory-only columns.
 */
export function subjectKeysForScope(scope: AnalysisScope, keys: string[]): string[] {
  if (scope === "full") return [...keys];
  const out: string[] = [];
  for (const k of keys) {
    if (scope === "midterm" && hasMidtermKeyword(k)) out.push(k);
    if (scope === "quiz" && hasQuizKeyword(k)) out.push(k);
    if (scope === "theory" && hasTheoryKeyword(k)) out.push(k);
    if (scope === "year_work" && (hasMidtermKeyword(k) || hasQuizKeyword(k))) out.push(k);
  }
  return [...new Set(out)];
}

export function applyAnalysisScope(students: StudentRecord[], scope: AnalysisScope): StudentRecord[] {
  return students.map((s) => {
    const keys =
      scope === "full"
        ? Object.keys(s.subjects)
        : subjectKeysForScope(scope, Object.keys(s.subjects));
    const average = weightedAverageFromSubjectEntries(s.subjects, keys);
    const status: StudentStatus = average >= PASS_THRESHOLD ? "Pass" : "Fail";
    return { ...s, average, status };
  });
}

/** Fixed order for UI: every scope is shown; use `subjectKeysForScope` to know if the file has matching columns. */
export const ANALYSIS_SCOPE_ORDER: AnalysisScope[] = [
  "full",
  "year_work",
  "midterm",
  "quiz",
  "theory",
];

export function getAvailableAnalysisScopes(allKeys: string[]): AnalysisScope[] {
  return ANALYSIS_SCOPE_ORDER.filter((scope) => subjectKeysForScope(scope, allKeys).length > 0);
}

export const ANALYSIS_SCOPE_LABELS: Record<AnalysisScope, { title: string; description: string }> = {
  full: {
    title: "Full course",
    description: "All components weighted by point values in parentheses in headers (e.g. 100 total).",
  },
  year_work: {
    title: "Year work (quiz + midterm)",
    description: "Continuous assessment only: columns detected as quiz and/or midterm.",
  },
  midterm: {
    title: "Midterm only",
    description: "Headers that match midterm (English or Arabic labels).",
  },
  quiz: {
    title: "Quiz only",
    description: "Headers that match quiz (English or Arabic labels).",
  },
  theory: {
    title: "Theory / written only",
    description: "Headers that match theory or final written exam.",
  },
};
