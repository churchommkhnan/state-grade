export type StudentStatus = "Pass" | "Fail";

export interface StudentRecord {
  /** Stable row id for React / sorting (not necessarily the sheet student number). */
  id: string;
  /** Student number / seat / national id from the sheet when detected (shown next to name). */
  sheetStudentId?: string;
  name: string;
  average: number;
  status: StudentStatus;
  subjects: Record<string, number>;
}

export interface SubjectAverage {
  subject: string;
  average: number;
}

export interface GradeBand {
  grade: string;
  count: number;
}

export interface ScorePoint {
  index: number;
  score: number;
}

export interface DashboardMetrics {
  classAverage: number;
  successRate: number;
  failRate: number;
  highestScore: number;
  lowestScore: number;
  gradeDistribution: GradeBand[];
  subjectAverages: SubjectAverage[];
  scoreTrend: ScorePoint[];
  totalStudents: number;
}
