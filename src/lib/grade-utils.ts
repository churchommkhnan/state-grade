import * as XLSX from "xlsx";
import { DashboardMetrics, GradeBand, StudentRecord } from "@/types/student";
import {
  AR_FOOTER_MARKERS,
  AR_HEADER_MARKERS,
  AR_ID_MARKERS,
  AR_META_ORDER,
  AR_META_SERIAL,
  AR_MIDTERM_PLAIN,
  AR_MIDTERM_PLAIN_SHORT,
  AR_MIDTERM_WITH_AL,
  AR_MIDTERM_WITH_AL_ALT,
  AR_PERCENT_PHRASE_LONG,
  AR_PERCENT_PHRASE_SHORT,
  AR_PERCENT_SIGN,
  AR_PERCENT_WORD,
  AR_PERCENT_WORD_ALT,
  AR_STATUS_FAIL,
  AR_STATUS_NOT_PASS,
  AR_STATUS_PASS,
  AR_STUDENT_NAME_ALT,
  AR_STUDENT_NAME_PRIMARY,
  AR_STUDENT_STATUS,
  AR_THEORY_DEFINITE,
  AR_THEORY_SHORT,
  AR_TOTAL_ALT,
  AR_TOTAL_PRIMARY,
  RE_GRADE_COMPONENT_HINT,
  RE_MIDTERM_COMPACT_SPELLING,
  RE_MIDTERM_HALF_SEMESTER,
  RE_MIDTERM_SPLIT_SPELLING,
  RE_QUIZ_EXTENDED,
  RE_QUIZ_WORD,
  RE_THEORY_BEFORE_PAREN,
  RE_THEORY_WRITTEN_PHRASES,
  RE_YEAR_WORK_BANNER,
} from "@/lib/sheetLocaleStrings";

export const PASS_THRESHOLD = 50;

const gradeScale = [
  { grade: "A", min: 90 },
  { grade: "B", min: 80 },
  { grade: "C", min: 70 },
  { grade: "D", min: 60 },
  { grade: "E", min: 50 },
  { grade: "F", min: 0 },
];

const normalizeKey = (value: string) => value.trim().toLowerCase();

/** Cell text for header / footer detection */
const cellString = (value: unknown): string =>
  value === null || value === undefined ? "" : String(value).replace(/\u00a0/g, " ").trim();

const nfkc = (s: string) => s.normalize("NFKC");

/** Arabic-Indic / Persian digits → ASCII so Number() parses sheet cells. */
function normalizeEasternDigitsInString(s: string): string {
  let out = "";
  for (const ch of s) {
    const cp = ch.codePointAt(0) ?? 0;
    if (cp >= 0x0660 && cp <= 0x0669) {
      out += String(cp - 0x0660);
    } else if (cp >= 0x06f0 && cp <= 0x06f9) {
      out += String(cp - 0x06f0);
    } else {
      out += ch;
    }
  }
  return out;
}

function headerHasStudentName(h: string): boolean {
  const L = nfkc(cellString(h));
  return L.includes(AR_STUDENT_NAME_PRIMARY) || L.includes(AR_STUDENT_NAME_ALT);
}

/** True if header text names a midterm (English or common non-English sheet spellings). */
export function hasMidtermKeyword(text: string): boolean {
  const s = nfkc(cellString(text));
  if (!s.trim()) return false;
  const lower = s.toLowerCase();
  if (/\bmid[\s_-]?term\b/i.test(lower)) return true;
  if (lower.includes("midterm") || lower.includes("mid-term") || lower.includes("mid term")) {
    return true;
  }
  const compact = s.replace(/\s+/g, "");
  if (
    compact.includes(AR_MIDTERM_WITH_AL) ||
    compact.includes(AR_MIDTERM_WITH_AL_ALT) ||
    compact.includes(AR_MIDTERM_PLAIN) ||
    compact.includes(AR_MIDTERM_PLAIN_SHORT)
  ) {
    return true;
  }
  const spaced = s.replace(/\s+/g, " ");
  if (RE_MIDTERM_SPLIT_SPELLING.test(spaced)) return true;
  if (RE_MIDTERM_COMPACT_SPELLING.test(spaced)) return true;
  if (RE_MIDTERM_HALF_SEMESTER.test(spaced)) return true;
  return false;
}

export function hasQuizKeyword(text: string): boolean {
  const s = nfkc(cellString(text));
  const lower = s.toLowerCase();
  if (/\bquiz(zes)?\b/i.test(lower)) return true;
  if (RE_QUIZ_WORD.test(s)) return true;
  if (RE_QUIZ_EXTENDED.test(s)) return true;
  if (/\bassign(ment)?s?\b/i.test(lower)) return true;
  return false;
}

/** Theory / final written — distinct from midterm and quiz. */
export function hasTheoryKeyword(text: string): boolean {
  const s = nfkc(cellString(text));
  if (s.includes(AR_THEORY_DEFINITE)) return true;
  const lower = s.toLowerCase();
  if (/\btheor(y|etical)\b/i.test(lower)) return true;
  if (/\bwritten\b/i.test(lower) && /exam|final/i.test(lower)) return true;
  if (RE_THEORY_BEFORE_PAREN.test(s)) return true;
  if (RE_THEORY_WRITTEN_PHRASES.test(s)) return true;
  if (/\bfinal\s*exam\b/i.test(lower) || /\bwritten\s*final\b/i.test(lower)) return true;
  return false;
}

/** Banner row: year-work / continuous assessment (parent of quiz + midterm in some sheets). */
export function hasYearWorkBannerKeyword(text: string): boolean {
  const s = nfkc(cellString(text));
  return RE_YEAR_WORK_BANNER.test(s);
}

/** Last "(digits)" in label — nested weights like component (20) under block (100). */
export function parseMaxPointsFromHeader(label: string): number | null {
  const L = cellString(label);
  const matches = [...L.matchAll(/\((\d+(?:\.\d+)?)\)/g)];
  if (!matches.length) return null;
  const n = Number(matches[matches.length - 1][1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

function isIdOrMetaColumn(label: string): boolean {
  const L = nfkc(cellString(label));
  if (!L) return false;
  const lower = L.toLowerCase();
  if (lower === "id" || lower === "#") return true;
  const markers = [
    ...AR_ID_MARKERS,
    "seat number",
    "student id",
    "student number",
    "student no",
    "student code",
    "univ id",
    "university id",
    "registration",
    "academic id",
    "national id",
  ];
  if (markers.some((m) => L.includes(m))) return true;
  if (/\breg\.?\s*no\.?\b/i.test(lower) || /\bstd\.?\s*id\b/i.test(lower)) return true;
  return false;
}

/** Header looks like a scored component (not an ID column). */
function looksLikeGradeColumnHeader(raw: string, merged: string): boolean {
  const pick = nfkc(`${raw} ${merged}`);
  if (!pick.trim()) return false;
  if (hasMidtermKeyword(pick)) return true;
  if (/\(\d+(?:\.\d+)?\)/.test(pick)) return true;
  return RE_GRADE_COMPONENT_HINT.test(pick);
}

function matrixRowWidth(row: unknown[] | undefined): number {
  return row?.length ?? 0;
}

function maxColInRange(matrix: Matrix, r0: number, r1: number): number {
  let w = 0;
  for (let r = r0; r <= r1; r++) {
    const len = matrixRowWidth(matrix[r]);
    if (len > w) w = len;
  }
  return w;
}

/** Merged headers: stack text from rows above + header row; then LTR forward-fill gaps in that row. */
function buildEffectiveHeaderLabels(matrix: Matrix, headerRowIndex: number, lookback = 10): string[] {
  const rStart = Math.max(0, headerRowIndex - lookback);
  const width = maxColInRange(matrix, rStart, headerRowIndex);
  const labels: string[] = [];

  for (let c = 0; c < width; c++) {
    const stack: string[] = [];
    const seen = new Set<string>();
    for (let r = rStart; r <= headerRowIndex; r++) {
      const t = nfkc(cellString(matrix[r]?.[c]));
      if (t && !seen.has(t)) {
        seen.add(t);
        stack.push(t);
      }
    }
    labels[c] = stack.join(" › ");
  }

  const rowH = matrix[headerRowIndex] ?? [];
  let carry = "";
  for (let c = 0; c < width; c++) {
    const t = nfkc(cellString(rowH[c]));
    if (t) carry = t;
    if (!labels[c] && carry) labels[c] = carry;
  }

  return labels;
}

const numericFrom = (value: unknown, options?: { allowPercent?: boolean }): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    let raw = value.replace(/\u00a0/g, " ").trim();
    if (raw === "") return null;
    raw = normalizeEasternDigitsInString(raw);
    if (options?.allowPercent !== false && raw.includes("%")) {
      const withoutPct = raw.replace(/%/g, "").replace(/,/g, ".").trim();
      const n = Number(withoutPct);
      return Number.isFinite(n) ? n : null;
    }
    const normalized = raw.replace(/,/g, ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};

/** Parse a cell value as a numeric grade (used by PDF/CSV helpers). */
export function parseCellNumber(value: unknown, options?: { allowPercent?: boolean }): number | null {
  return numericFrom(value, options);
}

type Matrix = unknown[][];

function sheetToMatrix(worksheet: XLSX.WorkSheet): Matrix {
  return XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: "",
    raw: false,
  }) as Matrix;
}

/**
 * Excel sometimes stores a stale/short `!ref`; SheetJS then returns too few rows/columns.
 * Recompute bounds from all cell addresses so every student row is included.
 */
function expandSheetRangeToUsedBounds(worksheet: XLSX.WorkSheet): void {
  let minR = 0;
  let minC = 0;
  let maxR = 0;
  let maxC = 0;
  let found = false;

  for (const key of Object.keys(worksheet)) {
    if (key[0] === "!") continue;
    const addr = key.replace(/\$/g, "");
    if (!/^[A-Za-z]{1,4}\d+$/.test(addr)) continue;
    try {
      const { r, c } = XLSX.utils.decode_cell(addr);
      if (!found) {
        minR = maxR = r;
        minC = maxC = c;
        found = true;
      } else {
        if (r < minR) minR = r;
        if (c < minC) minC = c;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }
    } catch {
      continue;
    }
  }

  if (!found) return;

  if (worksheet["!ref"]) {
    try {
      const b = XLSX.utils.decode_range(worksheet["!ref"]);
      minR = Math.min(minR, b.s.r);
      minC = Math.min(minC, b.s.c);
      maxR = Math.max(maxR, b.e.r);
      maxC = Math.max(maxC, b.e.c);
    } catch {
      /* ignore */
    }
  }

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { r: minR, c: minC },
    e: { r: maxR, c: maxC },
  });
}

/** Sheet !ref width (columns s.c … e.c inclusive). */
function decodeSheetColumnWidth(worksheet: XLSX.WorkSheet): number {
  const ref = worksheet["!ref"];
  if (!ref) return 0;
  const d = XLSX.utils.decode_range(ref);
  return d.e.c - d.s.c + 1;
}

function maxUsedColumnIndex(row: unknown[]): number {
  let m = row.length;
  for (const k of Object.keys(row)) {
    const n = Number.parseInt(k, 10);
    if (!Number.isNaN(n)) m = Math.max(m, n + 1);
  }
  return m;
}

/**
 * SheetJS can return sparse rows (holes) when columns are hidden or unset;
 * short arrays then miss `row[nameIdx]`. Pad every row to full sheet width.
 */
function densifyMatrix(matrix: Matrix, worksheet: XLSX.WorkSheet): Matrix {
  let width = decodeSheetColumnWidth(worksheet);
  for (const row of matrix) {
    if (!row) continue;
    width = Math.max(width, maxUsedColumnIndex(row));
  }
  if (width <= 0) return matrix;
  return matrix.map((row) => {
    const r = row ?? [];
    const out: unknown[] = new Array(width);
    for (let c = 0; c < width; c++) {
      const v = r[c];
      out[c] = v === undefined || v === null ? "" : v;
    }
    return out;
  });
}

function rowContainsHeaderKeyword(row: unknown[]): boolean {
  return row.some((cell) => {
    const s = nfkc(cellString(cell));
    return AR_HEADER_MARKERS.some((m) => s.includes(m));
  });
}

function rowHasStudentNameHeader(row: unknown[]): boolean {
  return row.some((cell) => headerHasStudentName(cellString(cell)));
}

function rowHasMidtermHeaderCell(row: unknown[]): boolean {
  return row.some((cell) => hasMidtermKeyword(cellString(cell)));
}

function headerRowQualityScore(matrix: Matrix, r: number): number {
  const row = matrix[r] ?? [];
  let score = 0;
  for (const cell of row) {
    if (cellString(cell) !== "") score += 1;
  }
  if (row.some((cell) => hasMidtermKeyword(cellString(cell)))) score += 12;
  if (row.some((cell) => isIdOrMetaColumn(cellString(cell)))) score += 8;
  if (row.some((cell) => /\(\d+(?:\.\d+)?\)/.test(cellString(cell)))) score += 6;
  return score;
}

/**
 * Among rows that contain the student-name header, pick the row that looks like the real grade table
 * (many columns, midterm/id headers), not an earlier form/title row. On a tie, prefer a lower row.
 */
function findHeaderRowIndex(matrix: Matrix): number {
  let best = -1;
  let bestQ = -1;
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    if (!rowHasStudentNameHeader(row)) continue;
    const q = headerRowQualityScore(matrix, r);
    if (q > bestQ || (q === bestQ && r > best)) {
      bestQ = q;
      best = r;
    }
  }
  if (best >= 0) return best;

  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    if (rowHasMidtermHeaderCell(row) && row.filter((cell) => cellString(cell) !== "").length >= 4) {
      return r;
    }
  }
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    if (rowContainsHeaderKeyword(row)) return r;
  }
  return -1;
}

function headerIncludes(h: string, fragment: string): boolean {
  return cellString(h).includes(fragment);
}

type ColumnMap = {
  nameIdx: number;
  /** Best-effort student / seat / national id column for display (not used in averages). */
  studentIdIdx: number | null;
  totalIdx: number | null;
  pctIdx: number | null;
  statusIdx: number | null;
  subjectIdxs: number[];
  /** Column indices whose header matches midterm (Arabic/English) — preferred for overall % when no sheet %/total. */
  midtermIdxs: number[];
};

function columnLooksLikeId(
  c: number,
  matrix: Matrix,
  headerRowIndex: number,
  columnLabels: string[]
): boolean {
  const raw = cellString(matrix[headerRowIndex]?.[c]);
  const merged = columnLabels[c] ?? "";
  return (
    isIdOrMetaColumn(raw) ||
    isIdOrMetaColumn(merged) ||
    isIdOrMetaColumn(`${raw} ${merged}`)
  );
}

function pickStudentDisplayIdColumn(
  matrix: Matrix,
  headerRowIndex: number,
  columnLabels: string[],
  width: number,
  reserved: Set<number>,
  nameIdx: number
): number | null {
  const rowH = matrix[headerRowIndex] ?? [];
  let best: number | null = null;
  let bestScore = -1;
  for (let c = 0; c < width; c++) {
    if (reserved.has(c)) continue;
    if (!columnLooksLikeId(c, matrix, headerRowIndex, columnLabels)) continue;
    const rawH = cellString(rowH[c]);
    const merged = columnLabels[c] ?? "";
    const mergedFull = nfkc(`${merged} ${rawH}`);
    if (headerHasStudentName(rawH) || headerHasStudentName(merged) || headerHasStudentName(mergedFull)) {
      continue;
    }
    if (looksLikeGradeColumnHeader(rawH, merged)) continue;

    let nameLikeData = 0;
    let dataSamples = 0;
    for (let r = headerRowIndex + 1; r < matrix.length && dataSamples < 40; r++) {
      const v = cellString(matrix[r]?.[c]);
      if (v.length < 2) continue;
      dataSamples++;
      if (!/^[\d\s./-]+$/.test(normalizeEasternDigitsInString(v))) nameLikeData++;
    }
    const label = mergedFull.trim().toLowerCase();
    let score = 20;
    if (headerIncludes(label, AR_META_SERIAL) || headerIncludes(label, AR_META_ORDER)) score = 5;
    if (/\bstudent\s*(id|no|number)\b/.test(label)) score = 120;
    if (/\bseat\b/.test(label)) score = 70;
    for (let i = 0; i < AR_ID_MARKERS.length - 2; i++) {
      const m = AR_ID_MARKERS[i];
      if (m && headerIncludes(label, m)) score = Math.max(score, 90 - i);
    }
    if (c < nameIdx) score += 8;
    score -= nameLikeData * 3;
    if (score > bestScore || (score === bestScore && best !== null && c < best)) {
      bestScore = score;
      best = c;
    }
  }
  return best;
}

function headerCellIsPercentageColumn(raw: string): boolean {
  const L = nfkc(cellString(raw)).trim();
  if (!L) return false;
  if (L === "%" || L === AR_PERCENT_SIGN) return true;
  if (headerIncludes(L, AR_PERCENT_PHRASE_LONG)) return true;
  if (headerIncludes(L, AR_PERCENT_PHRASE_SHORT)) return true;
  if (L === AR_PERCENT_WORD || L === AR_PERCENT_WORD_ALT) return true;
  if (
    headerIncludes(L, AR_PERCENT_WORD) &&
    !headerIncludes(L, AR_THEORY_DEFINITE) &&
    !headerIncludes(L, AR_THEORY_SHORT)
  ) {
    return true;
  }
  return false;
}

function headerCellIsTotalColumn(raw: string): boolean {
  const L = nfkc(cellString(raw)).trim();
  if (!L) return false;
  return headerIncludes(L, AR_TOTAL_PRIMARY) || headerIncludes(L, AR_TOTAL_ALT);
}

function headerCellIsStatusColumn(raw: string): boolean {
  return headerIncludes(nfkc(cellString(raw)), AR_STUDENT_STATUS);
}

/**
 * If "percentage" column values look like small component marks (e.g. midterm out of 20),
 * it was mis-assigned — drop it so the column can be scored as a subject.
 */
function refinePctIdx(
  matrix: Matrix,
  headerRowIndex: number,
  nameIdx: number,
  pctIdx: number | null
): number | null {
  if (pctIdx === null) return null;
  const vals: number[] = [];
  for (let r = headerRowIndex + 1; r < matrix.length && vals.length < 80; r++) {
    const row = matrix[r];
    if (!row || isRowEmpty(row) || isFooterOrSummaryRow(row)) continue;
    if (!cellString(row[nameIdx])) continue;
    const v = numericFrom(row[pctIdx], { allowPercent: true });
    if (v === null) continue;
    vals.push(v);
  }
  if (vals.length < 10) return pctIdx;
  const max = Math.max(...vals);
  const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
  const allInt = vals.every((x) => Number.isInteger(x));
  /** Typical mis-assignment: midterm/quiz raw marks (e.g. ≤20), not real % columns. */
  if (max <= 26 && mean < 19 && allInt) return null;
  return pctIdx;
}

function pickBestNameColumn(matrix: Matrix, headerRowIndex: number, candidates: number[]): number {
  if (candidates.length === 1) return candidates[0];
  let best = candidates[0]!;
  let bestScore = -1e9;
  for (const c of candidates) {
    let score = 0;
    let samples = 0;
    let digitLike = 0;
    for (let r = headerRowIndex + 1; r < matrix.length; r++) {
      const v = cellString(matrix[r]?.[c]);
      if (v.length < 2) continue;
      samples++;
      if (/^\d+$/.test(normalizeEasternDigitsInString(v))) digitLike++;
      else if (/^[\d\s./-]+$/.test(normalizeEasternDigitsInString(v))) digitLike += 0.65;
      if (/^\d+$/.test(normalizeEasternDigitsInString(v))) continue;
      if (/^[\d\s./-]+$/.test(normalizeEasternDigitsInString(v))) continue;
      if (/[\u0600-\u06FF\u0750-\u077F]/.test(v)) score += 4;
      else if (/[a-zA-Z]/.test(v)) score += 2;
      else score += 1;
    }
    const noise = samples > 0 ? (digitLike / samples) * 25 : 0;
    const colBonus = -c * 0.001;
    const finalScore = score - noise + colBonus;
    if (finalScore > bestScore) {
      bestScore = finalScore;
      best = c;
    }
  }
  return best;
}

function buildColumnMap(
  matrix: Matrix,
  headerRowIndex: number,
  columnLabels: string[]
): ColumnMap | null {
  const rowH = matrix[headerRowIndex] ?? [];
  const width = Math.max(columnLabels.length, rowH.length);

  const nameCandidates: number[] = [];
  let totalIdx: number | null = null;
  let pctIdx: number | null = null;
  let statusIdx: number | null = null;

  for (let c = 0; c < width; c++) {
    const raw = cellString(rowH[c]);
    const merged = columnLabels[c] ?? "";
    const pick = raw || merged;
    if (!nfkc(pick)) continue;

    if (headerHasStudentName(pick)) {
      if (!columnLooksLikeId(c, matrix, headerRowIndex, columnLabels)) {
        nameCandidates.push(c);
      }
    }
    /** Use deepest row text only for totals/%/status so merged multi-row titles cannot hijack columns. */
    if (totalIdx === null && raw && headerCellIsTotalColumn(raw)) totalIdx = c;
    if (pctIdx === null && raw && headerCellIsPercentageColumn(raw)) pctIdx = c;
    if (statusIdx === null && raw && headerCellIsStatusColumn(raw)) statusIdx = c;
  }

  if (!nameCandidates.length) {
    for (let c = 0; c < width; c++) {
      const raw = cellString(rowH[c]);
      const merged = columnLabels[c] ?? "";
      const pick = raw || merged;
      if (!nfkc(pick)) continue;
      if (headerHasStudentName(pick)) nameCandidates.push(c);
    }
  }
  if (!nameCandidates.length) return null;
  const nameIdx = pickBestNameColumn(matrix, headerRowIndex, nameCandidates);

  pctIdx = refinePctIdx(matrix, headerRowIndex, nameIdx, pctIdx);

  const reserved = new Set<number>([nameIdx]);
  if (totalIdx !== null) reserved.add(totalIdx);
  if (pctIdx !== null) reserved.add(pctIdx);
  if (statusIdx !== null) reserved.add(statusIdx);

  const subjectIdxs: number[] = [];
  for (let c = 0; c < width; c++) {
    if (reserved.has(c)) continue;
    if (columnLooksLikeId(c, matrix, headerRowIndex, columnLabels)) continue;
    const rawH = cellString(rowH[c]);
    const merged = columnLabels[c] ?? "";
    const label = (merged || rawH).trim();
    if (!label && !looksLikeGradeColumnHeader(rawH, merged)) continue;
    const lower = label.toLowerCase();
    if (lower === "id" || headerIncludes(label, AR_META_SERIAL) || headerIncludes(label, AR_META_ORDER)) continue;
    subjectIdxs.push(c);
  }

  const midtermIdxs = subjectIdxs.filter((c) => {
    const rawH = cellString(rowH[c]);
    const merged = columnLabels[c] ?? "";
    return hasMidtermKeyword(rawH) || hasMidtermKeyword(merged);
  });

  const studentIdIdx = pickStudentDisplayIdColumn(matrix, headerRowIndex, columnLabels, width, reserved, nameIdx);

  return { nameIdx, studentIdIdx, totalIdx, pctIdx, statusIdx, subjectIdxs, midtermIdxs };
}

function rowTextBlob(row: unknown[]): string {
  return row.map((c) => cellString(c)).join(" ");
}

/** Match whole-word-ish English footers to avoid false positives in data. */
const FOOTER_SUBSTRINGS_EN = [/\btotal\b/i, /\bstatistics\b/i];
function isFooterOrSummaryRow(row: unknown[]): boolean {
  const blob = rowTextBlob(row);
  if (FOOTER_SUBSTRINGS_EN.some((re) => re.test(blob))) return true;
  if (AR_FOOTER_MARKERS.some((s) => blob.includes(s))) return true;
  return false;
}

function isRowEmpty(row: unknown[]): boolean {
  return row.every((c) => cellString(c) === "");
}

function statusFromLocalizedCell(value: unknown): "Pass" | "Fail" | null {
  const s = cellString(value).toLowerCase();
  if (!s) return null;
  if (s.includes(AR_STATUS_NOT_PASS) && s.includes(AR_STATUS_PASS)) return "Fail";
  if (s.includes(AR_STATUS_PASS)) return "Pass";
  if (s.includes(AR_STATUS_FAIL)) return "Fail";
  if (s.includes("pass") || s.includes("success")) return "Pass";
  if (s.includes("fail")) return "Fail";
  return null;
}

const gradeFor = (score: number): string => {
  return gradeScale.find((band) => score >= band.min)?.grade ?? "F";
};

function parseMatrixWithStructuredHeader(
  matrix: Matrix,
  headerRowIndex: number,
  colMap: ColumnMap,
  columnLabels: string[]
): StudentRecord[] {
  const students: StudentRecord[] = [];
  let outIdx = 0;

  for (let r = headerRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    if (isRowEmpty(row) || isFooterOrSummaryRow(row)) continue;

    const name = cellString(row[colMap.nameIdx]);
    if (!name) continue;

    const sheetStudentId =
      colMap.studentIdIdx !== null ? cellString(row[colMap.studentIdIdx]).trim() : "";

    const totalVal =
      colMap.totalIdx !== null ? numericFrom(row[colMap.totalIdx], { allowPercent: true }) : null;
    const pctVal =
      colMap.pctIdx !== null ? numericFrom(row[colMap.pctIdx], { allowPercent: true }) : null;

    const midtermCols = new Set(colMap.midtermIdxs);

    let earned = 0;
    let possible = 0;
    let midEarned = 0;
    let midPossible = 0;
    const subjects: Record<string, number> = {};

    const rowHasEnteredPositiveElsewhere = (excludeC: number): boolean => {
      for (const c2 of colMap.subjectIdxs) {
        if (c2 === excludeC) continue;
        if (cellString(row[c2]) === "") continue;
        const v = numericFrom(row[c2], { allowPercent: true });
        if (v !== null && v > 0) return true;
      }
      return false;
    };

    for (const c of colMap.subjectIdxs) {
      const rawHeader = cellString(matrix[headerRowIndex]?.[c]);
      const mergedLabel = columnLabels[c] ?? "";
      const displayLabel = mergedLabel || rawHeader || `Column ${c + 1}`;
      const maxPts =
        parseMaxPointsFromHeader(rawHeader) ?? parseMaxPointsFromHeader(mergedLabel);

      const cellText = cellString(row[c]);
      if (cellText === "") continue;

      const raw = numericFrom(row[c], { allowPercent: true });
      if (raw === null) continue;

      if (
        maxPts !== null &&
        maxPts > 0 &&
        raw === 0 &&
        rowHasEnteredPositiveElsewhere(c)
      ) {
        continue;
      }

      if (maxPts !== null && maxPts > 0) {
        const piece = Math.max(0, Math.min(maxPts, raw));
        subjects[displayLabel] = (piece / maxPts) * 100;
        earned += piece;
        possible += maxPts;
        if (midtermCols.has(c)) {
          midEarned += piece;
          midPossible += maxPts;
        }
      } else {
        const piece = Math.max(0, Math.min(100, raw));
        subjects[displayLabel] = piece;
        earned += piece;
        possible += 100;
        if (midtermCols.has(c)) {
          midEarned += piece;
          midPossible += 100;
        }
      }
    }

    const componentPercent = possible > 0 ? (earned / possible) * 100 : null;
    const midtermPercent = midPossible > 0 ? (midEarned / midPossible) * 100 : null;

    let average: number | null = null;
    if (pctVal !== null && Number.isFinite(pctVal)) {
      average = Math.max(0, Math.min(100, pctVal));
    } else if (totalVal !== null && Number.isFinite(totalVal)) {
      if (possible > 0) {
        const tol = Math.max(0.75, possible * 0.04);
        const totalMatchesRawSum = Math.abs(earned - totalVal) <= tol;
        if (totalMatchesRawSum || totalVal > 100.5) {
          average = Math.max(0, Math.min(100, (totalVal / possible) * 100));
        } else if (totalVal <= 100) {
          if (
            componentPercent !== null &&
            Math.abs(totalVal - componentPercent) <= 12
          ) {
            average = Math.max(0, Math.min(100, totalVal));
          } else if (componentPercent === null) {
            average = Math.max(0, Math.min(100, totalVal));
          } else {
            average = Math.max(0, Math.min(100, (totalVal / possible) * 100));
          }
        } else {
          average = Math.max(0, Math.min(100, (totalVal / possible) * 100));
        }
      } else if (totalVal <= 100) {
        average = Math.max(0, Math.min(100, totalVal));
      } else if (componentPercent !== null) {
        average = Math.max(0, Math.min(100, componentPercent));
      } else {
        average = Math.max(0, Math.min(100, totalVal));
      }
    } else if (midtermPercent !== null && colMap.midtermIdxs.length > 0) {
      average = Math.max(0, Math.min(100, midtermPercent));
    } else if (componentPercent !== null) {
      average = Math.max(0, Math.min(100, componentPercent));
    } else {
      average = 0;
    }

    const fromSheet = colMap.statusIdx !== null ? statusFromLocalizedCell(row[colMap.statusIdx]) : null;
    const status: StudentRecord["status"] =
      fromSheet ?? (average >= PASS_THRESHOLD ? "Pass" : "Fail");

    outIdx += 1;
    students.push({
      id: `${outIdx}-${name.replace(/\s+/g, "-").slice(0, 80)}`,
      ...(sheetStudentId ? { sheetStudentId } : {}),
      name,
      average,
      status,
      subjects,
    });
  }

  return students;
}

function parseLegacyFlatRows(rows: Record<string, unknown>[]): StudentRecord[] {
  return rows
    .map((row, idx) => {
      const entries = Object.entries(row);
      const nameCell = entries.find(([key]) => {
        const normalized = normalizeKey(key);
        return (
          normalized.includes("name") ||
          (normalized.includes("student") &&
            !normalized.includes("id") &&
            !normalized.includes("number") &&
            !normalized.includes("no"))
        );
      });

      const idCell = entries.find(([key]) => {
        const normalized = normalizeKey(key);
        return (
          normalized === "id" ||
          normalized.includes("student id") ||
          normalized.includes("student number") ||
          normalized.includes("student no") ||
          normalized.includes("seat") ||
          normalized.includes("univ")
        );
      });

      const subjects: Record<string, number> = {};

      for (const [key, value] of entries) {
        const normalized = normalizeKey(key);
        if (normalized.includes("name")) continue;
        if (
          normalized.includes("student") &&
          !normalized.includes("id") &&
          !normalized.includes("number") &&
          !normalized.includes("no")
        ) {
          continue;
        }
        if (idCell && key === idCell[0]) continue;
        if (normalized === "id") continue;
        const score = numericFrom(value, { allowPercent: true });
        if (score !== null) {
          subjects[key] = Math.max(0, Math.min(100, score));
        }
      }

      const subjectScores = Object.values(subjects);
      if (!subjectScores.length) return null;

      const average =
        subjectScores.reduce((sum, score) => sum + score, 0) / subjectScores.length;

      const name =
        typeof nameCell?.[1] === "string" && nameCell[1].trim()
          ? nameCell[1].trim()
          : `Student ${idx + 1}`;

      const sheetId =
        idCell && typeof idCell[1] === "string" && idCell[1].trim() ? idCell[1].trim() : undefined;

      return {
        id: `${idx + 1}-${name.replace(/\s+/g, "-").toLowerCase()}`,
        ...(sheetId ? { sheetStudentId: sheetId } : {}),
        name,
        average,
        status: average >= PASS_THRESHOLD ? "Pass" : "Fail",
        subjects,
      } as StudentRecord;
    })
    .filter((student): student is StudentRecord => Boolean(student));
}

export type ParseStudentFileOptions = {
  /** For .xlsx / .xls: must match a name in `readWorkbookSheetNames`. Ignored for .csv. */
  sheetName?: string;
};

/** Returns tab names for Excel workbooks (single entry for typical CSV). */
export function readWorkbookSheetNames(file: File): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        if (!buffer) {
          reject(new Error("Unable to read the uploaded file."));
          return;
        }
        const workbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true,
          codepage: 65001,
        });
        resolve([...workbook.SheetNames]);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Failed to read the file."));
    reader.readAsArrayBuffer(file);
  });
}

export function parseStudentFile(
  file: File,
  options?: ParseStudentFileOptions
): Promise<StudentRecord[]> {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".pdf")) {
    return import("@/lib/pdf-grade-parser").then((m) => m.parseStudentPdf(file));
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        if (!buffer) {
          reject(new Error("Unable to read the uploaded file."));
          return;
        }

        const workbook = XLSX.read(buffer, {
          type: "array",
          cellDates: true,
          codepage: 65001,
        });
        const requested = options?.sheetName?.trim();
        if (requested && !workbook.SheetNames.includes(requested)) {
          reject(new Error(`Worksheet not found: "${requested}".`));
          return;
        }
        const sheetName = requested ?? workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          reject(new Error("Worksheet is empty or missing."));
          return;
        }
        expandSheetRangeToUsedBounds(worksheet);
        const matrix = densifyMatrix(sheetToMatrix(worksheet), worksheet);

        const headerRowIndex = findHeaderRowIndex(matrix);
        const columnLabels =
          headerRowIndex >= 0 ? buildEffectiveHeaderLabels(matrix, headerRowIndex) : [];
        const colMap =
          headerRowIndex >= 0 ? buildColumnMap(matrix, headerRowIndex, columnLabels) : null;

        let students: StudentRecord[];

        if (headerRowIndex >= 0 && colMap) {
          students = parseMatrixWithStructuredHeader(matrix, headerRowIndex, colMap, columnLabels);
        } else {
          const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
            defval: "",
            raw: false,
          });
          students = parseLegacyFlatRows(rows);
        }

        resolve(students);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("Failed to process the file."));
    reader.readAsArrayBuffer(file);
  });
}

export type CalculateMetricsOptions = {
  /** Limit “Subject comparison” bars to these column labels (e.g. current analysis scope). */
  includeSubjectKey?: (subjectKey: string) => boolean;
  /**
   * Show score metrics on 0…`scoreDisplayMax` instead of 0…100.
   * Student `average` values are still 0–100 internally; they are multiplied by `scoreDisplayMax / 100` here.
   */
  scoreDisplayMax?: number;
};

export function calculateMetrics(
  students: StudentRecord[],
  options?: CalculateMetricsOptions
): DashboardMetrics {
  if (!students.length) {
    return {
      classAverage: 0,
      successRate: 0,
      failRate: 0,
      highestScore: 0,
      lowestScore: 0,
      gradeDistribution: [],
      subjectAverages: [],
      scoreTrend: [],
      totalStudents: 0,
    };
  }

  const rawMax = options?.scoreDisplayMax;
  const scale =
    rawMax !== undefined && Number.isFinite(rawMax) && rawMax > 0 ? rawMax / 100 : 1;

  const averages = students.map((student) => student.average);
  const classAverage = averages.reduce((sum, score) => sum + score, 0) / students.length;
  const passCount = students.filter((student) => student.status === "Pass").length;

  const gradeMap = students.reduce<Record<string, number>>((acc, student) => {
    const grade = gradeFor(student.average);
    acc[grade] = (acc[grade] ?? 0) + 1;
    return acc;
  }, {});

  const gradeDistribution: GradeBand[] = ["A", "B", "C", "D", "E", "F"].map((grade) => ({
    grade,
    count: gradeMap[grade] ?? 0,
  }));

  const subjectTotals: Record<string, { total: number; count: number }> = {};
  const include = options?.includeSubjectKey;

  for (const student of students) {
    for (const [subject, score] of Object.entries(student.subjects)) {
      if (include && !include(subject)) continue;
      if (!subjectTotals[subject]) {
        subjectTotals[subject] = { total: 0, count: 0 };
      }
      subjectTotals[subject].total += score;
      subjectTotals[subject].count += 1;
    }
  }

  const subjectAverages = Object.entries(subjectTotals).map(([subject, values]) => ({
    subject,
    average: Number(((values.total / values.count) * scale).toFixed(2)),
  }));

  const scoreTrend = [...students]
    .sort((a, b) => b.average - a.average)
    .map((student, index) => ({
      index: index + 1,
      score: Number((student.average * scale).toFixed(2)),
    }));

  return {
    classAverage: Number((classAverage * scale).toFixed(2)),
    successRate: Number(((passCount / students.length) * 100).toFixed(1)),
    failRate: Number((((students.length - passCount) / students.length) * 100).toFixed(1)),
    highestScore: Number((Math.max(...averages) * scale).toFixed(2)),
    lowestScore: Number((Math.min(...averages) * scale).toFixed(2)),
    gradeDistribution,
    subjectAverages,
    scoreTrend,
    totalStudents: students.length,
  };
}
