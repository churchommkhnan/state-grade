import type { StudentRecord } from "@/types/student";
import { PASS_THRESHOLD, parseCellNumber } from "@/lib/grade-utils";

type TextItem = { str: string; transform: number[] };

function splitTableLine(line: string): string[] {
  if (line.includes("\t")) return line.split("\t").map((s) => s.trim()).filter(Boolean);
  const byGap = line.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
  if (byGap.length >= 2) return byGap;
  return line.trim().split(/\s+/).filter(Boolean);
}

function isHeaderLine(line: string): boolean {
  const l = line.toLowerCase();
  return (
    /\bname\b/.test(l) ||
    /\bstudent\b/.test(l) ||
    /\bid\b/.test(l) ||
    /\bgrade\b/.test(l) ||
    /\bscore\b/.test(l) ||
    /\u0627\u0633\u0645/.test(line)
  );
}

function linesFromPdfTextContent(textContent: { items: unknown[] }): string[] {
  const items = textContent.items
    .map((it) => it as Partial<TextItem>)
    .filter((it): it is TextItem => typeof it.str === "string" && it.str.trim() !== "" && Array.isArray(it.transform))
    .map((it) => ({
      str: it.str.trim(),
      x: it.transform[4] ?? 0,
      y: it.transform[5] ?? 0,
    }));
  if (!items.length) return [];
  items.sort((a, b) => {
    if (Math.abs(a.y - b.y) > 4) return b.y - a.y;
    return a.x - b.x;
  });
  const lineParts: typeof items[] = [];
  let row: typeof items = [];
  let lastY: number | null = null;
  for (const it of items) {
    if (lastY !== null && Math.abs(it.y - lastY) > 4) {
      if (row.length) lineParts.push(row);
      row = [];
    }
    row.push(it);
    lastY = it.y;
  }
  if (row.length) lineParts.push(row);
  return lineParts.map((parts) => parts.map((p) => p.str).join(" ").replace(/\s+/g, " ").trim());
}

function parseGradeLines(lines: string[]): StudentRecord[] {
  let headerIndex = lines.findIndex(isHeaderLine);
  if (headerIndex < 0) headerIndex = 0;
  const headerCells = splitTableLine(lines[headerIndex] ?? "");
  const students: StudentRecord[] = [];
  let outIdx = 0;

  for (let i = headerIndex + 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line || line.length < 2) continue;
    const cells = splitTableLine(line);
    if (cells.length < 2) continue;
    const nums = cells.map((c) => parseCellNumber(c, { allowPercent: true })).filter((n): n is number => n !== null);
    if (nums.length === 0) continue;

    let sheetStudentId = "";
    let name = "";
    let scoreCells: string[] = [];

    if (cells.length >= 3 && /^\d+$/.test(cells[0].replace(/\s/g, ""))) {
      sheetStudentId = cells[0].trim();
      name = cells[1].trim();
      scoreCells = cells.slice(2);
    } else {
      name = cells[0].trim();
      scoreCells = cells.slice(1);
    }
    if (!name || /^[\d.%\s]+$/.test(name)) continue;

    const subjects: Record<string, number> = {};
    const headerScoreLabels = headerCells.slice(sheetStudentId ? 2 : 1);
    scoreCells.forEach((raw, j) => {
      const v = parseCellNumber(raw, { allowPercent: true });
      if (v === null) return;
      const label =
        headerScoreLabels[j]?.trim() && !/^[\d.%\s]+$/.test(headerScoreLabels[j])
          ? headerScoreLabels[j].trim()
          : `Part ${j + 1}`;
      subjects[label] = Math.max(0, Math.min(100, v <= 130 ? v : Math.min(100, v)));
    });
    if (!Object.keys(subjects).length) continue;

    const vals = Object.values(subjects);
    const average = vals.reduce((a, b) => a + b, 0) / vals.length;
    outIdx += 1;
    students.push({
      id: `pdf-${outIdx}-${name.replace(/\s+/g, "-").slice(0, 60)}`,
      ...(sheetStudentId ? { sheetStudentId } : {}),
      name,
      average,
      status: average >= PASS_THRESHOLD ? "Pass" : "Fail",
      subjects,
    });
  }
  return students;
}

/**
 * Best-effort table extraction from text-based PDFs (scanned PDFs are not supported).
 * Expects rows with optional numeric id, a name, and numeric score columns separated by tabs or wide spaces.
 */
export async function parseStudentPdf(file: File): Promise<StudentRecord[]> {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjs.getDocument({ data, useSystemFonts: true }).promise;
  const allLines: string[] = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    allLines.push(...linesFromPdfTextContent(content));
  }
  const students = parseGradeLines(allLines);
  if (!students.length) {
    throw new Error(
      "No student rows found in this PDF. Export as text-based PDF or use Excel/CSV; tables must use tabs or multiple spaces between columns."
    );
  }
  return students;
}
