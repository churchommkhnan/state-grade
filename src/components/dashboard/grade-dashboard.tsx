"use client";

import { useEffect, useMemo, useState } from "react";
import { useDropzone } from "react-dropzone";
import { AnimatePresence, motion } from "framer-motion";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  ColumnDef,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  useReactTable,
} from "@tanstack/react-table";
import {
  BarChart3,
  CheckCircle2,
  FileSpreadsheet,
  Filter,
  GraduationCap,
  Layers,
  Search,
  Trophy,
  UploadCloud,
  User2,
  XCircle,
} from "lucide-react";
import {
  ANALYSIS_SCOPE_LABELS,
  ANALYSIS_SCOPE_ORDER,
  applyAnalysisScope,
  collectAllSubjectKeys,
  getAvailableAnalysisScopes,
  type AnalysisScope,
  subjectKeysForScope,
  weightedAverageFromSubjectEntries,
} from "@/lib/analysis-scope";
import { calculateMetrics, parseStudentFile, PASS_THRESHOLD, readWorkbookSheetNames } from "@/lib/grade-utils";
import { DashboardMetrics, StudentRecord, StudentStatus } from "@/types/student";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";

const pieColors = ["#6366f1", "#ef4444"];

/** Letter-grade slice colors (A → F) */
const GRADE_SLICE_COLORS: Record<string, string> = {
  A: "#22c55e",
  B: "#84cc16",
  C: "#eab308",
  D: "#f97316",
  E: "#f43f5e",
  F: "#991b1b",
};

const SCORE_RANGE_BANDS: { label: string; min: number; max: number }[] = [
  { label: "90–100%", min: 90, max: 100 },
  { label: "80–89%", min: 80, max: 89.999 },
  { label: "70–79%", min: 70, max: 79.999 },
  { label: "60–69%", min: 60, max: 69.999 },
  { label: "50–59%", min: 50, max: 59.999 },
  { label: "0–49%", min: 0, max: 49.999 },
];

const RANGE_SLICE_COLORS = ["#22c55e", "#84cc16", "#eab308", "#f97316", "#f43f5e", "#991b1b"];

type TableStudentRow = StudentRecord & { classRank: number };

const metricSeed: DashboardMetrics = {
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

export function GradeDashboard() {
  const [mounted, setMounted] = useState(false);
  const [parsedStudents, setParsedStudents] = useState<StudentRecord[]>([]);
  const [analysisScope, setAnalysisScope] = useState<AnalysisScope>("full");
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState<"all" | "Pass" | "Fail">("all");
  const [selectedStudent, setSelectedStudent] = useState<StudentRecord | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [sheetPickerFile, setSheetPickerFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheetName, setSelectedSheetName] = useState<string>("");
  const [gradePieMode, setGradePieMode] = useState<"letter" | "range">("letter");
  const [manualColumnMode, setManualColumnMode] = useState(false);
  const [selectedManualKeys, setSelectedManualKeys] = useState<string[]>([]);

  useEffect(() => {
    setMounted(true);
  }, []);

  const isExcelFile = (file: File) => {
    const n = file.name.toLowerCase();
    return n.endsWith(".xlsx") || n.endsWith(".xls");
  };

  const isPdfFile = (file: File) => file.name.toLowerCase().endsWith(".pdf");

  const runParse = async (file: File, sheetName?: string) => {
    const parsed = await parseStudentFile(file, sheetName ? { sheetName } : undefined);
    setParsedStudents(parsed);
    const keys = collectAllSubjectKeys(parsed);
    setSelectedManualKeys(keys);
    setManualColumnMode(false);
    const avail = getAvailableAnalysisScopes(keys);
    setAnalysisScope(avail.includes("full") ? "full" : avail[0] ?? "full");
    setSheetPickerFile(null);
    setSheetNames([]);
    setSelectedSheetName("");
  };

  const subjectKeysUnion = useMemo(
    () => collectAllSubjectKeys(parsedStudents),
    [parsedStudents]
  );

  const scopesWithData = useMemo(
    () => getAvailableAnalysisScopes(subjectKeysUnion),
    [subjectKeysUnion]
  );

  useEffect(() => {
    if (!scopesWithData.length) return;
    if (!scopesWithData.includes(analysisScope)) {
      setAnalysisScope(scopesWithData.includes("full") ? "full" : scopesWithData[0]);
    }
  }, [scopesWithData, analysisScope]);

  const scopedSubjectKeys = useMemo(
    () => subjectKeysForScope(analysisScope, subjectKeysUnion),
    [analysisScope, subjectKeysUnion]
  );

  const displayStudents = useMemo(() => {
    if (!parsedStudents.length) return [];
    if (manualColumnMode && selectedManualKeys.length > 0) {
      return parsedStudents.map((s) => {
        const average = weightedAverageFromSubjectEntries(s.subjects, selectedManualKeys);
        const status: StudentStatus = average >= PASS_THRESHOLD ? "Pass" : "Fail";
        return { ...s, average, status };
      });
    }
    return applyAnalysisScope(parsedStudents, analysisScope);
  }, [parsedStudents, analysisScope, manualColumnMode, selectedManualKeys]);

  const metrics = useMemo((): DashboardMetrics => {
    if (!displayStudents.length) return metricSeed;
    const includeSubjectKey =
      manualColumnMode && selectedManualKeys.length > 0
        ? (k: string) => selectedManualKeys.includes(k)
        : analysisScope === "full"
          ? undefined
          : (k: string) => scopedSubjectKeys.includes(k);
    return calculateMetrics(displayStudents, { includeSubjectKey });
  }, [displayStudents, analysisScope, scopedSubjectKeys, manualColumnMode, selectedManualKeys]);

  const letterGradePieData = useMemo(() => {
    const total = metrics.totalStudents || 1;
    return metrics.gradeDistribution
      .filter((g) => g.count > 0)
      .map((g) => ({
        name: g.grade,
        count: g.count,
        value: g.count,
        pct: Number(((g.count / total) * 100).toFixed(1)),
      }));
  }, [metrics.gradeDistribution, metrics.totalStudents]);

  const scoreRangePieData = useMemo(() => {
    const total = displayStudents.length;
    if (!total) return [];
    return SCORE_RANGE_BANDS.map((band, i) => {
      const count = displayStudents.filter(
        (s) => s.average >= band.min && s.average <= band.max
      ).length;
      return {
        name: band.label,
        count,
        value: count,
        pct: Number(((count / total) * 100).toFixed(1)),
        fill: RANGE_SLICE_COLORS[i % RANGE_SLICE_COLORS.length],
      };
    }).filter((d) => d.value > 0);
  }, [displayStudents]);

  const onDrop = async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    if (!file) return;

    try {
      setUploadError(null);
      setSheetPickerFile(null);
      setSheetNames([]);
      setSelectedSheetName("");

      if (isPdfFile(file)) {
        await runParse(file);
        return;
      }

      if (isExcelFile(file)) {
        const names = await readWorkbookSheetNames(file);
        if (names.length > 1) {
          setSheetPickerFile(file);
          setSheetNames(names);
          setSelectedSheetName(names[0] ?? "");
          return;
        }
        await runParse(file, names[0]);
        return;
      }

      await runParse(file);
    } catch (err) {
      setUploadError(
        err instanceof Error
          ? err.message
          : "Could not parse this file. Use CSV, XLSX, XLS, or a text-based PDF with clear columns."
      );
    }
  };

  const confirmSheetSelection = async () => {
    if (!sheetPickerFile || !selectedSheetName) return;
    try {
      setUploadError(null);
      await runParse(sheetPickerFile, selectedSheetName);
    } catch (err) {
      setUploadError(
        err instanceof Error
          ? err.message
          : "Could not parse this file. Use CSV, XLSX, XLS, or a text-based PDF with clear columns."
      );
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    maxFiles: 1,
    accept: {
      "text/csv": [".csv"],
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "application/pdf": [".pdf"],
    },
  });

  const rankByStudentId = useMemo(() => {
    const sorted = [...displayStudents].sort((a, b) => {
      if (b.average !== a.average) return b.average - a.average;
      return a.name.localeCompare(b.name);
    });
    const map = new Map<string, number>();
    for (let i = 0; i < sorted.length; i++) {
      const prev = sorted[i - 1];
      const cur = sorted[i];
      if (!cur) continue;
      const r =
        i > 0 && prev && cur.average === prev.average ? (map.get(prev.id) ?? i + 1) : i + 1;
      map.set(cur.id, r);
    }
    return map;
  }, [displayStudents]);

  const tableRows = useMemo((): TableStudentRow[] => {
    return displayStudents
      .filter((student) => {
        const matchQuery = student.name.toLowerCase().includes(query.toLowerCase());
        const matchId = (student.sheetStudentId ?? "")
          .toLowerCase()
          .includes(query.toLowerCase());
        const matchQ = matchQuery || matchId;
        const matchStatus = statusFilter === "all" ? true : student.status === statusFilter;
        return matchQ && matchStatus;
      })
      .map((s) => ({ ...s, classRank: rankByStudentId.get(s.id) ?? 0 }));
  }, [displayStudents, query, statusFilter, rankByStudentId]);

  const columns = useMemo<ColumnDef<TableStudentRow>[]>(
    () => [
      {
        id: "rank",
        header: "Rank",
        cell: ({ row }) => <span className="tabular-nums text-slate-200">{row.original.classRank}</span>,
      },
      {
        id: "sheetId",
        header: "ID",
        cell: ({ row }) => (
          <span className="text-slate-300">{row.original.sheetStudentId?.trim() || "—"}</span>
        ),
      },
      { accessorKey: "name", header: "Student" },
      {
        accessorKey: "average",
        header: "Average",
        cell: ({ row }) => `${row.original.average.toFixed(1)}%`,
      },
      {
        accessorKey: "status",
        header: "Status",
        cell: ({ row }) => (
          <Badge className={row.original.status === "Pass" ? "bg-emerald-500/20 text-emerald-200" : "bg-rose-500/20 text-rose-200"}>
            {row.original.status}
          </Badge>
        ),
      },
      {
        id: "subjects",
        header: "Subjects",
        cell: ({ row }) => Object.keys(row.original.subjects).length,
      },
    ],
    []
  );

  // eslint-disable-next-line react-hooks/incompatible-library
  const table = useReactTable({
    data: tableRows,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
  });

  const successData = [
    { name: "Pass", value: metrics.successRate },
    { name: "Fail", value: metrics.failRate },
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-indigo-950 text-slate-100">
      <div className="mx-auto flex max-w-7xl gap-6 p-4 md:p-6">
        <aside className="sticky top-4 hidden h-[calc(100vh-2rem)] w-64 flex-col rounded-2xl border border-white/15 bg-slate-900/60 p-5 backdrop-blur-xl lg:flex">
          <div className="flex items-center gap-3">
            <div className="rounded-xl bg-indigo-500/20 p-2">
              <GraduationCap className="size-5 text-indigo-300" />
            </div>
            <div>
              <p className="text-xs text-slate-400">Premium Suite</p>
              <p className="font-semibold">Grade Insights</p>
            </div>
          </div>
          <nav className="mt-8 space-y-2 text-sm text-slate-300">
            <p className="rounded-lg bg-white/10 px-3 py-2">Dashboard</p>
            <p className="px-3 py-2">Analytics</p>
            <p className="px-3 py-2">Students</p>
          </nav>
          <div className="mt-auto rounded-xl bg-indigo-500/15 p-4 text-xs text-indigo-100">
            Upload class sheets to generate instant performance analytics.
          </div>
        </aside>

        <main className="flex-1 space-y-6">
          <div className="flex flex-col gap-2">
            <h1 className="text-2xl font-semibold md:text-3xl">Student Grade Analysis Dashboard</h1>
            <p className="text-sm text-slate-400">Modern analytics for educational performance tracking.</p>
          </div>

          <Card className="p-5">
            <div
              {...getRootProps()}
              className={`cursor-pointer rounded-2xl border-2 border-dashed p-8 text-center transition-all ${
                isDragActive ? "border-indigo-400 bg-indigo-500/10" : "border-slate-600 hover:border-indigo-400"
              }`}
            >
              <input {...getInputProps()} />
              <UploadCloud className="mx-auto mb-3 size-8 text-indigo-300" />
              <p className="font-medium">Drag & drop Excel, CSV, or PDF</p>
              <p className="mt-1 text-sm text-slate-400">
                PDF must be text-based with columns separated by tabs or multiple spaces.
              </p>
            </div>
            {uploadError ? <p className="mt-3 text-sm text-rose-300">{uploadError}</p> : null}

            {sheetPickerFile && sheetNames.length > 1 ? (
              <div className="mt-5 rounded-xl border border-white/10 bg-slate-950/50 p-4">
                <p className="mb-2 text-sm font-medium text-slate-200">This workbook has multiple sheets</p>
                <p className="mb-3 text-xs text-slate-400">Choose which sheet to analyze, then load the data.</p>
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
                  <label className="sr-only" htmlFor="sheet-select">
                    Sheet
                  </label>
                  <select
                    id="sheet-select"
                    value={selectedSheetName}
                    onChange={(e) => setSelectedSheetName(e.target.value)}
                    className="w-full min-w-0 flex-1 rounded-lg border border-white/15 bg-slate-900 px-3 py-2 text-sm text-slate-100 outline-none ring-indigo-400/40 focus:ring-2 sm:max-w-md"
                  >
                    {sheetNames.map((name) => (
                      <option key={name} value={name}>
                        {name}
                      </option>
                    ))}
                  </select>
                  <Button type="button" onClick={confirmSheetSelection} className="shrink-0">
                    Load sheet
                  </Button>
                </div>
              </div>
            ) : null}
          </Card>

          {parsedStudents.length > 0 ? (
            <Card className="border-indigo-500/20 bg-slate-900/40 p-5">
              <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
                <div className="flex items-start gap-3">
                  <div className="rounded-lg bg-indigo-500/15 p-2">
                    <Layers className="size-5 text-indigo-300" />
                  </div>
                  <div>
                    <h2 className="text-base font-semibold text-slate-100">Analysis scope</h2>
                    <p className="text-xs text-slate-400">
                      Pick which grade components the average and charts use. Weights come from numbers in parentheses in
                      column headers (for example Midterm (20) is out of 20 points).
                    </p>
                  </div>
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                {ANALYSIS_SCOPE_ORDER.map((scope) => {
                  const meta = ANALYSIS_SCOPE_LABELS[scope];
                  const keyCount = subjectKeysForScope(scope, subjectKeysUnion).length;
                  const hasData = keyCount > 0;
                  const active = analysisScope === scope;
                  return (
                    <button
                      key={scope}
                      type="button"
                      disabled={!hasData}
                      onClick={() => hasData && setAnalysisScope(scope)}
                      title={
                        hasData
                          ? meta.description
                          : "No columns in this file match this scope (check header names like Midterm, Quiz, Theory)."
                      }
                      className={`rounded-xl border px-3 py-2 text-left text-sm transition-all ${
                        !hasData
                          ? "cursor-not-allowed border-white/5 bg-white/[0.02] text-slate-500 opacity-60"
                          : active
                            ? "border-indigo-400 bg-indigo-500/20 text-indigo-100 shadow-sm shadow-indigo-900/40"
                            : "border-white/10 bg-white/[0.03] text-slate-300 hover:border-white/20 hover:bg-white/[0.06]"
                      }`}
                    >
                      <span className="block font-medium">{meta.title}</span>
                      <span className="block text-[11px] text-slate-400">
                        {hasData ? `${keyCount} column${keyCount === 1 ? "" : "s"}` : "No matching columns"}
                      </span>
                    </button>
                  );
                })}
              </div>
              {scopesWithData.length === 1 && scopesWithData[0] === "full" ? (
                <div className="mt-3 rounded-lg border border-amber-500/25 bg-amber-500/10 px-3 py-2 text-[11px] text-amber-100/90">
                  <p className="font-medium text-amber-200">Why can&apos;t I select Midterm, Quiz, or Theory?</p>
                  <p className="mt-1 text-amber-100/85">
                    Those modes need column titles that name each part of the grade. The app scans headers for words like{" "}
                    <strong>Midterm</strong>, <strong>Quiz</strong>, <strong>Theory</strong>, or common Arabic header spellings
                    for those terms (often with points in parentheses, e.g. <strong>Midterm (20)</strong>). Your file
                    only contributed <strong>{subjectKeysUnion.length}</strong> scored column label
                    {subjectKeysUnion.length === 1 ? "" : "s"} that didn&apos;t match those patterns. Edit the header row in
                    Excel, then upload again.
                  </p>
                </div>
              ) : null}
              <p className="mt-3 text-[11px] text-slate-500">{ANALYSIS_SCOPE_LABELS[analysisScope].description}</p>
            </Card>
          ) : null}

          {parsedStudents.length > 0 && subjectKeysUnion.length > 0 ? (
            <Card className="border-white/10 bg-slate-900/40 p-5">
              <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h2 className="text-base font-semibold text-slate-100">Columns for class average</h2>
                  <p className="text-xs text-slate-400">
                    Turn on custom selection to choose exactly which score columns drive the average, charts, pass/fail,
                    and rank. When off, the analysis scope above is used.
                  </p>
                </div>
                <label className="flex cursor-pointer items-center gap-2 text-sm text-slate-300">
                  <input
                    type="checkbox"
                    checked={manualColumnMode}
                    onChange={(e) => {
                      setManualColumnMode(e.target.checked);
                      if (e.target.checked && selectedManualKeys.length === 0) {
                        setSelectedManualKeys([...subjectKeysUnion]);
                      }
                    }}
                    className="size-4 rounded border-white/20 bg-slate-900"
                  />
                  Custom columns
                </label>
              </div>
              {manualColumnMode ? (
                <>
                  <div className="mb-2 flex flex-wrap gap-2">
                    <Button
                      type="button"
                      size="sm"
                      variant="outline"
                      onClick={() => setSelectedManualKeys([...subjectKeysUnion])}
                    >
                      Select all
                    </Button>
                    <Button type="button" size="sm" variant="outline" onClick={() => setSelectedManualKeys([])}>
                      Clear
                    </Button>
                  </div>
                  <div className="max-h-48 space-y-2 overflow-y-auto rounded-lg border border-white/10 bg-slate-950/50 p-3">
                    {subjectKeysUnion.map((key) => {
                      const checked = selectedManualKeys.includes(key);
                      return (
                        <label
                          key={key}
                          className="flex cursor-pointer items-start gap-2 text-sm text-slate-200"
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => {
                              setSelectedManualKeys((prev) =>
                                prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]
                              );
                            }}
                            className="mt-0.5 size-4 shrink-0 rounded border-white/20 bg-slate-900"
                          />
                          <span className="break-all">{key}</span>
                        </label>
                      );
                    })}
                  </div>
                  {selectedManualKeys.length === 0 ? (
                    <p className="mt-2 text-xs text-amber-200/90">Select at least one column, or turn off custom mode.</p>
                  ) : null}
                </>
              ) : null}
            </Card>
          ) : null}

          <section className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-4">
            <StatCard icon={BarChart3} title="Class Average" value={`${metrics.classAverage}%`} />
            <StatCard icon={CheckCircle2} title="Success Rate" value={`${metrics.successRate}%`} />
            <StatCard icon={Trophy} title="Highest Score" value={`${metrics.highestScore}%`} />
            <StatCard icon={XCircle} title="Lowest Score" value={`${metrics.lowestScore}%`} />
          </section>

          <section className="flex flex-col gap-4 xl:flex-row">
            <Card className="min-h-80 flex-1 p-4 xl:min-w-0">
              <p className="mb-1 text-sm text-slate-300">Score ranking curve</p>
              <p className="mb-3 text-[11px] text-slate-500">
                Sorted averages for the analysis scope you selected above.
              </p>
              {mounted ? (
                <ResponsiveContainer width="100%" height={320}>
                  <AreaChart data={metrics.scoreTrend}>
                    <defs>
                      <linearGradient id="colorScore" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#818cf8" stopOpacity={0.8} />
                        <stop offset="100%" stopColor="#818cf8" stopOpacity={0.1} />
                      </linearGradient>
                    </defs>
                    <XAxis dataKey="index" stroke="#94a3b8" tick={{ fontSize: 11 }} />
                    <YAxis stroke="#94a3b8" domain={[0, 100]} tick={{ fontSize: 11 }} />
                    <Tooltip
                      formatter={(v) => [`${v}%`, "Average"]}
                    />
                    <Area type="monotone" dataKey="score" stroke="#818cf8" fill="url(#colorScore)" />
                  </AreaChart>
                </ResponsiveContainer>
              ) : null}
            </Card>

            <div className="flex w-full flex-col gap-4 xl:w-[min(100%,420px)] xl:shrink-0">
              <Card className="flex min-h-[220px] flex-1 flex-col p-4">
                <p className="mb-1 text-sm font-medium text-slate-200">Pass vs Fail</p>
                <p className="mb-2 text-[11px] text-slate-500">
                  Pass vs fail at the 50% threshold for the displayed scope average.
                </p>
                {mounted && metrics.totalStudents > 0 ? (
                  <ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie
                        data={successData}
                        innerRadius={48}
                        outerRadius={72}
                        dataKey="value"
                        paddingAngle={4}
                        nameKey="name"
                        label={({ name, value }) => `${name}: ${value}%`}
                      >
                        {successData.map((entry, index) => (
                          <Cell key={entry.name} fill={pieColors[index % pieColors.length]} />
                        ))}
                      </Pie>
                      <Tooltip formatter={(v) => [`${v}%`, "Share"]} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : null}
              </Card>

              <Card className="flex min-h-[320px] flex-1 flex-col p-4">
                <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <p className="text-sm font-medium text-slate-200">Student count by grade</p>
                    <p className="text-[11px] text-slate-500">
                      How many students fall in each bucket for the same analysis scope.
                    </p>
                  </div>
                  <div className="flex shrink-0 gap-1 rounded-lg border border-white/10 bg-slate-950/80 p-0.5">
                    <button
                      type="button"
                      onClick={() => setGradePieMode("letter")}
                      className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors ${
                        gradePieMode === "letter"
                          ? "bg-indigo-500/30 text-indigo-100"
                          : "text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      A–F
                    </button>
                    <button
                      type="button"
                      onClick={() => setGradePieMode("range")}
                      className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors ${
                        gradePieMode === "range"
                          ? "bg-indigo-500/30 text-indigo-100"
                          : "text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      % ranges
                    </button>
                  </div>
                </div>
                {mounted && metrics.totalStudents > 0 ? (
                  <div className="flex min-h-[240px] flex-1 flex-col">
                    {gradePieMode === "letter" && letterGradePieData.length > 0 ? (
                      <ResponsiveContainer width="100%" height={260}>
                        <PieChart>
                          <Pie
                            data={letterGradePieData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={52}
                            outerRadius={88}
                            paddingAngle={2}
                            label={({ name, percent }) =>
                              `${name}: ${((percent ?? 0) * 100).toFixed(1)}%`
                            }
                          >
                            {letterGradePieData.map((entry) => (
                              <Cell key={entry.name} fill={GRADE_SLICE_COLORS[entry.name] ?? "#64748b"} />
                            ))}
                          </Pie>
                          <Tooltip
                            content={({ active, payload }) => {
                              if (!active || !payload?.[0]) return null;
                              const p = payload[0].payload as {
                                name: string;
                                count: number;
                                pct: number;
                              };
                              return (
                                <div className="rounded-lg border border-white/10 bg-slate-950 px-3 py-2 text-xs shadow-xl">
                                  <p className="font-medium text-slate-100">Grade {p.name}</p>
                                  <p className="text-slate-400">
                                    {p.count} students · {p.pct}% of class
                                  </p>
                                </div>
                              );
                            }}
                          />
                          <Legend
                            verticalAlign="bottom"
                            formatter={(value) => <span className="text-slate-300">{value}</span>}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    ) : null}
                    {gradePieMode === "letter" && letterGradePieData.length === 0 ? (
                      <p className="py-8 text-center text-sm text-slate-500">No grade data to show.</p>
                    ) : null}
                    {gradePieMode === "range" && scoreRangePieData.length > 0 ? (
                      <ResponsiveContainer width="100%" height={260}>
                        <PieChart>
                          <Pie
                            data={scoreRangePieData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={52}
                            outerRadius={88}
                            paddingAngle={2}
                            label={({ name, percent }) =>
                              `${name}: ${((percent ?? 0) * 100).toFixed(1)}%`
                            }
                          >
                            {scoreRangePieData.map((entry, index) => (
                              <Cell key={entry.name} fill={entry.fill ?? RANGE_SLICE_COLORS[index % 6]} />
                            ))}
                          </Pie>
                          <Tooltip
                            content={({ active, payload }) => {
                              if (!active || !payload?.[0]) return null;
                              const p = payload[0].payload as {
                                name: string;
                                count: number;
                                pct: number;
                              };
                              return (
                                <div className="rounded-lg border border-white/10 bg-slate-950 px-3 py-2 text-xs shadow-xl">
                                  <p className="font-medium text-slate-100">{p.name}</p>
                                  <p className="mt-1 text-slate-400">
                                    {p.count} students · {p.pct}% of class
                                  </p>
                                </div>
                              );
                            }}
                          />
                          <Legend verticalAlign="bottom" />
                        </PieChart>
                      </ResponsiveContainer>
                    ) : null}
                    {gradePieMode === "range" && scoreRangePieData.length === 0 ? (
                      <p className="py-8 text-center text-sm text-slate-500">No students in these ranges.</p>
                    ) : null}
                  </div>
                ) : null}
              </Card>
            </div>
          </section>

          <Card className="h-80 p-4">
            <p className="mb-3 text-sm text-slate-300">Subject Comparison</p>
            {mounted ? (
              <ResponsiveContainer width="100%" height="90%">
                <BarChart data={metrics.subjectAverages}>
                  <XAxis dataKey="subject" stroke="#94a3b8" />
                  <YAxis stroke="#94a3b8" domain={[0, 100]} />
                  <Tooltip />
                  <Bar dataKey="average" fill="#6366f1" radius={[6, 6, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : null}
          </Card>

          <Card className="p-4">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center">
              <div className="relative w-full md:max-w-sm">
                <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
                <Input
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  className="pl-9"
                  placeholder="Search by name or ID..."
                />
              </div>
              <div className="flex gap-2">
                <Filter className="mt-2 size-4 text-slate-400" />
                {(["all", "Pass", "Fail"] as const).map((item) => (
                  <Button
                    key={item}
                    size="sm"
                    variant={statusFilter === item ? "default" : "outline"}
                    onClick={() => setStatusFilter(item)}
                  >
                    {item}
                  </Button>
                ))}
              </div>
            </div>

            <div className="overflow-x-auto rounded-xl border border-white/10" dir="auto">
              <table className="w-full text-start text-sm">
                <thead className="bg-white/5 text-slate-300">
                  {table.getHeaderGroups().map((headerGroup) => (
                    <tr key={headerGroup.id}>
                      {headerGroup.headers.map((header) => (
                        <th key={header.id} className="px-4 py-3 font-medium">
                          {flexRender(header.column.columnDef.header, header.getContext())}
                        </th>
                      ))}
                    </tr>
                  ))}
                </thead>
                <tbody>
                  {table.getRowModel().rows.map((row) => (
                    <tr
                      key={row.id}
                      onClick={() => setSelectedStudent(row.original)}
                      className="cursor-pointer border-t border-white/10 transition-colors hover:bg-white/5"
                    >
                      {row.getVisibleCells().map((cell) => (
                        <td key={cell.id} className="px-4 py-3">
                          {flexRender(cell.column.columnDef.cell, cell.getContext())}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </main>
      </div>

      <AnimatePresence>
        {selectedStudent ? (
          <motion.div
            className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setSelectedStudent(null)}
          >
            <motion.div
              initial={{ y: 20, opacity: 0 }}
              animate={{ y: 0, opacity: 1 }}
              exit={{ y: 10, opacity: 0 }}
              onClick={(e) => e.stopPropagation()}
              className="w-full max-w-lg rounded-2xl border border-white/15 bg-slate-900/90 p-5 backdrop-blur-xl"
            >
              <div className="mb-4 flex items-start justify-between">
                <div>
                  <p className="text-xs text-slate-400">Student Insight</p>
                  <h3 className="text-xl font-semibold">{selectedStudent.name}</h3>
                  <p className="mt-1 text-xs text-slate-500">
                    ID: {selectedStudent.sheetStudentId?.trim() || "—"} · Rank:{" "}
                    {rankByStudentId.get(selectedStudent.id) ?? "—"} (by current average)
                  </p>
                </div>
                <Button variant="ghost" size="sm" onClick={() => setSelectedStudent(null)}>
                  Close
                </Button>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <MiniStat label="Rank" value={String(rankByStudentId.get(selectedStudent.id) ?? "—")} icon={Trophy} />
                <MiniStat label="Average" value={`${selectedStudent.average.toFixed(1)}%`} icon={FileSpreadsheet} />
                <MiniStat label="Status" value={selectedStudent.status} icon={User2} />
                <MiniStat
                  label="Sheet ID"
                  value={selectedStudent.sheetStudentId?.trim() || "—"}
                  icon={FileSpreadsheet}
                />
              </div>

              <div className="mt-4 space-y-2">
                {Object.entries(selectedStudent.subjects).map(([subject, score]) => (
                  <div key={subject} className="flex items-center justify-between rounded-lg bg-white/5 px-3 py-2 text-sm">
                    <span className="text-slate-300">{subject}</span>
                    <span className="font-medium text-indigo-200">{score.toFixed(1)}%</span>
                  </div>
                ))}
              </div>
            </motion.div>
          </motion.div>
        ) : null}
      </AnimatePresence>
    </div>
  );
}

function StatCard({
  icon: Icon,
  title,
  value,
}: {
  icon: React.ComponentType<{ className?: string }>;
  title: string;
  value: string;
}) {
  return (
    <Card className="p-4">
      <div className="mb-2 flex items-center justify-between">
        <p className="text-sm text-slate-300">{title}</p>
        <Icon className="size-4 text-indigo-300" />
      </div>
      <p className="text-2xl font-semibold">{value}</p>
    </Card>
  );
}

function MiniStat({
  icon: Icon,
  label,
  value,
}: {
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  value: string;
}) {
  return (
    <div className="rounded-xl border border-white/10 bg-white/5 p-3">
      <div className="mb-1 flex items-center gap-2 text-slate-300">
        <Icon className="size-4 text-indigo-300" />
        <span className="text-xs">{label}</span>
      </div>
      <p className="text-lg font-semibold">{value}</p>
    </div>
  );
}
