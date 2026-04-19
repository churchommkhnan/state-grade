/**
 * Header/status fragments that appear on Arabic university grade sheets.
 * Stored only as \\u escapes so the rest of the codebase stays free of Arabic script.
 */

/** "Student name" (common header) */
export const AR_STUDENT_NAME_PRIMARY =
  "\u0627\u0633\u0645\u0020\u0627\u0644\u0637\u0627\u0644\u0628";
/** Alternate spelling with hamza on alif */
export const AR_STUDENT_NAME_ALT =
  "\u0625\u0633\u0645\u0020\u0627\u0644\u0637\u0627\u0644\u0628";

/** "Student status" */
export const AR_STUDENT_STATUS =
  "\u062d\u0627\u0644\u0629\u0020\u0627\u0644\u0637\u0627\u0644\u0628";

/** "Percentage" (column title) */
export const AR_PERCENT_WORD =
  "\u0627\u0644\u0646\u0633\u0628\u0629";
/** Alternate spelling (ta marbuta vs ha) */
export const AR_PERCENT_WORD_ALT =
  "\u0627\u0644\u0646\u0633\u0628\u0647";

/** Arabic percent sign */
export const AR_PERCENT_SIGN = "\u066a";

/** "Percentage" (full phrase) */
export const AR_PERCENT_PHRASE_LONG =
  "\u0627\u0644\u0646\u0633\u0628\u0629\u0020\u0627\u0644\u0645\u0626\u0648\u064a\u0629";
export const AR_PERCENT_PHRASE_SHORT =
  "\u0646\u0633\u0628\u0629\u0020\u0645\u0626\u0648\u064a\u0629";

/** "Total" */
export const AR_TOTAL_PRIMARY =
  "\u0627\u0644\u0645\u062c\u0645\u0648\u0639";
export const AR_TOTAL_ALT =
  "\u0625\u062c\u0645\u0627\u0644\u064a";

/** "Theoretical" (theory exam) */
export const AR_THEORY_DEFINITE =
  "\u0627\u0644\u0646\u0638\u0631\u064a";
/** "Theory" without article */
export const AR_THEORY_SHORT =
  "\u0646\u0638\u0631\u064a";

/** Midterm spellings (substring checks) */
export const AR_MIDTERM_WITH_AL =
  "\u0627\u0644\u0645\u064a\u062f\u062a\u064a\u0631\u0645";
export const AR_MIDTERM_WITH_AL_ALT =
  "\u0627\u0644\u0645\u064a\u062f\u062a\u0631\u0645";
export const AR_MIDTERM_PLAIN =
  "\u0645\u064a\u062f\u062a\u064a\u0631\u0645";
export const AR_MIDTERM_PLAIN_SHORT =
  "\u0645\u064a\u062f\u062a\u0631\u0645";

/** Quiz */
export const AR_QUIZ =
  "\u0643\u0648\u064a\u0632";
export const AR_QUIZ_PLURAL =
  "\u0643\u0648\u064a\u0632\u0627\u062a";

/** "Serial" / "sequence" row index columns */
export const AR_META_SERIAL = "\u0645\u0633\u0644\u0633\u0644";
export const AR_META_ORDER = "\u062a\u0633\u0644\u0633\u0644";

/** ID column title fragments */
export const AR_ID_MARKERS: readonly string[] = [
  "\u0631\u0642\u0645\u0020\u0627\u0644\u0637\u0627\u0644\u0628",
  "\u0631\u0642\u0645\u0020\u0627\u0644\u0637\u0627\u0644\u0628\u0647",
  "\u0643\u0648\u062f\u0020\u0627\u0644\u0637\u0627\u0644\u0628",
  "\u0627\u0644\u0631\u0642\u0645\u0020\u0627\u0644\u062c\u0627\u0645\u0639\u064a",
  "\u0627\u0644\u0631\u0642\u0645\u0020\u0627\u0644\u062c\u0627\u0645\u0639\u0649",
  "\u0631\u0642\u0645\u0020\u0627\u0644\u062c\u0627\u0645\u0639\u064a",
  "\u0631\u0642\u0645\u0020\u0627\u0644\u062c\u0627\u0645\u0639\u0649",
  "\u0631\u0642\u0645\u0020\u0627\u0644\u062c\u0644\u0648\u0633",
  "\u0627\u0644\u062c\u0644\u0648\u0633",
  "\u0631\u0642\u0645\u0020\u0627\u0644\u0642\u0648\u0645\u064a",
  "\u0627\u0644\u0631\u0642\u0645\u0020\u0627\u0644\u0642\u0648\u0645\u064a",
  "\u0627\u0644\u0631\u0642\u0645\u0020\u0627\u0644\u0648\u0637\u0646\u064a",
  AR_META_SERIAL,
  AR_META_ORDER,
];

/** Row markers used to locate the header row */
export const AR_HEADER_MARKERS: readonly string[] = [
  AR_STUDENT_NAME_PRIMARY,
  AR_STUDENT_STATUS,
  AR_PERCENT_WORD,
];

/** Footer / summary row substrings */
export const AR_FOOTER_MARKERS: readonly string[] = [
  "\u0625\u062c\u0645\u0627\u0644\u064a\u0020\u0627\u0644\u0637\u0644\u0627\u0628",
  "\u0625\u062c\u0645\u0627\u0644\u064a\u0020\u0627\u0644\u062f\u0631\u062c\u0627\u062a",
  "\u0627\u0644\u0625\u062d\u0635\u0627\u0626\u064a\u0627\u062a",
  "\u0625\u062d\u0635\u0627\u0626\u064a\u0627\u062a",
  "\u0627\u062d\u0635\u0627\u0626\u064a\u0627\u062a",
  "\u0645\u0644\u062e\u0635",
];

/** Pass/fail cell fragments */
export const AR_STATUS_NOT_PASS =
  "\u063a\u064a\u0631";
export const AR_STATUS_PASS =
  "\u0646\u0627\u062c\u062d";
export const AR_STATUS_FAIL =
  "\u0631\u0627\u0633\u0628";

/** Regex: colloquial split spelling of "midterm" */
export const RE_MIDTERM_SPLIT_SPELLING = new RegExp(
  "\u0645\u064a\u062f\\s*\u064a?\\s*\u062a?\\s*\u064a?\\s*\u0631\u0645"
);
export const RE_MIDTERM_COMPACT_SPELLING = new RegExp("\u0645\u064a\u062f\\s*\u062a\u0631\u0645");

/** Regex: half-semester / mid-semester exam phrases */
export const RE_MIDTERM_HALF_SEMESTER = new RegExp(
  "\u0645\u0646\u062a\u0635\u0641\\s*\u0627\u0644\u0641\u0635\u0644|\u0627\u0645\u062a\u062d\u0627\u0646\\s*\u0645\u0646\u062a\u0635\u0641|\u0627\u062e\u062a\u0628\u0627\u0631\\s*\u0645\u0646\u062a\u0635\u0641|\u0646\u0635\u0641\\s*\u0627\u0644\u0641\u0635\u0644|\u0646\u0635\u0641\\s*\u0627\u0644\u062a\u0631\u0645|\u0627\u0644\u0646\u0635\u0641\u064a|\u0627\u062e\u062a\u0628\u0627\u0631\\s*\u0646\u0635\u0641\u064a",
  "i"
);

export const RE_QUIZ_WORD = new RegExp(
  "\u0643\u0648\u064a\u0632|\u0643\u0648\u064a\u0632\u0627\u062a"
);

/** Quiz / homework-style column titles */
export const RE_QUIZ_EXTENDED = new RegExp(
  "\u0627\u062e\u062a\u0628\u0627\u0631\\s*\u0642\u0635\u064a\u0631|\u062a\u0642\u064a\u064a\u0645\\s*\u0645\u0633\u062a\u0645\u0631|\u0648\u0627\u062c\u0628\u0627\u062a?|\u062a\u0643\u0627\u0644\u064a\u0641|\u0646\u0634\u0627\u0637|\u0627\u0645\u062a\u062d\u0627\u0646\\s*\u0634\u0647\u0631\u064a|short\\s*test|class\\s*work|homework",
  "i"
);

/** "Theory" before opening paren */
export const RE_THEORY_BEFORE_PAREN = new RegExp(
  "(^|[\\s\u203a])\u0646\u0638\u0631\u064a\\s*\\("
);

/** Written final / theory exam (Arabic) */
export const RE_THEORY_WRITTEN_PHRASES = new RegExp(
  "\u062a\u062d\u0631\u064a\u0631\u064a|\u0627\u0644\u062a\u062d\u0631\u064a\u0631\u064a|\u0627\u0645\u062a\u062d\u0627\u0646\\s*\u0646\u0647\u0627\u0626\u064a|\u0627\u0644\u0627\u0645\u062a\u062d\u0627\u0646\\s*\u0627\u0644\u0646\u0647\u0627\u0626\u064a|\u0646\u0647\u0627\u0626\u064a\\s*\u0627\u0644\u0646\u0638\u0631\u064a|\u062f\u0631\u062c\u0629\\s*\u0627\u0644\u0646\u0638\u0631\u064a",
  "i"
);

/** Year-work / continuous assessment banner */
export const RE_YEAR_WORK_BANNER = new RegExp(
  "\u0623\u0639\u0645\u0627\u0644\\s*\u0627\u0644\u0633\u0646\u0629|\u0627\u0639\u0645\u0627\u0644\\s*\u0627\u0644\u0633\u0646\u0629|year\\s*work|continuous\\s*assessment",
  "i"
);

/**
 * Any of these substrings suggests a scored component column (Arabic university sheets).
 */
export const RE_GRADE_COMPONENT_HINT = new RegExp(
  [
    "\u0646\u0638\u0631\u064a",
    "\u0639\u0645\u0644\u064a",
    "\u0645\u064a\u062f\u062a\u064a\u0631\u0645",
    "\u0645\u064a\u062f \u062a\u0631\u0645",
    "\u0645\u064a\u062f\u062a\u0631\u0645",
    "\u0643\u0648\u064a\u0632",
    "\u062a\u062d\u0631\u064a\u0631\u064a",
    "\u0634\u0641\u0648\u064a",
    "\u0623\u0639\u0645\u0627\u0644",
    "\u0627\u0639\u0645\u0627\u0644",
    "\u0633\u0646\u0648\u064a",
    "\u0633\u0646\u0647",
    "\u0627\u0645\u062a\u062d\u0627\u0646",
    "\u062a\u0637\u0628\u064a\u0642\u064a",
    "\u0634\u0627\u0645\u0644",
    "\u0641\u0635\u0644\u064a",
  ].join("|")
);
