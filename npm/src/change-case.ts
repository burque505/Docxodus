// Change Case's pure part: the new text for a case mode. The editor (DocxEditor.changeCase)
// reads the selected text run by run and writes back only the characters that changed, so each
// keeps its formatting. Only letters change, one for one: the text keeps its length (a letter
// whose other case is two letters, as ß's capital SS, stays as it is).
//
// Headline Case is AhkWord's (rte_headline.cs): every word capitalized as Chicago does it, except
//   - the small words (a, the, and, of, to ...) inside a line: the first and the last word of a
//     line, and the first after : ; ? ! a dash or a sentence's period, are capitalized whatever
//     they are;
//   - words that are always lower case (v., vs., et al., a/k/a, e.g.), wherever they are but first;
//   - words that are always as written (LLC, L.L.C., Inc., the state codes that are no English word);
//   - words already in capitals (USA, NOTE) and in mixed case (McDonald, iPhone) kept, unless the
//     whole line is in capitals ("DEED OF TRUST" becomes "Deed of Trust"); initials with periods
//     (u.s.) and Roman numerals (iii) in capitals; ordinals (1st, 2ND) in lower case after the
//     number; each part of a hyphenated or slashed word on its own (Self-Employed, Run-of-the-Mill,
//     and/or); O' and D' (O'Brien).

/** Word's Change Case modes, and AhkWord's Headline Case. */
export type CaseMode = "upper" | "lower" | "sentence" | "title" | "toggle" | "headline";

export const CASE_MODES: readonly CaseMode[] = ["sentence", "lower", "upper", "title", "toggle", "headline"];

/** Headline Case's word lists. */
export interface HeadlineCaseLists {
  /** Lower case inside a line. */
  small: string[];
  /** Always lower case, except as a line's first word. */
  lower: string[];
  /** Always as written. */
  exact: string[];
}

/** The lists as a host keeps them in a file to be edited (AhkWord's headline-case.txt): sections
 *  `[small]`, `[lower]` and `[exact]`, words separated by spaces or lines, `#` starts a comment. */
export const DEFAULT_HEADLINE_CASE_TEXT = `# Headline Case's word lists (Change Case > Headline Case). Edit them as you like; the file is
# read again when it has changed. Words are separated by spaces or lines; # starts a comment.

[small]
# Lower case inside a line; capitalized as the first or last word of a line, or after : . ; ? ! or a dash.
a an the
and but for nor or
as at by en from in into of off on onto out over per re than to up upon via with

[lower]
# Always lower case, wherever they are (except as a line's first word).
v. vs. et al. etc. e.g. i.e. a/k/a d/b/a f/k/a n/k/a c/o a.m. p.m.

[exact]
# Always as written here, wherever they are.
LLC L.L.C. LLP L.L.P. PLLC Inc. Inc Corp. Co. Ltd. L.P. LP P.C. P.A. N.A. FSB
ISAOA ATIMA MERS HUD FHA VA USDA HOA PUD IRS EIN SSN UCC DBA AKA FKA
U.S. USA U.S.A. D.C. DC
AK AZ AR CA CT FL GA IA KS KY MI MN MO MT NC ND NE NH NJ NM NV NY RI SC SD TN TX UT VT WA WI WV WY PR
`;

/** Read the lists from a file's text (see {@link DEFAULT_HEADLINE_CASE_TEXT}). */
export function parseHeadlineCaseLists(text: string): HeadlineCaseLists {
  const out: HeadlineCaseLists = { small: [], lower: [], exact: [] };
  let section = "";
  for (const raw of text.split("\n")) {
    const line = raw.split("#")[0]!.trim();
    if (line.startsWith("[") && line.endsWith("]")) { section = line.slice(1, -1).trim().toLowerCase(); continue; }
    for (const word of line.split(/\s+/).filter(Boolean)) {
      if (section === "small" || section === "lower" || section === "exact") out[section].push(word);
    }
  }
  return out;
}

export const DEFAULT_HEADLINE_CASE_LISTS: HeadlineCaseLists = parseHeadlineCaseLists(DEFAULT_HEADLINE_CASE_TEXT);

// One character's other case, or the character itself when that would change the text's length.
const up = (c: string): string => { const u = c.toUpperCase(); return u.length === c.length ? u : c; };
const low = (c: string): string => { const l = c.toLowerCase(); return l.length === c.length ? l : c; };
const mapChars = (s: string, f: (c: string) => string): string => s.split("").map(f).join("");
const isUpper = (c: string): boolean => c !== c.toLowerCase() && c === c.toUpperCase();
const isLower = (c: string): boolean => c !== c.toUpperCase() && c === c.toLowerCase();
const isLetter = (c: string): boolean => /\p{L}/u.test(c);

/** Lines: paragraphs (`\n`) and anything below a space but a tab or a line break (`\v`). */
function eachLine(text: string, f: (line: string, offset: number) => void): void {
  let start = 0;
  for (let i = 0; i <= text.length; i++) {
    if (i < text.length && (text[i]! >= " " || text[i] === "\v" || text[i] === "\t")) continue;
    f(text.slice(start, i), start);
    start = i + 1;
  }
}

/**
 * Word's Shift+F3: the mode that comes next for `text` — all in capitals goes to lower case, all in
 * lower case to Capitalize Each Word, anything else to capitals.
 */
export function nextCycleCase(text: string): CaseMode {
  const letters = text.split("").filter(isLetter);
  if (letters.length > 0 && letters.every((c) => !isLower(c))) return "lower";
  if (letters.length > 0 && letters.every((c) => !isUpper(c))) return "title";
  return "upper";
}

/** `text` in `mode`, the same length; `hidden` marks characters to pass over (left as they are,
 *  read as spaces). */
export function changeCase(text: string, mode: CaseMode, lists: HeadlineCaseLists = DEFAULT_HEADLINE_CASE_LISTS, hidden?: boolean[]): string {
  const seen = hidden ? text.split("").map((c, i) => (hidden[i] ? " " : c)).join("") : text;
  let out: string;
  switch (mode) {
    case "upper": out = mapChars(seen, up); break;
    case "lower": out = mapChars(seen, low); break;
    case "toggle": out = mapChars(seen, (c) => (isUpper(c) ? low(c) : isLower(c) ? up(c) : c)); break;
    case "title": out = titleCase(seen); break;
    case "sentence": out = sentenceCase(seen); break;
    case "headline": out = headlineCase(seen, lists); break;
  }
  if (!hidden) return out;
  return out.split("").map((c, i) => (hidden[i] ? text[i]! : c)).join("");
}

// Capitalize Each Word: a word's first letter in capitals, the rest in lower case (as Word does it).
function titleCase(text: string): string {
  let inWord = false;
  return text.split("").map((c) => {
    const letter = /[\p{L}\p{N}\p{M}'’]/u.test(c);
    const r = letter ? (inWord ? low(c) : up(c)) : c;
    inWord = letter;
    return r;
  }).join("");
}

// Sentence case: lower case, a sentence's first letter in capitals (the selection's and each
// line's start, and after . ! ? and a space). "I" on its own stays a capital.
function sentenceCase(text: string): string {
  const chars = mapChars(text, low).split("");
  let start = true;
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i]!;
    if (c === "\n" || c === "\v" || c < " " && c !== "\t") { start = true; continue; }
    if (isLetter(c)) {
      if (start) chars[i] = up(c);
      start = false;
    } else if (/[.!?]/.test(c)) {
      start = i + 1 >= chars.length || /\s/.test(chars[i + 1]!);
    }
  }
  // (the pronoun I)
  return chars.join("").replace(/(^|[^\p{L}\p{N}])i(?=$|[\s,;:!?)])/gu, (m, before: string) => `${before}I`)
    .replace(/(^|[^\p{L}\p{N}])i(?=['’](m|ll|ve|d)\b)/gu, (m, before: string) => `${before}I`);
}

const WORD = /[\p{L}\p{N}]+(?:['’./&-][\p{L}\p{N}]+)*\.?/gu;
const ROMAN = /^(?=[ivx]{2,}$)x{0,3}(ix|iv|v?i{0,3})$/i;
const ORDINAL = /^\d+(st|nd|rd|th|d)\.?$/i;
const INITIALS = /^(\p{L}\.){2,}$|^(\p{L}\.)+\p{L}$/u;

function headlineCase(text: string, lists: HeadlineCaseLists): string {
  const small = new Set(lists.small.map((w) => w.toLowerCase()));
  const lower = new Set(lists.lower.map((w) => w.toLowerCase()));
  const exact = new Map(lists.exact.map((w) => [w.toLowerCase(), w]));
  const result = text.split("");
  const lowerS = (s: string) => mapChars(s, low);
  const upperS = (s: string) => mapChars(s, up);

  // In capitals (USA), or in mixed case (McDonald, iPhone): as written.
  const kept = (word: string): boolean => {
    let u = false, l = false, upperAfterFirst = false;
    for (let i = 0; i < word.length; i++) {
      if (isUpper(word[i]!)) { u = true; if (i > 0) upperAfterFirst = true; }
      else if (isLower(word[i]!)) l = true;
    }
    return (u && !l && word.split("").filter(isLetter).length > 1) || (l && upperAfterFirst);
  };

  const single = (word: string, edge: boolean): string => {
    const bare = word.replace(/\.+$/, "");
    const ex = exact.get(word.toLowerCase()) ?? exact.get(bare.toLowerCase());
    if (ex !== undefined) return ex + word.slice(ex.length);
    if (!edge && (small.has(word.toLowerCase()) || small.has(bare.toLowerCase()))) return lowerS(word);
    if (ROMAN.test(bare)) return upperS(word);
    const chars = lowerS(word).split("");
    chars[0] = up(chars[0]!);
    // O'Brien, D'Angelo
    if (chars.length > 2 && (chars[1] === "'" || chars[1] === "’") && "OD".includes(chars[0]!)) chars[2] = up(chars[2]!);
    return chars.join("");
  };

  const caseWord = (word: string, first: boolean, edge: boolean, allCaps: boolean): string => {
    const bare = word.endsWith(".") ? word.slice(0, -1) : word;
    const key = word.toLowerCase(), bareKey = bare.toLowerCase();
    const ex = exact.get(key);
    if (ex !== undefined) return ex;
    if (bare !== word && exact.has(bareKey)) return exact.get(bareKey)! + ".";
    if (ORDINAL.test(word)) return lowerS(word);
    if (/\p{N}/u.test(word[0]!)) return word;
    if (!allCaps && kept(bare)) return word;
    if (lower.has(key) || lower.has(bareKey)) return first ? single(word, true) : lowerS(word);
    if (INITIALS.test(word)) return upperS(word);
    if (bare.search(/[-/]/) > 0) {
      // Each part on its own: the first as the word would be, the last capitalized after a hyphen
      // (Made-Up, Built-In), those between unless small (Run-of-the-Mill).
      let parts = "";
      let start = 0;
      for (let i = 0; i <= word.length; i++) {
        if (i < word.length && word[i] !== "-" && word[i] !== "/") continue;
        const part = word.slice(start, i);
        parts += part.length === 0 ? part : single(part, start === 0 ? edge : i === word.length && word[start - 1] === "-");
        if (i < word.length) parts += word[i];
        start = i + 1;
      }
      return parts;
    }
    return single(word, edge);
  };

  eachLine(text, (line, offset) => {
    const words = Array.from(line.matchAll(WORD));
    if (words.length === 0) return;
    const allCaps = !line.split("").some(isLower);
    let phraseStart = true;
    for (let w = 0; w < words.length; w++) {
      const m = words[w]!;
      if (w > 0) {
        const prev = words[w - 1]!;
        const between = line.slice(prev.index! + prev[0].length, m.index!);
        if (/[:;?!—–]/.test(between) || between.includes(" - ")) phraseStart = true;
      }
      const cased = caseWord(m[0], w === 0, phraseStart || w === words.length - 1, allCaps);
      for (let k = 0; k < cased.length; k++) result[offset + m.index! + k] = cased[k]!;
      // A period that ends a sentence starts a phrase: not an abbreviation's, an initial's or one in the lists.
      const v = m[0], vk = v.toLowerCase();
      phraseStart = v.endsWith(".") && v.length > 2 && !v.slice(0, -1).includes(".")
        && !exact.has(vk) && !lower.has(vk) && !small.has(vk);
    }
  });
  return result.join("");
}
