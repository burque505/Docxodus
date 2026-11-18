/**
 * Find options as in Word's Find and Replace (More >>): a query and its options compile to one
 * regular expression that runs over a "stream" of a story's text, in which
 *
 *   "\r"    is a paragraph mark (between two paragraphs of the same story; never after the last)
 *   "\v"    is a manual line break (Word's ^l, ^11)
 *   "\f"    is a page break (^m, ^12)
 *   "\x0E"  is a column break (^n, ^14)
 *   "\t"    is a tab
 *
 * the same characters Word itself uses for them, so ^13 / ^p and a regex's \r mean the same mark.
 *
 * Three ways to read a query, as in Word:
 *  - plain text, with Word's ^ codes (^p ^t ^? ^# ^$ ^w ...) and the options Match case, Find whole
 *    words only, Match prefix, Match suffix, Ignore punctuation, Ignore white-space;
 *  - Word wildcards (Use wildcards): ? * @ < > [ ] [! ] {n,m} ( ) \ — always case-sensitive, and the
 *    other options do not apply, as in Word;
 *  - regular expressions (JavaScript's, which AutoHotkey users know as PCRE): ^ and $ are paragraph
 *    starts and ends, \n and \R a paragraph mark, an AutoHotkey "i)" options prefix is accepted.
 *    Match case still applies; the other options do not.
 *
 * Pure: no DOM. The editor builds the stream and maps hits back to blocks.
 */

export interface FindOptions {
  matchCase?: boolean;
  /** Find whole words only. */
  wholeWord?: boolean;
  /** The hit starts a word. */
  matchPrefix?: boolean;
  /** The hit ends a word. */
  matchSuffix?: boolean;
  /** Punctuation in the text and in the query is skipped: "e-mail" finds "email" and "e.mail". */
  ignorePunctuation?: boolean;
  /** White space in the text and in the query is skipped: "ice cream" finds "icecream". */
  ignoreSpace?: boolean;
  /** Word's wildcards. */
  wildcards?: boolean;
  /** Regular expressions. Wins over wildcards when both are set. */
  regex?: boolean;
}

/** The option names, in the order the Options menu lists them. */
export const FIND_OPTION_KEYS = [
  "matchCase", "wholeWord", "matchPrefix", "matchSuffix", "ignorePunctuation", "ignoreSpace", "wildcards", "regex",
] as const satisfies ReadonlyArray<keyof FindOptions>;

/** Which options a query mode leaves usable (Word greys the rest out). */
export function findOptionApplies(key: keyof FindOptions, options: FindOptions): boolean {
  if (key === "wildcards" || key === "regex") return true;
  if (options.regex) return key === "matchCase";
  if (options.wildcards) return false;
  return true;
}

export type CompiledFind = { ok: true; regex: RegExp } | { ok: false; error: string };

/** A Word ^ code as the Special menu offers it. */
export interface SpecialCode {
  code: string;
  label: string;
}

/** The Special menu for the Find field, by mode (Word's lists, the ones this editor can find). */
export function specialFindCodes(options: FindOptions): SpecialCode[] {
  if (options.regex) {
    return [
      { code: "\\n", label: "Paragraph Mark" },
      { code: "\\t", label: "Tab Character" },
      { code: "\\v", label: "Manual Line Break" },
      { code: "\\f", label: "Manual Page Break" },
      { code: ".", label: "Any Character" },
      { code: "\\d", label: "Any Digit" },
      { code: "\\p{L}", label: "Any Letter" },
      { code: "\\s+", label: "White Space" },
      { code: "\\b", label: "Word Boundary" },
      { code: "^", label: "Start of Paragraph" },
      { code: "$", label: "End of Paragraph" },
      { code: "( )", label: "Group" },
    ];
  }
  if (options.wildcards) {
    return [
      { code: "^13", label: "Paragraph Mark" },
      { code: "^t", label: "Tab Character" },
      { code: "^l", label: "Manual Line Break" },
      { code: "^m", label: "Manual Page Break" },
      { code: "?", label: "Any Character" },
      { code: "*", label: "Any Characters" },
      { code: "[0-9]", label: "Any Digit" },
      { code: "[A-Za-z]", label: "Any Letter" },
      { code: "@", label: "One or More of the Previous" },
      { code: "{1,3}", label: "Occurrences of the Previous" },
      { code: "[ ]", label: "In Range" },
      { code: "[! ]", label: "Not" },
      { code: "<", label: "Beginning of Word" },
      { code: ">", label: "End of Word" },
      { code: "( )", label: "Expression" },
      { code: "\\", label: "Escape Character" },
    ];
  }
  return [
    { code: "^p", label: "Paragraph Mark" },
    { code: "^t", label: "Tab Character" },
    { code: "^?", label: "Any Character" },
    { code: "^#", label: "Any Digit" },
    { code: "^$", label: "Any Letter" },
    { code: "^^", label: "Caret Character" },
    { code: "^+", label: "Em Dash" },
    { code: "^=", label: "En Dash" },
    { code: "^l", label: "Manual Line Break" },
    { code: "^m", label: "Manual Page Break" },
    { code: "^n", label: "Column Break" },
    { code: "^~", label: "Nonbreaking Hyphen" },
    { code: "^s", label: "Nonbreaking Space" },
    { code: "^-", label: "Optional Hyphen" },
    { code: "^w", label: "White Space" },
  ];
}

/** The Special menu for the Replace field, by mode. */
export function specialReplaceCodes(options: FindOptions): SpecialCode[] {
  if (options.regex) {
    return [
      { code: "\\n", label: "Paragraph Mark" },
      { code: "\\t", label: "Tab Character" },
      { code: "\\v", label: "Manual Line Break" },
      { code: "$&", label: "Find What Text" },
      { code: "$1", label: "Group 1" },
      { code: "$$", label: "Dollar Sign" },
    ];
  }
  const common: SpecialCode[] = [
    { code: "^p", label: "Paragraph Mark" },
    { code: "^t", label: "Tab Character" },
    { code: "^^", label: "Caret Character" },
    { code: "^&", label: "Find What Text" },
    { code: "^+", label: "Em Dash" },
    { code: "^=", label: "En Dash" },
    { code: "^l", label: "Manual Line Break" },
    { code: "^~", label: "Nonbreaking Hyphen" },
    { code: "^s", label: "Nonbreaking Space" },
    { code: "^-", label: "Optional Hyphen" },
  ];
  if (options.wildcards) common.splice(4, 0, { code: "\\1", label: "Expression 1" });
  return common;
}

// ── character classes ────────────────────────────────────────────────────────

const WORD_CHAR = "[\\p{L}\\p{N}\\p{M}_]";
/** Punctuation and symbols Word's Ignore punctuation skips. */
const PUNCT_CLASS = "\\p{P}\\p{S}";
/** White space inside a paragraph (a paragraph mark is never skipped). */
const SPACE_CLASS = " \\t\\v\\u00a0\\u1680\\u2000-\\u200b\\u202f\\u205f\\u3000";
const PUNCT_RE = /[\p{P}\p{S}]/u;
const SPACE_RE = /[ \t\v\u00a0\u1680\u2000-\u200b\u202f\u205f\u3000]/u;

function escapeLiteral(ch: string): string {
  return /[\\^$.*+?()[\]{}|/]/.test(ch) ? "\\" + ch : ch;
}

/** A character code as a regex atom (\uXXXX, or \u{...} past the BMP). */
function codeAtom(cp: number): string {
  return cp > 0xffff ? `\\u{${cp.toString(16)}}` : `\\u${cp.toString(16).padStart(4, "0")}`;
}

// ── plain text with ^ codes ──────────────────────────────────────────────────

/** One unit of a plain query: a literal character, or a pattern a ^ code stands for. */
type Token = { lit: string } | { re: string };

/** Read ^nnn (decimal, ^0nnn too) or ^unnnn (Unicode) at `query[i]` (just after the ^). */
function readNumericCode(query: string, i: number): { cp: number; next: number } | null {
  if (query[i] === "u" || query[i] === "U") {
    const m = /^[0-9]{1,6}/.exec(query.slice(i + 1));
    if (!m) return null;
    return { cp: Number.parseInt(m[0], 10), next: i + 1 + m[0].length };
  }
  const m = /^[0-9]{1,4}/.exec(query.slice(i));
  if (!m) return null;
  let cp = Number.parseInt(m[0], 10);
  // Word's ^0nnn is a Windows (ANSI) character: ^0151 is an em dash.
  if (m[0].startsWith("0") && cp >= 0x80 && cp <= 0x9f) cp = CP1252[cp - 0x80] || cp;
  return { cp, next: i + m[0].length };
}

/** Windows-1252's 0x80-0x9F, as Unicode code points (0 where it has no character). */
const CP1252 = [
  0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017d, 0,
  0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0, 0x017e, 0x0178,
];

/** The character a code stands for in the stream (shared by find and replace). */
const CODE_CHARS: Record<string, string> = {
  p: "\r", t: "\t", l: "\v", m: "\f", n: "\x0E", "^": "^", "+": "\u2014", "=": "\u2013",
  "~": "\u2011", s: "\u00a0", "-": "\u00ad",
};

/** Word's numeric codes for the stream's marks map to themselves (^13 = \r, ^11 = \v, ^12 = \f). */
function numericChar(cp: number): string | null {
  if (cp < 0 || cp > 0x10ffff) return null;
  return String.fromCodePoint(cp);
}

function tokenizePlain(query: string): Token[] | string {
  const out: Token[] = [];
  for (let i = 0; i < query.length; i++) {
    const ch = query[i];
    if (ch !== "^" || i === query.length - 1) {
      // Astral characters stay whole.
      const cp = query.codePointAt(i)!;
      const whole = String.fromCodePoint(cp);
      out.push({ lit: whole });
      i += whole.length - 1;
      continue;
    }
    const code = query[i + 1];
    const lower = code.toLowerCase();
    if (/[0-9uU]/.test(code)) {
      const num = readNumericCode(query, i + 1);
      const char = num ? numericChar(num.cp) : null;
      if (!num || char === null) return `^${code} is not a code Find knows`;
      out.push({ lit: char });
      i = num.next - 1;
      continue;
    }
    i++;
    if (code === "?") out.push({ re: "[^\\r]" });
    else if (code === "#") out.push({ re: "[0-9]" });
    else if (code === "$") out.push({ re: "\\p{L}" });
    else if (lower === "w") out.push({ re: `[${SPACE_CLASS}]+` });
    else if (CODE_CHARS[lower] !== undefined) out.push({ lit: CODE_CHARS[lower] });
    else return `^${code} is not a code Find knows`;
  }
  return out;
}

function compilePlain(query: string, options: FindOptions): CompiledFind {
  const tokens = tokenizePlain(query);
  if (typeof tokens === "string") return { ok: false, error: tokens };
  const skipPunct = !!options.ignorePunctuation, skipSpace = !!options.ignoreSpace;
  const kept = tokens.filter((t) => !("lit" in t)
    || !((skipPunct && PUNCT_RE.test(t.lit)) || (skipSpace && SPACE_RE.test(t.lit))));
  if (kept.length === 0) return { ok: false, error: "nothing to find once punctuation and spaces are ignored" };
  const gap = skipPunct || skipSpace
    ? `[${skipPunct ? PUNCT_CLASS : ""}${skipSpace ? SPACE_CLASS : ""}]*`
    : "";
  let body = kept.map((t) => ("lit" in t ? Array.from(t.lit).map(escapeLiteral).join("") : t.re)).join(gap);
  const startsWord = options.wholeWord || options.matchPrefix;
  const endsWord = options.wholeWord || options.matchSuffix;
  if (startsWord) body = `(?<!${WORD_CHAR})(?:${body})`;
  if (endsWord) body = `(?:${body})(?!${WORD_CHAR})`;
  return build(body, options.matchCase ? "gmu" : "gimu");
}

// ── Word wildcards ───────────────────────────────────────────────────────────

function compileWildcards(query: string): CompiledFind {
  let out = "";
  for (let i = 0; i < query.length; i++) {
    const ch = query[i];
    if (ch === "\\") {
      if (i === query.length - 1) return { ok: false, error: "\\ needs a character after it" };
      const next = query[++i];
      // \1..\9 in Find repeat a group, as in Word.
      out += /[1-9]/.test(next) ? `\\${next}` : escapeLiteral(next);
    } else if (ch === "^") {
      const code = query[i + 1];
      if (code === undefined) { out += "\\^"; continue; }
      if (/[0-9uU]/.test(code)) {
        const num = readNumericCode(query, i + 1);
        const char = num ? numericChar(num.cp) : null;
        if (!num || char === null) return { ok: false, error: `^${code} is not a code Find knows` };
        out += codeAtom(char.codePointAt(0)!);
        i = num.next - 1;
      } else if (code.toLowerCase() === "w") {
        out += `[${SPACE_CLASS}]+`;
        i++;
      } else if (CODE_CHARS[code.toLowerCase()] !== undefined) {
        out += codeAtom(CODE_CHARS[code.toLowerCase()].codePointAt(0)!);
        i++;
      } else {
        return { ok: false, error: `^${code} is not a code Find knows` };
      }
    } else if (ch === "?") {
      out += "[^\\r]";
    } else if (ch === "*") {
      out += "[^\\r]*?";
    } else if (ch === "@") {
      out += "+";
    } else if (ch === "<") {
      out += `(?<!${WORD_CHAR})(?=${WORD_CHAR})`;
    } else if (ch === ">") {
      out += `(?<=${WORD_CHAR})(?!${WORD_CHAR})`;
    } else if (ch === "{") {
      const m = /^\{(\d*)(?:([,;])(\d*))?\}/.exec(query.slice(i));
      if (!m || (m[1] === "" && m[3] === undefined)) return { ok: false, error: "{ } needs a count, as {2} or {1,3}" };
      out += m[2] ? `{${m[1] || "0"},${m[3] ?? ""}}` : `{${m[1]}}`;
      i += m[0].length - 1;
    } else if (ch === "[") {
      const close = findClassEnd(query, i);
      if (close < 0) return { ok: false, error: "[ has no closing ]" };
      let inner = query.slice(i + 1, close);
      let negate = false;
      if (inner.startsWith("!")) { negate = true; inner = inner.slice(1); }
      let cls = "";
      for (let k = 0; k < inner.length; k++) {
        const c = inner[k];
        if (c === "\\" && k + 1 < inner.length) cls += "\\" + inner[++k];
        else if (c === "-" && k > 0 && k < inner.length - 1) cls += "-";
        else cls += /[\\\]\[^-]/.test(c) ? "\\" + c : c;
      }
      // Word's [!...] never matches a paragraph mark either.
      out += negate ? `[^\\r${cls}]` : `[${cls}]`;
      i = close;
    } else if (ch === "(") {
      out += "(";
    } else if (ch === ")") {
      out += ")";
    } else {
      const cp = query.codePointAt(i)!;
      const whole = String.fromCodePoint(cp);
      out += Array.from(whole).map(escapeLiteral).join("");
      i += whole.length - 1;
    }
  }
  return build(out, "gmu");
}

/** Index of the ] closing the class opened at `open` (a ] right after [ or [! is literal). */
function findClassEnd(query: string, open: number): number {
  let k = open + 1;
  if (query[k] === "!") k++;
  if (query[k] === "]") k++;
  for (; k < query.length; k++) {
    if (query[k] === "\\") { k++; continue; }
    if (query[k] === "]") return k;
  }
  return -1;
}

// ── regular expressions ──────────────────────────────────────────────────────

/** \n and \R (a newline in AutoHotkey's PCRE) are the stream's paragraph mark, \r. */
function translateRegex(source: string): string {
  let out = "";
  for (let i = 0; i < source.length; i++) {
    const ch = source[i];
    if (ch !== "\\" || i === source.length - 1) { out += ch; continue; }
    const next = source[++i];
    if (next === "n" || next === "R") out += "\\r";
    else out += "\\" + next;
  }
  return out;
}

function compileRegex(query: string, options: FindOptions): CompiledFind {
  let source = query;
  let caseless = !options.matchCase;
  let dotAll = false;
  // AutoHotkey's options prefix: "i)abc". Only the options JavaScript can honour.
  const prefix = /^([imsxU`]*)\)/.exec(source);
  if (prefix && prefix[1].length > 0 && /^[ims]+$/.test(prefix[1])) {
    if (prefix[1].includes("i")) caseless = true;
    if (prefix[1].includes("s")) dotAll = true;
    source = source.slice(prefix[0].length);
  }
  if (!source) return { ok: false, error: "nothing to find" };
  const flags = "gm" + (caseless ? "i" : "") + (dotAll ? "s" : "");
  const translated = translateRegex(source);
  // Unicode mode first (\p{L} works); patterns written for PCRE that it rejects (\- outside a
  // class, a lone {) still run without it.
  const unicode = build(translated, flags + "u");
  return unicode.ok ? unicode : build(translated, flags);
}

function build(source: string, flags: string): CompiledFind {
  try {
    return { ok: true, regex: new RegExp(source, flags) };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message.replace(/^Invalid regular expression: /, "") : String(err) };
  }
}

/** Compile `query` with `options`. An empty query is an error (nothing to find). */
export function compileFind(query: string, options: FindOptions = {}): CompiledFind {
  if (!query) return { ok: false, error: "nothing to find" };
  if (options.regex) return compileRegex(query, options);
  if (options.wildcards) return compileWildcards(query);
  return compilePlain(query, options);
}

// ── replacement ──────────────────────────────────────────────────────────────

/**
 * The text a hit is replaced with, in the stream's characters ("\r" a new paragraph, "\v" a line
 * break, "\t" a tab). `found` is the hit as `RegExp.exec` gave it.
 *
 *  - plain text: ^p ^t ^l ^^ ^& ^+ ^= ^~ ^s ^- ^nnn ^unnnn;
 *  - wildcards: the same, and \1..\9 for the groups;
 *  - regex: $& $1..$99 $<name> $$, \1..\9, and \n \r (paragraph) \t \v \\.
 */
export function expandReplacement(replacement: string, found: RegExpExecArray, options: FindOptions = {}): string {
  let out = "";
  const group = (n: number) => found[n] ?? "";
  for (let i = 0; i < replacement.length; i++) {
    const ch = replacement[i];
    const next = replacement[i + 1];
    if (options.regex) {
      if (ch === "$" && next !== undefined) {
        if (next === "$") { out += "$"; i++; continue; }
        if (next === "&") { out += found[0]; i++; continue; }
        if (next === "<") {
          const close = replacement.indexOf(">", i + 2);
          if (close > 0) {
            out += found.groups?.[replacement.slice(i + 2, close)] ?? "";
            i = close;
            continue;
          }
        }
        if (next === "{") {
          const m = /^\{(\w+)\}/.exec(replacement.slice(i + 1));
          if (m) {
            out += /^\d+$/.test(m[1]) ? group(Number(m[1])) : found.groups?.[m[1]] ?? "";
            i += m[0].length;
            continue;
          }
        }
        const digits = /^\d{1,2}/.exec(replacement.slice(i + 1));
        if (digits) {
          // $12 with only one group is $1 then "2", as in JavaScript.
          let d = digits[0];
          if (d.length === 2 && Number(d) >= found.length) d = d[0];
          out += group(Number(d));
          i += d.length;
          continue;
        }
      }
      if (ch === "\\" && next !== undefined) {
        i++;
        if (next === "n" || next === "r") out += "\r";
        else if (next === "t") out += "\t";
        else if (next === "v") out += "\v";
        else if (/[0-9]/.test(next)) out += group(Number(next));
        else out += next;
        continue;
      }
      out += ch;
      continue;
    }
    if (options.wildcards && ch === "\\" && next !== undefined) {
      i++;
      out += /[1-9]/.test(next) ? group(Number(next)) : next;
      continue;
    }
    if (ch === "^" && next !== undefined) {
      if (next === "&") { out += found[0]; i++; continue; }
      if (/[0-9uU]/.test(next)) {
        const num = readNumericCode(replacement, i + 1);
        const char = num ? numericChar(num.cp) : null;
        if (num && char !== null) { out += char; i = num.next - 1; continue; }
      }
      const mapped = CODE_CHARS[next.toLowerCase()];
      if (mapped !== undefined && next.toLowerCase() !== "m" && next.toLowerCase() !== "n") {
        out += mapped;
        i++;
        continue;
      }
    }
    out += ch;
  }
  return out;
}

/** The text Find shows for the options that are on ("Match case, Whole words"), for a tooltip. */
export function describeFindOptions(options: FindOptions, labels: Record<keyof FindOptions, string>): string {
  return FIND_OPTION_KEYS.filter((k) => options[k] && findOptionApplies(k, options)).map((k) => labels[k]).join(", ");
}

/** Only the options that are set (true), for storing and recording. */
export function normalizeFindOptions(options: FindOptions | null | undefined): FindOptions {
  const out: FindOptions = {};
  if (!options) return out;
  for (const key of FIND_OPTION_KEYS) if (options[key]) out[key] = true;
  if (out.regex && out.wildcards) delete out.wildcards;
  return out;
}
