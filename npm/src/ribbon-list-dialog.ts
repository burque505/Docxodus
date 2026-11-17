/**
 * The ribbon's List Levels dialog — Word's Define New Multilevel List, laid out as a table: for
 * each of the nine levels its numbering (1, a, A, i, I, bullet …), punctuation (1. 1) (1) 1), start,
 * and where its number and its text start. "Legal numbering" makes every level decimal and its
 * number the levels above it and its own (1., 1.1, 1.1.1), with the periods where the user wants
 * them. OK writes all nine levels into the caret's list (making the paragraphs a numbered list
 * when they are not one) as one `setListLevels`; "Also for new documents" hands them to the host
 * as the levels new lists start from.
 */
import type { ListLevelInfo, ListLevelSpec } from "./types.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

type Punctuation = "." | ")" | "()" | "" | "custom";
type Periods = "top" | "all" | "none";

const FORMATS: ReadonlyArray<[string, string]> = [
  ["decimal", "1, 2, 3"], ["lowerLetter", "a, b, c"], ["upperLetter", "A, B, C"],
  ["lowerRoman", "i, ii, iii"], ["upperRoman", "I, II, III"], ["decimalZero", "01, 02, 03"],
  ["bullet", "• Bullet"], ["none", "(no number)"],
];
const PUNCTUATION: ReadonlyArray<[Punctuation, string]> = [[".", "1."], [")", "1)"], ["()", "(1)"], ["", "1"]];

/** The texts the dialog writes into its rows (beside its markup's), for a host checking its translations. */
export const LIST_DIALOG_TEXTS: readonly string[] = [...FORMATS.map(([, t]) => t), "(as it is)"];

/** A level's number text for a punctuation. */
function textFor(level: number, p: Punctuation): string {
  const n = `%${level + 1}`;
  return p === "." ? `${n}.` : p === ")" ? `${n})` : p === "()" ? `(${n})` : n;
}

/** The punctuation a level's number text has, or "custom" for any other text. */
function punctuationOf(level: number, text: string): Punctuation {
  for (const [p] of PUNCTUATION) if (textFor(level, p) === text) return p;
  return "custom";
}

/** The legal text of a level (%1.%2 …) with its period. */
function legalText(level: number, periods: Periods): string {
  const number = Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join(".");
  return number + (periods === "all" || (periods === "top" && level === 0) ? "." : "");
}

/** Whether the levels are legal numbering, and with which periods. */
function legalOf(levels: ListLevelInfo[]): Periods | null {
  for (const periods of ["top", "all", "none"] as const) {
    if (levels.length === 9 && levels.every((l) => l.format === "decimal" && l.text === legalText(l.level, periods))) return periods;
  }
  return null;
}

/** The levels a list has before anyone changes them (Docxodus's own: 1. a. i. …). */
export function plainLevels(): ListLevelInfo[] {
  const cycle = ["decimal", "lowerLetter", "lowerRoman"];
  return Array.from({ length: 9 }, (_, level) => ({
    level, format: cycle[level % 3], text: `%${level + 1}.`, start: 1,
    numberPositionTwips: 360 + 720 * level, textPositionTwips: 720 + 720 * level, legal: false,
  }));
}

/** A number in a level's format, for the preview. */
function numberIn(format: string, n: number): string {
  const roman = (v: number) => {
    const map: Array<[number, string]> = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
    let out = "";
    for (const [value, s] of map) while (v >= value) { out += s; v -= value; }
    return out;
  };
  const letter = (v: number) => String.fromCharCode(96 + (((v - 1) % 26) + 1)).repeat(Math.floor((v - 1) / 26) + 1);
  switch (format) {
    case "lowerLetter": return letter(n);
    case "upperLetter": return letter(n).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "decimalZero": return n < 10 ? `0${n}` : String(n);
    case "none": return "";
    default: return String(n);
  }
}

export const LIST_DIALOG_CSS = `
.dxr-listdialog { width: min(660px, calc(100vw - 32px)); }
.dxr-listdialog .dxr-ldbody { padding: 10px 16px 4px; }
.dxr-listdialog .dxr-ldlegal { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 16px; margin-bottom: 8px; }
.dxr-listdialog .dxr-ldlegal label { display: flex; align-items: center; gap: 6px; color: var(--dxr-muted); }
.dxr-listdialog table { width: 100%; border-collapse: collapse; }
.dxr-listdialog th { padding: 0 4px 4px; color: var(--dxr-faint); font-size: 9.5px; font-weight: 600; letter-spacing: .08em; text-align: left; text-transform: uppercase; }
.dxr-listdialog td { padding: 2px 4px; }
.dxr-listdialog td:first-child { width: 2.5em; color: var(--dxr-muted); text-align: center; }
.dxr-listdialog td select, .dxr-listdialog td input { width: 100%; height: 26px; }
.dxr-listdialog td input[data-ld-field="start"] { max-width: 4.5em; }
.dxr-listdialog .dxr-ldpreview {
  height: 84px; overflow: hidden; margin-top: 8px; padding: 6px 10px; border: 1px solid var(--dxr-rule);
  border-radius: 6px; background: var(--dxr-sheet); font: 11px/1.5 var(--dxr-serif); color: var(--dxr-ink);
}
.dxr-listdialog .dxr-ldpreview div { position: relative; white-space: nowrap; }
.dxr-listdialog .dxr-ldpreview span { position: absolute; }
.dxr-listdialog .dxr-ldbar { display: inline-block; width: 180px; height: 5px; border-radius: 3px; background: var(--dxr-rule-strong); vertical-align: middle; }
.dxr-listdialog .dxr-ldfoot { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 6px 16px 14px; }
.dxr-listdialog .dxr-ldfoot .dxr-dcheck { margin: 0; }
`;

export const LIST_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-listdialog" data-dxr="listdialog">
  <form method="dialog">
    <h2 data-ld-title>List levels</h2>
    <div class="dxr-ldbody">
      <div class="dxr-ldlegal">
        <label><input type="checkbox" data-ld="legal" /> Legal numbering (1.1, 1.1.1)</label>
        <label>Periods
          <select data-ld="periods">
            <option value="top">1.  1.1  1.1.1</option>
            <option value="all">1.  1.1.  1.1.1.</option>
            <option value="none">1  1.1  1.1.1</option>
          </select>
        </label>
      </div>
      <table>
        <thead><tr><th>Level</th><th>Numbering</th><th>Punctuation</th><th>Start</th><th>Number at</th><th>Text at</th></tr></thead>
        <tbody data-ld-rows></tbody>
      </table>
      <div class="dxr-ldpreview" aria-hidden="true" data-ld-preview></div>
    </div>
    <div class="dxr-derror" role="alert"><span data-ld-msg="length" hidden>Type a measurement, such as 1.5" or 2 cm.</span><span data-ld-msg="start" hidden>Type a start number, 0 or more.</span></div>
    <div class="dxr-ldfoot">
      <label class="dxr-dcheck" data-ld-defaultsrow hidden><input type="checkbox" data-ld="defaults" /> Also for new documents</label>
      <span></span>
      <div class="dxr-dfoot" style="padding:0">
        <button type="submit" value="ok">OK</button>
        <button type="submit" value="cancel" formnovalidate>Cancel</button>
      </div>
    </div>
  </form>
</dialog>
`;

export interface ListDialogHost {
  /** The caret's list's levels, or null when the caret is in no list. */
  read(): ListLevelInfo[] | null;
  /** The levels new lists start from, when the host keeps any (shown for a paragraph in no list). */
  defaults(): ListLevelSpec[] | null;
  /** Write the levels (OK). */
  apply(levels: ListLevelSpec[]): void;
  /** "Also for new documents" was ticked: keep these as the levels new lists start from. Without
   *  this, the dialog has no such box. */
  keepDefaults?: (levels: ListLevelSpec[]) => void;
}

/** The List Levels dialog over its markup (`LIST_DIALOG_HTML`, already in the page). */
export class ListDialog {
  private unit: MeasurementUnit;
  /** The texts levels had when the dialog opened, kept for a "custom" punctuation. */
  private texts: string[] = [];

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: ListDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    const rows = dialog.querySelector("[data-ld-rows]")!;
    for (let level = 0; level < 9; level++) {
      const tr = document.createElement("tr");
      tr.dataset.level = String(level);
      tr.innerHTML = `<td>${level + 1}</td>`
        + `<td><select data-ld-field="format" aria-label="Numbering">${FORMATS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("")}<option value="other" hidden>(as it is)</option></select></td>`
        + `<td><select data-ld-field="punctuation" aria-label="Punctuation">${PUNCTUATION.map(([v, t]) => `<option value="${v}">${t}</option>`).join("")}<option value="custom" hidden>(as it is)</option></select></td>`
        + `<td><input type="text" autocomplete="off" data-ld-field="start" aria-label="Start" /></td>`
        + `<td><input type="text" autocomplete="off" data-ld-field="number" aria-label="Number at" /></td>`
        + `<td><input type="text" autocomplete="off" data-ld-field="text" aria-label="Text at" /></td>`;
      rows.append(tr);
    }
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-ld-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
      this.preview();
    });
    form.addEventListener("change", () => { this.syncLegal(); this.preview(); });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      const levels = this.collect();
      if (!levels) return;
      const keep = this.field<HTMLInputElement>("defaults").checked && !!this.host.keepDefaults;
      this.dialog.close("ok");
      this.host.apply(levels);
      if (keep) this.host.keepDefaults!(levels);
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the caret's list (or, off a list, on the levels a new list would get). */
  show(): void {
    const own = this.host.read();
    const defaults = this.host.defaults();
    const base = own ?? plainLevels();
    // Off a list, the host's defaults show over the plain levels.
    const levels = base.map((l) => {
      const d = own ? undefined : defaults?.find((x) => x.level === l.level);
      return d ? { ...l, ...Object.fromEntries(Object.entries(d).filter(([, v]) => v !== undefined)) } as ListLevelInfo : l;
    });
    this.texts = Array.from({ length: 9 }, (_, i) => levels.find((l) => l.level === i)?.text ?? `%${i + 1}.`);
    const legal = legalOf(levels);
    this.field<HTMLInputElement>("legal").checked = legal !== null;
    this.field<HTMLSelectElement>("periods").value = legal ?? "top";
    for (let level = 0; level < 9; level++) {
      const l = levels.find((x) => x.level === level) ?? plainLevels()[level];
      const row = this.row(level);
      const format = row.querySelector<HTMLSelectElement>('[data-ld-field="format"]')!;
      format.value = FORMATS.some(([v]) => v === l.format) ? l.format : "other";
      row.querySelector<HTMLSelectElement>('[data-ld-field="punctuation"]')!.value = l.format === "bullet" ? "." : punctuationOf(level, l.text);
      this.input(level, "start").value = String(l.start);
      this.input(level, "number").value = formatLength(l.numberPositionTwips, this.unit);
      this.input(level, "text").value = formatLength(l.textPositionTwips, this.unit);
    }
    this.dialog.querySelector<HTMLElement>("[data-ld-defaultsrow]")!.hidden = !this.host.keepDefaults;
    this.field<HTMLInputElement>("defaults").checked = false;
    this.showError(null);
    this.syncLegal();
    this.preview();
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.row(0).querySelector<HTMLElement>("select")!.focus();
  }

  /** The nine levels the fields say, or null (the bad field marked) when one is not valid. */
  collect(): ListLevelSpec[] | null {
    const legal = this.field<HTMLInputElement>("legal").checked;
    const periods = this.field<HTMLSelectElement>("periods").value as Periods;
    const levels: ListLevelSpec[] = [];
    for (let level = 0; level < 9; level++) {
      const start = Number(this.input(level, "start").value.trim());
      if (!Number.isInteger(start) || start < 0) return this.bad(level, "start", "start");
      const numberAt = parseLength(this.input(level, "number").value, this.unit);
      if (numberAt === null || Math.abs(numberAt) > 31680) return this.bad(level, "number", "length");
      const textAt = parseLength(this.input(level, "text").value, this.unit);
      if (textAt === null || Math.abs(textAt) > 31680) return this.bad(level, "text", "length");
      const spec: ListLevelSpec = { level, start, numberPositionTwips: numberAt, textPositionTwips: textAt, legal: legal && level > 0 };
      if (legal) {
        spec.format = "decimal";
        spec.text = legalText(level, periods);
      } else {
        const format = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="format"]')!.value;
        const p = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="punctuation"]')!.value as Punctuation;
        if (format !== "other") spec.format = format;
        if (format === "bullet") spec.text = this.texts[level].includes("%") ? "•" : this.texts[level];
        else if (format === "none") spec.text = "";
        else if (p !== "custom") spec.text = textFor(level, p);
        else if (this.texts[level].includes("%")) spec.text = this.texts[level];
        else spec.text = textFor(level, ".");    // (was a bullet, now a number with no punctuation chosen)
      }
      levels.push(spec);
    }
    return levels;
  }

  // ── internals ──

  private field<T extends HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-ld="${name}"]`)!;
  }

  private row(level: number): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`tr[data-level="${level}"]`)!;
  }

  private input(level: number, field: string): HTMLInputElement {
    return this.row(level).querySelector<HTMLInputElement>(`[data-ld-field="${field}"]`)!;
  }

  private bad(level: number, field: string, message: "length" | "start"): null {
    const el = this.input(level, field);
    el.setAttribute("aria-invalid", "true");
    el.focus();
    this.showError(message);
    return null;
  }

  private showError(which: string | null): void {
    for (const msg of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-ld-msg]"))) msg.hidden = msg.dataset.ldMsg !== which;
  }

  /** Legal numbering decides every level's numbering and punctuation; periods mean something only then. */
  private syncLegal(): void {
    const legal = this.field<HTMLInputElement>("legal").checked;
    this.field<HTMLSelectElement>("periods").disabled = !legal;
    for (let level = 0; level < 9; level++) {
      const format = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="format"]')!;
      const p = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="punctuation"]')!;
      format.disabled = legal;
      p.disabled = legal || format.value === "bullet" || format.value === "none";
    }
  }

  /** The first three levels, then the top level again, as the fields would number them. */
  private preview(): void {
    const box = this.dialog.querySelector<HTMLElement>("[data-ld-preview]")!;
    const levels = this.collectQuietly();
    const scale = 28 / 1440;    // 28 px an inch
    const counters = levels.map((l) => l.start ?? 1);
    const label = (level: number) => {
      const l = levels[level];
      if (l.format === "bullet") return l.text || "•";
      return (l.text ?? "").replace(/%(\d)/g, (_, d: string) => {
        const i = Number(d) - 1;
        const fmt = l.legal && i < level ? "decimal" : levels[i]?.format ?? "decimal";
        return numberIn(fmt, counters[i]);
      });
    };
    const lines: string[] = [];
    for (const level of [0, 1, 2]) {
      const l = levels[level];
      lines.push(`<div style="height:1.5em"><span style="left:${Math.max(0, (l.numberPositionTwips ?? 0) * scale)}px">${escapeHtml(label(level))}</span>`
        + `<span style="left:${Math.max(0, (l.textPositionTwips ?? 0) * scale)}px"><i class="dxr-ldbar"></i></span></div>`);
    }
    counters[0]++;
    counters[1] = levels[1].start ?? 1;
    const l0 = levels[0];
    lines.push(`<div style="height:1.5em"><span style="left:${Math.max(0, (l0.numberPositionTwips ?? 0) * scale)}px">${escapeHtml(label(0))}</span>`
      + `<span style="left:${Math.max(0, (l0.textPositionTwips ?? 0) * scale)}px"><i class="dxr-ldbar"></i></span></div>`);
    box.innerHTML = lines.join("");
  }

  /** The fields as levels, bad entries taken as their old values, for the preview. */
  private collectQuietly(): ListLevelSpec[] {
    const legal = this.field<HTMLInputElement>("legal").checked;
    const periods = this.field<HTMLSelectElement>("periods").value as Periods;
    return Array.from({ length: 9 }, (_, level) => {
      const format = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="format"]')!.value;
      const p = this.row(level).querySelector<HTMLSelectElement>('[data-ld-field="punctuation"]')!.value as Punctuation;
      const start = Number(this.input(level, "start").value.trim());
      return {
        level,
        format: legal ? "decimal" : format === "other" ? "decimal" : format,
        text: legal ? legalText(level, periods)
          : format === "bullet" ? (this.texts[level].includes("%") ? "•" : this.texts[level])
          : format === "none" ? "" : p === "custom" ? this.texts[level] : textFor(level, p),
        start: Number.isInteger(start) && start >= 0 ? start : 1,
        numberPositionTwips: parseLength(this.input(level, "number").value, this.unit) ?? 0,
        textPositionTwips: parseLength(this.input(level, "text").value, this.unit) ?? 0,
        legal: legal && level > 0,
      };
    });
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}
