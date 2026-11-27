/**
 * The ribbon's Rules menu and dialog — Mailings > Rules, as Word has it: Ask…, Fill-in…,
 * If…Then…Else…, Merge Record #, Merge Sequence #, Next Record, Next Record If…, Skip Record If….
 * Each writes Word's field (ASK, FILLIN, IF, MERGEREC, MERGESEQ, NEXT, NEXTIF, SKIPIF); a condition
 * compares a merge field with a value (or tests it blank). Ask can show its answer where it is
 * asked too (a REF to its bookmark). The code is shown as it is chosen.
 */
import { makeDialogMovable } from "./ribbon-paragraph-dialog.js";

export type RuleKind = "ask" | "fillin" | "if" | "mergerec" | "mergeseq" | "next" | "nextif" | "skipif";

export const RULES_MENU_HTML = `
<div class="dxr-pop dxr-filemenu" data-dxr="mergerulesmenu" role="menu" aria-label="Rules">
  <button type="button" role="menuitem" data-rule="ask" title="Ask for a value once a record (or once), kept in a bookmark"><span>Ask&#8230;</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="fillin" title="Ask for text where the field is"><span>Fill-in&#8230;</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="if" title="One text or another, as a merge field compares with a value"><span>If&#8230;Then&#8230;Else&#8230;</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="mergerec" title="The data source row's number"><span>Merge Record #</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="mergeseq" title="The merged record's number (records left out not counted)"><span>Merge Sequence #</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="next" title="The fields after this one take the next row (labels: several records a page)"><span>Next Record</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="nextif" title="Moves to the next row within the same copy (labels), when a merge field compares with a value. It does not leave records out: to filter, use Skip Record If"><span>Next Record If&#8230;</span><kbd aria-hidden="true"></kbd></button>
  <button type="button" role="menuitem" data-rule="skipif" title="Leaves the record out of the merge when a merge field compares with a value: the way to filter recipients"><span>Skip Record If&#8230;</span><kbd aria-hidden="true"></kbd></button>
</div>
`;

export const RULE_DIALOG_CSS = `
.dxr-ruledialog { width: min(520px, calc(100vw - 32px)); }
.dxr-ruledialog .dxr-rulebody { display: flex; flex-direction: column; gap: 8px; padding: 10px 16px 4px; }
.dxr-ruledialog label { display: flex; flex-direction: column; gap: 3px; }
.dxr-ruledialog label.dxr-dcheck { flex-direction: row; align-items: center; gap: 6px; }
.dxr-ruledialog .dxr-rulerow { display: flex; gap: 8px; }
.dxr-ruledialog .dxr-rulerow > label { flex: 1 1 0; min-width: 0; }
.dxr-ruledialog input[type="text"], .dxr-ruledialog select { width: 100%; box-sizing: border-box; font: inherit; }
.dxr-ruledialog .dxr-rulecode { font-family: Consolas, "Courier New", monospace; font-size: 11.5px; word-break: break-all;
  padding: 4px 6px; border: 1px solid var(--dxr-rule); border-radius: 4px; min-height: 1.4em; background: var(--dxr-surface, #fff); }
.dxr-ruledialog [hidden] { display: none !important; }
`;

export const RULE_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-ruledialog" data-dxr="ruledialog">
  <form method="dialog" novalidate>
    <h2 data-rd-title><span data-rd-for="ask">Insert Word Field: Ask</span><span data-rd-for="fillin">Insert Word Field: Fill-in</span><span data-rd-for="if">Insert Word Field: IF</span><span data-rd-for="nextif">Insert Word Field: Next Record If</span><span data-rd-for="skipif">Insert Word Field: Skip Record If</span></h2>
    <div class="dxr-rulebody">
      <div class="dxr-rulerow" data-rd-for="if nextif skipif">
        <label>Field name <select data-rd="field"></select></label>
        <label>Comparison
          <select data-rd="op">
            <option value="=">Equal to</option>
            <option value="&lt;&gt;">Not equal to</option>
            <option value="&lt;">Less than</option>
            <option value="&gt;">Greater than</option>
            <option value="&lt;=">Less than or equal</option>
            <option value="&gt;=">Greater than or equal</option>
            <option value="contains">Contains</option>
            <option value="notcontains">Does not contain</option>
            <option value="starts">Starts with</option>
            <option value="ends">Ends with</option>
            <option value="blank">Is blank</option>
            <option value="notblank">Is not blank</option>
          </select>
        </label>
        <label data-rd-row="value">Compare to <input type="text" data-rd="value" /></label>
      </div>
      <label data-rd-for="if">Insert this text <input type="text" data-rd="then" /></label>
      <label data-rd-for="if">Otherwise insert this text <input type="text" data-rd="else" /></label>
      <label data-rd-for="ask">Bookmark <input type="text" data-rd="bookmark" spellcheck="false" autocomplete="off" /></label>
      <label data-rd-for="ask fillin">Prompt <input type="text" data-rd="prompt" /></label>
      <label data-rd-for="ask fillin">Default text <input type="text" data-rd="default" /></label>
      <label class="dxr-dcheck" data-rd-for="ask fillin"><input type="checkbox" data-rd="once" /> Ask once</label>
      <label class="dxr-dcheck" data-rd-for="ask"><input type="checkbox" data-rd="ref" checked /> Show the answer here too</label>
      <span>Field code</span>
      <div class="dxr-rulecode" data-rd="code"></div>
    </div>
    <div class="dxr-derror" role="alert"><span data-rd-msg="bookmark" hidden>A bookmark name starts with a letter and has only letters, digits and _ (up to 40).</span><span data-rd-msg="field" hidden>Choose a field to compare.</span><span data-rd-msg="refused" hidden>The field cannot be put there.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok" data-rd="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

const quote = (s: string) => `"${s.replace(/"/g, '\\"')}"`;
const fieldName = (name: string) => (/[\s"\\{}]/.test(name) ? quote(name) : name);

export interface RuleChoice {
  kind: RuleKind;
  field?: string;
  op?: string;
  value?: string;
  then?: string;
  else?: string;
  bookmark?: string;
  prompt?: string;
  default?: string;
  once?: boolean;
}

/**
 * A condition: `{ MERGEFIELD X } op "value"` (blank: compared with ""). Word has no "contains": it
 * compares with wildcards (* any text, ? any character) after = and <>, so contains is = "*value*".
 */
function condition(c: RuleChoice): string {
  // (in quotes, as Word wants it: a value with spaces is then one value, not several words)
  const left = `"{ MERGEFIELD ${fieldName(c.field ?? "")} }"`;
  const v = c.value ?? "";
  if (c.op === "blank") return `${left} = ""`;
  if (c.op === "notblank") return `${left} <> ""`;
  if (c.op === "contains") return `${left} = ${quote(`*${v}*`)}`;
  if (c.op === "notcontains") return `${left} <> ${quote(`*${v}*`)}`;
  if (c.op === "starts") return `${left} = ${quote(`${v}*`)}`;
  if (c.op === "ends") return `${left} = ${quote(`*${v}`)}`;
  return `${left} ${c.op ?? "="} ${quote(c.value ?? "")}`;
}

/** The field code for a rule (nested fields in `{ }`, as insertField takes it). */
export function ruleCode(c: RuleChoice): string {
  const prompt = quote(c.prompt ?? "");
  const def = c.default ? ` \\d ${quote(c.default)}` : "";
  const once = c.once ? " \\o" : "";
  switch (c.kind) {
    case "ask": return `ASK ${c.bookmark ?? ""} ${prompt}${def}${once}`;
    case "fillin": return `FILLIN ${prompt}${def}${once}`;
    case "if": return `IF ${condition(c)} ${quote(c.then ?? "")} ${quote(c.else ?? "")}`;
    case "nextif": return `NEXTIF ${condition(c)}`;
    case "skipif": return `SKIPIF ${condition(c)}`;
    case "mergerec": return "MERGEREC";
    case "mergeseq": return "MERGESEQ";
    default: return "NEXT";
  }
}

export interface RuleDialogHost {
  /** The field names to compare: the data source's columns, then the document's merge fields. */
  names(): string[];
  /** Insert `codes` at the caret, in order (Ask and its REF: two). */
  insert(codes: string[]): Promise<boolean>;
  closed(): void;
}

export class RuleDialog {
  private kind: RuleKind = "if";

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: RuleDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-rd-title]")!);
    for (const el of Array.from(dialog.querySelectorAll<HTMLElement>("input, select")))
      el.addEventListener(el.tagName === "SELECT" || (el as HTMLInputElement).type === "checkbox" ? "change" : "input", () => this.update());
    const form = dialog.querySelector("form")!;
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter && submitter.value !== "ok") { this.close(); return; }
      void this.ok();
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.close(); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  show(kind: RuleKind): void {
    this.kind = kind;
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-rd-for]")))
      el.hidden = !(el.dataset.rdFor ?? "").split(" ").includes(kind);
    const field = this.select("field");
    const keep = field.value;
    const names = this.host.names();
    field.replaceChildren(...names.map((n) => new Option(n, n)));
    if (names.includes(keep)) field.value = keep;
    this.showError(null);
    this.update();
    if (!this.dialog.open) this.dialog.showModal();
    (kind === "ask" ? this.input("bookmark") : kind === "fillin" ? this.input("prompt") : field).focus();
  }

  close(): void {
    if (!this.dialog.open) return;
    this.dialog.close();
    this.host.closed();
  }

  choice(): RuleChoice {
    return {
      kind: this.kind,
      field: this.select("field").value,
      op: this.select("op").value,
      value: this.input("value").value,
      then: this.input("then").value,
      else: this.input("else").value,
      bookmark: this.input("bookmark").value.trim(),
      prompt: this.input("prompt").value,
      default: this.input("default").value,
      once: this.input("once").checked,
    };
  }

  /** The codes OK inserts: the rule's, and for Ask with "Show the answer here", a REF to its bookmark. */
  codes(): string[] {
    const c = this.choice();
    const codes = [ruleCode(c)];
    if (c.kind === "ask" && this.input("ref").checked && c.bookmark) codes.push(`REF ${c.bookmark}`);
    return codes;
  }

  private problem(): "bookmark" | "field" | null {
    const c = this.choice();
    if (c.kind === "ask" && !/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(c.bookmark ?? "")) return "bookmark";
    if ((c.kind === "if" || c.kind === "nextif" || c.kind === "skipif") && !c.field) return "field";
    return null;
  }

  private update(): void {
    const op = this.select("op").value;
    const valueRow = this.dialog.querySelector<HTMLElement>('[data-rd-row="value"]');
    if (valueRow) valueRow.hidden = op === "blank" || op === "notblank";
    this.part("code").textContent = this.codes().map((c) => `{ ${c} }`).join(" ");
    this.button("ok").disabled = this.problem() !== null;
  }

  private async ok(): Promise<void> {
    const problem = this.problem();
    if (problem) { this.showError(problem); return; }
    const codes = this.codes();
    this.dialog.close();
    if (!(await this.host.insert(codes))) { this.dialog.showModal(); this.showError("refused"); return; }
    this.host.closed();
  }

  private showError(which: "bookmark" | "field" | "refused" | null): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-rd-msg]"))) el.hidden = el.dataset.rdMsg !== which;
  }

  private part(key: string): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`[data-rd="${key}"]`)!;
  }

  private button(key: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`button[data-rd="${key}"]`)!;
  }

  private input(key: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`input[data-rd="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`select[data-rd="${key}"]`)!;
  }
}
