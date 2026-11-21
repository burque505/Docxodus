/**
 * The ribbon's Tabs dialog — Word's Tabs (Paragraph > Tabs…, or double-click a tab stop on the
 * ruler): the paragraph's tab stops with their alignment and leader, Set / Clear / Clear All, and
 * the document's default tab interval. OK is one `setTabStops` (and a `setDefaultTabStop` when
 * the interval changed).
 */
import type { TabStop, TabStopAlignment, TabStopLeader, TabStopsInfo } from "./types.js";
import { formatLength, makeDialogMovable, parseLength, restoreDialogPlace, type MeasurementUnit } from "./ribbon-paragraph-dialog.js";

export const TABS_DIALOG_CSS = `
.dxr-tabsdialog { width: min(420px, calc(100vw - 32px)); }
.dxr-tabsdialog .dxr-tabsbody { padding: 10px 16px 4px; display: grid; grid-template-columns: 1fr 1fr; gap: 6px 16px; }
.dxr-tabsdialog .dxr-tabspos { display: flex; flex-direction: column; gap: 4px; }
.dxr-tabsdialog .dxr-tabspos input { display: block; width: 100%; box-sizing: border-box; margin-top: 2px; }
.dxr-tabsdialog select[size] { width: 100%; min-height: 7.5em; }
.dxr-tabsdialog fieldset { border: 0; margin: 0; padding: 0; min-width: 0; display: flex; flex-direction: column; gap: 3px; }
.dxr-tabsdialog legend { padding: 0; margin-bottom: 2px; font-weight: 600; }
.dxr-tabsdialog .dxr-tabsside { display: flex; flex-direction: column; gap: 6px; }
.dxr-tabsdialog .dxr-tabsside input { display: block; width: 6em; margin-top: 2px; }
.dxr-tabsdialog .dxr-tabsclear { color: var(--dxr-muted); font-size: 12px; min-height: 1.2em; }
.dxr-tabsdialog .dxr-tabsbuttons { grid-column: 1 / -1; display: flex; gap: 6px; flex-wrap: wrap; }
`;

export const TABS_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-tabsdialog" data-dxr="tabsdialog">
  <form method="dialog" novalidate>
    <h2 data-tabs-title>Tabs</h2>
    <div class="dxr-tabsbody">
      <div class="dxr-tabspos">
        <label>Tab stop position <input type="text" autocomplete="off" data-tabs="position" /></label>
        <select size="6" data-tabs="list" aria-label="Tab stops"></select>
      </div>
      <div class="dxr-tabsside">
        <label>Default tab stops <input type="text" autocomplete="off" data-tabs="default" /></label>
        <span>Tab stops to be cleared</span>
        <span class="dxr-tabsclear" data-tabs="cleared"></span>
      </div>
      <fieldset>
        <legend>Alignment</legend>
        <label><input type="radio" name="dxrtabalign" value="left" /> Left</label>
        <label><input type="radio" name="dxrtabalign" value="center" /> Center</label>
        <label><input type="radio" name="dxrtabalign" value="right" /> Right</label>
        <label><input type="radio" name="dxrtabalign" value="decimal" /> Decimal</label>
        <label><input type="radio" name="dxrtabalign" value="bar" /> Bar</label>
      </fieldset>
      <fieldset>
        <legend>Leader</legend>
        <label><input type="radio" name="dxrtablead" value="none" /> 1 None</label>
        <label><input type="radio" name="dxrtablead" value="dot" /> 2 .......</label>
        <label><input type="radio" name="dxrtablead" value="hyphen" /> 3 -------</label>
        <label><input type="radio" name="dxrtablead" value="underscore" /> 4 _______</label>
        <label><input type="radio" name="dxrtablead" value="middleDot" /> 5 ·······</label>
      </fieldset>
      <div class="dxr-tabsbuttons">
        <button type="button" data-tabs="set">Set</button>
        <button type="button" data-tabs="clear">Clear</button>
        <button type="button" data-tabs="clearall">Clear All</button>
      </div>
    </div>
    <div class="dxr-derror" role="alert"><span data-tabs-msg="position" hidden>Type a tab stop position, such as 1.5".</span><span data-tabs-msg="default" hidden>Type a default tab interval greater than 0, such as 0.5".</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
  </form>
</dialog>
`;

export interface TabsDialogHost {
  /** The caret paragraph's tab stops, or null when there is no paragraph. */
  read(): TabStopsInfo | null;
  apply(stops: TabStop[], defaultTabStopTwips: number | null): void;
}

/** The Tabs dialog over its markup (`TABS_DIALOG_HTML`, already in the page). */
export class TabsDialog {
  private unit: MeasurementUnit;
  private stops: Required<TabStop>[] = [];
  private cleared: number[] = [];
  private initialDefault = 720;

  constructor(private readonly dialog: HTMLDialogElement, private readonly host: TabsDialogHost, unit: MeasurementUnit = "in") {
    this.unit = unit;
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-tabs-title]")!);
    const form = dialog.querySelector("form")!;
    form.addEventListener("input", (event) => {
      (event.target as HTMLElement).removeAttribute("aria-invalid");
      this.showError(null);
    });
    this.list.addEventListener("change", () => this.pick(Number(this.list.value)));
    this.field("set").addEventListener("click", () => this.set());
    this.field("clear").addEventListener("click", () => this.clear());
    this.field("clearall").addEventListener("click", () => {
      this.cleared.push(...this.stops.map((s) => s.positionTwips));
      this.stops = [];
      this.render();
      this.input("position").value = "";
    });
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
      if (submitter?.value === "cancel") { this.dialog.close("cancel"); return; }
      // A position typed but not Set is set by OK, as in Word.
      if (this.input("position").value.trim() && !this.set()) return;
      const defaultText = this.input("default").value.trim();
      const defaultTwips = parseLength(defaultText, this.unit);
      if (defaultTwips === null || defaultTwips <= 0) { this.invalid("default"); return; }
      this.dialog.close("ok");
      this.host.apply(this.stops.map((s) => ({ ...s })), defaultTwips !== this.initialDefault ? defaultTwips : null);
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); this.dialog.close("cancel"); });
  }

  get open(): boolean {
    return this.dialog.open;
  }

  setUnit(unit: MeasurementUnit): void {
    this.unit = unit;
  }

  /** Open on the caret paragraph's tab stops, `selectTwips` chosen when given. False when there
   *  is no paragraph. */
  show(selectTwips?: number): boolean {
    const info = this.host.read();
    if (!info) return false;
    this.stops = info.stops.map((s) => ({ ...s }));
    this.cleared = [];
    this.initialDefault = info.defaultTabStopTwips;
    this.input("default").value = formatLength(info.defaultTabStopTwips, this.unit);
    this.render();
    const chosen = this.stops.findIndex((s) => s.positionTwips === selectTwips);
    if (chosen >= 0 || this.stops.length > 0) this.pick(chosen >= 0 ? chosen : 0);
    else {
      this.input("position").value = "";
      this.radio("dxrtabalign", "left");
      this.radio("dxrtablead", "none");
    }
    this.showError(null);
    this.dialog.returnValue = "";
    this.dialog.showModal();
    restoreDialogPlace(this.dialog);
    this.input("position").focus();
    return true;
  }

  private get list(): HTMLSelectElement {
    return this.field<HTMLSelectElement>("list");
  }

  private render(): void {
    this.list.replaceChildren(...this.stops.map((stop, index) => {
      const option = this.dialog.ownerDocument.createElement("option");
      option.value = String(index);
      option.textContent = formatLength(stop.positionTwips, this.unit);
      return option;
    }));
    this.field("cleared").textContent = this.cleared
      .filter((pos) => !this.stops.some((s) => s.positionTwips === pos))
      .filter((pos, i, all) => all.indexOf(pos) === i)
      .sort((a, b) => a - b).map((pos) => formatLength(pos, this.unit)).join(", ");
  }

  private pick(index: number): void {
    const stop = this.stops[index];
    if (!stop) return;
    this.list.value = String(index);
    this.input("position").value = formatLength(stop.positionTwips, this.unit);
    this.radio("dxrtabalign", stop.alignment);
    this.radio("dxrtablead", stop.leader);
  }

  /** Set (add or change) the stop the position box names. False when the position is not valid. */
  private set(): boolean {
    const twips = parseLength(this.input("position").value.trim(), this.unit);
    if (twips === null || Math.abs(twips) > 31680) { this.invalid("position"); return false; }
    const alignment = (this.checked("dxrtabalign") ?? "left") as TabStopAlignment;
    const leader = (this.checked("dxrtablead") ?? "none") as TabStopLeader;
    this.stops = this.stops.filter((s) => s.positionTwips !== twips);
    this.stops.push({ positionTwips: twips, alignment, leader });
    this.stops.sort((a, b) => a.positionTwips - b.positionTwips);
    this.render();
    this.pick(this.stops.findIndex((s) => s.positionTwips === twips));
    return true;
  }

  private clear(): void {
    const twips = parseLength(this.input("position").value.trim(), this.unit);
    const index = this.stops.findIndex((s) => s.positionTwips === twips);
    if (index < 0) return;
    this.cleared.push(this.stops[index].positionTwips);
    this.stops.splice(index, 1);
    this.render();
    if (this.stops.length > 0) this.pick(Math.min(index, this.stops.length - 1));
    else this.input("position").value = "";
  }

  private radio(name: string, value: string): void {
    const input = this.dialog.querySelector<HTMLInputElement>(`input[name="${name}"][value="${value}"]`);
    if (input) input.checked = true;
  }

  private checked(name: string): string | undefined {
    return this.dialog.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`)?.value;
  }

  private invalid(name: string): null {
    const field = this.input(name);
    field.setAttribute("aria-invalid", "true");
    this.showError(name);
    field.focus();
    return null;
  }

  private showError(message: string | null): void {
    for (const span of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-tabs-msg]")))
      span.hidden = span.dataset.tabsMsg !== message;
  }

  private field<T extends HTMLElement = HTMLElement>(name: string): T {
    return this.dialog.querySelector<T>(`[data-tabs="${name}"]`)!;
  }

  private input(name: string): HTMLInputElement {
    return this.field<HTMLInputElement>(name);
  }
}
