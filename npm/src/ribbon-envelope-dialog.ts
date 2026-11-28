/**
 * The ribbon's Envelopes dialog — Mailings > Envelopes…: an envelope of a standard size (#10, DL,
 * C5, C6, Monarch) or a custom one, the return address at its top left and the delivery address
 * lower right. The delivery address is the Address Block's lines for a party («LeFull» …) and can
 * be edited: each «Name» in it becomes a merge field, other text stays text. The return address is
 * the host's to keep. The document becomes the envelope when it is empty; else the envelope is put
 * in front of it as a section of its own, as Word's Add to Document does.
 */
import { makeDialogMovable } from "./ribbon-paragraph-dialog.js";
import { mergeFieldCode } from "./ribbon-mergefields-pane.js";
import { addressLines, addressPrefixes, matchAddress, type Line } from "./ribbon-address-dialog.js";

export interface EnvelopeSize { key: string; label: string; widthTwips: number; heightTwips: number }

const IN = 1440;
const MM = 1440 / 25.4;
/** Word's envelope sizes, the long side as the width (the envelope prints landscape). */
export const ENVELOPE_SIZES: EnvelopeSize[] = [
  { key: "no10", label: "Size 10 (4 1/8 x 9 1/2 in)", widthTwips: Math.round(9.5 * IN), heightTwips: Math.round(4.125 * IN) },
  { key: "dl", label: "DL (110 x 220 mm)", widthTwips: Math.round(220 * MM), heightTwips: Math.round(110 * MM) },
  { key: "c5", label: "C5 (162 x 229 mm)", widthTwips: Math.round(229 * MM), heightTwips: Math.round(162 * MM) },
  { key: "c6", label: "C6 (114 x 162 mm)", widthTwips: Math.round(162 * MM), heightTwips: Math.round(114 * MM) },
  { key: "monarch", label: "Monarch (3 7/8 x 7 1/2 in)", widthTwips: Math.round(7.5 * IN), heightTwips: Math.round(3.875 * IN) },
];

/** The texts the dialog writes from code (the sizes), for a host's translation tables. */
export const ENVELOPE_DIALOG_TEXTS: readonly string[] = [...ENVELOPE_SIZES.map((s) => s.label), "Custom size"];

/** Where an envelope's addresses go, from its top left corner (twips). */
export interface EnvelopePlaces { returnLeft: number; returnTop: number; deliveryLeft: number; deliveryTop: number }

/** Word's places for a size: the return address half an inch in (a printer cannot print at the very edge), the
 *  delivery address about 4 x 2 inches in on a #10 envelope, in proportion on the others. */
export function defaultPlaces(widthTwips: number, heightTwips: number): EnvelopePlaces {
  return {
    returnLeft: Math.round(0.5 * IN), returnTop: Math.round(0.5 * IN),
    deliveryLeft: Math.round(widthTwips * 0.42), deliveryTop: Math.round(heightTwips * 0.48),
  };
}

/** The page margins and the delivery address's indent and space above, for places (twips). */
export function envelopeLayout(widthTwips: number, heightTwips: number, places = defaultPlaces(widthTwips, heightTwips)) {
  return {
    marginLeft: places.returnLeft, marginTop: places.returnTop, marginRight: Math.round(0.5 * IN), marginBottom: Math.round(0.4 * IN),
    deliveryIndent: Math.max(0, places.deliveryLeft - places.returnLeft),
    deliveryTop: Math.max(0, places.deliveryTop - places.returnTop),
  };
}

/** A text box's lines as block lines: each «Name» a merge field, the rest text. */
export function linesOf(text: string): Line[] {
  return text.replace(/\r\n?/g, "\n").split("\n").filter((l, i, all) => l.trim() || (i > 0 && i < all.length - 1)).map((line) => {
    const pieces: Line = [];
    let at = 0;
    for (const m of line.matchAll(/«([^«»{}\\"]{1,60})»/g)) {
      if (m.index! > at) pieces.push({ text: line.slice(at, m.index) });
      pieces.push({ field: m[1].trim() });
      at = m.index! + m[0].length;
    }
    if (at < line.length) pieces.push({ text: line.slice(at) });
    return pieces;
  });
}

/** Lines as a text box shows them («Name» for a field). */
export function textOf(lines: Line[]): string {
  return lines.map((l) => l.map((p) => ("field" in p ? `«${p.field}»` : p.text)).join("")).join("\n");
}

const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const pt = (twips: number) => `${(twips / 20).toFixed(1)}pt`;

/** A line as HTML: a field a span the paste makes a field again, text escaped. */
export function lineHtml(line: Line): string {
  return line.map((piece) => "field" in piece
    ? `<span data-field-code="${escapeHtml(mergeFieldCode({ name: piece.field, format: "text" }))}">«${escapeHtml(piece.field)}»</span>`
    : escapeHtml(piece.text)).join("") || "&#8203;";
}

/** The envelope's content: the return address's lines, then the delivery address's, indented and lower down. */
export function envelopeHtml(ret: Line[], delivery: Line[], widthTwips: number, heightTwips: number, places?: EnvelopePlaces): string {
  const { deliveryIndent, deliveryTop } = envelopeLayout(widthTwips, heightTwips, places);
  const retLines = ret.map((l) => `<p style="margin:0">${lineHtml(l)}</p>`).join("");
  // (the delivery address's top: its space before, less what the return address's lines take)
  const used = ret.length * 276;
  const deliveryLines = delivery.map((l, i) =>
    `<p style="margin:0;margin-left:${pt(deliveryIndent)}${i === 0 ? `;margin-top:${pt(Math.max(0, deliveryTop - used))}` : ""}">${lineHtml(l)}</p>`).join("");
  return `<html><body>${retLines}${deliveryLines}</body></html>`;
}

export const ENVELOPE_DIALOG_CSS = `
.dxr-envdialog { width: min(640px, calc(100vw - 32px)); }
.dxr-envdialog .dxr-envbody { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 10px 16px; padding: 10px 16px 4px; }
.dxr-envdialog label { display: flex; flex-direction: column; gap: 3px; }
.dxr-envdialog label.dxr-dcheck { flex-direction: row; align-items: center; gap: 6px; }
.dxr-envdialog select, .dxr-envdialog input[type="text"], .dxr-envdialog textarea { width: 100%; box-sizing: border-box; font: inherit; }
.dxr-envdialog textarea { resize: vertical; min-height: 6em; font-family: Calibri, "Segoe UI", sans-serif; }
.dxr-envdialog .dxr-envrow { display: flex; gap: 8px; }
.dxr-envdialog .dxr-envrow > label { flex: 1 1 0; min-width: 0; }
.dxr-envdialog .dxr-envwide { grid-column: 1 / -1; }
.dxr-envdialog .dxr-envpreview { position: relative; border: 1px solid var(--dxr-rule-strong); background: #fff; color: #111;
  box-shadow: 0 1px 3px rgba(0,0,0,.15); margin: 4px auto 0; overflow: hidden; font-family: Calibri, "Segoe UI", sans-serif; }
.dxr-envdialog .dxr-envpreview div { position: absolute; white-space: pre; line-height: 1.15; }
.dxr-envdialog .dxr-envnote { color: var(--dxr-muted); font-size: 11.5px; }
.dxr-envdialog .dxr-envplaces { display: grid; grid-template-columns: max-content repeat(2, minmax(0, 6em)); gap: 4px 8px; align-items: center; }
.dxr-envdialog [hidden] { display: none !important; }
`;

export const ENVELOPE_DIALOG_HTML = `
<dialog class="dxr-dialog dxr-envdialog" data-dxr="envelopedialog">
  <form method="dialog" novalidate>
    <h2 data-ev-title>Envelopes</h2>
    <div class="dxr-envbody">
      <div class="dxr-envrow dxr-envwide">
        <label>Envelope size <select data-ev="size"></select></label>
        <label data-ev-row="custom">Width (in) <input type="text" data-ev="width" value="9.5" /></label>
        <label data-ev-row="custom">Height (in) <input type="text" data-ev="height" value="4.125" /></label>
      </div>
      <div>
        <label data-ev-row="prefix">Columns of <select data-ev="prefix"></select></label>
        <label>Delivery address <textarea data-ev="delivery" spellcheck="false" title="«Name» is a merge field; other text stays text"></textarea></label>
      </div>
      <div>
        <label>Return address <textarea data-ev="return" title="Kept for the next envelope"></textarea></label>
        <label class="dxr-dcheck"><input type="checkbox" data-ev="omit" /> Omit</label>
      </div>
      <div class="dxr-envwide dxr-envplaces">
        <span></span><span>From left (in)</span><span>From top (in)</span>
        <span>Return address</span><input type="text" data-ev="retleft" aria-label="Return address from left" /><input type="text" data-ev="rettop" aria-label="Return address from top" />
        <span>Delivery address</span><input type="text" data-ev="delleft" aria-label="Delivery address from left" /><input type="text" data-ev="deltop" aria-label="Delivery address from top" />
      </div>
      <div class="dxr-envwide">
        <span>Preview</span>
        <div class="dxr-envpreview" data-ev="preview"></div>
      </div>
      <span class="dxr-envnote dxr-envwide" data-ev="where"></span>
    </div>
    <div class="dxr-derror" role="alert"><span data-ev-msg="size" hidden>Give the custom size in inches (2 to 22 wide, 2 to 14 high).</span><span data-ev-msg="places" hidden>Give each address's place in inches, inside the envelope.</span><span data-ev-msg="empty" hidden>Type a delivery address, or choose the columns of one.</span><span data-ev-msg="refused" hidden>The envelope could not be made here.</span></div>
    <div class="dxr-dfoot">
      <button type="submit" value="ok" data-ev="ok">OK</button>
      <button type="submit" value="cancel" formnovalidate>Cancel</button>
    </div>
    <span hidden data-ev-text="allcolumns">(all columns)</span>
    <span hidden data-ev-text="new">This document becomes the envelope.</span>
    <span hidden data-ev-text="add">The envelope goes in front of the document, as a section of its own.</span>
  </form>
</dialog>
`;

export interface EnvelopeDialogHost {
  columns(): string[];
  /** The return address kept from last time. */
  returnAddress(): string;
  /** The return address typed, to keep. */
  keepReturnAddress(text: string): void;
  /** Whether the document has anything in it yet (then the envelope goes in front of it). */
  hasContent(): boolean;
  /** Make the envelope: the page `size`, the content `html` (the return address's `place.returnLines`, then the
   *  delivery address's `place.deliveryLines`), the delivery address indented and lower down as `place` says. */
  make(size: { widthTwips: number; heightTwips: number; marginLeft: number; marginTop: number; marginRight: number; marginBottom: number }, html: string,
    place: { returnLines: number; deliveryLines: number; indentTwips: number; beforeTwips: number }): Promise<boolean>;
  translate(text: string): string;
  closed(): void;
}

export class EnvelopeDialog {
  constructor(private readonly dialog: HTMLDialogElement, private readonly host: EnvelopeDialogHost) {
    makeDialogMovable(dialog, dialog.querySelector<HTMLElement>("[data-ev-title]")!);
    for (const el of Array.from(dialog.querySelectorAll<HTMLElement>("input, select, textarea")))
      el.addEventListener(el.tagName === "SELECT" || (el as HTMLInputElement).type === "checkbox" ? "change" : "input", () => this.update());
    this.select("prefix").addEventListener("change", () => { this.fillDelivery(); this.update(); });
    // (another size: Word's places for it)
    this.select("size").addEventListener("change", () => { this.fillPlaces(); this.update(); });
    for (const key of ["width", "height"]) this.input(key).addEventListener("change", () => { this.fillPlaces(); this.update(); });
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

  show(): void {
    const size = this.select("size");
    const keep = size.value;
    size.replaceChildren(...ENVELOPE_SIZES.map((s) => new Option(this.host.translate(s.label), s.key)), new Option(this.host.translate("Custom size"), "custom"));
    size.value = keep || "no10";
    const columns = this.host.columns();
    const prefixes = addressPrefixes(columns);
    const prefix = this.select("prefix");
    const all = this.dialog.querySelector("[data-ev-text=allcolumns]")?.textContent ?? "";
    const keptPrefix = prefix.value;
    prefix.replaceChildren(...(prefixes.length ? prefixes : [""]).map((p) => new Option(p ? `${p}…` : all, p)));
    if (Array.from(prefix.options).some((o) => o.value === keptPrefix)) prefix.value = keptPrefix;
    const row = this.dialog.querySelector<HTMLElement>('[data-ev-row="prefix"]');
    if (row) row.hidden = !(prefixes.length > 1 || (prefixes.length === 1 && prefixes[0] !== ""));
    this.fillDelivery();
    this.fillPlaces();
    this.textarea("return").value = this.host.returnAddress();
    this.part("where").textContent = this.dialog.querySelector(`[data-ev-text=${this.host.hasContent() ? "add" : "new"}]`)?.textContent ?? "";
    this.showError(null);
    this.update();
    if (!this.dialog.open) this.dialog.showModal();
    this.textarea("delivery").focus();
  }

  close(): void {
    if (!this.dialog.open) return;
    this.dialog.close();
    this.host.closed();
  }

  /** The envelope's size in twips, or null for a custom size out of range. */
  size(): EnvelopeSize | null {
    const key = this.select("size").value;
    const known = ENVELOPE_SIZES.find((s) => s.key === key);
    if (known) return known;
    const w = Number(this.input("width").value.replace(",", ".")), h = Number(this.input("height").value.replace(",", "."));
    if (!(w >= 2 && w <= 22 && h >= 2 && h <= 14)) return null;
    // (the long side across: an envelope prints landscape)
    const [long, short] = w >= h ? [w, h] : [h, w];
    return { key: "custom", label: "", widthTwips: Math.round(long * IN), heightTwips: Math.round(short * IN) };
  }

  private fillPlaces(): void {
    const size = this.size();
    if (!size) return;
    const p = defaultPlaces(size.widthTwips, size.heightTwips);
    const inches = (t: number) => String(Math.round((t / IN) * 100) / 100);
    this.input("retleft").value = inches(p.returnLeft);
    this.input("rettop").value = inches(p.returnTop);
    this.input("delleft").value = inches(p.deliveryLeft);
    this.input("deltop").value = inches(p.deliveryTop);
  }

  /** The addresses' places as typed (twips), or null for one outside the envelope. */
  places(size: EnvelopeSize): EnvelopePlaces | null {
    const v = (key: string) => Number(this.input(key).value.replace(",", "."));
    const p = { returnLeft: v("retleft"), returnTop: v("rettop"), deliveryLeft: v("delleft"), deliveryTop: v("deltop") };
    const w = size.widthTwips / IN, h = size.heightTwips / IN;
    if ([p.returnLeft, p.deliveryLeft].some((x) => !(x >= 0.1 && x < w - 1)) || [p.returnTop, p.deliveryTop].some((y) => !(y >= 0.1 && y < h - 0.5))) return null;
    return { returnLeft: Math.round(p.returnLeft * IN), returnTop: Math.round(p.returnTop * IN), deliveryLeft: Math.round(p.deliveryLeft * IN), deliveryTop: Math.round(p.deliveryTop * IN) };
  }

  private fillDelivery(): void {
    const match = matchAddress(this.host.columns(), this.select("prefix").value);
    this.textarea("delivery").value = textOf(addressLines(match, { name: match.full ? "full" : "titlefirstlast", company: true, address: true, country: false }));
  }

  private update(): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>('[data-ev-row="custom"]'))) el.hidden = this.select("size").value !== "custom";
    const size = this.size();
    const preview = this.part("preview");
    const doc = this.dialog.ownerDocument;
    const places = size ? this.places(size) : null;
    if (size && places) {
      const scale = 300 / size.widthTwips;
      preview.style.width = `${Math.round(size.widthTwips * scale)}px`;
      preview.style.height = `${Math.round(size.heightTwips * scale)}px`;
      preview.style.fontSize = `${Math.max(5, 220 * scale)}px`;
      const ret = doc.createElement("div");
      ret.style.left = `${Math.round(places.returnLeft * scale)}px`;
      ret.style.top = `${Math.round(places.returnTop * scale)}px`;
      ret.textContent = this.input("omit").checked ? "" : this.textarea("return").value;
      const del = doc.createElement("div");
      del.style.left = `${Math.round(places.deliveryLeft * scale)}px`;
      del.style.top = `${Math.round(places.deliveryTop * scale)}px`;
      del.textContent = this.textarea("delivery").value;
      preview.replaceChildren(ret, del);
    }
    this.button("ok").disabled = !size || !places || !this.textarea("delivery").value.trim();
  }

  private async ok(): Promise<void> {
    const size = this.size();
    if (!size) { this.showError("size"); return; }
    const places = this.places(size);
    if (!places) { this.showError("places"); return; }
    const delivery = linesOf(this.textarea("delivery").value);
    if (!delivery.length) { this.showError("empty"); return; }
    const returnText = this.textarea("return").value;
    this.host.keepReturnAddress(returnText);
    const ret = this.input("omit").checked ? [] : linesOf(returnText).map((l) => l.map((p) => ("field" in p ? { text: `«${p.field}»` } : p)));
    const html = envelopeHtml(ret, delivery, size.widthTwips, size.heightTwips, places);
    const layout = envelopeLayout(size.widthTwips, size.heightTwips, places);
    // (the delivery address's top: its space before, less what the return address's lines take, at about 13.8pt a line)
    const place = { returnLines: ret.length, deliveryLines: delivery.length, indentTwips: layout.deliveryIndent, beforeTwips: Math.max(0, layout.deliveryTop - ret.length * 276) };
    this.dialog.close();
    if (!(await this.host.make({ widthTwips: size.widthTwips, heightTwips: size.heightTwips, marginLeft: layout.marginLeft, marginTop: layout.marginTop,
      marginRight: layout.marginRight, marginBottom: layout.marginBottom }, html, place))) {
      this.dialog.showModal();
      this.showError("refused");
      return;
    }
    this.host.closed();
  }

  private showError(which: "size" | "places" | "empty" | "refused" | null): void {
    for (const el of Array.from(this.dialog.querySelectorAll<HTMLElement>("[data-ev-msg]"))) el.hidden = el.dataset.evMsg !== which;
  }

  private part(key: string): HTMLElement {
    return this.dialog.querySelector<HTMLElement>(`[data-ev="${key}"]`)!;
  }

  private button(key: string): HTMLButtonElement {
    return this.dialog.querySelector<HTMLButtonElement>(`button[data-ev="${key}"]`)!;
  }

  private input(key: string): HTMLInputElement {
    return this.dialog.querySelector<HTMLInputElement>(`input[data-ev="${key}"]`)!;
  }

  private textarea(key: string): HTMLTextAreaElement {
    return this.dialog.querySelector<HTMLTextAreaElement>(`textarea[data-ev="${key}"]`)!;
  }

  private select(key: string): HTMLSelectElement {
    return this.dialog.querySelector<HTMLSelectElement>(`select[data-ev="${key}"]`)!;
  }
}
