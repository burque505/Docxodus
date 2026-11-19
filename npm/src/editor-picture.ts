/**
 * The editor's picture selection: Word's outline and eight sizing handles over a selected
 * picture, and resizing it with the mouse. A corner keeps the picture's shape; an edge stretches
 * it one way. While dragging, a dotted outline shows the new size; letting go reports it once
 * (one undo step), as a ratio of the size the picture had, so the page's zoom never enters.
 *
 * The overlay is fixed-position in the page's own document and follows the picture through
 * scrolling and resizing, so it works the same in continuous and page view.
 */

export const PICTURE_SELECTION_CSS = `
.docx-picture-overlay { position: fixed; z-index: 2147483000; pointer-events: none; box-sizing: border-box;
  outline: 1px solid #1a73e8; outline-offset: 0; }
.docx-picture-overlay[data-dragging] { outline-style: dotted; }
.docx-picture-overlay .docx-picture-ghost { position: absolute; left: 0; top: 0; box-sizing: border-box;
  border: 1px dotted #1a73e8; display: none; }
.docx-picture-overlay[data-dragging] .docx-picture-ghost { display: block; }
.docx-picture-overlay .docx-picture-handle { position: absolute; width: 9px; height: 9px; margin: -5px 0 0 -5px;
  box-sizing: border-box; background: #fff; border: 1px solid #1a73e8; pointer-events: auto; }
.docx-picture-overlay .docx-picture-handle[data-h="n"], .docx-picture-overlay .docx-picture-handle[data-h="s"] { cursor: ns-resize; }
.docx-picture-overlay .docx-picture-handle[data-h="e"], .docx-picture-overlay .docx-picture-handle[data-h="w"] { cursor: ew-resize; }
.docx-picture-overlay .docx-picture-handle[data-h="nw"], .docx-picture-overlay .docx-picture-handle[data-h="se"] { cursor: nwse-resize; }
.docx-picture-overlay .docx-picture-handle[data-h="ne"], .docx-picture-overlay .docx-picture-handle[data-h="sw"] { cursor: nesw-resize; }
.docx-picture-overlay[data-readonly] .docx-picture-handle { display: none; }
`;

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
/** Where each handle sits, as fractions of the picture's box. */
const HANDLE_AT: Record<Handle, [number, number]> = {
  nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5],
};

export interface PictureOverlayHost {
  /** The user let go of a handle: the picture's new size as ratios of its present one. */
  resize(widthRatio: number, heightRatio: number): void;
  /** The user dragged a floating picture and let go: how far it went, in screen pixels. */
  move(dxPixels: number, dyPixels: number): void;
}

/** How far the pointer must go before a press on a picture becomes a move. */
const MOVE_THRESHOLD = 3;

/** The smallest a drag makes a picture, in screen pixels. */
const MIN_PIXELS = 8;

export class PictureOverlay {
  private readonly overlay: HTMLDivElement;
  private readonly ghost: HTMLDivElement;
  private target: HTMLElement | null = null;
  private frame = 0;

  constructor(private readonly doc: Document, private readonly host: PictureOverlayHost) {
    ensureStyles(doc);
    this.overlay = doc.createElement("div");
    this.overlay.className = "docx-picture-overlay";
    this.overlay.hidden = true;
    this.overlay.setAttribute("aria-hidden", "true");
    this.ghost = doc.createElement("div");
    this.ghost.className = "docx-picture-ghost";
    this.overlay.append(this.ghost);
    for (const h of HANDLES) {
      const handle = doc.createElement("div");
      handle.className = "docx-picture-handle";
      handle.dataset.h = h;
      handle.style.left = `${HANDLE_AT[h][0] * 100}%`;
      handle.style.top = `${HANDLE_AT[h][1] * 100}%`;
      handle.addEventListener("pointerdown", (ev) => this.startDrag(h, ev));
      // Keep the selection: a handle press is not a click in the document.
      handle.addEventListener("mousedown", (ev) => { ev.preventDefault(); ev.stopPropagation(); });
      this.overlay.append(handle);
    }
    (doc.body ?? doc.documentElement).append(this.overlay);
    doc.addEventListener("scroll", this.place, true);
    doc.defaultView?.addEventListener("resize", this.place);
  }

  /** The picture the overlay is on, or null. */
  get picture(): HTMLElement | null {
    return this.target;
  }

  /** Whether `node` is part of the overlay (a handle press is not a click elsewhere). */
  owns(node: Node | null): boolean {
    return !!node && this.overlay.contains(node);
  }

  show(picture: HTMLElement, resizable: boolean): void {
    this.target = picture;
    this.overlay.toggleAttribute("data-readonly", !resizable);
    this.overlay.hidden = false;
    this.place();
  }

  hide(): void {
    this.target = null;
    this.overlay.hidden = true;
    this.overlay.removeAttribute("data-dragging");
  }

  destroy(): void {
    this.hide();
    this.doc.removeEventListener("scroll", this.place, true);
    this.doc.defaultView?.removeEventListener("resize", this.place);
    this.overlay.remove();
  }

  /** Put the overlay over the picture (after layout moved it). */
  readonly place = (): void => {
    const target = this.target;
    if (!target) return;
    if (!target.isConnected) { this.hide(); return; }
    const r = target.getBoundingClientRect();
    Object.assign(this.overlay.style, { left: `${r.left}px`, top: `${r.top}px`, width: `${r.width}px`, height: `${r.height}px` });
  };

  /** Keep the overlay on the picture while the page settles (after a remount, fonts, images). */
  follow(): void {
    cancelAnimationFrame(this.frame);
    let frames = 0;
    const tick = (): void => {
      this.place();
      if (++frames < 10 && this.target) this.frame = requestAnimationFrame(tick);
    };
    this.frame = requestAnimationFrame(tick);
  }

  /**
   * A press on a floating picture moves it: a dotted outline follows the pointer, and letting go
   * reports the distance once (one undo step). A press that hardly moves is just a click;
   * Escape while dragging puts it back.
   */
  startMove(ev: MouseEvent): void {
    const target = this.target;
    if (!target || ev.button !== 0 || this.overlay.hasAttribute("data-readonly")) return;
    const start = target.getBoundingClientRect();
    const view = this.doc.defaultView;
    let dx = 0, dy = 0, moving = false;
    const move = (e: MouseEvent): void => {
      dx = e.clientX - ev.clientX;
      dy = e.clientY - ev.clientY;
      if (!moving && Math.hypot(dx, dy) < MOVE_THRESHOLD) return;
      if (!moving) { moving = true; this.overlay.setAttribute("data-dragging", ""); }
      e.preventDefault();
      Object.assign(this.ghost.style, { left: `${dx}px`, top: `${dy}px`, width: `${start.width}px`, height: `${start.height}px` });
    };
    // The browser's own image drag would swallow the release: never let it start.
    const noNativeDrag = (e: Event): void => e.preventDefault();
    const finish = (drop: boolean): void => {
      view?.removeEventListener("mousemove", move, true);
      view?.removeEventListener("pointermove", move as (e: Event) => void, true);
      view?.removeEventListener("mouseup", release, true);
      view?.removeEventListener("pointerup", release, true);
      view?.removeEventListener("pointercancel", cancel, true);
      view?.removeEventListener("dragstart", noNativeDrag, true);
      view?.removeEventListener("keydown", escape, true);
      view?.removeEventListener("blur", cancel);
      this.overlay.removeAttribute("data-dragging");
      if (drop && moving) this.host.move(dx, dy);
    };
    const release = (e: MouseEvent): void => {
      if (moving) e.preventDefault();
      finish(true);
    };
    const cancel = (): void => finish(false);
    const escape = (e: KeyboardEvent): void => {
      if (e.key !== "Escape") return;
      e.preventDefault();
      e.stopPropagation();
      finish(false);
    };
    view?.addEventListener("mousemove", move, true);
    view?.addEventListener("pointermove", move as (e: Event) => void, true);
    view?.addEventListener("mouseup", release, true);
    view?.addEventListener("pointerup", release, true);
    view?.addEventListener("pointercancel", cancel, true);
    view?.addEventListener("dragstart", noNativeDrag, true);
    view?.addEventListener("keydown", escape, true);
    view?.addEventListener("blur", cancel);
  }

  private startDrag(h: Handle, ev: PointerEvent): void {
    const target = this.target;
    if (!target || ev.button !== 0 || this.overlay.hasAttribute("data-readonly")) return;
    ev.preventDefault();
    ev.stopPropagation();
    const start = target.getBoundingClientRect();
    if (start.width <= 0 || start.height <= 0) return;
    const handle = ev.currentTarget as HTMLElement;
    handle.setPointerCapture?.(ev.pointerId);
    const [fx, fy] = HANDLE_AT[h];
    const horizontal = fx !== 0.5;
    const vertical = fy !== 0.5;
    const corner = horizontal && vertical;
    let width = start.width, height = start.height;
    this.overlay.setAttribute("data-dragging", "");
    const draw = (): void => {
      // The ghost grows away from the opposite edge.
      const left = fx === 0 ? start.width - width : 0;
      const top = fy === 0 ? start.height - height : 0;
      Object.assign(this.ghost.style, { left: `${left}px`, top: `${top}px`, width: `${width}px`, height: `${height}px` });
    };
    draw();
    const move = (e: PointerEvent): void => {
      const dx = (e.clientX - ev.clientX) * (fx === 0 ? -1 : 1);
      const dy = (e.clientY - ev.clientY) * (fy === 0 ? -1 : 1);
      let w = horizontal ? Math.max(MIN_PIXELS, start.width + dx) : start.width;
      let h2 = vertical ? Math.max(MIN_PIXELS, start.height + dy) : start.height;
      if (corner) {
        // A corner keeps the shape: follow whichever way the pointer moved further.
        const ratio = Math.max(w / start.width, h2 / start.height);
        w = Math.max(MIN_PIXELS, start.width * ratio);
        h2 = Math.max(MIN_PIXELS, start.height * ratio);
      }
      width = w;
      height = h2;
      draw();
    };
    const end = (e: PointerEvent): void => {
      handle.removeEventListener("pointermove", move);
      handle.removeEventListener("pointerup", end);
      handle.removeEventListener("pointercancel", end);
      handle.releasePointerCapture?.(e.pointerId);
      this.overlay.removeAttribute("data-dragging");
      if (e.type === "pointercancel") return;
      const wr = width / start.width, hr = height / start.height;
      if (Math.abs(wr - 1) < 0.002 && Math.abs(hr - 1) < 0.002) return;
      this.host.resize(wr, hr);
    };
    handle.addEventListener("pointermove", move);
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  }
}

const styled = new WeakSet<Document>();
function ensureStyles(doc: Document): void {
  if (styled.has(doc)) return;
  styled.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusPicture = "true";
  style.textContent = PICTURE_SELECTION_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}
