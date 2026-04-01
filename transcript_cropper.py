# Transcript PDF Cropper - GUI Version
# Dependencies: pymupdf, pillow, tkinter (built-in)

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import fitz


# ── Pixel edge detection ──────────────────────────────────────────
def find_content_right_edge(page, dark_threshold=100, min_dark_rows=8):
    mat = fitz.Matrix(1.0, 1.0)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
    w_px, h_px = pix.width, pix.height
    samples = pix.samples
    for x in range(w_px - 1, w_px // 4, -1):
        dark_count = sum(
            1 for y in range(h_px)
            if samples[y * w_px + x] < dark_threshold
        )
        if dark_count >= min_dark_rows:
            return float(x) + 2.0
    return page.rect.width * 0.85


def render_page_image(page, max_width=700):
    scale = max_width / page.rect.width
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img, scale


# ── Processing ───────────────────────────────────────────────────
def process_pdfs(input_folder, output_folder, dark_threshold, min_dark_rows,
                 progress_cb, log_cb, done_cb):
    os.makedirs(output_folder, exist_ok=True)
    log_path = os.path.join(output_folder, "_crop_log.txt")

    pdf_paths = sorted([
        os.path.join(input_folder, f)
        for f in os.listdir(input_folder)
        if f.lower().endswith(".pdf")
    ])

    if not pdf_paths:
        log_cb("[ERROR] No PDF files found.")
        done_cb(0, 0)
        return

    total = len(pdf_paths)
    log_cb(f"[INFO] Found {total} PDF files. Processing...\n")

    log_lines = [
        "=== TRANSCRIPT CROP LOG ===",
        f"Input:  {input_folder}",
        f"Output: {output_folder}",
        f"Dark threshold: {dark_threshold}  Min rows: {min_dark_rows}",
        "",
        f"{'File':<50} {'PageW':<8} {'SplitX':<8} {'SideW':<8} Status",
        "-" * 100
    ]

    done = 0
    errors = 0

    for idx, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        stem = os.path.splitext(filename)[0]
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc[page_num]
                rect = page.rect
                w = rect.width
                suffix = f"_p{page_num+1}" if len(doc) > 1 else ""

                split_x = find_content_right_edge(page, dark_threshold, min_dark_rows)
                side_w = w - split_x

                for label, crop_rect in [
                    ("MAIN", fitz.Rect(rect.x0, rect.y0, rect.x0 + split_x, rect.y1)),
                    ("SIDE", fitz.Rect(rect.x0 + split_x, rect.y0, rect.x1, rect.y1)),
                ]:
                    out_path = os.path.join(output_folder, f"{stem}{suffix}_{label}.pdf")
                    out = fitz.open()
                    out.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    # Clamp cropbox to actual page bounds
                    pr = out[0].rect
                    safe = fitz.Rect(
                        max(crop_rect.x0, pr.x0),
                        max(crop_rect.y0, pr.y0),
                        min(crop_rect.x1, pr.x1),
                        min(crop_rect.y1, pr.y1)
                    )
                    if safe.width > 1 and safe.height > 1:
                        out[0].set_cropbox(safe)
                    out.save(out_path)
                    out.close()

                log_lines.append(
                    f"{filename:<50} {w:<8.1f} {split_x:<8.1f} {side_w:<8.1f} OK"
                )
                done += 1

            doc.close()
        except Exception as e:
            log_lines.append(f"{filename:<50} ERROR: {e}")
            errors += 1

        progress_cb(idx + 1, total)

    log_lines += ["", "=== SUMMARY ===",
                  f"Processed: {done}", f"Errors: {errors}"]

    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    log_cb(f"\n[DONE] Processed: {done}  Errors: {errors}")
    log_cb(f"[INFO] Log saved: {log_path}")
    done_cb(done, errors)


# ── GUI ──────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Transcript PDF Cropper")
        self.geometry("1100x780")
        self.minsize(900, 650)
        self.configure(bg="#1e1e2e")
        self.resizable(True, True)

        # State
        self.input_var  = tk.StringVar()
        self.output_var = tk.StringVar()
        self.threshold_var = tk.IntVar(value=100)
        self.minrows_var   = tk.IntVar(value=8)
        self.preview_doc   = None
        self.preview_page  = 0
        self.preview_scale = 1.0
        self.split_x_pt    = None
        self.drag_x        = None
        self._tk_img       = None

        self._build_ui()

    # ── UI Construction ──────────────────────────────────────────
    def _build_ui(self):
        DARK   = "#1e1e2e"
        PANEL  = "#2a2a3e"
        ACCENT = "#7c6af7"
        TEXT   = "#cdd6f4"
        MUTED  = "#6c7086"
        GREEN  = "#a6e3a1"
        RED    = "#f38ba8"

        self._colors = dict(DARK=DARK, PANEL=PANEL, ACCENT=ACCENT,
                            TEXT=TEXT, MUTED=MUTED, GREEN=GREEN, RED=RED)

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame",       background=PANEL)
        style.configure("Dark.TFrame",  background=DARK)
        style.configure("TLabel",       background=PANEL,  foreground=TEXT,   font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=DARK,   foreground=TEXT,   font=("Segoe UI", 13, "bold"))
        style.configure("Muted.TLabel", background=PANEL,  foreground=MUTED,  font=("Segoe UI", 9))
        style.configure("TButton",      background=ACCENT, foreground="white",
                        font=("Segoe UI", 10, "bold"), borderwidth=0, padding=8)
        style.map("TButton",
                  background=[("active", "#6a5af0"), ("disabled", MUTED)])
        style.configure("Run.TButton",  background=GREEN,  foreground=DARK,
                        font=("Segoe UI", 12, "bold"), padding=12)
        style.map("Run.TButton",
                  background=[("active", "#8ed992"), ("disabled", MUTED)])
        style.configure("TScale",       background=PANEL,  troughcolor=DARK)
        style.configure("TProgressbar", background=ACCENT, troughcolor=DARK, borderwidth=0)

        # ── Header ───────────────────────────────────────────────
        hdr = tk.Frame(self, bg=DARK, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✂  Transcript PDF Cropper",
                 bg=DARK, fg=TEXT, font=("Segoe UI", 16, "bold")).pack(side="left", padx=20)
        tk.Label(hdr, text="Government of Manitoba — Microfiche Processing",
                 bg=DARK, fg=MUTED, font=("Segoe UI", 10)).pack(side="left", padx=4)

        # ── Main layout ───────────────────────────────────────────
        body = tk.Frame(self, bg=DARK)
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        body.columnconfigure(0, weight=0, minsize=300)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # ── Left panel ───────────────────────────────────────────
        left = ttk.Frame(body, padding=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(left, text="FOLDERS", style="Muted.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6)); row += 1

        # Input
        ttk.Label(left, text="Input").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(left, textvariable=self.input_var, width=22).grid(
            row=row, column=1, sticky="ew", padx=6)
        ttk.Button(left, text="…", width=3,
                   command=self._browse_input).grid(row=row, column=2); row += 1

        # Output
        ttk.Label(left, text="Output").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(left, textvariable=self.output_var, width=22).grid(
            row=row, column=1, sticky="ew", padx=6)
        ttk.Button(left, text="…", width=3,
                   command=self._browse_output).grid(row=row, column=2); row += 1

        ttk.Separator(left, orient="horizontal").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=12); row += 1

        # Settings
        ttk.Label(left, text="DETECTION SETTINGS", style="Muted.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6)); row += 1

        # Dark threshold
        ttk.Label(left, text="Dark Threshold").grid(row=row, column=0, sticky="w")
        self._thr_lbl = ttk.Label(left, text="100")
        self._thr_lbl.grid(row=row, column=2, sticky="e"); row += 1
        thr_scale = ttk.Scale(left, from_=30, to=200, variable=self.threshold_var,
                              orient="horizontal", command=self._on_threshold)
        thr_scale.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 8)); row += 1
        ttk.Label(left, text="Lower = stricter detection", style="Muted.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w"); row += 1

        # Min rows
        ttk.Label(left, text="Min Dark Rows").grid(row=row, column=0, sticky="w", pady=(10, 0))
        self._rows_lbl = ttk.Label(left, text="8")
        self._rows_lbl.grid(row=row, column=2, sticky="e", pady=(10, 0)); row += 1
        rows_scale = ttk.Scale(left, from_=2, to=40, variable=self.minrows_var,
                               orient="horizontal", command=self._on_minrows)
        rows_scale.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 8)); row += 1
        ttk.Label(left, text="Higher = fewer false edges", style="Muted.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w"); row += 1

        ttk.Separator(left, orient="horizontal").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=12); row += 1

        # Preview controls
        ttk.Label(left, text="PREVIEW", style="Muted.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6)); row += 1

        ttk.Button(left, text="Load Sample Page",
                   command=self._load_preview).grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=2); row += 1
        ttk.Button(left, text="↺ Re-detect Split Line",
                   command=self._redetect).grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=2); row += 1

        nav = ttk.Frame(left)
        nav.grid(row=row, column=0, columnspan=3, sticky="ew", pady=4); row += 1
        ttk.Button(nav, text="◀ Prev", command=self._prev_page).pack(side="left", expand=True, fill="x")
        self._page_lbl = ttk.Label(nav, text="—")
        self._page_lbl.pack(side="left", padx=8)
        ttk.Button(nav, text="Next ▶", command=self._next_page).pack(side="right", expand=True, fill="x")

        ttk.Separator(left, orient="horizontal").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=12); row += 1

        # Run button
        self._run_btn = ttk.Button(left, text="▶  Run on All Files",
                                   style="Run.TButton", command=self._run)
        self._run_btn.grid(row=row, column=0, columnspan=3, sticky="ew", pady=4); row += 1

        # Progress
        self._progress = ttk.Progressbar(left, mode="determinate")
        self._progress.grid(row=row, column=0, columnspan=3, sticky="ew", pady=4); row += 1

        self._status_lbl = ttk.Label(left, text="Ready", style="Muted.TLabel")
        self._status_lbl.grid(row=row, column=0, columnspan=3, sticky="w"); row += 1

        # ── Right panel ───────────────────────────────────────────
        right = ttk.Frame(body, padding=0)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # Preview canvas
        ttk.Label(right, text="Preview  (drag the purple line to adjust split)",
                  style="Muted.TLabel").grid(row=0, column=0, sticky="w", padx=8, pady=(0, 4))

        canvas_frame = tk.Frame(right, bg=DARK, bd=1, relief="flat")
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(canvas_frame, bg="#111122",
                                 highlightthickness=0, cursor="sb_h_double_arrow")
        self._canvas.grid(row=0, column=0, sticky="nsew")

        sb_v = ttk.Scrollbar(canvas_frame, orient="vertical",   command=self._canvas.yview)
        sb_h = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self._canvas.xview)
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h.grid(row=1, column=0, sticky="ew")
        self._canvas.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)

        self._canvas.bind("<ButtonPress-1>",   self._drag_start)
        self._canvas.bind("<B1-Motion>",       self._drag_move)
        self._canvas.bind("<ButtonRelease-1>", self._drag_end)

        # Info bar below canvas
        self._info_lbl = tk.Label(right, text="Load a sample page to preview the split line",
                                  bg=DARK, fg=MUTED, font=("Segoe UI", 9), anchor="w")
        self._info_lbl.grid(row=2, column=0, sticky="ew", padx=8, pady=4)

        # Log area
        ttk.Label(right, text="LOG", style="Muted.TLabel").grid(
            row=3, column=0, sticky="w", padx=8, pady=(8, 2))
        log_frame = tk.Frame(right, bg=DARK)
        log_frame.grid(row=4, column=0, sticky="ew", padx=0, pady=(0, 0))
        log_frame.columnconfigure(0, weight=1)

        self._log_text = tk.Text(log_frame, height=6, bg="#111122", fg=GREEN,
                                 font=("Consolas", 9), relief="flat",
                                 state="disabled", wrap="word")
        self._log_text.grid(row=0, column=0, sticky="ew")
        log_sb = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        log_sb.grid(row=0, column=1, sticky="ns")
        self._log_text.configure(yscrollcommand=log_sb.set)

    # ── Helpers ──────────────────────────────────────────────────
    def _log(self, msg):
        self._log_text.configure(state="normal")
        self._log_text.insert("end", msg + "\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    def _browse_input(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_var.set(folder)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_var.set(folder)

    def _on_threshold(self, val):
        v = int(float(val))
        self.threshold_var.set(v)
        self._thr_lbl.configure(text=str(v))

    def _on_minrows(self, val):
        v = int(float(val))
        self.minrows_var.set(v)
        self._rows_lbl.configure(text=str(v))

    # ── Preview ──────────────────────────────────────────────────
    def _load_preview(self):
        folder = self.input_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("No Input", "Set an input folder first.")
            return
        pdfs = sorted([f for f in os.listdir(folder) if f.lower().endswith(".pdf")])
        if not pdfs:
            messagebox.showwarning("No PDFs", "No PDF files found in input folder.")
            return
        if self.preview_doc:
            self.preview_doc.close()
        self.preview_doc = fitz.open(os.path.join(folder, pdfs[0]))
        self.preview_page = 0
        self._render_preview()

    def _render_preview(self):
        if not self.preview_doc:
            return
        page = self.preview_doc[self.preview_page]
        total = len(self.preview_doc)
        self._page_lbl.configure(
            text=f"{self.preview_page+1}/{total}")

        # Detect split
        thr = self.threshold_var.get()
        rows = self.minrows_var.get()
        self.split_x_pt = find_content_right_edge(page, thr, rows)

        # Render
        canvas_w = max(self._canvas.winfo_width(), 700)
        img, scale = render_page_image(page, canvas_w - 20)
        self.preview_scale = scale
        self._tk_img = ImageTk.PhotoImage(img)

        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=self._tk_img)
        self._canvas.configure(scrollregion=(0, 0, img.width, img.height))

        self._draw_split_line()

        w_pt  = page.rect.width
        side  = w_pt - self.split_x_pt
        pct   = (self.split_x_pt / w_pt) * 100
        self._info_lbl.configure(
            text=f"Page width: {w_pt:.0f}pt  |  Split at: {self.split_x_pt:.1f}pt ({pct:.1f}%)  |  "
                 f"MAIN: {self.split_x_pt:.1f}pt   SIDE: {side:.1f}pt   "
                 f"  Drag line to adjust  |  Threshold: {self.threshold_var.get()}  Rows: {self.minrows_var.get()}"
        )

    def _draw_split_line(self):
        if self.split_x_pt is None or not self._tk_img:
            return
        self._canvas.delete("splitline", "splitlabel")
        x_px = self.split_x_pt * self.preview_scale
        h = self._tk_img.height()

        # Shaded right region
        self._canvas.create_rectangle(
            x_px, 0, self._tk_img.width(), h,
            fill="#7c6af720", outline="", tags="splitline")

        # Line
        self._canvas.create_line(
            x_px, 0, x_px, h,
            fill="#7c6af7", width=2, dash=(6, 3), tags="splitline")

        # Label
        self._canvas.create_text(
            x_px + 6, 16,
            text=f"  SIDE ▶", fill="#7c6af7",
            font=("Segoe UI", 9, "bold"), anchor="w", tags="splitlabel")
        self._canvas.create_text(
            x_px - 6, 16,
            text="◀ MAIN  ", fill="#a6e3a1",
            font=("Segoe UI", 9, "bold"), anchor="e", tags="splitlabel")

    def _redetect(self):
        self._render_preview()

    def _prev_page(self):
        if self.preview_doc and self.preview_page > 0:
            self.preview_page -= 1
            self._render_preview()

    def _next_page(self):
        if self.preview_doc and self.preview_page < len(self.preview_doc) - 1:
            self.preview_page += 1
            self._render_preview()

    # ── Drag split line ──────────────────────────────────────────
    def _drag_start(self, event):
        if self.split_x_pt is None:
            return
        x_px = self.split_x_pt * self.preview_scale
        if abs(event.x - x_px) < 16:
            self.drag_x = event.x

    def _drag_move(self, event):
        if self.drag_x is None:
            return
        self.drag_x = event.x
        self.split_x_pt = event.x / self.preview_scale
        self._draw_split_line()
        if self.preview_doc:
            page = self.preview_doc[self.preview_page]
            w_pt = page.rect.width
            side = w_pt - self.split_x_pt
            pct  = (self.split_x_pt / w_pt) * 100
            self._info_lbl.configure(
                text=f"Page width: {w_pt:.0f}pt  |  Split at: {self.split_x_pt:.1f}pt ({pct:.1f}%)  |  "
                     f"MAIN: {self.split_x_pt:.1f}pt   SIDE: {side:.1f}pt   "
                     f"  [MANUAL OVERRIDE]"
            )

    def _drag_end(self, event):
        self.drag_x = None

    # ── Run ──────────────────────────────────────────────────────
    def _run(self):
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip()

        if not inp or not os.path.isdir(inp):
            messagebox.showerror("Error", "Set a valid input folder.")
            return
        if not out:
            messagebox.showerror("Error", "Set an output folder.")
            return

        self._run_btn.configure(state="disabled")
        self._progress["value"] = 0
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

        thr  = self.threshold_var.get()
        rows = self.minrows_var.get()

        def progress_cb(done, total):
            pct = (done / total) * 100
            self.after(0, lambda: self._progress.configure(value=pct))
            self.after(0, lambda: self._status_lbl.configure(
                text=f"{done}/{total} files processed"))

        def log_cb(msg):
            self.after(0, lambda: self._log(msg))

        def done_cb(done, errors):
            self.after(0, lambda: self._run_btn.configure(state="normal"))
            self.after(0, lambda: self._status_lbl.configure(
                text=f"✓ Done — {done} files, {errors} errors"))
            if errors == 0:
                self.after(0, lambda: messagebox.showinfo(
                    "Complete", f"Processed {done} files.\nOutput: {out}"))
            else:
                self.after(0, lambda: messagebox.showwarning(
                    "Done with errors",
                    f"Processed {done} files with {errors} errors.\nCheck _crop_log.txt"))

        thread = threading.Thread(
            target=process_pdfs,
            args=(inp, out, thr, rows, progress_cb, log_cb, done_cb),
            daemon=True
        )
        thread.start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
