#!/usr/bin/env python3
"""
3MF Converter GUI - Convert old BambuStudio 3MF files to Orca Slicer format.
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path

from convert_3mf import convert_3mf, classify_3mf
import zipfile


# --- Theme colors ---
BG = "#1e1e2e"
BG_LIGHT = "#2a2a3d"
BG_CARD = "#313145"
FG = "#cdd6f4"
FG_DIM = "#6c7086"
ACCENT = "#89b4fa"
ACCENT_HOVER = "#b4d0fb"
SUCCESS = "#a6e3a1"
ERROR = "#f38ba8"
WARNING = "#fab387"
BORDER = "#45475a"


class FileItem:
    """Represents a file in the conversion queue."""
    def __init__(self, path):
        self.path = Path(path)
        self.status = "pending"  # pending, converting, done, skipped, error
        self.message = ""
        self.output_path = None


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3MF Converter for Orca Slicer")
        self.root.geometry("720x650")
        self.root.configure(bg=BG)
        self.root.minsize(600, 550)

        self.files = []
        self.converting = False
        self.output_dir = None

        self._build_ui()
        self._setup_drop_target()

    def _log(self, msg, level="info"):
        """Append a message to the debug log."""
        color_tag = level  # "info", "error", "success", "warn"
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n", color_tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _log_threadsafe(self, msg, level="info"):
        """Log from a worker thread via root.after."""
        self.root.after(0, lambda: self._log(msg, level))

    def _build_ui(self):
        # --- Header ---
        header = tk.Frame(self.root, bg=BG, padx=20, pady=16)
        header.pack(fill=tk.X)

        title = tk.Label(header, text="3MF Converter", font=("Segoe UI", 20, "bold"),
                         bg=BG, fg=FG)
        title.pack(anchor=tk.W)

        subtitle = tk.Label(header, text="Convert old BambuStudio 3MF files to Orca Slicer format",
                            font=("Segoe UI", 10), bg=BG, fg=FG_DIM)
        subtitle.pack(anchor=tk.W)

        # --- Output folder picker ---
        output_frame = tk.Frame(self.root, bg=BG, padx=20)
        output_frame.pack(fill=tk.X, pady=(0, 8))

        tk.Label(output_frame, text="Output folder:", font=("Segoe UI", 9),
                 bg=BG, fg=FG_DIM).pack(side=tk.LEFT)

        self.output_label = tk.Label(output_frame, text="Same as input file",
                                     font=("Segoe UI", 9), bg=BG, fg=FG, anchor=tk.W)
        self.output_label.pack(side=tk.LEFT, padx=(8, 8), fill=tk.X, expand=True)

        self.output_browse_btn = tk.Button(output_frame, text="Browse", font=("Segoe UI", 9),
                                           bg=BG_CARD, fg=FG, activebackground=BG_LIGHT,
                                           activeforeground=FG, relief=tk.FLAT, padx=10, pady=2,
                                           cursor="hand2", command=self._browse_output)
        self.output_browse_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.output_browse_btn.bind("<Enter>", lambda e: self.output_browse_btn.configure(bg=BG_LIGHT))
        self.output_browse_btn.bind("<Leave>", lambda e: self.output_browse_btn.configure(bg=BG_CARD))

        self.output_reset_btn = tk.Label(output_frame, text="Reset", font=("Segoe UI", 8),
                                         bg=BG, fg=ACCENT, cursor="hand2")
        self.output_reset_btn.pack(side=tk.LEFT)
        self.output_reset_btn.bind("<Button-1>", lambda e: self._reset_output())
        self.output_reset_btn.bind("<Enter>", lambda e: self.output_reset_btn.configure(fg=ACCENT_HOVER))
        self.output_reset_btn.bind("<Leave>", lambda e: self.output_reset_btn.configure(fg=ACCENT))

        # --- Options row ---
        options_frame = tk.Frame(self.root, bg=BG, padx=20)
        options_frame.pack(fill=tk.X, pady=(0, 6))

        self.strip_settings_var = tk.BooleanVar(value=False)
        self.strip_cb = tk.Checkbutton(
            options_frame, text="Strip slicer settings (geometry + colors only)",
            variable=self.strip_settings_var, font=("Segoe UI", 9),
            bg=BG, fg=FG_DIM, selectcolor=BG_CARD, activebackground=BG,
            activeforeground=FG, highlightthickness=0)
        self.strip_cb.pack(side=tk.LEFT)

        # --- Drop zone / file list area ---
        self.list_frame = tk.Frame(self.root, bg=BG, padx=20)
        self.list_frame.pack(fill=tk.BOTH, expand=True)

        # Drop zone (shown when no files)
        self.drop_zone = tk.Frame(self.list_frame, bg=BG_CARD, highlightbackground=BORDER,
                                  highlightthickness=2, padx=40, pady=40)

        drop_icon = tk.Label(self.drop_zone, text="+", font=("Segoe UI", 36, "bold"),
                             bg=BG_CARD, fg=ACCENT)
        drop_icon.pack()

        drop_text = tk.Label(self.drop_zone, text="Drop 3MF files here\nor click to browse",
                             font=("Segoe UI", 12), bg=BG_CARD, fg=FG_DIM, justify=tk.CENTER)
        drop_text.pack(pady=(4, 0))

        self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.drop_zone.bind("<Button-1>", lambda e: self._browse_files())
        drop_icon.bind("<Button-1>", lambda e: self._browse_files())
        drop_text.bind("<Button-1>", lambda e: self._browse_files())

        # File list (shown when files added)
        self.file_list_container = tk.Frame(self.list_frame, bg=BG)

        list_header = tk.Frame(self.file_list_container, bg=BG)
        list_header.pack(fill=tk.X, pady=(0, 6))

        tk.Label(list_header, text="Files", font=("Segoe UI", 11, "bold"),
                 bg=BG, fg=FG).pack(side=tk.LEFT)

        self.clear_btn = tk.Label(list_header, text="Clear all", font=("Segoe UI", 9),
                                  bg=BG, fg=ACCENT, cursor="hand2")
        self.clear_btn.pack(side=tk.RIGHT)
        self.clear_btn.bind("<Button-1>", lambda e: self._clear_files())
        self.clear_btn.bind("<Enter>", lambda e: self.clear_btn.configure(fg=ACCENT_HOVER))
        self.clear_btn.bind("<Leave>", lambda e: self.clear_btn.configure(fg=ACCENT))

        # Scrollable file list
        self.canvas = tk.Canvas(self.file_list_container, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.file_list_container, orient=tk.VERTICAL,
                                  command=self.canvas.yview)
        self.scrollable = tk.Frame(self.canvas, bg=BG)

        self.scrollable.bind("<Configure>",
                             lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # --- Progress bar ---
        self.progress_frame = tk.Frame(self.root, bg=BG, padx=20)
        self.progress_frame.pack(fill=tk.X)

        self.progress_label = tk.Label(self.progress_frame, text="", font=("Segoe UI", 9),
                                       bg=BG, fg=FG_DIM)
        self.progress_canvas = tk.Canvas(self.progress_frame, height=8, bg=BG_CARD,
                                         highlightthickness=0, bd=0)
        self.progress_canvas.bind("<Configure>", lambda e: self._draw_progress())
        self.progress_value = 0.0
        self.progress_visible = False

        # --- Debug log panel ---
        log_frame = tk.Frame(self.root, bg=BG, padx=20)
        log_frame.pack(fill=tk.X, pady=(4, 0))

        log_header = tk.Frame(log_frame, bg=BG)
        log_header.pack(fill=tk.X)
        tk.Label(log_header, text="Log", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=FG_DIM).pack(side=tk.LEFT)

        self.log_text = tk.Text(log_frame, height=6, bg=BG_CARD, fg=FG,
                                font=("Consolas", 9), relief=tk.FLAT, wrap=tk.WORD,
                                state=tk.DISABLED, highlightthickness=0,
                                insertbackground=FG, selectbackground=ACCENT)
        self.log_text.pack(fill=tk.X, pady=(4, 4))

        # Configure color tags
        self.log_text.tag_configure("info", foreground=FG_DIM)
        self.log_text.tag_configure("error", foreground=ERROR)
        self.log_text.tag_configure("success", foreground=SUCCESS)
        self.log_text.tag_configure("warn", foreground=WARNING)
        self.log_text.tag_configure("accent", foreground=ACCENT)

        # --- Bottom bar ---
        bottom = tk.Frame(self.root, bg=BG, padx=20, pady=10)
        bottom.pack(fill=tk.X)

        self.status_label = tk.Label(bottom, text="", font=("Segoe UI", 9),
                                     bg=BG, fg=FG_DIM)
        self.status_label.pack(side=tk.LEFT)

        btn_frame = tk.Frame(bottom, bg=BG)
        btn_frame.pack(side=tk.RIGHT)

        self.add_btn = tk.Button(btn_frame, text="Add Files", font=("Segoe UI", 10),
                                 bg=BG_CARD, fg=FG, activebackground=BG_LIGHT,
                                 activeforeground=FG, relief=tk.FLAT, padx=16, pady=6,
                                 cursor="hand2", command=self._browse_files)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.convert_btn = tk.Button(btn_frame, text="Convert All", font=("Segoe UI", 10, "bold"),
                                     bg=ACCENT, fg="#1e1e2e", activebackground=ACCENT_HOVER,
                                     activeforeground="#1e1e2e", relief=tk.FLAT, padx=20, pady=6,
                                     cursor="hand2", command=self._start_conversion)
        self.convert_btn.pack(side=tk.LEFT)

        self.add_btn.bind("<Enter>", lambda e: self.add_btn.configure(bg=BG_LIGHT))
        self.add_btn.bind("<Leave>", lambda e: self.add_btn.configure(bg=BG_CARD))
        self.convert_btn.bind("<Enter>", lambda e: self.convert_btn.configure(bg=ACCENT_HOVER))
        self.convert_btn.bind("<Leave>", lambda e: self.convert_btn.configure(bg=ACCENT))

    def _setup_drop_target(self):
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except ImportError:
            pass

    def _on_drop(self, event):
        raw = event.data
        paths = []
        i = 0
        while i < len(raw):
            if raw[i] == '{':
                end = raw.index('}', i)
                paths.append(raw[i + 1:end])
                i = end + 2
            elif raw[i] == ' ':
                i += 1
            else:
                end = raw.find(' ', i)
                if end == -1:
                    end = len(raw)
                paths.append(raw[i:end])
                i = end + 1

        for p in paths:
            if p.lower().endswith(".3mf"):
                self._add_file(p)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir = Path(folder)
            display = str(self.output_dir)
            if len(display) > 50:
                display = "..." + display[-47:]
            self.output_label.configure(text=display, fg=ACCENT)
            self._log(f"Output folder set: {self.output_dir}", "info")

    def _reset_output(self):
        self.output_dir = None
        self.output_label.configure(text="Same as input file", fg=FG)
        self._log("Output folder reset to: same as input", "info")

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select 3MF files",
            filetypes=[("3MF files", "*.3mf"), ("All files", "*.*")]
        )
        for p in paths:
            self._add_file(p)

    def _add_file(self, path):
        path = Path(path)
        if any(f.path == path for f in self.files):
            self._log(f"Skipping duplicate: {path.name}", "warn")
            return

        item = FileItem(path)
        self._log(f"Added: {path.name}  ({path})", "info")

        try:
            with zipfile.ZipFile(path, "r") as z:
                names = z.namelist()
                self._log(f"  ZIP contents: {len(names)} entries", "info")
                if "3D/3dmodel.model" in names:
                    model_size = z.getinfo("3D/3dmodel.model").file_size
                    self._log(f"  Model size: {model_size/1024/1024:.1f}MB uncompressed", "info")
                    # Read head+tail for classification (handles large files)
                    with z.open("3D/3dmodel.model") as f:
                        if model_size <= 10 * 1024 * 1024:
                            classify_bytes = f.read()
                        else:
                            head = f.read(50000)
                            tail = b""
                            while True:
                                chunk = f.read(1024 * 1024)
                                if not chunk:
                                    break
                                tail = chunk[-50000:] if len(chunk) >= 50000 else (tail + chunk)[-50000:]
                            classify_bytes = head + tail
                    fmt = classify_3mf(classify_bytes, zip_names=names)
                    self._log(f"  Format: {fmt}", "info")
                    if fmt == "orca_ready":
                        item.status = "skipped"
                        item.message = "Already in Orca format"
                        self._log(f"  -> Marked as SKIPPED", "warn")
                    elif fmt == "sliced":
                        item.status = "error"
                        item.message = "Sliced .gcode.3mf — no geometry to convert"
                        self._log(f"  -> Sliced file, no model data", "error")
                    elif fmt == "unknown":
                        item.status = "error"
                        item.message = "Unrecognized 3MF format"
                        self._log(f"  -> Unrecognized format", "error")
                    else:
                        self._log(f"  -> Ready to convert ({fmt})", "success")
                else:
                    # Check for sliced gcode files
                    has_gcode = any(n.endswith(".gcode") for n in names)
                    if has_gcode:
                        item.status = "error"
                        item.message = "Sliced .gcode.3mf — no geometry to convert"
                        self._log(f"  -> Sliced file with no model data", "error")
                    else:
                        item.status = "error"
                    item.message = "Not a valid 3MF (no model file)"
                    self._log(f"  -> No 3D/3dmodel.model found", "error")
        except Exception as e:
            self._log(f"  Error inspecting file: {e}", "error")

        self.files.append(item)
        self._refresh_list()

    def _clear_files(self):
        if self.converting:
            return
        self.files.clear()
        self._refresh_list()
        self._log("Cleared all files", "info")

    def _refresh_list(self):
        for w in self.scrollable.winfo_children():
            w.destroy()

        if not self.files:
            self.file_list_container.pack_forget()
            self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            self.status_label.configure(text="")
            return

        self.drop_zone.pack_forget()
        self.file_list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        for i, item in enumerate(self.files):
            row = tk.Frame(self.scrollable, bg=BG_CARD, padx=12, pady=8)
            row.pack(fill=tk.X, pady=2)

            if item.status == "done":
                indicator_color = SUCCESS
                indicator_text = "done"
            elif item.status == "error":
                indicator_color = ERROR
                indicator_text = "err"
            elif item.status == "skipped":
                indicator_color = WARNING
                indicator_text = "skip"
            elif item.status == "converting":
                indicator_color = ACCENT
                indicator_text = "..."
            else:
                indicator_color = FG_DIM
                indicator_text = " "

            dot = tk.Label(row, text=indicator_text, font=("Segoe UI", 8),
                           bg=BG_CARD, fg=indicator_color, width=4)
            dot.pack(side=tk.LEFT, padx=(0, 8))

            info = tk.Frame(row, bg=BG_CARD)
            info.pack(side=tk.LEFT, fill=tk.X, expand=True)

            name = tk.Label(info, text=item.path.name, font=("Segoe UI", 10),
                            bg=BG_CARD, fg=FG, anchor=tk.W)
            name.pack(fill=tk.X)

            if item.message:
                msg_color = SUCCESS if item.status == "done" else (
                    ERROR if item.status == "error" else (
                        WARNING if item.status == "skipped" else FG_DIM))
                msg = tk.Label(info, text=item.message, font=("Segoe UI", 8),
                               bg=BG_CARD, fg=msg_color, anchor=tk.W)
                msg.pack(fill=tk.X)

            if not self.converting:
                remove = tk.Label(row, text="x", font=("Segoe UI", 10),
                                  bg=BG_CARD, fg=FG_DIM, cursor="hand2", padx=6)
                remove.pack(side=tk.RIGHT)
                idx = i
                remove.bind("<Button-1>", lambda e, idx=idx: self._remove_file(idx))
                remove.bind("<Enter>", lambda e, w=remove: w.configure(fg=ERROR))
                remove.bind("<Leave>", lambda e, w=remove: w.configure(fg=FG_DIM))

        pending = sum(1 for f in self.files if f.status == "pending")
        done = sum(1 for f in self.files if f.status == "done")
        skipped = sum(1 for f in self.files if f.status == "skipped")
        errors = sum(1 for f in self.files if f.status == "error")
        total = len(self.files)

        parts = []
        if pending:
            parts.append(f"{pending} pending")
        if done:
            parts.append(f"{done} converted")
        if skipped:
            parts.append(f"{skipped} skipped")
        if errors:
            parts.append(f"{errors} failed")
        self.status_label.configure(text=f"{total} files: " + ", ".join(parts), fg=FG_DIM)

        self.scrollable.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _show_progress(self):
        if not self.progress_visible:
            self.progress_label.pack(fill=tk.X, pady=(0, 4))
            self.progress_canvas.pack(fill=tk.X, pady=(0, 8))
            self.progress_visible = True

    def _hide_progress(self):
        if self.progress_visible:
            self.progress_label.pack_forget()
            self.progress_canvas.pack_forget()
            self.progress_visible = False

    def _draw_progress(self):
        self.progress_canvas.delete("all")
        w = self.progress_canvas.winfo_width()
        h = self.progress_canvas.winfo_height()
        if w <= 1:
            return
        fill_w = max(0, int(w * self.progress_value))
        if fill_w > 0:
            self.progress_canvas.create_rectangle(0, 0, fill_w, h, fill=ACCENT, outline="")

    def _update_progress(self, done_count, total_count, current_name=""):
        if total_count == 0:
            self.progress_value = 0.0
        else:
            self.progress_value = done_count / total_count

        pct = int(self.progress_value * 100)
        if current_name:
            self.progress_label.configure(
                text=f"Converting {done_count + 1} of {total_count}  —  {pct}%  —  {current_name}",
                fg=ACCENT)
        else:
            self.progress_label.configure(
                text=f"{done_count} of {total_count} complete  —  {pct}%",
                fg=SUCCESS if pct == 100 else FG_DIM)

        self._draw_progress()

    def _remove_file(self, index):
        if self.converting or index >= len(self.files):
            return
        self.files.pop(index)
        self._refresh_list()

    def _start_conversion(self):
        if self.converting:
            self._log("Already converting, ignoring click", "warn")
            return

        pending = [f for f in self.files if f.status == "pending"]
        if not pending:
            self._log("No pending files to convert", "warn")
            return

        self._log(f"--- Starting conversion of {len(pending)} file(s) ---", "accent")

        self.converting = True
        self.convert_btn.configure(text="Converting...", state=tk.DISABLED, bg=BG_CARD)
        self.add_btn.configure(state=tk.DISABLED)

        self.progress_value = 0.0
        self._show_progress()
        self._update_progress(0, len(pending))

        thread = threading.Thread(target=self._convert_worker, daemon=True)
        thread.start()
        self._log("Worker thread started", "info")

    def _convert_worker(self):
        pending = [f for f in self.files if f.status == "pending"]
        total = len(pending)
        done_count = 0

        self._log_threadsafe(f"Worker: {total} files to process", "info")

        for item in self.files:
            if item.status != "pending":
                continue

            item.status = "converting"
            fname = item.path.name
            self._log_threadsafe(f"[{done_count+1}/{total}] Converting: {fname}", "accent")

            self.root.after(0, self._refresh_list)
            self.root.after(0, lambda d=done_count, t=total, n=fname:
                            self._update_progress(d, t, n))

            try:
                out_dir = self.output_dir if self.output_dir else item.path.parent
                output_path = out_dir / f"{item.path.stem}_orca.3mf"

                self._log_threadsafe(f"  Input:  {item.path}", "info")
                self._log_threadsafe(f"  Output: {output_path}", "info")

                strip = self.strip_settings_var.get()
                success = convert_3mf(item.path, output_path, force=True,
                                      strip_settings=strip)

                if success:
                    item.status = "done"
                    item.output_path = output_path
                    item.message = f"Saved: {output_path.name}"
                    self._log_threadsafe(f"  -> SUCCESS: {output_path.name}", "success")
                else:
                    item.status = "error"
                    item.message = "Conversion returned False"
                    self._log_threadsafe(f"  -> FAILED: convert_3mf returned False", "error")

            except Exception as e:
                item.status = "error"
                item.message = str(e)[:80]
                tb = traceback.format_exc()
                self._log_threadsafe(f"  -> EXCEPTION: {e}", "error")
                self._log_threadsafe(tb, "error")

            done_count += 1
            self.root.after(0, self._refresh_list)
            self.root.after(0, lambda d=done_count, t=total:
                            self._update_progress(d, t))

        self._log_threadsafe(f"--- Conversion complete: {done_count} processed ---", "accent")
        self.converting = False
        self.root.after(0, self._conversion_done)

    def _conversion_done(self):
        self.convert_btn.configure(text="Convert All", state=tk.NORMAL, bg=ACCENT)
        self.add_btn.configure(state=tk.NORMAL)
        self._refresh_list()

        done = sum(1 for f in self.files if f.status == "done")
        errors = sum(1 for f in self.files if f.status == "error")
        if done:
            self.progress_value = 1.0
            self._draw_progress()
            self.progress_label.configure(
                text=f"All done!  {done} file{'s' if done != 1 else ''} converted",
                fg=SUCCESS)
            self.status_label.configure(fg=SUCCESS)
        if errors:
            self._log(f"{errors} file(s) had errors - check log above", "error")


def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Vertical.TScrollbar",
                    background=BG_LIGHT, troughcolor=BG,
                    bordercolor=BG, arrowcolor=FG_DIM)

    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
