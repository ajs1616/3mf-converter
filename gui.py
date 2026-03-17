#!/usr/bin/env python3
"""
3MF Converter GUI - Convert old BambuStudio 3MF files to Orca Slicer format.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path

from convert_3mf import convert_3mf, needs_conversion
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
        self.root.geometry("720x560")
        self.root.configure(bg=BG)
        self.root.minsize(600, 450)

        self.files = []
        self.converting = False

        self._build_ui()
        self._setup_drop_target()

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

        # Mouse wheel scrolling
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # --- Bottom bar ---
        bottom = tk.Frame(self.root, bg=BG, padx=20, pady=14)
        bottom.pack(fill=tk.X)

        # Status label
        self.status_label = tk.Label(bottom, text="", font=("Segoe UI", 9),
                                     bg=BG, fg=FG_DIM)
        self.status_label.pack(side=tk.LEFT)

        # Buttons
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

        # Hover effects
        self.add_btn.bind("<Enter>", lambda e: self.add_btn.configure(bg=BG_LIGHT))
        self.add_btn.bind("<Leave>", lambda e: self.add_btn.configure(bg=BG_CARD))
        self.convert_btn.bind("<Enter>", lambda e: self.convert_btn.configure(bg=ACCENT_HOVER))
        self.convert_btn.bind("<Leave>", lambda e: self.convert_btn.configure(bg=ACCENT))

    def _setup_drop_target(self):
        """Try to enable drag-and-drop via tkinterdnd2, fall back gracefully."""
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except ImportError:
            pass

    def _on_drop(self, event):
        """Handle files dropped onto the window."""
        # Parse the dropped file paths (tkinterdnd2 format)
        raw = event.data
        paths = []
        # Handle {path with spaces} and regular paths
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

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select 3MF files",
            filetypes=[("3MF files", "*.3mf"), ("All files", "*.*")]
        )
        for p in paths:
            self._add_file(p)

    def _add_file(self, path):
        path = Path(path)
        # Don't add duplicates
        if any(f.path == path for f in self.files):
            return

        item = FileItem(path)

        # Quick check if it needs conversion
        try:
            with zipfile.ZipFile(path, "r") as z:
                names = z.namelist()
                has_objects = any(n.startswith("3D/Objects/") for n in names)
                has_rels = "3D/_rels/3dmodel.model.rels" in names
                if has_objects and has_rels:
                    model_xml = z.read("3D/3dmodel.model")
                    if not needs_conversion(model_xml):
                        item.status = "skipped"
                        item.message = "Already in Orca format"
        except Exception:
            pass

        self.files.append(item)
        self._refresh_list()

    def _clear_files(self):
        if self.converting:
            return
        self.files.clear()
        self._refresh_list()

    def _refresh_list(self):
        """Rebuild the file list UI."""
        # Clear existing widgets
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

            # Status indicator
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

            # File info
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

            # Remove button (only when not converting)
            if not self.converting:
                remove = tk.Label(row, text="x", font=("Segoe UI", 10),
                                  bg=BG_CARD, fg=FG_DIM, cursor="hand2", padx=6)
                remove.pack(side=tk.RIGHT)
                idx = i
                remove.bind("<Button-1>", lambda e, idx=idx: self._remove_file(idx))
                remove.bind("<Enter>", lambda e, w=remove: w.configure(fg=ERROR))
                remove.bind("<Leave>", lambda e, w=remove: w.configure(fg=FG_DIM))

        # Update status
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
        self.status_label.configure(text=f"{total} files: " + ", ".join(parts))

        # Update canvas scroll region
        self.scrollable.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _remove_file(self, index):
        if self.converting or index >= len(self.files):
            return
        self.files.pop(index)
        self._refresh_list()

    def _start_conversion(self):
        if self.converting:
            return
        pending = [f for f in self.files if f.status == "pending"]
        if not pending:
            return

        self.converting = True
        self.convert_btn.configure(text="Converting...", state=tk.DISABLED, bg=BG_CARD)
        self.add_btn.configure(state=tk.DISABLED)

        thread = threading.Thread(target=self._convert_worker, daemon=True)
        thread.start()

    def _convert_worker(self):
        for item in self.files:
            if item.status != "pending":
                continue

            item.status = "converting"
            self.root.after(0, self._refresh_list)

            try:
                output_path = item.path.parent / f"{item.path.stem}_orca.3mf"
                success = convert_3mf(item.path, output_path, force=True)
                if success:
                    item.status = "done"
                    item.output_path = output_path
                    item.message = f"Saved: {output_path.name}"
                else:
                    item.status = "error"
                    item.message = "Conversion failed"
            except Exception as e:
                item.status = "error"
                item.message = str(e)[:80]

            self.root.after(0, self._refresh_list)

        self.converting = False
        self.root.after(0, self._conversion_done)

    def _conversion_done(self):
        self.convert_btn.configure(text="Convert All", state=tk.NORMAL, bg=ACCENT)
        self.add_btn.configure(state=tk.NORMAL)
        self._refresh_list()

        done = sum(1 for f in self.files if f.status == "done")
        if done:
            self.status_label.configure(
                text=self.status_label.cget("text") + "  —  All done!",
                fg=SUCCESS
            )


def main():
    # Try to use tkinterdnd2 for drag-and-drop support
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    # Set window icon if possible
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    # Style the ttk scrollbar
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Vertical.TScrollbar",
                    background=BG_LIGHT, troughcolor=BG,
                    bordercolor=BG, arrowcolor=FG_DIM)

    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
