#!/usr/bin/env python3
"""Simple GUI for converting a DOC/DOCX file to RTF.

The RTF file is saved next to the input file with the same name,
e.g. C:\\Docs\\report.docx -> C:\\Docs\\report.rtf
"""
from __future__ import annotations

import gc
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from doc2rtf import ConversionError, convert_file

FILE_TYPES = [("Word documents", "*.doc *.docx"), ("All files", "*.*")]


def _convert_worker(src: Path, results: queue.Queue):
    """Run in a background thread so the window stays responsive."""
    try:
        results.put(("ok", convert_file(src)))  # saved next to the input file
    except ConversionError as e:
        results.put(("error", str(e)))
    except Exception as e:  # unexpected problems should still reach the user
        results.put(("error", f"Unexpected error: {e}"))


class Doc2RtfApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.input_path: Path | None = None
        self._results: queue.Queue = queue.Queue()

        root.title("DOC/DOCX to RTF Converter")
        root.resizable(True, False)
        root.minsize(520, 0)

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Input file:").grid(row=0, column=0, sticky="w")
        self.path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.path_var, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=6)
        self.browse_btn = ttk.Button(frame, text="Browse...", command=self.browse)
        self.browse_btn.grid(row=0, column=2)

        ttk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.output_var = tk.StringVar(value="-")
        ttk.Label(frame, textvariable=self.output_var, foreground="gray").grid(
            row=1, column=1, columnspan=2, sticky="w", padx=6, pady=(6, 0))

        self.convert_btn = ttk.Button(frame, text="Convert to RTF",
                                      command=self.convert, state="disabled")
        self.convert_btn.grid(row=2, column=0, columnspan=3, pady=(12, 6))

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.progress.grid_remove()  # only shown while converting

        self.status_var = tk.StringVar(value="Choose a .doc or .docx file to convert.")
        ttk.Label(frame, textvariable=self.status_var, wraplength=480).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))

    # -- actions -----------------------------------------------------------
    def browse(self):
        filename = filedialog.askopenfilename(
            parent=self.root, title="Select a Word document", filetypes=FILE_TYPES)
        if filename:
            self.set_input(filename)

    def set_input(self, filename: str):
        self.input_path = Path(filename)
        self.path_var.set(str(self.input_path))
        self.output_var.set(str(self.output_path()))
        self.convert_btn.config(state="normal")
        self.status_var.set("Ready. Click 'Convert to RTF'.")

    def output_path(self) -> Path:
        assert self.input_path is not None
        return self.input_path.with_suffix(".rtf")

    def convert(self):
        if self.input_path is None:
            return
        out = self.output_path()
        if out.exists() and not messagebox.askyesno(
                "File exists", f"{out.name} already exists.\nDo you want to replace it?",
                parent=self.root):
            self.status_var.set("Conversion cancelled.")
            return

        self._set_busy(True)
        self.status_var.set(f"Converting {self.input_path.name}...")
        src = self.input_path
        # The worker gets only the path and the queue, never `self`: Tk objects
        # must not be touched (or garbage-collected) outside the main thread.
        threading.Thread(target=_convert_worker, args=(src, self._results),
                         daemon=True).start()
        self.root.after(100, self._poll)

    # -- result handling (main thread) ----------------------------------------
    def _poll(self):
        try:
            kind, value = self._results.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll)
            return
        self._set_busy(False)
        if kind == "ok":
            self.status_var.set(f"Done! Saved to: {value}")
            messagebox.showinfo("Conversion complete", f"Saved to:\n{value}", parent=self.root)
        else:
            self.status_var.set("Conversion failed.")
            messagebox.showerror("Conversion failed", value, parent=self.root)

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.convert_btn.config(state=state)
        self.browse_btn.config(state=state)
        if busy:
            self.progress.grid()
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.grid_remove()


def main():
    root = tk.Tk()
    app = Doc2RtfApp(root)
    root.mainloop()
    # Free all Tk objects here on the main thread. If the window was closed
    # while a conversion was still running, letting the garbage collector free
    # them later from that worker thread would make Tcl abort the process.
    del app, root
    gc.collect()


if __name__ == "__main__":
    main()
