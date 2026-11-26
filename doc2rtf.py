#!/usr/bin/env python3
"""Convert DOC and DOCX files to RTF using LibreOffice in headless mode.

Usage:
    python doc2rtf.py                  # open the GUI (also: --gui)
    python doc2rtf.py input.docx
    python doc2rtf.py file1.doc file2.docx -o out_dir
    python doc2rtf.py some_folder -r -o out_dir
"""
from __future__ import annotations

import argparse
import gc
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

try:  # Tkinter is only needed for the GUI; the CLI works without it.
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # pragma: no cover - depends on the Python build
    tk = None

SUPPORTED_EXTENSIONS = {".doc", ".docx"}


class ConversionError(RuntimeError):
    """Raised when a file cannot be converted."""


def find_soffice() -> str:
    """Locate the LibreOffice executable."""
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).is_file():
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    raise ConversionError(
        "LibreOffice was not found. Install it or set SOFFICE_PATH to the soffice executable."
    )


def convert_file(src: str | Path, out_dir: str | Path | None = None,
                 overwrite: bool = True, timeout: int = 120) -> Path:
    """Convert one .doc/.docx file to .rtf and return the output path."""
    src = Path(src).resolve()
    if not src.is_file():
        raise ConversionError(f"Input file not found: {src}")
    if src.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ConversionError(f"Unsupported file type (expected .doc/.docx): {src}")

    out_dir = Path(out_dir).resolve() if out_dir else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / (src.stem + ".rtf")
    if target.exists() and not overwrite:
        raise ConversionError(f"Output already exists (use --overwrite): {target}")

    soffice = find_soffice()
    # A private profile dir lets several conversions run at once and avoids
    # clashing with a LibreOffice instance the user may already have open.
    with tempfile.TemporaryDirectory(prefix="doc2rtf_") as tmp:
        profile = Path(tmp, "profile").as_uri()
        work_out = Path(tmp, "out")
        work_out.mkdir()
        cmd = [
            soffice, f"-env:UserInstallation={profile}",
            "--headless", "--norestore", "--nolockcheck",
            "--convert-to", "rtf", "--outdir", str(work_out), str(src),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError(f"Timed out converting {src}") from exc

        produced = work_out / (src.stem + ".rtf")
        if proc.returncode != 0 or not produced.is_file():
            detail = (proc.stderr or proc.stdout).strip()
            raise ConversionError(f"LibreOffice failed to convert {src}: {detail}")
        shutil.move(str(produced), str(target))
    return target


def collect_inputs(paths: list[str], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            it = p.rglob("*") if recursive else p.iterdir()
            files.extend(sorted(f for f in it
                                if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
                                and not f.name.startswith("~$")))
        else:
            files.append(p)
    return files


# ---------------------------------------------------------------------------
# Graphical interface
# ---------------------------------------------------------------------------
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


def run_gui() -> int:
    """Open the converter window."""
    if tk is None:
        print("Tkinter is not available, so the GUI cannot start. Install it "
              "(e.g. 'python3-tk' on Debian/Ubuntu) or pass files on the command line.",
              file=sys.stderr)
        return 1
    root = tk.Tk()
    app = Doc2RtfApp(root)
    root.mainloop()
    # Free all Tk objects here on the main thread. If the window was closed
    # while a conversion was still running, letting the garbage collector free
    # them later from that worker thread would make Tcl abort the process.
    del app, root
    gc.collect()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert DOC/DOCX files to RTF.")
    parser.add_argument("inputs", nargs="*",
                        help="Files or folders to convert (omit to open the GUI)")
    parser.add_argument("--gui", action="store_true", help="Open the graphical interface")
    parser.add_argument("-o", "--output-dir", help="Output folder (default: next to each input)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Search folders recursively")
    parser.add_argument("--no-overwrite", action="store_true", help="Skip files whose RTF already exists")
    parser.add_argument("--timeout", type=int, default=120, help="Seconds allowed per file")
    args = parser.parse_args(argv)
    if args.gui or not args.inputs:
        return run_gui()

    files = collect_inputs(args.inputs, args.recursive)
    if not files:
        print("No .doc/.docx files found.", file=sys.stderr)
        return 1

    failures = 0
    for f in files:
        try:
            out = convert_file(f, args.output_dir, overwrite=not args.no_overwrite,
                               timeout=args.timeout)
            print(f"OK    {f} -> {out}")
        except ConversionError as e:
            failures += 1
            print(f"FAIL  {e}", file=sys.stderr)
    print(f"\n{len(files) - failures} converted, {failures} failed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
